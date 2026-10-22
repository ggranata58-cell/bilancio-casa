#!/usr/bin/env python3
"""Apply the bidirectional Excel-ID improvement to the packaged PWA archive."""
import os
import re
import sys
import tempfile
import zipfile

archive = sys.argv[1] if len(sys.argv) > 1 else "Bilancio_Casa_PWA_v0_2_0.zip"
target = "bilancio-casa-pwa/app.js"

with zipfile.ZipFile(archive, "r") as src:
    bad = src.testzip()
    if bad:
        raise SystemExit(f"Input ZIP is corrupt: {bad}")
    files = {info.filename: (info, src.read(info.filename)) for info in src.infolist() if not info.is_dir()}

if target not in files:
    raise SystemExit(f"App source not found in archive: {target}")

info, raw = files[target]
app = raw.decode("utf-8")
def replace_once(old, new, label):
    global app
    count = app.count(old)
    if count == 0:
        return
    if count != 1:
        raise SystemExit(f"Expected one occurrence for {label}, found {count}")
    app = app.replace(old, new, 1)

if "withStableId" not in app:
    replace_once(
    "          notes: find(['note', 'notes'])\n",
    "          notes: find(['note', 'notes']),\n          id: find(['idmovimento', 'movementid', 'movementuuid', 'id'])\n",
    "Excel ID header detection")
    replace_once(
    "      payload.push({\n        user_id: currentUser.id,\n        movement_date: date,",
    "      const suppliedId = indexes.id >= 0 ? String(row[indexes.id] ?? '').trim() : '';\n      const hasStableId = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(suppliedId);\n      payload.push({\n        ...(hasStableId ? { id: suppliedId } : {}),\n        user_id: currentUser.id,\n        movement_date: date,",
    "stable ID import")
    replace_once(
    "        import_key: `xlsx:${fileHashValue}:${sheetName}:${r + 1}`\n      });",
    "        ...(hasStableId ? {} : { import_key: `xlsx:${fileHashValue}:${sheetName}:${r + 1}` })\n      });",
    "stable import key selection")
    old_batch = """    let insertedBatches = 0;
     for (let i = 0; i < payload.length; i += 150) {
       const part = payload.slice(i, i + 150);
       const { error } = await db.from('bilancio_movements').upsert(part, { onConflict: 'user_id,import_key', ignoreDuplicates: true });
       if (error) throw error;
       insertedBatches++;
     }
    """
    new_batch = """    const withStableId = payload.filter(m => m.id);
        const withoutStableId = payload.filter(m => !m.id);
        for (let i = 0; i < withStableId.length; i += 150) {
          const part = withStableId.slice(i, i + 150);
          const { error } = await db.from('bilancio_movements').upsert(part, { onConflict: 'id' });
          if (error) throw error;
        }
        for (let i = 0; i < withoutStableId.length; i += 150) {
          const part = withoutStableId.slice(i, i + 150);
          const { error } = await db.from('bilancio_movements').upsert(part, { onConflict: 'user_id,import_key', ignoreDuplicates: true });
          if (error) throw error;
        }
    """
    batch_pattern = r"    let insertedBatches = 0;.*?insertedBatches\+\+;\s*\}"
    app, batch_count = re.subn(batch_pattern, new_batch.rstrip(), app, count=1, flags=re.S)
    if batch_count != 1:
        raise SystemExit(f"Expected one legacy import batch block, found {batch_count}")
    replace_once(
    "L'importazione non modifica né cancella movimenti già registrati. Continuare?",
    "Le righe con ID movimento vengono aggiornate senza creare doppioni; le righe senza ID vengono aggiunte senza cancellare movimenti esistenti. Continuare?",
    "import confirmation")
    replace_once(
    "Se il file era già stato importato, le stesse righe non vengono duplicate.",
    "Le righe con ID movimento mantengono la propria identità tra esportazioni e reimportazioni; le righe senza ID usano il sistema di importazione precedente.",
    "import completion notice")
    replace_once(
    "    const rows = [['Data', 'Giorno', 'Concetto', 'Categoria', 'Tipo', 'Importo', 'Importo +/-', 'Saldo', 'Mese', 'Note']];",
    "    const rows = [['ID movimento', 'Data', 'Giorno', 'Concetto', 'Categoria', 'Tipo', 'Importo', 'Importo +/-', 'Saldo', 'Mese', 'Note']];",
    "CSV export header")
    replace_once(
    "    rows.push([csvDate(settings.opening_date), '', 'Saldo conto', 'Altro', 'Entrata', csvNumber(settings.opening_balance), csvNumber(settings.opening_balance), csvNumber(balance), '', 'Saldo iniziale importato dalle impostazioni']);",
    "    rows.push(['', csvDate(settings.opening_date), '', 'Saldo conto', 'Altro', 'Entrata', csvNumber(settings.opening_balance), csvNumber(settings.opening_balance), csvNumber(balance), '', 'Saldo iniziale importato dalle impostazioni']);",
    "CSV opening-balance row")
    replace_once(
    "      rows.push([csvDate(m.movement_date), '', m.concept, m.category, m.movement_type, csvNumber(m.amount), csvNumber(signed(m)), csvNumber(balance), monthName, m.notes || '']);",
    "      rows.push([m.id || '', csvDate(m.movement_date), '', m.concept, m.category, m.movement_type, csvNumber(m.amount), csvNumber(signed(m)), csvNumber(balance), monthName, m.notes || '']);",
    "CSV movement ID export")
else:
    # Existing archive already has stable-ID import. Repair its known key-overwrite bug.
    tick = chr(96)
    lines = app.splitlines()
    spread = next((j for j, line in enumerate(lines) if "...(hasStableId ?" in line and "excel-id:" in line), None)
    if spread is not None:
        lines[spread] = lines[spread].replace(", import_key: " + tick + "excel-id:${suppliedId}" + tick, "")
    duplicate = next((j for j, line in enumerate(lines) if line.strip().startswith("import_key:") and "xlsx:" in line and "hasStableId" not in line), None)
    if duplicate is not None:
        del lines[duplicate]
    app = "\n".join(lines) + "\n"
    if "...(hasStableId ? { id: suppliedId } : { import_key:" not in app:
        raise SystemExit("Stable-ID conditional import_key assignment not found.")
    if "upsert(part, { onConflict: 'id' })" not in app:
        raise SystemExit("Stable-ID upsert behavior is missing.")


replace_once(
"  function renderDashboard() {",
"""  const CASH_BALANCE_KEY_PREFIX = 'bilancio-casa-cash-balance-v1:';
  function cashBalanceKey() { return CASH_BALANCE_KEY_PREFIX + (currentUser?.id || 'local'); }
  function readCashBalance() {
    if (cloudCashBalance !== null) return cloudCashBalance;
    try {
      const raw = localStorage.getItem(cashBalanceKey());
      if (raw === null) return 50;
      const value = Number(raw);
      return Number.isFinite(value) ? value : 50;
    } catch (_) { return 50; }
  }
  async function loadCashBalance() {
    const key = cashBalanceKey();
    let localValue = null;
    try {
      const raw = localStorage.getItem(key);
      if (raw !== null && Number.isFinite(Number(raw))) localValue = Math.max(0, Number(raw));
    } catch (_) {}
    const { data, error } = await db.from('bilancio_cash_settings').select('cash_balance').eq('user_id', currentUser.id).maybeSingle();
    if (error) {
      cloudCashBalance = localValue === null ? 50 : localValue;
      cashCloudAvailable = false;
      renderCashPanel();
      return;
    }
    cashCloudAvailable = true;
    if (data) cloudCashBalance = Math.max(0, Number(data.cash_balance) || 0);
    else {
      cloudCashBalance = localValue === null ? 50 : localValue;
      const { error: saveError } = await db.from('bilancio_cash_settings').upsert({
        user_id: currentUser.id, cash_balance: Math.round(cloudCashBalance * 100) / 100, updated_at: new Date().toISOString()
      }, { onConflict: 'user_id' });
      if (saveError) cashCloudAvailable = false;
    }
    try { localStorage.setItem(key, cloudCashBalance.toFixed(2)); } catch (_) {}
    renderCashPanel();
  }
  function renderCashPanel() {
    const table = $('monthlyRows')?.closest('table');
    if (!table) return;
    let panel = $('cashSummaryPanel');
    if (!panel) {
      table.insertAdjacentHTML('afterend', `
        <section id="cashSummaryPanel" class="settings-card" style="margin-top:1rem">
          <h3>Contanti e prelievi</h3>
          <p class="small-note">I prelievi restano nel saldo bancario, ma non sono conteggiati come spese ordinarie. Non registrare qui i singoli acquisti in contanti.</p>
          <div class="form-grid">
            <label for="cashBalanceInput">Saldo residuo contanti (€)
              <input id="cashBalanceInput" type="number" inputmode="decimal" min="0" step="0.01" />
            </label>
            <div class="form-actions"><button id="saveCashBalanceBtn" type="button">Salva saldo contanti</button></div>
          </div>
          <div style="margin-top:1rem;display:flex;flex-direction:column;gap:1rem">
            <div style="display:flex;flex-direction:column;align-items:flex-start;gap:.5rem;padding:.75rem;border:1px solid var(--border-color,#ddd);border-radius:.5rem;min-width:0"><span class="small-note" style="display:block">Prelievi nel mese selezionato</span><strong id="cashMonthWithdrawals" style="display:block;width:100%;margin:0">—</strong></div>
            <div style="display:flex;flex-direction:column;align-items:flex-start;gap:.5rem;padding:.75rem;border:1px solid var(--border-color,#ddd);border-radius:.5rem;min-width:0"><span class="small-note" style="display:block">Prelievi complessivi registrati</span><strong id="cashTotalWithdrawals" style="display:block;width:100%;margin:0">—</strong></div>
          </div>
          <p id="cashSyncStatus" class="small-note">Caricamento del saldo contanti…</p>
        </section>`);
    }
    const ym = $('periodMonth').value || todayMonth();
    const isCashWithdrawal = m => m.category === 'Prelievo contanti' && m.movement_type === 'Uscita';
    const monthTotal = monthMovements(ym).filter(isCashWithdrawal).reduce((sum, m) => sum + Number(m.amount || 0), 0);
    const allTotal = movements.filter(m => inAccountingPeriod(m) && isCashWithdrawal(m)).reduce((sum, m) => sum + Number(m.amount || 0), 0);
    $('cashBalanceInput').value = readCashBalance().toFixed(2);
    $('cashMonthWithdrawals').textContent = money(monthTotal);
    $('cashTotalWithdrawals').textContent = money(allTotal);
    const status = $('cashSyncStatus');
    if (status) status.textContent = cashCloudAvailable
      ? 'Saldo sincronizzato con il cloud: disponibile sugli altri dispositivi con lo stesso account.'
      : 'Saldo salvato solo su questo dispositivo. Per la sincronizzazione cloud occorre creare la tabella dedicata.';
  }
  async function saveCashBalance() {
    const input = $('cashBalanceInput');
    const value = Number(String(input.value).replace(',', '.'));
    if (!Number.isFinite(value) || value < 0) {
      showNotice('Inserisci un saldo contanti valido, pari o superiore a zero.', 'error');
      return;
    }
    const rounded = Math.round(value * 100) / 100;
    if (!cashCloudAvailable) {
      try {
        localStorage.setItem(cashBalanceKey(), rounded.toFixed(2));
        cloudCashBalance = rounded;
        input.value = rounded.toFixed(2);
        showNotice('Saldo contanti salvato su questo dispositivo. La sincronizzazione cloud non è attiva.', 'info');
        renderCashPanel();
      } catch (_) { showNotice('Non riesco a salvare il saldo contanti nel browser.', 'error'); }
      return;
    }
    const button = $('saveCashBalanceBtn');
    const controller = new AbortController();
    let saveTimeout;
    setButtonBusy(button, true, 'Salvataggio…');
    try {
      saveTimeout = setTimeout(() => {
        controller.abort();
        setButtonBusy(button, false);
        showNotice('Il salvataggio non ha risposto entro 15 secondi. Il pulsante è stato riattivato; verifica il saldo prima di riprovare.', 'error');
      }, 15000);
      const { error } = await db.from('bilancio_cash_settings').upsert({
        user_id: currentUser.id, cash_balance: rounded, updated_at: new Date().toISOString()
      }, { onConflict: 'user_id' }).abortSignal(controller.signal);
      if (error) throw error;
      cloudCashBalance = rounded;
      try { localStorage.setItem(cashBalanceKey(), rounded.toFixed(2)); } catch (_) {}
      input.value = rounded.toFixed(2);
      renderCashPanel();
      showNotice('Saldo contanti salvato e sincronizzato.');
    } catch (error) {
      if (error?.name === 'AbortError' || controller.signal.aborted) {
        showNotice('Il salvataggio ha impiegato troppo tempo. Il pulsante è stato riattivato; verifica il saldo prima di riprovare.', 'error');
      } else {
        showNotice('Errore nel salvataggio cloud del saldo contanti: ' + (error?.message || error), 'error');
      }
    } finally {
      clearTimeout(saveTimeout);
      setButtonBusy(button, false);
    }
  }

  function renderDashboard() {""",
"cash balance and withdrawals UI")
replace_once(
"  let noticeTimer = null;",
"  let noticeTimer = null;\n  let cloudCashBalance = null;\n  let cashCloudAvailable = false;",
"cash sync state")
replace_once(
"      await loadSettings();\n      await loadMovements();",
"      await loadSettings();\n      await loadCashBalance();\n      await loadMovements();",
"load cash balance after login")
replace_once(
"""    } finally {
      clearTimeout(saveTimeout);
      setButtonBusy(button, false);
    }
  }

  function renderDashboard() {""",
"""    } finally {
      clearTimeout(saveTimeout);
      if (button) {
        button.textContent = 'Salva saldo contanti';
        button.disabled = false;
        delete button.dataset.originalText;
      }
    }
  }

  function renderDashboard() {""",
"force cash save button reset")

replace_once(
"    $('exportBtn').addEventListener('click', exportCsv);",
"    $('exportBtn').addEventListener('click', exportCsv);\n    document.addEventListener('click', event => { if (event.target.closest('#saveCashBalanceBtn')) saveCashBalance(); });",
"cash balance save action")


replace_once(
"    $('statTaxReserveCaption').textContent = `Accantonamento per ${s.dayCount} giorni`; ",
"    $('statTaxReserveCaption').textContent = `Accantonamento per ${s.dayCount} giorni`; \n    renderCashPanel();",
"render cash summary")
replace_once(
"    $('exportBtn').addEventListener('click', exportCsv);",
"    $('exportBtn').addEventListener('click', exportCsv);\n    document.addEventListener('click', event => { if (event.target.closest('#saveCashBalanceBtn')) saveCashBalance(); });",
"cash balance save action")

replace_once(
"    'Finanz. Pannelli solari', 'Pensione', 'Imposte e tasse', 'Banca', 'Contanti', 'Altro'\n  ];\n  const SPECIAL_CATEGORY = 'Tasas y Intereses';",
"    'Finanz. Pannelli solari', 'Pensione', 'Imposte e tasse', 'Banca', 'Contanti', 'Altro',\n    'Movimenti esclusi', 'Prelievo contanti'\n  ];\n  const SPECIAL_CATEGORY = 'Tasas y Intereses';\n  const EXCLUDED_CATEGORIES = new Set(['Movimenti esclusi', 'Prelievo contanti']);\n  function isExcludedMovement(m) { return EXCLUDED_CATEGORIES.has(m.category); }",
"excluded movement categories")
replace_once(
"    const expenses = rows.filter(m => m.movement_type === 'Uscita').reduce((s, m) => s + m.amount, 0);",
"    const ordinaryOutflows = rows.filter(m => m.movement_type === 'Uscita' && !isExcludedMovement(m)).reduce((s, m) => s + m.amount, 0);\n    const interestCredits = rows.filter(m => m.movement_type === 'Entrata' && m.category === 'Interessi attivi').reduce((s, m) => s + m.amount, 0);\n    const expenses = Math.round((ordinaryOutflows - interestCredits + Number.EPSILON) * 100) / 100;\n    const excludedOutflows = rows.filter(m => m.movement_type === 'Uscita' && isExcludedMovement(m)).reduce((s, m) => s + m.amount, 0);",
"separate ordinary and excluded expenses")
replace_once(
"return { ym, rows, income, expenses, budgetAllocated, taxReserve, budgetRemainder, monthOpen, closing, dayCount };",
"return { ym, rows, income, expenses, excludedOutflows, budgetAllocated, taxReserve, budgetRemainder, monthOpen, closing, dayCount };",
"return excluded total")
replace_once(
"    for (const m of rows) {\n      if (m.category === SPECIAL_CATEGORY) {\n        totals.set(m.category, (totals.get(m.category) || 0) + (m.movement_type === 'Uscita' ? m.amount : -m.amount));\n      } else if (m.movement_type === 'Uscita') {\n        totals.set(m.category, (totals.get(m.category) || 0) + m.amount);\n      }\n    }",
"    for (const m of rows) {\n      if (isExcludedMovement(m)) {\n        totals.set(m.category, (totals.get(m.category) || 0) + (m.movement_type === 'Uscita' ? m.amount : -m.amount));\n      } else if (m.category === SPECIAL_CATEGORY) {\n        totals.set(m.category, (totals.get(m.category) || 0) + (m.movement_type === 'Uscita' ? m.amount : -m.amount));\n      } else if (m.movement_type === 'Uscita') {\n        totals.set(m.category, (totals.get(m.category) || 0) + m.amount);\n      }\n    }",
"separate category totals")
replace_once(
"    $('statExpenses').textContent = money(s.expenses);",
"    $('statExpenses').textContent = money(s.expenses);\n    $('statExpenses').title = `Spese ordinarie: ${money(s.expenses)}. Uscite escluse dal budget: ${money(s.excludedOutflows)}.`;",
"excluded expense tooltip")
replace_once(
"      rows.push([m.id || '', csvDate(m.movement_date), '', m.concept, m.category, m.movement_type, csvNumber(m.amount), csvNumber(signed(m)), csvNumber(balance), monthName, m.notes || '']);",
"      rows.push([m.id || '', csvDate(m.movement_date), '', m.concept, m.category, m.movement_type, csvNumber(m.amount), csvNumber(signed(m)), csvNumber(balance), monthName, [m.notes || '', isExcludedMovement(m) ? 'ESCLUSO DALLE SPESE ORDINARIE' : ''].filter(Boolean).join(' | ')]);",
"CSV excluded note");

files[target] = (info, app.encode("utf-8"))
fd, temp_path = tempfile.mkstemp(suffix=".zip", dir=os.path.dirname(os.path.abspath(archive)))
os.close(fd)
try:
    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as dst:
        for name, (old_info, content) in files.items():
            dst.writestr(old_info, content)
    with zipfile.ZipFile(temp_path, "r") as check:
        bad = check.testzip()
        if bad:
            raise SystemExit(f"Patched ZIP is corrupt: {bad}")
        patched = check.read(target).decode("utf-8")
        if "withStableId" not in patched or "ID movimento" not in patched:
            raise SystemExit("Patch verification failed.")
        for required in ["function renderCashPanel()", "cashBalanceInput", "cashMonthWithdrawals", "cashTotalWithdrawals", "localStorage.setItem(cashBalanceKey()", "renderCashPanel();"]:
            if required not in patched:
                raise SystemExit(f"Cash summary patch verification failed: {required}")
    os.replace(temp_path, archive)
finally:
    if os.path.exists(temp_path):
        os.remove(temp_path)
print("Applied and verified stable Excel movement IDs.")
