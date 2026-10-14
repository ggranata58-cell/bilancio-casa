(function () {
  'use strict';

  const CATEGORIES = [
    'Tabacchi', 'Abbigliamento', 'Materiali casa', 'Ristoranti e Bar', 'Carburante',
    'Antifurto', 'Assicurazione', 'Sanitarie', 'Veterinarie', 'Super e alim.',
    'Varie auto e moto', 'Rete e Tel.', 'Energia elettrica', 'Acqua - Facsa',
    'Amazon Prime + Temu', 'Varie', 'Divertimento', 'Tasas y Intereses',
    'Finanz. Pannelli solari', 'Pensione', 'Imposte e tasse', 'Banca', 'Contanti', 'Altro',
    'Movimenti esclusi', 'Prelievo contanti'
  ];
  const SPECIAL_CATEGORY = 'Tasas y Intereses';
  const EXCLUDED_CATEGORIES = new Set(['Movimenti esclusi', 'Prelievo contanti']);
  function isExcludedMovement(m) { return EXCLUDED_CATEGORIES.has(m.category); }
  const DEFAULTS = { opening_date: '2026-09-01', opening_balance: 17727.21, daily_budget: 81.71, tax_reserve_daily: 19.50 };
  const $ = id => document.getElementById(id);
  const euro = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR' });
  const dateFmt = new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' });
  let db = null;
  let currentUser = null;
  let settings = { ...DEFAULTS };
  let quotaHistory = [];
  let movements = [];
  let authMode = 'signin';
  let pendingInstallPrompt = null;
  let noticeTimer = null;

  function money(value) { return euro.format(Number(value) || 0); }
  function normText(value) {
    return String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  }
  function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  }
  function isoToday() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }
  function todayMonth() { return isoToday().slice(0, 7); }
  function fmtDate(value) {
    if (!value) return '—';
    const d = new Date(`${String(value).slice(0, 10)}T12:00:00`);
    return Number.isNaN(d.getTime()) ? String(value) : dateFmt.format(d);
  }
  function monthLabel(ym) {
    const [y, m] = ym.split('-').map(Number);
    return new Intl.DateTimeFormat('it-IT', { month: 'long', year: 'numeric' }).format(new Date(y, m - 1, 15));
  }
  function monthStart(ym) { return `${ym}-01`; }
  function isoLocalDate(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }
  function monthEndDate(ym) {
    const [y, m] = ym.split('-').map(Number);
    return new Date(y, m, 0).getDate();
  }
  function signed(m) { return m.movement_type === 'Entrata' ? Number(m.amount) : -Number(m.amount); }
  function inAccountingPeriod(m) { return m.movement_date >= settings.opening_date; }
  function showNotice(message, type = 'ok') {
    const node = $('globalMessage');
    node.textContent = message;
    node.className = `message${type === 'error' ? ' error' : type === 'info' ? ' info' : ''}`;
    node.hidden = false;
    clearTimeout(noticeTimer);
    noticeTimer = setTimeout(() => { node.hidden = true; }, 6500);
  }
  function setAuthMessage(message, isError = false) {
    const n = $('authMessage');
    n.textContent = message || '';
    n.style.color = isError ? 'var(--red)' : 'var(--muted)';
  }
  function setButtonBusy(button, busy, busyText = 'Attendi…') {
    if (!button) return;
    if (busy) {
      button.dataset.originalText = button.textContent;
      button.textContent = busyText;
      button.disabled = true;
    } else {
      button.textContent = button.dataset.originalText || button.textContent;
      button.disabled = false;
      delete button.dataset.originalText;
    }
  }
  function setView(name) {
    document.querySelectorAll('.view-panel').forEach(el => { el.hidden = el.id !== `view-${name}`; });
    document.querySelectorAll('.nav-button').forEach(el => el.classList.toggle('active', el.dataset.view === name));
    if (name === 'movements') $('movementConcept').focus({ preventScroll: true });
  }

  function fillCategoryOptions() {
    $('movementCategory').innerHTML = CATEGORIES.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  }

  function defaultMovementForm() {
    $('movementId').value = '';
    $('movementDate').value = isoToday();
    $('movementType').value = 'Uscita';
    $('movementConcept').value = '';
    $('movementCategory').value = 'Altro';
    $('movementAmount').value = '';
    $('movementNotes').value = '';
    $('movementFormTitle').textContent = 'Nuovo movimento';
    $('saveMovementBtn').textContent = 'Salva movimento';
    $('cancelEditBtn').hidden = true;
  }

  function setAuthenticatedPanels(isAuthenticated) {
    $('authPanel').hidden = isAuthenticated;
    $('appPanel').hidden = !isAuthenticated;
    $('configPanel').hidden = true;
  }

  function showSetupRequired() {
    $('configPanel').hidden = false;
    $('authPanel').hidden = true;
    $('appPanel').hidden = true;
  }

  async function handleAuthSubmit(event) {
    event.preventDefault();
    const email = $('authEmail').value.trim();
    const password = $('authPassword').value;
    const button = $('authSubmit');
    setButtonBusy(button, true);
    setAuthMessage('');
    try {
      if (authMode === 'signup') {
        const { data, error } = await db.auth.signUp({ email, password });
        if (error) throw error;
        if (data.session) {
          currentUser = data.user;
          await afterLogin();
        } else {
          setAuthMessage('Account creato. Controlla la tua email per confermare la registrazione, poi torna qui per accedere.');
        }
      } else {
        const { data, error } = await db.auth.signInWithPassword({ email, password });
        if (error) throw error;
        currentUser = data.user;
        await afterLogin();
      }
    } catch (error) {
      setAuthMessage(error?.message || 'Non è stato possibile completare l’accesso.', true);
    } finally {
      setButtonBusy(button, false);
    }
  }

  function switchAuthMode(mode) {
    authMode = mode;
    const signup = mode === 'signup';
    $('authTitle').textContent = signup ? 'Crea il tuo account' : 'Accedi al tuo bilancio';
    $('authSubmit').textContent = signup ? 'Crea account' : 'Accedi';
    $('authPassword').autocomplete = signup ? 'new-password' : 'current-password';
    $('showSignupBtn').hidden = signup;
    $('showSigninBtn').hidden = !signup;
    setAuthMessage(signup ? 'Usa la tua email personale e una password dedicata di almeno 8 caratteri.' : '');
  }

  async function afterLogin() {
    if (!currentUser) return;
    setAuthenticatedPanels(true);
    $('signedInEmail').textContent = currentUser.email || 'Account collegato';
    showNotice('Accesso riuscito. Sto caricando movimenti e riepiloghi…', 'info');
    try {
      await loadSettings();
      await loadMovements();
      defaultMovementForm();
      $('periodMonth').value = todayMonth();
      renderAll();
      setView('dashboard');
      $('globalMessage').hidden = true;
    } catch (error) {
      showNotice(`Errore nel caricamento dei dati: ${error?.message || error}`, 'error');
    }
  }

  async function loadSettings() {
    const { data, error } = await db.from('bilancio_settings').select('*').eq('user_id', currentUser.id).maybeSingle();
    if (error) throw error;
    if (data) {
      settings = {
        opening_date: String(data.opening_date).slice(0, 10),
        opening_balance: Number(data.opening_balance),
        daily_budget: Number(data.daily_budget)
      };
    } else {
      const record = { user_id: currentUser.id, opening_date: DEFAULTS.opening_date, opening_balance: DEFAULTS.opening_balance, daily_budget: DEFAULTS.daily_budget };
      const { data: inserted, error: insertError } = await db.from('bilancio_settings').insert(record).select().single();
      if (insertError) throw insertError;
      settings = { opening_date: String(inserted.opening_date).slice(0, 10), opening_balance: Number(inserted.opening_balance), daily_budget: Number(inserted.daily_budget) };
    }
    await loadQuotaHistory();
    renderSettings();
  }

  async function loadQuotaHistory() {
    const { data, error } = await db.from('bilancio_quota_storico')
      .select('user_id,valid_from,daily_budget,tax_reserve_daily,created_at,updated_at')
      .eq('user_id', currentUser.id).order('valid_from', { ascending: true });
    if (error) throw error;
    quotaHistory = (data || []).map(row => ({
      ...row,
      valid_from: String(row.valid_from).slice(0, 10),
      daily_budget: Number(row.daily_budget),
      tax_reserve_daily: Number(row.tax_reserve_daily)
    }));
    if (!quotaHistory.length) {
      const seed = {
        user_id: currentUser.id,
        valid_from: settings.opening_date,
        daily_budget: Number(settings.daily_budget) || DEFAULTS.daily_budget,
        tax_reserve_daily: DEFAULTS.tax_reserve_daily
      };
      const { error: seedError } = await db.from('bilancio_quota_storico').upsert(seed, { onConflict: 'user_id,valid_from' });
      if (seedError) throw seedError;
      const { data: seeded, error: reloadError } = await db.from('bilancio_quota_storico')
        .select('user_id,valid_from,daily_budget,tax_reserve_daily,created_at,updated_at')
        .eq('user_id', currentUser.id).order('valid_from', { ascending: true });
      if (reloadError) throw reloadError;
      quotaHistory = (seeded || []).map(row => ({ ...row, valid_from: String(row.valid_from).slice(0, 10), daily_budget: Number(row.daily_budget), tax_reserve_daily: Number(row.tax_reserve_daily) }));
    }
  }

  function ratesForDate(date) {
    if (!quotaHistory.length) return { daily_budget: DEFAULTS.daily_budget, tax_reserve_daily: DEFAULTS.tax_reserve_daily };
    let selected = null;
    for (const item of quotaHistory) {
      if (item.valid_from <= date) selected = item;
      else break;
    }
    if (!selected) selected = quotaHistory[0];
    return { daily_budget: Number(selected.daily_budget) || 0, tax_reserve_daily: Number(selected.tax_reserve_daily) || 0 };
  }

  function sumDailyRate(fromDate, toDate, field) {
    if (!fromDate || !toDate || toDate < fromDate) return 0;
    let sum = 0;
    const d = new Date(`${fromDate}T12:00:00`);
    const end = new Date(`${toDate}T12:00:00`);
    for (; d <= end; d.setDate(d.getDate() + 1)) {
      const date = isoLocalDate(d);
      sum += Number(ratesForDate(date)[field]) || 0;
    }
    return Math.round((sum + Number.EPSILON) * 100) / 100;
  }

  async function loadMovements() {
    let all = [];
    const pageSize = 1000;
    for (let from = 0; ; from += pageSize) {
      const { data, error } = await db.from('bilancio_movements').select('id,user_id,movement_date,concept,category,movement_type,amount,notes,import_key,created_at')
        .eq('user_id', currentUser.id)
        .order('movement_date', { ascending: true })
        .order('created_at', { ascending: true })
        .range(from, from + pageSize - 1);
      if (error) throw error;
      all = all.concat(data || []);
      if (!data || data.length < pageSize) break;
    }
    movements = all.map(m => ({ ...m, movement_date: String(m.movement_date).slice(0, 10), amount: Number(m.amount) }));
  }

  async function refreshData() {
    await loadSettings();
    await loadMovements();
    renderAll();
  }

  function balanceAt(cutoffDate) {
    let total = Number(settings.opening_balance) || 0;
    for (const m of movements) {
      if (m.movement_date < settings.opening_date || m.movement_date > cutoffDate) continue;
      total += signed(m);
    }
    return total;
  }

  function monthMovements(ym) {
    const from = monthStart(ym);
    const last = `${ym}-${String(monthEndDate(ym)).padStart(2, '0')}`;
    return movements.filter(m => inAccountingPeriod(m) && m.movement_date >= from && m.movement_date <= last);
  }

  function monthlyStats(ym) {
    const rows = monthMovements(ym);
    const income = rows.filter(m => m.movement_type === 'Entrata').reduce((s, m) => s + m.amount, 0);
    const expenses = rows.filter(m => m.movement_type === 'Uscita' && !isExcludedMovement(m)).reduce((s, m) => s + m.amount, 0);
    const excludedOutflows = rows.filter(m => m.movement_type === 'Uscita' && isExcludedMovement(m)).reduce((s, m) => s + m.amount, 0);
    const [y, mo] = ym.split('-').map(Number);
    const now = new Date();
    let dayCount = monthEndDate(ym);
    const openingMonth = settings.opening_date.slice(0, 7);
    if (ym < openingMonth) dayCount = 0;
    else if (y === now.getFullYear() && mo === now.getMonth() + 1) dayCount = now.getDate();
    else if (new Date(y, mo - 1, 1) > new Date(now.getFullYear(), now.getMonth(), 1)) dayCount = 0;
    const finalDate = dayCount === 0 ? `${ym}-00` : `${ym}-${String(dayCount).padStart(2, '0')}`;
    const quotaStart = monthStart(ym) < settings.opening_date ? settings.opening_date : monthStart(ym);
    const budgetAllocated = dayCount === 0 ? 0 : sumDailyRate(quotaStart, finalDate, 'daily_budget');
    const taxReserve = dayCount === 0 ? 0 : sumDailyRate(quotaStart, finalDate, 'tax_reserve_daily');
    const budgetRemainder = budgetAllocated - expenses;
    const previousMonthDay = new Date(y, mo - 1, 0);
    const monthOpen = ym < openingMonth ? 0 : ym === openingMonth
      ? Number(settings.opening_balance)
      : balanceAt(isoLocalDate(previousMonthDay));
    const cutoff = dayCount === 0 ? '0000-00-00' : finalDate;
    const monthSigned = rows.filter(m => m.movement_date <= cutoff).reduce((s, m) => s + signed(m), 0);
    const closing = monthOpen + monthSigned;
    return { ym, rows, income, expenses, excludedOutflows, budgetAllocated, taxReserve, budgetRemainder, monthOpen, closing, dayCount };
  }

  function categoryStatsForRows(rows) {
    const totals = new Map(CATEGORIES.map(c => [c, 0]));
    for (const m of rows) {
      if (isExcludedMovement(m)) {
        totals.set(m.category, (totals.get(m.category) || 0) + (m.movement_type === 'Uscita' ? m.amount : -m.amount));
      } else if (m.category === SPECIAL_CATEGORY) {
        totals.set(m.category, (totals.get(m.category) || 0) + (m.movement_type === 'Uscita' ? m.amount : -m.amount));
      } else if (m.movement_type === 'Uscita') {
        totals.set(m.category, (totals.get(m.category) || 0) + m.amount);
      }
    }
    return [...totals.entries()].map(([category, total]) => ({ category, total })).filter(x => Math.abs(x.total) >= 0.005).sort((a, b) => b.total - a.total);
  }


  const CASH_BALANCE_KEY_PREFIX = 'bilancio-casa-cash-balance-v1:';
  function cashBalanceKey() { return CASH_BALANCE_KEY_PREFIX + (currentUser?.id || 'local'); }
  function readCashBalance() {
    try {
      const raw = localStorage.getItem(cashBalanceKey());
      if (raw === null) {
        localStorage.setItem(cashBalanceKey(), '50.00');
        return 50;
      }
      const value = Number(raw);
      return Number.isFinite(value) ? value : 50;
    } catch (_) { return 50; }
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
          <div class="summary-grid" style="margin-top:.75rem">
            <div><span class="small-note">Prelievi nel mese selezionato</span><strong id="cashMonthWithdrawals">—</strong></div>
            <div><span class="small-note">Prelievi complessivi registrati</span><strong id="cashTotalWithdrawals">—</strong></div>
          </div>
          <p class="small-note">Il saldo contanti viene salvato su questo dispositivo e non si sincronizza automaticamente con altri dispositivi.</p>
        </section>`);
      panel = $('cashSummaryPanel');
    }
    const ym = $('periodMonth').value || todayMonth();
    const isCashWithdrawal = m => m.category === 'Prelievo contanti' && m.movement_type === 'Uscita';
    const monthTotal = monthMovements(ym).filter(isCashWithdrawal).reduce((sum, m) => sum + Number(m.amount || 0), 0);
    const allTotal = movements.filter(m => inAccountingPeriod(m) && isCashWithdrawal(m)).reduce((sum, m) => sum + Number(m.amount || 0), 0);
    $('cashBalanceInput').value = readCashBalance().toFixed(2);
    $('cashMonthWithdrawals').textContent = money(monthTotal);
    $('cashTotalWithdrawals').textContent = money(allTotal);
  }

  function saveCashBalance() {
    const input = $('cashBalanceInput');
    const value = Number(String(input.value).replace(',', '.'));
    if (!Number.isFinite(value) || value < 0) {
      showNotice('Inserisci un saldo contanti valido, pari o superiore a zero.', 'error');
      return;
    }
    try {
      localStorage.setItem(cashBalanceKey(), (Math.round(value * 100) / 100).toFixed(2));
      input.value = (Math.round(value * 100) / 100).toFixed(2);
      showNotice('Saldo contanti aggiornato su questo dispositivo.');
    } catch (_) {
      showNotice('Non riesco a salvare il saldo contanti nel browser.', 'error');
    }
  }

  function renderDashboard() {
    const ym = $('periodMonth').value || todayMonth();
    const s = monthlyStats(ym);
    $('statBalance').textContent = money(balanceAt(isoToday()));
    $('statIncome').textContent = money(s.income);
    $('statExpenses').textContent = money(s.expenses);
    $('statExpenses').title = `Spese ordinarie: ${money(s.expenses)}. Uscite escluse dal budget: ${money(s.excludedOutflows)}.`;
    $('statSavings').textContent = money(s.budgetRemainder);
    $('statSavings').classList.toggle('amount-out', s.budgetRemainder < 0);
    $('statSavingsCaption').textContent = `${money(s.budgetAllocated)} di quote giornaliere − uscite`;
    $('statTaxReserve').textContent = money(s.taxReserve);
    $('statTaxReserveCaption').textContent = `Accantonamento per ${s.dayCount} giorni`; 
    renderCashPanel();
    renderMovementList($('recentMovements'), [...movements].filter(inAccountingPeriod).sort(sortLatest).slice(0, 8), false);

    const startMonth = settings.opening_date.slice(0, 7);
    const current = todayMonth();
    const [sy, sm] = startMonth.split('-').map(Number);
    const [cy, cm] = current.split('-').map(Number);
    const count = Math.max(0, (cy - sy) * 12 + (cm - sm));
    const ymList = [];
    for (let i = 0; i <= Math.min(count, 120); i++) {
      const d = new Date(sy, sm - 1 + i, 1);
      ymList.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
    }
    $('monthlyRows').innerHTML = ymList.map(m => {
      const ms = monthlyStats(m);
      const displayBalance = ms.dayCount === 0 ? balanceAt(`${m}-00`) : ms.closing;
      return `<tr><td>${escapeHtml(monthLabel(m))}</td><td class="money">${money(ms.income)}</td><td class="money">${money(ms.expenses)}</td><td class="money">${money(ms.taxReserve)}</td><td class="money ${ms.budgetRemainder < 0 ? 'amount-out' : ''}">${money(ms.budgetRemainder)}</td><td class="money">${money(displayBalance)}</td></tr>`;
    }).join('') || '<tr><td colspan="5" class="table-empty">Nessun riepilogo disponibile.</td></tr>';
  }

  function sortLatest(a, b) {
    const dateCompare = b.movement_date.localeCompare(a.movement_date);
    if (dateCompare !== 0) return dateCompare;
    return String(b.created_at || '').localeCompare(String(a.created_at || ''));
  }

  function renderMovementList(container, rows, withActions) {
    if (!rows.length) {
      container.innerHTML = '<p class="empty-state">Non ci sono movimenti da mostrare.</p>';
      return;
    }
    container.innerHTML = rows.map(m => {
      const incoming = m.movement_type === 'Entrata';
      const amt = `${incoming ? '+' : '−'} ${money(m.amount)}`;
      return `<div class="movement-item">
        <div class="movement-icon" aria-hidden="true">${incoming ? '↙' : '↗'}</div>
        <div class="movement-main"><div class="movement-title">${escapeHtml(m.concept)}</div>
          <div class="movement-meta">${escapeHtml(fmtDate(m.movement_date))} · ${escapeHtml(m.category)}${m.notes ? ` · ${escapeHtml(m.notes)}` : ''}</div></div>
        <div class="movement-right"><span class="movement-amount ${incoming ? 'amount-in' : 'amount-out'}">${escapeHtml(amt)}</span>
          ${withActions ? `<div class="row-actions"><button class="icon-button" type="button" data-edit="${escapeHtml(m.id)}">Modifica</button><button class="icon-button" type="button" data-delete="${escapeHtml(m.id)}">Elimina</button></div>` : ''}
        </div>
      </div>`;
    }).join('');
  }

  function renderAllMovements() {
    const rows = [...movements].filter(inAccountingPeriod).sort(sortLatest);
    $('movementCount').textContent = `${rows.length} movimenti registrati`;
    renderMovementList($('allMovements'), rows.slice(0, 100), true);
    if (rows.length > 100) {
      $('allMovements').insertAdjacentHTML('beforeend', `<p class="small-note">Mostrati gli ultimi 100 movimenti su ${rows.length}. Tutti i movimenti restano nell'archivio ed entrano nei riepiloghi.</p>`);
    }
  }

  function renderCategoryList(container, rows, totalNode) {
    const total = rows.reduce((s, x) => s + x.total, 0);
    totalNode.textContent = money(total);
    if (!rows.length) {
      container.innerHTML = '<p class="empty-state">Nessuna spesa per questo periodo.</p>';
      return;
    }
    const max = Math.max(...rows.map(x => Math.max(x.total, 0)), 0.01);
    container.innerHTML = rows.map(r => {
      const pct = total > 0 ? r.total / total * 100 : 0;
      const width = Math.max(0, Math.min(100, r.total / max * 100));
      return `<div class="category-row">
        <span class="category-label">${escapeHtml(r.category)}</span>
        <span class="category-amount ${r.total < 0 ? 'amount-in' : ''}">${money(r.total)}</span>
        <div class="category-bar"><div class="bar-track"><div class="bar-fill" style="width:${width.toFixed(2)}%"></div></div><span class="category-percent">${total > 0 ? `${pct.toLocaleString('it-IT', { maximumFractionDigits: 1 })}%` : '—'}</span></div>
      </div>`;
    }).join('');
  }

  function renderCategories() {
    const ym = $('periodMonth').value || todayMonth();
    $('categoryMonthTitle').textContent = `Spese di ${monthLabel(ym)}`;
    const ymRows = categoryStatsForRows(monthMovements(ym));
    renderCategoryList($('monthlyCategories'), ymRows, $('categoryMonthTotal'));
    const year = Number(ym.slice(0, 4));
    const start = `${year}-01-01`;
    const end = `${year}-12-31`;
    const annualRows = movements.filter(m => inAccountingPeriod(m) && m.movement_date >= start && m.movement_date <= end);
    $('categoryYearTitle').textContent = `Totali ${year}`;
    renderCategoryList($('annualCategories'), categoryStatsForRows(annualRows), $('categoryYearTotal'));
  }

  function renderSettings() {
    $('openingDate').value = settings.opening_date;
    $('openingBalance').value = Number(settings.opening_balance).toFixed(2);
    if (!$('quotaEffectiveDate').value) $('quotaEffectiveDate').value = isoToday();
    const rates = ratesForDate($('quotaEffectiveDate').value || isoToday());
    $('dailyBudget').value = Number(rates.daily_budget).toFixed(2);
    $('taxReserveDaily').value = Number(rates.tax_reserve_daily).toFixed(2);
    $('quotaHistoryRows').innerHTML = quotaHistory.slice().reverse().map(item => `<tr><td>${escapeHtml(fmtDate(item.valid_from))}</td><td class="money">${money(item.daily_budget)}</td><td class="money">${money(item.tax_reserve_daily)}</td></tr>`).join('') || '<tr><td colspan="3" class="table-empty">Nessun parametro registrato.</td></tr>';
  }

  function updateQuotaFormForDate() {
    const date = $('quotaEffectiveDate').value;
    if (!date) return;
    const rates = ratesForDate(date);
    $('dailyBudget').value = Number(rates.daily_budget).toFixed(2);
    $('taxReserveDaily').value = Number(rates.tax_reserve_daily).toFixed(2);
  }

  function renderAll() {
    renderDashboard();
    renderAllMovements();
    renderCategories();
    renderSettings();
  }

  async function saveMovement(event) {
    event.preventDefault();
    const id = $('movementId').value;
    const movementDate = $('movementDate').value;
    const concept = $('movementConcept').value.trim();
    const category = $('movementCategory').value;
    const movementType = $('movementType').value;
    const amount = Number(String($('movementAmount').value).replace(',', '.'));
    const notes = $('movementNotes').value.trim();
    if (!movementDate || !concept || !category || !['Entrata', 'Uscita'].includes(movementType) || !Number.isFinite(amount) || amount <= 0) {
      showNotice('Controlla data, concetto, tipo e importo. L’importo deve essere positivo.', 'error');
      return;
    }
    if (movementDate < settings.opening_date) {
      showNotice(`La data è precedente all'inizio della contabilità (${fmtDate(settings.opening_date)}). Modifica prima la data iniziale nelle impostazioni se vuoi registrare movimenti precedenti.`, 'error');
      return;
    }
    const payload = { user_id: currentUser.id, movement_date: movementDate, concept, category, movement_type: movementType, amount: Math.round(amount * 100) / 100, notes };
    const button = $('saveMovementBtn');
    setButtonBusy(button, true, 'Salvataggio…');
    try {
      let query;
      if (id) query = db.from('bilancio_movements').update(payload).eq('id', id).eq('user_id', currentUser.id);
      else query = db.from('bilancio_movements').insert(payload);
      const { error } = await query;
      if (error) throw error;
      defaultMovementForm();
      await loadMovements();
      renderAll();
      showNotice(id ? 'Movimento aggiornato.' : 'Movimento salvato. Riepiloghi aggiornati.');
    } catch (error) {
      showNotice(`Non sono riuscito a salvare il movimento: ${error?.message || error}`, 'error');
    } finally {
      setButtonBusy(button, false);
    }
  }

  function startEdit(id) {
    const m = movements.find(item => item.id === id);
    if (!m) return;
    $('movementId').value = m.id;
    $('movementDate').value = m.movement_date;
    $('movementType').value = m.movement_type;
    $('movementConcept').value = m.concept;
    if (!CATEGORIES.includes(m.category)) CATEGORIES.push(m.category);
    fillCategoryOptions();
    $('movementCategory').value = m.category;
    $('movementAmount').value = Number(m.amount).toFixed(2);
    $('movementNotes').value = m.notes || '';
    $('movementFormTitle').textContent = 'Modifica movimento';
    $('saveMovementBtn').textContent = 'Salva modifiche';
    $('cancelEditBtn').hidden = false;
    setView('movements');
    $('movementConcept').focus();
  }

  async function deleteMovement(id) {
    const m = movements.find(item => item.id === id);
    if (!m) return;
    if (!window.confirm(`Vuoi eliminare il movimento “${m.concept}” del ${fmtDate(m.movement_date)} di ${money(m.amount)}?`)) return;
    try {
      const { error } = await db.from('bilancio_movements').delete().eq('id', id).eq('user_id', currentUser.id);
      if (error) throw error;
      await loadMovements();
      renderAll();
      showNotice('Movimento eliminato e riepiloghi aggiornati.');
    } catch (error) {
      showNotice(`Non è stato possibile eliminare il movimento: ${error?.message || error}`, 'error');
    }
  }

  async function saveSettings(event) {
    event.preventDefault();
    const openingDate = $('openingDate').value;
    const openingBalance = Number(String($('openingBalance').value).replace(',', '.'));
    if (!openingDate || !Number.isFinite(openingBalance)) {
      showNotice('Controlla la data iniziale e il saldo iniziale.', 'error');
      return;
    }
    if (openingDate !== settings.opening_date && movements.some(m => m.movement_date < openingDate)) {
      const ok = window.confirm('La nuova data iniziale renderà esclusi dal saldo i movimenti precedenti. Vuoi continuare?');
      if (!ok) return;
    }
    const newSettings = { user_id: currentUser.id, opening_date: openingDate, opening_balance: Math.round(openingBalance * 100) / 100, daily_budget: Number(settings.daily_budget) || DEFAULTS.daily_budget, updated_at: new Date().toISOString() };
    try {
      const { error } = await db.from('bilancio_settings').upsert(newSettings, { onConflict: 'user_id' });
      if (error) throw error;
      settings = { opening_date: openingDate, opening_balance: newSettings.opening_balance, daily_budget: newSettings.daily_budget };
      renderAll();
      showNotice('Parametri del conto salvati. Riepiloghi aggiornati.');
    } catch (error) {
      showNotice(`Non è stato possibile salvare i parametri del conto: ${error?.message || error}`, 'error');
    }
  }

  async function saveQuotaParameters(event) {
    event.preventDefault();
    const validFrom = $('quotaEffectiveDate').value;
    const dailyBudget = Number(String($('dailyBudget').value).replace(',', '.'));
    const taxReserveDaily = Number(String($('taxReserveDaily').value).replace(',', '.'));
    if (!validFrom || !Number.isFinite(dailyBudget) || dailyBudget < 0 || !Number.isFinite(taxReserveDaily) || taxReserveDaily < 0) {
      showNotice('Inserisci una data valida e quote giornaliere non negative.', 'error');
      return;
    }
    if (validFrom < settings.opening_date) {
      showNotice(`La data di efficacia non può precedere l'inizio della contabilità (${fmtDate(settings.opening_date)}).`, 'error');
      return;
    }
    const exactExisting = quotaHistory.find(item => item.valid_from === validFrom);
    if (validFrom < isoToday()) {
      const ok = window.confirm(`La modifica avrà effetto dal ${fmtDate(validFrom)} e cambierà i riepiloghi storici da quella data. Vuoi continuare?`);
      if (!ok) return;
    }
    const record = {
      user_id: currentUser.id,
      valid_from: validFrom,
      daily_budget: Math.round(dailyBudget * 100) / 100,
      tax_reserve_daily: Math.round(taxReserveDaily * 100) / 100,
      updated_at: new Date().toISOString()
    };
    try {
      const { error } = await db.from('bilancio_quota_storico').upsert(record, { onConflict: 'user_id,valid_from' });
      if (error) throw error;
      await loadQuotaHistory();
      settings.daily_budget = Number(ratesForDate(isoToday()).daily_budget);
      renderAll();
      $('quotaEffectiveDate').value = validFrom;
      updateQuotaFormForDate();
      showNotice(exactExisting ? 'Quote aggiornate dalla data indicata.' : 'Nuove quote registrate con decorrenza dalla data indicata.');
    } catch (error) {
      showNotice(`Non è stato possibile salvare le quote: ${error?.message || error}`, 'error');
    }
  }

  function parseAmount(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    let s = String(value ?? '').trim().replace(/[€\s]/g, '');
    if (!s) return NaN;
    if (s.includes(',') && s.includes('.')) s = s.replace(/\./g, '').replace(',', '.');
    else if (s.includes(',')) s = s.replace(',', '.');
    s = s.replace(/[^0-9.\-]/g, '');
    const n = Number(s);
    return Number.isFinite(n) ? n : NaN;
  }

  function parseDateCell(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`;
    }
    if (typeof value === 'number' && window.XLSX?.SSF) {
      const d = XLSX.SSF.parse_date_code(value);
      if (d && d.y && d.m && d.d) return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
    }
    const s = String(value ?? '').trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const it = s.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})$/);
    if (it) return `${it[3]}-${String(it[2]).padStart(2, '0')}-${String(it[1]).padStart(2, '0')}`;
    if (s) {
      const d = new Date(s);
      if (!Number.isNaN(d.getTime())) return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }
    return '';
  }

  function headerName(v) { return normText(v).replace(/[^a-z0-9]/g, ''); }
  function findWorkbookMovementSheet(workbook) {
    const preferred = workbook.SheetNames.slice().sort((a, b) => {
      const aa = normText(a).includes('moviment') ? 0 : 1;
      const bb = normText(b).includes('moviment') ? 0 : 1;
      return aa - bb;
    });
    for (const sheetName of preferred) {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
      for (let r = 0; r < Math.min(rows.length, 40); r++) {
        const headers = rows[r].map(headerName);
        const find = names => headers.findIndex(h => names.includes(h));
        const indexes = {
          date: find(['data']),
          concept: find(['concetto', 'descrizione', 'concept']),
          category: find(['categoria', 'category']),
          type: find(['tipo', 'movementtype']),
          amount: find(['importo', 'amount']),
          signed: find(['importopiu', 'movimentofirmato', 'importofirmato', 'signedamount']),
          notes: find(['note', 'notes']),
          id: find(['idmovimento', 'movementid', 'movementuuid', 'id'])
        };
        if (indexes.date >= 0 && indexes.concept >= 0 && indexes.type >= 0 && (indexes.amount >= 0 || indexes.signed >= 0)) {
          return { sheetName, rows, headerRow: r, indexes };
        }
      }
    }
    return null;
  }

  async function fileHash(file, arrayBuffer) {
    try {
      if (window.crypto?.subtle) {
        const digest = await crypto.subtle.digest('SHA-256', arrayBuffer);
        return [...new Uint8Array(digest)].map(b => b.toString(16).padStart(2, '0')).join('');
      }
    } catch (_) { /* fallback deterministico */ }
    return `${file.name}-${file.size}-${file.lastModified}`.replace(/[^a-zA-Z0-9_-]/g, '_');
  }

  async function importWorkbook(file) {
    if (!window.XLSX) throw new Error('La libreria di lettura Excel non è disponibile. Controlla la connessione e riprova.');
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    const parsed = findWorkbookMovementSheet(workbook);
    if (!parsed) throw new Error('Non trovo un foglio con le colonne Data, Concetto, Tipo e Importo. Usa il file Excel di controllo bancario.');
    const { rows, headerRow, indexes, sheetName } = parsed;
    const fileHashValue = await fileHash(file, arrayBuffer);
    const payload = [];
    let opening = null;
    let skipped = 0;
    for (let r = headerRow + 1; r < rows.length; r++) {
      const row = rows[r];
      const date = parseDateCell(row[indexes.date]);
      const concept = String(row[indexes.concept] ?? '').trim();
      if (!date || !concept) { skipped++; continue; }
      const category = String(indexes.category >= 0 ? row[indexes.category] ?? 'Altro' : 'Altro').trim() || 'Altro';
      const typeRaw = normText(indexes.type >= 0 ? row[indexes.type] : '');
      const selectedAmount = indexes.amount >= 0 && row[indexes.amount] !== null ? parseAmount(row[indexes.amount]) : parseAmount(row[indexes.signed]);
      const signedAmount = indexes.signed >= 0 ? parseAmount(row[indexes.signed]) : selectedAmount;
      const amountRaw = Number.isFinite(selectedAmount) ? selectedAmount : signedAmount;
      if (!Number.isFinite(amountRaw) || Math.abs(amountRaw) < 0.005) { skipped++; continue; }
      const isOpening = /saldo\s*(conto|iniziale)|saldo conto|saldo iniziale/i.test(concept);
      if (isOpening) {
        if (!opening) opening = { opening_date: date, opening_balance: Math.abs(amountRaw) };
        continue;
      }
      let movementType = typeRaw.startsWith('entr') ? 'Entrata' : typeRaw.startsWith('usc') ? 'Uscita' : (signedAmount >= 0 ? 'Entrata' : 'Uscita');
      const absAmount = Math.round(Math.abs(amountRaw) * 100) / 100;
      if (date < (opening?.opening_date || settings.opening_date)) { skipped++; continue; }
      const notes = indexes.notes >= 0 ? String(row[indexes.notes] ?? '').trim() : '';
      const suppliedId = indexes.id >= 0 ? String(row[indexes.id] ?? '').trim() : '';
      const hasStableId = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(suppliedId);
      payload.push({
        ...(hasStableId ? { id: suppliedId } : {}),
        user_id: currentUser.id,
        movement_date: date,
        concept,
        category,
        movement_type: movementType,
        amount: absAmount,
        notes,
        ...(hasStableId ? {} : { import_key: `xlsx:${fileHashValue}:${sheetName}:${r + 1}` })
      });
    }
    if (!payload.length && !opening) throw new Error('Non ho trovato movimenti da importare. Controlla il foglio scelto e le intestazioni.');
    const openingDate = opening?.opening_date || settings.opening_date;
    if (openingDate && payload.some(m => m.movement_date < openingDate)) {
      for (let i = payload.length - 1; i >= 0; i--) if (payload[i].movement_date < openingDate) payload.splice(i, 1);
    }
    const explanation = `Foglio rilevato: ${sheetName}. Movimenti da importare: ${payload.length}${opening ? `; saldo iniziale: ${money(opening.opening_balance)} dal ${fmtDate(opening.opening_date)}` : ''}.\nLe righe con ID movimento vengono aggiornate senza creare doppioni; le righe senza ID vengono aggiunte senza cancellare movimenti esistenti. Continuare?`;
    if (!window.confirm(explanation)) return;
    if (opening) {
      const newSettings = { user_id: currentUser.id, opening_date: opening.opening_date, opening_balance: opening.opening_balance, daily_budget: settings.daily_budget, updated_at: new Date().toISOString() };
      const { error } = await db.from('bilancio_settings').upsert(newSettings, { onConflict: 'user_id' });
      if (error) throw error;
      settings = { opening_date: opening.opening_date, opening_balance: opening.opening_balance, daily_budget: settings.daily_budget };
    }
    const withStableId = payload.filter(m => m.id);
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
    await refreshData();
    setView('dashboard');
    showNotice(`Importazione completata. Elaborate ${payload.length} righe${skipped ? `; ignorate ${skipped} righe vuote o non utilizzabili` : ''}. Le righe con ID movimento mantengono la propria identità tra esportazioni e reimportazioni; le righe senza ID usano il sistema di importazione precedente.`, 'ok');
  }

  function csvCell(v) {
    const s = String(v ?? '');
    return `"${s.replace(/"/g, '""')}"`;
  }
  function csvNumber(v) { return Number(v || 0).toFixed(2).replace('.', ','); }
  function csvDate(v) {
    const [y, m, d] = String(v).slice(0, 10).split('-');
    return `${d}/${m}/${y}`;
  }
  function exportCsv() {
    const rows = [['ID movimento', 'Data', 'Giorno', 'Concetto', 'Categoria', 'Tipo', 'Importo', 'Importo +/-', 'Saldo', 'Mese', 'Note']];
    const all = [...movements].filter(inAccountingPeriod).sort((a, b) => a.movement_date.localeCompare(b.movement_date) || String(a.created_at).localeCompare(String(b.created_at)));
    let balance = Number(settings.opening_balance);
    rows.push(['', csvDate(settings.opening_date), '', 'Saldo conto', 'Altro', 'Entrata', csvNumber(settings.opening_balance), csvNumber(settings.opening_balance), csvNumber(balance), '', 'Saldo iniziale importato dalle impostazioni']);
    for (const m of all) {
      balance += signed(m);
      const monthName = new Intl.DateTimeFormat('it-IT', { month: 'long', year: 'numeric' }).format(new Date(`${m.movement_date}T12:00:00`));
      rows.push([m.id || '', csvDate(m.movement_date), '', m.concept, m.category, m.movement_type, csvNumber(m.amount), csvNumber(signed(m)), csvNumber(balance), monthName, [m.notes || '', isExcludedMovement(m) ? 'ESCLUSO DALLE SPESE ORDINARIE' : ''].filter(Boolean).join(' | ')]);
    }
    const content = '\ufeff' + rows.map(row => row.map(csvCell).join(';')).join('\r\n');
    const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Bilancio_Casa_${isoToday()}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showNotice('Esportazione CSV avviata. Conserva il file in un luogo sicuro.');
  }

  function bindEvents() {
    $('authForm').addEventListener('submit', handleAuthSubmit);
    $('showSignupBtn').addEventListener('click', () => switchAuthMode('signup'));
    $('showSigninBtn').addEventListener('click', () => switchAuthMode('signin'));
    $('signOutBtn').addEventListener('click', async () => {
      const { error } = await db.auth.signOut();
      if (error) showNotice(`Errore durante l'uscita: ${error.message}`, 'error');
      else {
        currentUser = null; movements = []; quotaHistory = []; settings = { ...DEFAULTS };
        setAuthenticatedPanels(false);
        $('authEmail').value = '';
        $('authPassword').value = '';
        switchAuthMode('signin');
        showNotice('Hai effettuato la disconnessione.');
      }
    });
    $('movementForm').addEventListener('submit', saveMovement);
    $('cancelEditBtn').addEventListener('click', defaultMovementForm);
    $('settingsForm').addEventListener('submit', saveSettings);
    $('quotaForm').addEventListener('submit', saveQuotaParameters);
    $('quotaEffectiveDate').addEventListener('change', updateQuotaFormForDate);
    $('periodMonth').addEventListener('change', () => { renderDashboard(); renderCategories(); });
    document.querySelectorAll('.nav-button').forEach(button => button.addEventListener('click', () => setView(button.dataset.view)));
    document.querySelectorAll('[data-go-view]').forEach(button => button.addEventListener('click', () => setView(button.dataset.goView)));
    $('allMovements').addEventListener('click', event => {
      const edit = event.target.closest('[data-edit]');
      const del = event.target.closest('[data-delete]');
      if (edit) startEdit(edit.dataset.edit);
      else if (del) deleteMovement(del.dataset.delete);
    });
    $('importBtn').addEventListener('click', () => $('importFile').click());
    $('importFile').addEventListener('change', async event => {
      const file = event.target.files?.[0];
      if (!file) return;
      const button = $('importBtn');
      setButtonBusy(button, true, 'Importazione…');
      try { await importWorkbook(file); }
      catch (error) { showNotice(`Importazione non riuscita: ${error?.message || error}`, 'error'); }
      finally { setButtonBusy(button, false); event.target.value = ''; }
    });
    $('exportBtn').addEventListener('click', exportCsv);
    document.addEventListener('click', event => { if (event.target.closest('#saveCashBalanceBtn')) saveCashBalance(); });
    window.addEventListener('beforeinstallprompt', event => {
      event.preventDefault();
      pendingInstallPrompt = event;
      $('installBtn').hidden = false;
    });
    $('installBtn').addEventListener('click', async () => {
      if (!pendingInstallPrompt) return;
      pendingInstallPrompt.prompt();
      await pendingInstallPrompt.userChoice;
      pendingInstallPrompt = null;
      $('installBtn').hidden = true;
    });
  }

  async function initialize() {
    fillCategoryOptions();
    bindEvents();
    $('periodMonth').value = todayMonth();
    $('movementDate').value = isoToday();
    const cfg = window.BILANCIO_CONFIG || {};
    const key = String(cfg.SUPABASE_PUBLISHABLE_KEY || '');
    const url = String(cfg.SUPABASE_URL || '');
    if (!url.startsWith('https://') || !key || key.includes('INCOLLA_QUI')) {
      showSetupRequired();
    } else if (!window.supabase?.createClient) {
      $('configPanel').hidden = false;
      $('configPanel').querySelector('h2').textContent = 'Libreria di collegamento non disponibile';
      $('configPanel').querySelectorAll('p')[0].textContent = 'Controlla la connessione internet e ricarica la pagina.';
      $('configPanel').querySelectorAll('p')[1].hidden = true;
    } else {
      db = window.supabase.createClient(url, key, { auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true } });
      const { data, error } = await db.auth.getSession();
      if (error) showNotice(error.message, 'error');
      if (data?.session?.user) {
        currentUser = data.session.user;
        await afterLogin();
      } else {
        setAuthenticatedPanels(false);
      }
      db.auth.onAuthStateChange((event, session) => {
        if (event === 'SIGNED_OUT' || !session) {
          currentUser = null;
          $('appPanel').hidden = true;
          $('authPanel').hidden = false;
        }
      });
    }
    if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
      navigator.serviceWorker.register('./service-worker.js').catch(() => {});
    }
  }

  document.addEventListener('DOMContentLoaded', initialize);
})();
