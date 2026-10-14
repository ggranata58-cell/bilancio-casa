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
"    'Finanz. Pannelli solari', 'Pensione', 'Imposte e tasse', 'Banca', 'Contanti', 'Altro'\n  ];\n  const SPECIAL_CATEGORY = 'Tasas y Intereses';",
"    'Finanz. Pannelli solari', 'Pensione', 'Imposte e tasse', 'Banca', 'Contanti', 'Altro',\n    'Movimenti esclusi', 'Prelievo contanti'\n  ];\n  const SPECIAL_CATEGORY = 'Tasas y Intereses';\n  const EXCLUDED_CATEGORIES = new Set(['Movimenti esclusi', 'Prelievo contanti']);\n  function isExcludedMovement(m) { return EXCLUDED_CATEGORIES.has(m.category); }",
"excluded movement categories")
replace_once(
"    const expenses = rows.filter(m => m.movement_type === 'Uscita').reduce((s, m) => s + m.amount, 0);",
"    const expenses = rows.filter(m => m.movement_type === 'Uscita' && !isExcludedMovement(m)).reduce((s, m) => s + m.amount, 0);\n    const excludedOutflows = rows.filter(m => m.movement_type === 'Uscita' && isExcludedMovement(m)).reduce((s, m) => s + m.amount, 0);",
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
    os.replace(temp_path, archive)
finally:
    if os.path.exists(temp_path):
        os.remove(temp_path)
print("Applied and verified stable Excel movement IDs.")
