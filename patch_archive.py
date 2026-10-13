#!/usr/bin/env python3
"""Apply the bidirectional Excel-ID improvement to the packaged PWA archive."""
import os
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
if "ID movimento" in app and "withStableId" in app:
    print("Bidirectional Excel-ID improvement already applied.")
    raise SystemExit(0)

def replace_once(old, new, label):
    global app
    count = app.count(old)
    if count != 1:
        raise SystemExit(f"Expected one occurrence for {label}, found {count}")
    app = app.replace(old, new, 1)

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
"        import_key: hasStableId ? `excel-id:${suppliedId}` : `xlsx:${fileHashValue}:${sheetName}:${r + 1}`\\n      });",
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
replace_once(old_batch, new_batch, "separate stable-ID and legacy import batches")
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
