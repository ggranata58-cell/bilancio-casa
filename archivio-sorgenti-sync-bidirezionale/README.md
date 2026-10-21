# Archivio sorgenti sincronizzazione bidirezionale

Copie di sicurezza dei sorgenti presenti nel ramo `test/sync-bidirezionale`, al 21 ottobre 2026. Sono archiviate per riferimento e recupero e non fanno parte della pubblicazione della PWA: non modificano l'app online.

- `app.js`: sorgente JavaScript dell'app di test.
- `SincronizzaBilancioSupabase.bas`: modulo VBA per la sincronizzazione Excel/Supabase.
- `AggiornaMovimentiDaApp.bas`: macro di importazione movimenti.
- `CreaFoglioContanti.bas`: macro per il foglio Contanti.
- `EventoThisWorkbook_AutoSync.txt`: codice evento da inserire in ThisWorkbook.

Per sicurezza, l'eventuale chiave publishable Supabase è stata sostituita con un segnaposto nel modulo VBA. Prima di riutilizzarlo, inserire la chiave pubblica corretta. Non archiviare password, refresh token o chiavi service-role.
