# Bilancio Casa

Applicazione web installabile (PWA) per la gestione del bilancio personale.

## Pubblicazione

La versione online viene pubblicata su GitHub Pages tramite GitHub Actions, dal branch `main`.

Il workflow si trova in `.github/workflows/deploy.yml`. Ricostruisce il pacchetto dell'app dai file `Bilancio_Casa_PWA_v0_2_0.zip.b64.part*`, verifica l'archivio e applica le modifiche previste da `patch_archive.py` prima della pubblicazione.

Per questo motivo, il pacchetto compresso e `patch_archive.py` sono parte della procedura di distribuzione: non eliminarli né modificarli senza verificare prima il workflow. Le modifiche al codice dell'app devono essere applicate alla sorgente effettivamente usata da questa procedura.

Il workflow si avvia automaticamente quando cambiano `Bilancio_Casa_PWA_v0_2_0.zip.b64.ready`, `patch_archive.py` o il file del workflow; è inoltre possibile avviarlo manualmente da GitHub Actions.

## Archivio

La cartella `archivio-sorgenti-sync-bidirezionale` contiene copie di riferimento del vecchio ramo di test. È un archivio di recupero e non fa parte della PWA pubblicata.

## Dati e sicurezza

Il repository non dovrebbe contenere dati bancari personali, password, refresh token o chiavi segrete. Non caricare backup personali su GitHub. La cartella di archivio può contenere sorgenti storici: prima di riutilizzarli, verificarne il contenuto e aggiornare eventuali credenziali pubbliche sostituite da segnaposto.
