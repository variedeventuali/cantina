# Cantina (web app per telefono) – v1.2, 21 ott 2026

Versione installabile (PWA) della tua `Cantina-app` (Mac, `server.py`): stessa interfaccia (Rossi, Non rossi, Dolci, Magnum, Bevuti, Statistiche, controllo doppioni, "bevuta", cantinetta) ma sul telefono.
**Il file Excel è l'unica fonte dei dati**: `Vino.xlsx` nel repo PRIVATO `scheda-dati`. Nel repo pubblico dell'app c'è solo il codice, nessun vino.

## Come funziona
- L'app tiene sul telefono una copia di Vino.xlsx e un registro delle modifiche (aggiungi, modifica, +1, bevuta, elimina). Funziona anche offline.
- **Sincronizza solo quando premi ⟳ in alto** (il numero rosso = modifiche da inviare). Alla pressione: scarica l'ultima versione dal repo, ci riapplica le tue modifiche, e la salva su GitHub. Così due telefoni non si sovrascrivono: ogni modifica viene applicata sopra l'ultimo file.
- Le bottiglie si riconoscono da Nome + Cantina + Anno + Acquisto. Se un'altra persona ha eliminato/cambiato quella riga nel frattempo, la tua modifica su quella riga viene scartata e l'app lo segnala.
- Stili, formule (`=SUM`), filtri, formati € e date del file restano (testato con il tuo Vino.xlsx). Nel file compare un foglio nascosto `_sync` (elenco delle operazioni già inviate, serve a non ripeterle se un invio si interrompe). Non lo vedi in Excel salvo "Mostra".
- Ogni sincronizzazione crea una versione nella cronologia di GitHub: è anche il tuo backup (puoi tornare a una versione precedente dal repo → Vino.xlsx → History).

## Messa online
1. **Dati**: nel repo privato `scheda-dati` → Add file → Upload files → carica `Vino.xlsx` (anche il tuo originale in Downloads) nella cartella principale.
2. **App**: nuovo repo pubblico `cantina` → carica i 7 file di `sito/` (index.html, exceljs.min.js, manifest.webmanifest, sw.js, 3 icone) → Settings → Pages → `main` / root.
3. Apri https://variedeventuali.github.io/cantina/ → Chrome ⋮ → Installa app.
4. In app: ⚙ → repo `variedeventuali/scheda-dati`, file `Vino.xlsx`, token (pulsante "Usa il token di palestra/casa") → Salva e sincronizza: scarica il file.
Alternativa senza GitHub per il primo avvio: ⚙ → "Carica Vino.xlsx dal telefono…", poi ⟳ lo salva nel repo.

## Sul Mac
`Cantina-app` (server.py) lavora su `~/Downloads/Vino.xlsx` locale e NON è sincronizzata con GitHub. Per usarla con i dati aggiornati: dall'app ⚙ → "Scarica l'Excel aggiornato" (o scarica il file dal repo) e sostituisci il file in Downloads. Meglio usare una sola via per volta.

## Aggiornare l'app
Ricarica i file di `sito/` su GitHub; cambia `APP_VERSION` in index.html e `V` in sw.js (cantina-v4…). Dopo l'aggiornamento chiudi e riapri l'app.

## Test
`test/cantina-test.js` (Node + jsdom + exceljs + finto GitHub, usa `py_read.py` per confrontare la lettura con il motore Python): 19 controlli – lettura identica a server.py su tutti i fogli, nessun invio automatico, invio con pulsante, stili e formule conservati, due telefoni con modifiche concorrenti, blocchi Dolci/Magnum, eliminazioni. Non provata su telefono reale né su GitHub vero.

## Limiti
- Il file viene riscritto da ExcelJS: se in futuro aggiungi grafici/immagini in Excel potrebbero andare persi (oggi il file non ne ha).
- Il filtro automatico di Excel mantiene il suo intervallo originale (come faceva anche server.py): righe nuove possono stare fuori dal filtro finché non lo riapplichi.
- Se modifichi Vino.xlsx a mano su GitHub/Excel mentre l'app ha modifiche in sospeso, al ⟳ le modifiche dell'app vengono riapplicate sopra il tuo file.
- File oltre 1 MB non sono leggibili con questo metodo (oggi 140 KB).

## Versioni
- 1.0: prima versione per telefono, sync su pulsante, Vino.xlsx nel repo privato.
- 1.1: il contatore sotto i filtri mostra le bottiglie (non le voci); con il filtro cantinetta mostra le bottiglie in cantinetta.
- 1.2: nelle Statistiche, sezione "Ultime bottiglie aggiunte" (le 6 con acquisto più recente tra Rossi e Non rossi; toccando una si apre la scheda).
