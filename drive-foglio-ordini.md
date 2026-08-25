# Foglio ordini su Google Drive (accesso solo su invito)

La pagina non può creare da sola un file sul tuo Drive: serve un piccolo script che il gruppo
autorizza una volta. Dieci minuti, si fa dal browser.

## 1. Crea il foglio
1. Su Google Drive (account del gruppo) crea un nuovo **Foglio Google**, nome es. `Ordini 50° — Castel Maggiore 1`.
2. Prima riga come intestazione:
   `Data | Codice | Nome | Contatto | Articolo | Colore | Taglia | Quantità | Prezzo | Note`
3. Condivisione: lascia **Con limitazioni** e aggiungi solo le persone invitate (i capi).
   Chi prenota non vede il foglio: riceve solo la conferma con il codice.

## 2. Collega lo script
Nel foglio: **Estensioni → Apps Script**, cancella tutto il codice presente e incolla questo, poi
salva (icona dischetto o Ctrl/Cmd+S).

Accetta i tre modi in cui la pagina può inviare i dati (JSON via `fetch`, modulo nascosto,
richiesta GET): serve perché il browser, quando la pagina è pubblicata (GitHub Pages o artefatto),
a volte perde il corpo della richiesta POST durante il redirect che Google fa verso `/exec`. La
pagina manda tutti e tre i canali insieme invece che uno alla volta (così se uno perde il corpo per
strada, gli altri due arrivano comunque); **proprio perché arrivano insieme**, lo script usa un
blocco (`LockService`) per evitare che due arrivi quasi simultanei dello stesso ordine finiscano
per scrivere due righe invece di una — senza il blocco, entrambi potrebbero leggere il foglio
"prima" che l'altro abbia scritto, e non vedersi a vicenda.

> **Se avevi già incollato una versione precedente di questo script**, sostituiscila con questa:
> aggiunge il blocco anti-doppioni e non si blocca più su un payload malformato (prima, un invio
> troncato o corrotto faceva fallire l'intera esecuzione invece di limitarsi a scartarlo). Toglie
> anche la colonna **Unità**, che il modulo di prenotazione non ha mai raccolto: la versione
> precedente scriveva comunque 11 valori su un foglio con 10 intestazioni, spostando di una
> colonna tutto quello che viene dopo "Nome" (contatto sotto "Articolo", articolo sotto "Colore",
> ecc.) — se hai già delle righe scritte con la vecchia versione, controllale a mano. Manda anche
> un'email con il PDF della prenotazione a un indirizzo fisso (variabile `EMAIL_DESTINATARIO` in
> cima allo script): cambia quel valore se vuoi che vada a un indirizzo diverso.

```js
var EMAIL_DESTINATARIO = 'gastaniscorrazza@gmail.com';

function handle(raw) {
  if (!raw) return ContentService.createTextOutput('vuoto');
  var d;
  try { d = JSON.parse(raw); } catch (e) { return ContentService.createTextOutput('payload non valido'); }
  if (!d || !d.codice) return ContentService.createTextOutput('payload non valido');

  // Blocco: i tre canali (fetch, modulo, GET) mandano lo stesso ordine quasi in
  // contemporanea. Senza questo lock, due esecuzioni potrebbero leggere il foglio
  // prima che l'altra abbia scritto e finire per duplicare la riga (e mandare due
  // email per lo stesso ordine).
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // già registrato? esci (evita doppioni dai tentativi multipli)
    var codici = sh.getRange(1, 2, Math.max(sh.getLastRow(), 1), 1).getValues();
    for (var i = 0; i < codici.length; i++) {
      if (codici[i][0] === d.codice) return ContentService.createTextOutput('già presente');
    }

    (d.articoli || []).forEach(function (a) {
      sh.appendRow([new Date(), d.codice, d.nome, d.contatto,
                    a.articolo, a.colore, a.taglia, a.quantita, a.prezzo, d.note]);
    });

    // Email con il PDF allegato: un fallimento qui (quota mail esaurita,
    // indirizzo non valido, ecc.) non deve mai far sparire la riga già scritta
    // sopra, quindi è isolato nel suo try/catch.
    try {
      var attachments = [];
      if (d.pdfBase64) {
        attachments.push(Utilities.newBlob(Utilities.base64Decode(d.pdfBase64), 'application/pdf', 'prenotazione-' + d.codice + '.pdf'));
      }
      MailApp.sendEmail({
        to: EMAIL_DESTINATARIO,
        subject: 'Nuova prenotazione 50° — ' + d.codice,
        body: d.resoconto || ('Codice: ' + d.codice),
        attachments: attachments
      });
    } catch (e) {}

    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  var raw = (e.postData && e.postData.contents) || '';
  if (e.parameter && e.parameter.payload) raw = e.parameter.payload;
  return handle(raw);
}

function doGet(e) {
  return handle(e.parameter && e.parameter.p);
}
```

## 3. Pubblica (rifallo ogni volta che modifichi lo script)
**Distribuisci → Gestisci distribuzioni → Modifica (icona matita) → Versione: Nuova versione → Distribuisci**
- Esegui come: **me** (l'account del gruppo)
- Chi ha accesso: **Tutti**  ← passaggio più importante: se resta "Solo io" o "Chiunque abbia un
  account Google", la pagina pubblica non può mai scrivere, e non te ne accorgi perché il browser
  non mostra l'errore (vedi sotto perché).

Copia l'URL `https://script.google.com/macros/s/.../exec` (deve finire con `/exec`, **non** `/dev`:
`/dev` funziona solo per te mentre sei loggato nell'editor, mai da una pagina pubblica).

> Se lo script viene modificato ma non ridistribuito con **Nuova versione**, l'URL continua a
> eseguire la versione vecchia: è la causa più comune di "funzionava e ora non più".

## 4. Verifica lo script DA SOLO, prima di collegarlo alla pagina
Questo passaggio salta il browser e l'HTML: apre l'URL direttamente, così scopri subito se il
problema è nella distribuzione (qui) o nella pagina.

1. Prendi l'URL `/exec` e incollalo nella barra degli indirizzi seguito da:
   `?p=%7B%22codice%22%3A%22TEST-1%22%2C%22nome%22%3A%22Prova%22%2C%22contatto%22%3A%22-%22%2C%22articoli%22%3A%5B%5D%7D`
   (è lo stesso payload di test, già codificato — copia e incolla l'URL intero).
2. Premi Invio. Guarda cosa restituisce la pagina:
   - **`ok`** in testo semplice → la distribuzione funziona, e dovresti trovare una riga `TEST-1`
     nel foglio (cancellala dopo). Questo test minimo non include il PDF (non c'è nella query di
     prova), ma l'email parte comunque, senza allegato: controlla che sia arrivata a
     `EMAIL_DESTINATARIO` come ulteriore conferma che tutta la catena funziona. Passa al punto 5.
   - **Una pagina di login Google / "Accedi"** → l'accesso della distribuzione non è su **Tutti**.
     Torna al punto 3.
   - **Errore Apps Script (schermata rossa, "Eccezione…")** → il codice ha un problema; controlla
     di aver incollato lo script del punto 2 senza modifiche e di aver salvato.
   - **404 / "pagina non trovata"** → l'URL è sbagliato: hai incollato il link del foglio
     (`docs.google.com/spreadsheets/...`) invece dell'URL dello script, oppure manca `/exec`.
3. Se hai visto `ok`, sei a posto: lo script funziona a prescindere dalla pagina.

## 5. Incolla l'URL nella pagina
Apri i **Tweaks** del listino → sezione *Ordini* → campo **endpointOrdini**: incolla l'URL `/exec`
verificato al punto 4.

## Se una conferma non arriva
La pagina lo dice solo se il browser è davvero offline: in tutti gli altri casi (deploy sbagliato,
accesso non su "Tutti", ecc.) la pagina non può accorgersi dell'errore — per il browser la
richiesta "è partita" anche se Google la rifiuta in silenzio. Ecco perché il punto 4 (verifica
diretta dell'URL, fuori dalla pagina) è il modo più veloce per capire se il problema è lo script o
qualcos'altro.

Cause tipiche, in ordine di frequenza:
- script non ridistribuito dopo una modifica (vedi punto 3);
- accesso della distribuzione non impostato su **Tutti**;
- URL incollato con `/dev` invece di `/exec`, o il link del foglio invece dello script.

## Se il foglio si aggiorna ma l'email no
L'invio email è isolato nel suo try/catch: un problema lì (quota giornaliera di `MailApp`
esaurita — circa 100 al giorno su un account Google normale, allegato troppo grande, indirizzo in
`EMAIL_DESTINATARIO` scritto male) non fa mai sparire la riga sul foglio, quindi il foglio resta la
fonte affidabile anche se qualche email non arriva. Controlla **Esecuzioni** nell'editor di Apps
Script: se l'esecuzione compare senza errori ma l'email non arriva, guarda anche nello spam del
destinatario.
