# Foglio ordini su Google Drive (accesso solo su invito)

La pagina non può creare da sola un file sul tuo Drive: serve un piccolo script che il gruppo
autorizza una volta. Dieci minuti, si fa dal browser.

## 1. Crea il foglio
1. Su Google Drive (account del gruppo) crea un nuovo **Foglio Google**, nome es. `Ordini 50° — Castel Maggiore 1`.
2. Prima riga come intestazione:
   `Data | Codice | Nome | Unità | Contatto | Articolo | Colore | Taglia | Quantità | Prezzo | Note`
3. Condivisione: lascia **Con limitazioni** e aggiungi solo le persone invitate (i capi).
   Chi prenota non vede il foglio: riceve solo la conferma con il codice.

## 2. Collega lo script
Nel foglio: **Estensioni → Apps Script**, incolla questo codice (sostituisce l'eventuale
versione precedente), salva.

Accetta i tre modi in cui la pagina può inviare i dati (JSON, modulo, immagine): serve perché
quando la pagina è pubblicata come artefatto o su GitHub Pages alcuni invii vengono bloccati
dal browser e si passa automaticamente al successivo. Il controllo sul codice evita righe doppie.

```js
function handle(raw) {
  if (!raw) return ContentService.createTextOutput('vuoto');
  var d = JSON.parse(raw);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  // già registrato? esci (evita doppioni dai tentativi multipli)
  var codici = sh.getRange(1, 2, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (var i = 0; i < codici.length; i++) {
    if (codici[i][0] === d.codice) return ContentService.createTextOutput('già presente');
  }

  (d.articoli || []).forEach(function (a) {
    sh.appendRow([new Date(), d.codice, d.nome, d.unita, d.contatto,
                  a.articolo, a.colore, a.taglia, a.quantita, a.prezzo, d.note]);
  });
  return ContentService.createTextOutput('ok');
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
- Chi ha accesso: **Tutti**  ← serve perché la pagina possa scrivere; il foglio resta privato,
  visibile solo alle persone invitate.

Copia l'URL `https://script.google.com/macros/s/.../exec` (deve finire con `/exec`).

> Se lo script viene modificato ma non ridistribuito, l'URL continua a eseguire la versione vecchia:
> è la causa più comune di "non funziona più".

## 4. Incolla l'URL nella pagina
Apri i **Tweaks** del listino → sezione *Ordini* → campo **endpointOrdini**: incolla l'URL.

## Se una conferma non arriva
La pagina lo dice: mostra comunque il codice con l'avviso che l'invio automatico non è passato.
In quel caso la prenotazione va comunicata ai capi con il codice. Cause tipiche:
- script non ridistribuito dopo una modifica (vedi punto 3);
- accesso della distribuzione non impostato su **Tutti**;
- URL incollato con `/dev` invece di `/exec`, o il link del foglio invece dello script.
