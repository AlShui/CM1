# Listino 50° — Gruppo Scout Castel Maggiore 1

Listino prezzi e prenotazioni per i capi del cinquantesimo anniversario.

**Pagina pubblica:** https://alshui.github.io/CM1/listino-50.html

## Come funziona

- `listino-50.html` è la pagina: catalogo, carrello, prenotazione. Usa il runtime in `support.js`,
  `ds-base.js`, `doc-page.js`, `image-slot.js` e la cartella `_ds/` — non toccarli separatamente,
  fanno parte dello stesso export.
- Quando qualcuno conferma una prenotazione, la pagina prova a scrivere una riga su un Google
  Sheet privato tramite uno script Google (Apps Script). Setup e risoluzione problemi:
  [`drive-foglio-ordini.md`](./drive-foglio-ordini.md).
- `.nojekyll` è necessario: senza questo file, GitHub Pages ignora per default qualsiasi cartella
  che inizia con `_` (qui `_ds/`), e la pagina si carica senza stili né script.

## Pubblicare su GitHub Pages

1. Impostazioni del repository → **Pages** → Source: **Deploy from a branch** → branch `main`,
   cartella `/ (root)` → Save.
2. Dopo un minuto la pagina è live su `https://<utente>.github.io/<repo>/listino-50.html`.

## Dominio

GitHub Pages dà già un indirizzo gratuito (`alshui.github.io/CM1/...`): funziona, è pubblico, non
richiede nulla da pagare. Se in futuro volete un dominio "vostro" (tipo `listino50.it` o simile),
non esistono servizi di dominio realmente gratuiti e affidabili a lungo termine: le opzioni serie
sono un dominio a pagamento (pochi euro l'anno, es. Cloudflare Registrar o un registrar `.it`)
puntato con un file `CNAME` a GitHub Pages, oppure un sottodominio gratuito di terzi tipo
`is-a.dev` (via richiesta su GitHub). Per un listino interno al gruppo, l'indirizzo
`github.io` è probabilmente sufficiente — è già gratuito, pubblico e stabile.
