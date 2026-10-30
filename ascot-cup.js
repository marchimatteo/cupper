/*
 * ascot-cup.js
 * Nessun build step: includere con <script src="ascot-cup.js"></script> a fine <body>
 * (dopo Bootstrap e dopo gli elementi della pagina).
 *
 * Elementi attesi nella pagina:
 *   #inputCaricaAscot  (input type="file", oppure un button: in tal caso viene creato un input file nascosto)
 *   #alertInCorso, #alertSuccess, #alertError  (tutti con attributo hidden)
 *   #buttonScarica
 */
(function () {
  'use strict';

  const SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
  const FOGLIO_DATI = 'EAPDSROC';
  const FOGLIO_TESTATA = 'Testata';

  /* ------------------------------------------------------------------ */
  /* Codice del worker (viene eseguito in un thread separato)           */
  /* ------------------------------------------------------------------ */
  function workerMain() {
    const COL_S = 18; // indice 0-based della colonna S
    const COL_T = 19; // colonna T (nuova)
    // CUP: 1 lettera + 2 cifre + 1 lettera + 11 cifre = 15 caratteri
    const CUP_REGEX = /(?<![A-Z0-9])[A-Z]\d{2}[A-Z]\d{11}(?![A-Z0-9])/gi;

    self.onmessage = function (e) {
      try {
        const { buffer, url, foglioDati, foglioTestata } = e.data;
        importScripts(url);

        const wb = XLSX.read(buffer, {
          type: 'array',
          cellStyles: true,
          cellNF: true,
          cellDates: true
        });

        if (!wb.SheetNames.includes(foglioTestata)) {
          throw new Error('Foglio "' + foglioTestata + '" non trovato nel file.');
        }
        const ws = wb.Sheets[foglioDati];
        if (!ws) throw new Error('Foglio "' + foglioDati + '" non trovato nel file.');
        if (!ws['!ref']) throw new Error('Il foglio "' + foglioDati + '" è vuoto.');

        const range = XLSX.utils.decode_range(ws['!ref']);
        const enc = XLSX.utils.encode_cell;

        /* 1) Sposta di una colonna a destra tutto da T in poi (da destra verso sinistra) */
        for (let r = range.s.r; r <= range.e.r; r++) {
          for (let c = range.e.c; c >= COL_T; c--) {
            const from = enc({ r, c });
            const to = enc({ r, c: c + 1 });
            if (ws[from]) {
              ws[to] = ws[from];
              delete ws[from];
            } else {
              delete ws[to];
            }
          }
        }

        /* 2) Aggiorna range, larghezze colonne, merge e autofilter */
        const nuovoUltimaCol = Math.max(range.e.c >= COL_T ? range.e.c + 1 : range.e.c, COL_T);
        ws['!ref'] = XLSX.utils.encode_range({ s: range.s, e: { r: range.e.r, c: nuovoUltimaCol } });

        if (ws['!cols']) {
          const cols = ws['!cols'];
          while (cols.length < COL_T) cols.push(undefined);
          cols.splice(COL_T, 0, cols[COL_S] ? Object.assign({}, cols[COL_S]) : undefined);
        }

        if (ws['!merges']) {
          ws['!merges'] = ws['!merges'].map(function (m) {
            const n = { s: { r: m.s.r, c: m.s.c }, e: { r: m.e.r, c: m.e.c } };
            if (n.s.c >= COL_T) n.s.c++;
            if (n.e.c >= COL_T) n.e.c++;
            return n;
          });
        }

        if (ws['!autofilter'] && ws['!autofilter'].ref) {
          const af = XLSX.utils.decode_range(ws['!autofilter'].ref);
          if (af.s.c >= COL_T) af.s.c++;
          if (af.e.c >= COL_T || af.e.c === COL_S) af.e.c++;
          ws['!autofilter'].ref = XLSX.utils.encode_range(af);
        }

        /* 3) Intestazione + estrazione CUP */
        const intestS = ws[enc({ r: range.s.r, c: COL_S })];
        const intest = { t: 's', v: 'CUP' };
        if (intestS && intestS.s) intest.s = intestS.s;
        ws[enc({ r: range.s.r, c: COL_T })] = intest;

        let totaleRigheConCup = 0;
        for (let r = range.s.r + 1; r <= range.e.r; r++) {
          const cellS = ws[enc({ r, c: COL_S })];
          let testo = '';
          if (cellS && cellS.v != null) testo = cellS.w != null ? String(cellS.w) : String(cellS.v);

          const trovati = testo.match(CUP_REGEX) || [];
          const unici = Array.from(new Set(trovati.map(function (x) { return x.toUpperCase(); })));

          const nuova = { t: 's', v: unici.join(';') };
          if (cellS && cellS.s) nuova.s = cellS.s; // stesso stile della colonna S
          ws[enc({ r, c: COL_T })] = nuova;
          if (unici.length) totaleRigheConCup++;
        }

        /* 4) Scrittura in formato xls (BIFF8) */
        const out = XLSX.write(wb, { bookType: 'biff8', type: 'array', cellStyles: true });
        self.postMessage({ ok: true, buffer: out, righeConCup: totaleRigheConCup }, [out.buffer || out]);
      } catch (err) {
        self.postMessage({ ok: false, error: (err && err.message) ? err.message : String(err) });
      }
    };
  }

  /* ------------------------------------------------------------------ */
  /* Interfaccia                                                        */
  /* ------------------------------------------------------------------ */
  function init() {
    const elInput = document.getElementById('inputCaricaAscot');
    const alertInCorso = document.getElementById('alertInCorso');
    const alertSuccess = document.getElementById('alertSuccess');
    const alertError = document.getElementById('alertError');
    const buttonScarica = document.getElementById('buttonScarica');

    if (!elInput || !alertInCorso || !alertSuccess || !alertError || !buttonScarica) {
      console.error('ascot-cup.js: elementi della pagina mancanti.');
      return;
    }

    let risultato = null; // { blob, nome }

    /**
     * @param el      L'elemento html
     * @param display Quando l'elemento non può essere nascosto con hidden perchè usa d-flex
     *                e viene quindi nascosto con d-none
     */
    function mostra(el, display) {
      el.hidden = false;
      el.classList.remove('d-none');
      if (display) el.classList.add(display);
    }
    function nascondi(el) {
      el.hidden = true;
    }
    function resetAlert() {
      nascondi(alertInCorso);
      nascondi(alertError);
      alertSuccess.hidden = true;
      alertSuccess.classList.remove('d-flex');
      alertSuccess.classList.add('d-none');
    }
    function mostraErrore(msg) {
      nascondi(alertInCorso);
      alertError.textContent = msg;
      mostra(alertError);
    }

    function elabora(file) {
      resetAlert();
      risultato = null;
      mostra(alertInCorso);

      const reader = new FileReader();
      reader.onerror = function () { mostraErrore('Impossibile leggere il file selezionato.'); };
      reader.onload = function () {
        const buffer = reader.result;
        const codice = '(' + workerMain.toString() + ')();';
        const workerUrl = URL.createObjectURL(new Blob([codice], { type: 'text/javascript' }));
        const worker = new Worker(workerUrl);

        function chiudi() {
          worker.terminate();
          URL.revokeObjectURL(workerUrl);
        }

        worker.onmessage = function (e) {
          chiudi();
          if (!e.data.ok) { mostraErrore(e.data.error); return; }
          const nomeBase = file.name.replace(/\.[^.]+$/, '');
          risultato = {
            blob: new Blob([e.data.buffer], { type: 'application/vnd.ms-excel' }),
            nome: nomeBase + '_CUP.xls'
          };
          nascondi(alertInCorso);
          mostra(alertSuccess, 'd-flex');
        };
        worker.onerror = function (e) {
          chiudi();
          mostraErrore(e.message || 'Errore durante l\'elaborazione.');
        };

        worker.postMessage(
          { buffer: buffer, url: SHEETJS_URL, foglioDati: FOGLIO_DATI, foglioTestata: FOGLIO_TESTATA },
          [buffer]
        );
      };
      reader.readAsArrayBuffer(file);
    }

    elInput.accept = '.xls,application/vnd.ms-excel';
    elInput.addEventListener('change', function () {
      const file = elInput.files && elInput.files[0];
      if (file) elabora(file);
    });

    buttonScarica.addEventListener('click', function () {
      if (!risultato) return;
      const url = URL.createObjectURL(risultato.blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = risultato.nome;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
