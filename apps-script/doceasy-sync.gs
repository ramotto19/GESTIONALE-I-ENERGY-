/**
 * Gestionale I-Energy — sincronizzazione automatica fatture da DocEasy.
 *
 * Gira dentro Apps Script (lato server, codice NON pubblico): un attivatore a tempo
 * lo esegue ogni giorno e importa nel gestionale le fatture passive e attive nuove,
 * SOLO I DATI (niente PDF/Drive), evitando i doppioni.
 *
 * ──────────────────────────────────────────────────────────────────────────
 *  SICUREZZA DELLE CHIAVI — da fare UNA volta, a mano, nell'editor di Apps Script:
 *  Impostazioni progetto  ▸  Proprietà script  ▸  Aggiungi proprietà
 *  (le chiavi restano sul server di Google, fuori dal codice e fuori da GitHub)
 *
 *    DOCEASY_BASE          https://webapide.aliasgrouplab.it   (URL base dell'API DocEasy)
 *    DOCEASY_API_KEY       API Key   (DocEasy ▸ Area Clienti ▸ Gestisci API Key)
 *    DOCEASY_API_SECRET    API Secret (stessa pagina)
 *    DOCEASY_PIVA          la partita IVA della tua azienda (header richiesto da DocEasy)
 *    SUPABASE_SERVICE_KEY  la chiave segreta del progetto Supabase (sb_secret_… oppure
 *                          la legacy service_role; Supabase ▸ Project Settings ▸ API Keys)
 *
 *  Nota sui dati importati: l'endpoint "elenco" di DocEasy restituisce numero, data,
 *  controparte, partita IVA e TOTALE del documento. Non espone imponibile, IVA e
 *  scadenza (quelli stanno nel dettaglio della singola fattura): per ora restano vuoti
 *  e si possono compilare a mano nella scheda, oppure si aggiunge in seguito una
 *  chiamata al dettaglio /api/documento{attivo|passivo}/{id}/json.
 *
 *  NON mettere queste chiavi nel codice, nel repository o nel frontend (index.html):
 *  index.html è pubblico. La service_role di Supabase in particolare scavalca le
 *  regole RLS: deve stare SOLO qui, nelle Proprietà script.
 * ──────────────────────────────────────────────────────────────────────────
 */

var SB_URL = 'https://uagjyykdrwvzhcemlcwk.supabase.co';
var LOOKBACK_DAYS = 7; // guarda indietro N giorni: se un giorno salta, si recupera (i doppioni vengono scartati)

/* ===== punto d'ingresso: lo chiama l'attivatore a tempo ===== */
function syncDocEasy() {
  var da = isoDaysAgo_(LOOKBACK_DAYS);
  var rep = { nuove: 0, gia: 0, errori: 0 };
  ['P', 'A'].forEach(function (dir) {
    try {
      var grezze = doceasyFetch_(dir, da);          // ← elenco fatture dall'API DocEasy
      grezze.forEach(function (raw) {
        try {
          var r = mapDocEasy_(raw, dir);             // → record nel formato del gestionale
          if (!r.numero || !r.data) { return; }
          var id = fattId_(r);                        // stessa chiave anti-doppione del frontend
          if (sbExists_('fatture/' + id)) { rep.gia++; return; }  // già presente → non tocco (resta il dirFix manuale)
          r.file = null;                              // solo i dati, nessun PDF
          sbUpsert_('fatture/' + id, r);
          rep.nuove++;
        } catch (e) { rep.errori++; }
      });
    } catch (e) {
      Logger.log('DocEasy ' + dir + ' errore: ' + e);
      rep.errori++;
    }
  });
  Logger.log('Sync DocEasy: ' + rep.nuove + ' nuove, ' + rep.gia + ' già presenti, ' + rep.errori + ' errori');
  return rep;
}

/* ===== DocEasy: elenco fatture passive/attive =====
 * Endpoint ufficiali (webapide.aliasgrouplab.it):
 *   GET /api/documentoattivo/elenco   → fatture emesse (attive)
 *   GET /api/documentopassivo/elenco  → fatture ricevute (passive)
 * L'elenco torna gli ultimi documenti; qui tengo solo quelli dalla finestra
 * (`dalGiorno`) in poi, i doppioni li scarta comunque fattId_.                */
function doceasyFetch_(dir, dalGiorno) {
  var base = getProp_('DOCEASY_BASE').replace(/\/+$/, '');
  var seg = (dir === 'P') ? 'documentopassivo' : 'documentoattivo';
  var res = UrlFetchApp.fetch(base + '/api/' + seg + '/elenco', {
    method: 'get', muteHttpExceptions: true, headers: doceasyHeaders_()
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('DocEasy HTTP ' + res.getResponseCode() + ' su ' + seg + ': ' + res.getContentText().slice(0, 300));
  }
  var body = JSON.parse(res.getContentText() || '[]');
  var lista = Array.isArray(body) ? body : (body.items || body.data || body.elenco || []);
  return lista.filter(function (d) { return isoDate_(d.data || d.dataDocumento) >= dalGiorno; });
}
/* autenticazione DocEasy: tutto negli header (APIKey + APISecret obbligatori,
 * PartitaIva per indicare l'azienda) — niente credenziali in URL */
function doceasyHeaders_() {
  return {
    'APIKey': getProp_('DOCEASY_API_KEY'),
    'APISecret': getProp_('DOCEASY_API_SECRET'),
    'PartitaIva': getProp_('DOCEASY_PIVA'),
    'Accept': 'application/json'
  };
}

/* ===== mappa una fattura DocEasy (elenco) nel record del gestionale =====
 * Campi dell'elenco DocEasy: numero, data, denominazione (o nome+cognome),
 * partitaIva/codiceFiscale, importo (totale), tipoDocumento.                       */
function mapDocEasy_(raw, dir) {
  var controparte = String(raw.denominazione || [raw.nome, raw.cognome].filter(Boolean).join(' ') || '').trim();
  var td = String(raw.tipoDocumento || raw.tipoFattura || '').toUpperCase();
  var segno = (td === 'TD04') ? -1 : 1;                   // nota di credito → importo negativo (come l'import XML)
  return {
    dir: dir,
    numero: String(raw.numero || '').trim(),
    data: isoDate_(raw.data || raw.dataDocumento),        // → 'YYYY-MM-DD' (deve combaciare con l'import XML)
    controparte: controparte,
    piva: String(raw.partitaIva || raw.codiceFiscale || '').trim(),
    imponibile: 0,                                         // non nell'elenco (vedi nota in testa al file)
    iva: 0,                                                // non nell'elenco
    tot: round2_(segno * num_(raw.importo)),
    scadenza: '',                                          // non nell'elenco: sta nel dettaglio della fattura
    descrizione: '',
    tipoDoc: td,
    pagata: false,
    importatoIl: today_()
  };
}

/* ===== chiave anti-doppione: IDENTICA al frontend (index.html) ===== */
/* fattId = 'f' + hashC(k) + hashC(reverse(k)),  k = dir|piva|numero|data          */
function fattId_(r) {
  var k = [r.dir, r.piva, r.numero, r.data].join('|');
  return 'f' + hashC_(k) + hashC_(k.split('').reverse().join(''));
}
function hashC_(t) {
  var h = 5381;
  String(t).split('').forEach(function (ch) { h = ((h * 33) ^ ch.charCodeAt(0)) >>> 0; });
  return h.toString(36);
}

/* ===== Supabase REST con la chiave service_role (scavalca la RLS) ===== */
function sbHead_() {
  var k = getProp_('SUPABASE_SERVICE_KEY');
  return { apikey: k, Authorization: 'Bearer ' + k, 'Content-Type': 'application/json' };
}
function sbExists_(path) {
  var url = SB_URL + '/rest/v1/docs?select=path&path=eq.' + encodeURIComponent(path);
  var res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: sbHead_() });
  if (res.getResponseCode() !== 200) throw new Error('Supabase read ' + res.getResponseCode());
  return JSON.parse(res.getContentText() || '[]').length > 0;
}
function sbUpsert_(path, data) {
  var url = SB_URL + '/rest/v1/docs?on_conflict=path';
  var payload = JSON.stringify({ path: path, data: data, updated_at: new Date().toISOString() });
  var res = UrlFetchApp.fetch(url, {
    method: 'post', muteHttpExceptions: true,
    headers: Object.assign(sbHead_(), { Prefer: 'resolution=merge-duplicates,return=minimal' }),
    payload: payload
  });
  if (res.getResponseCode() >= 300) throw new Error('Supabase write ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
}

/* ===== utilità ===== */
function getProp_(name) {
  var v = PropertiesService.getScriptProperties().getProperty(name);
  if (!v) throw new Error('Manca la proprietà script: ' + name + ' (impostala in Impostazioni progetto ▸ Proprietà script)');
  return v;
}
function num_(x) { var n = parseFloat(String(x == null ? '' : x).replace(',', '.')); return isNaN(n) ? 0 : n; }
function round2_(n) { return Math.round(n * 100) / 100; }
function today_() { return Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM-dd'); }
function isoDaysAgo_(d) { return Utilities.formatDate(new Date(Date.now() - d * 86400000), 'Europe/Rome', 'yyyy-MM-dd'); }
function isoDate_(x) {
  if (!x) return '';
  var s = String(x);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);        // già ISO
  var m = s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})/);          // dd/mm/yyyy
  if (m) return m[3] + '-' + m[2] + '-' + m[1];
  var dt = new Date(s);
  return isNaN(dt.getTime()) ? s.slice(0, 10) : Utilities.formatDate(dt, 'Europe/Rome', 'yyyy-MM-dd');
}

/* ===== prova manuale: eseguila UNA volta dall'editor e guarda il Log (Esecuzioni) =====
 * Non scrive niente nel gestionale: mostra solo cosa torna DocEasy e come viene
 * mappato, così verifichi che chiavi, P.IVA ed endpoint siano a posto prima del trigger. */
function testDocEasy() {
  ['P', 'A'].forEach(function (dir) {
    try {
      var l = doceasyFetch_(dir, isoDaysAgo_(60));
      Logger.log(dir + ': ' + l.length + ' documenti negli ultimi 60 giorni');
      if (l[0]) {
        Logger.log(dir + ' · grezzo DocEasy: ' + JSON.stringify(l[0]));
        Logger.log(dir + ' · mappato gestionale: ' + JSON.stringify(mapDocEasy_(l[0], dir)));
      }
    } catch (e) { Logger.log(dir + ' · ERRORE: ' + e); }
  });
}

/* ===== crea l'attivatore giornaliero (eseguila UNA volta a mano dall'editor) ===== */
function setupDocEasyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncDocEasy') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncDocEasy').timeBased().everyDays(1).atHour(6).create();
  Logger.log('Attivatore creato: syncDocEasy ogni giorno verso le 6:00');
}
