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
 *    DOCEASY_BASE          es. https://api.doceasy.it           (base URL dell'API)
 *    DOCEASY_API_KEY       la chiave che generi nella sezione "creazione API" di DocEasy
 *    SUPABASE_SERVICE_KEY  la chiave service_role del progetto Supabase
 *                          (Supabase ▸ Project Settings ▸ API ▸ service_role secret)
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

/* ===== DocEasy: l'UNICO pezzo da adattare ai TUOI endpoint =====
 * Deve restituire un array di fatture grezze per la direzione richiesta
 * ('P' = passive/ricevute, 'A' = attive/emesse) a partire dalla data `dalGiorno`.
 * Completa l'URL, i parametri e lo schema di autenticazione con quanto trovi
 * nella documentazione della tua sezione API di DocEasy.                      */
function doceasyFetch_(dir, dalGiorno) {
  var base = getProp_('DOCEASY_BASE');
  var key = getProp_('DOCEASY_API_KEY');
  // ⚠️ ADATTA questo blocco: nome endpoint, nome dei parametri e header di auth
  //    dipendono dall'API DocEasy. Esempio generico:
  var tipo = (dir === 'P') ? 'ricevute' : 'emesse';
  var url = base.replace(/\/+$/, '') + '/fatture?tipo=' + tipo + '&dal=' + dalGiorno;
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      // ⚠️ ADATTA: alcune API vogliono 'Authorization: Bearer <key>',
      //    altre 'X-Api-Key: <key>' o un parametro in query. Vedi la doc DocEasy.
      'Authorization': 'Bearer ' + key,
      'Accept': 'application/json'
    }
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('DocEasy HTTP ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
  }
  var body = JSON.parse(res.getContentText() || '{}');
  return body.items || body.fatture || body.data || (Array.isArray(body) ? body : []);
}

/* ===== mappa una fattura DocEasy nel record del gestionale =====
 * ⚠️ ADATTA i nomi dei campi di sinistra (raw.*) a quelli reali dell'API DocEasy.
 * I campi di destra sono quelli che la pagina Fatture del gestionale già usa.      */
function mapDocEasy_(raw, dir) {
  var imponibile = round2_(num_(raw.imponibile));
  var iva = round2_(num_(raw.iva));
  var tot = round2_(num_(raw.totale) || (imponibile + iva));
  return {
    dir: dir,
    numero: String(raw.numero || raw.numeroDocumento || '').trim(),
    data: isoDate_(raw.data || raw.dataDocumento),        // → 'YYYY-MM-DD' (deve combaciare con l'import XML)
    controparte: String(raw.controparte || raw.denominazione || raw.cedente || raw.cessionario || '').trim(),
    piva: String(raw.piva || raw.partitaIva || '').trim(),
    imponibile: imponibile,
    iva: iva,
    tot: tot,
    scadenza: raw.scadenza ? isoDate_(raw.scadenza) : '',
    descrizione: String(raw.causale || raw.descrizione || '').slice(0, 400),
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

/* ===== crea l'attivatore giornaliero (eseguila UNA volta a mano dall'editor) ===== */
function setupDocEasyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncDocEasy') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncDocEasy').timeBased().everyDays(1).atHour(6).create();
  Logger.log('Attivatore creato: syncDocEasy ogni giorno verso le 6:00');
}
