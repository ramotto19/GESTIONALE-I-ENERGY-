/**
 * Gestionale Impianti — backend Google Apps Script.
 *
 * Fa da intermediario tra l'app (index.html) e i dati dell'azienda:
 * - verifica l'identità di chi chiama tramite il suo ID token Google;
 * - applica le stesse regole di accesso dell'artifact originale
 *   (l'amministratore vede e modifica tutto; un dipendente vede e
 *   modifica solo le proprie ore e la propria scheda accessi, e legge
 *   l'elenco commesse);
 * - salva i dati in un Google Sheet (una scheda per collezione) e gli
 *   allegati in una cartella Google Drive.
 *
 * Setup: vedi il README nella cartella principale del repository.
 */

var COLS = ['clienti', 'commesse', 'dipendenti', 'squadre', 'assegnazioni', 'assegnazioniG', 'tipiLavoro', 'categorie', 'docAzienda', 'assicurazioni', 'mezzi', 'documenti', 'verbali', 'scadenze', 'riservato', 'accessi', 'impostazioni', 'commesseElenco', 'registri'];

/* Incolla qui il Client ID OAuth creato su Google Cloud Console
   (Credenziali → ID client OAuth → tipo Applicazione web).
   Deve essere lo STESSO valore messo in CONFIG.CLIENT_ID dentro index.html. */
/* (non più usato per l'accesso: l'identità arriva da Supabase) */
var CLIENT_ID = '993000663967-5fv7cntdsjrcl0urmv8j2uscmju29ieu.apps.googleusercontent.com';

/* ---------- setup iniziale (da eseguire una sola volta dall'editor di Apps Script) ---------- */
function setup() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SHEET_ID')) {
    var ss = SpreadsheetApp.create('Gestionale Impianti — Dati');
    props.setProperty('SHEET_ID', ss.getId());
    Logger.log('Creato Google Sheet: ' + ss.getUrl());
  }
  if (!props.getProperty('FOLDER_ID')) {
    var folder = DriveApp.createFolder('Gestionale Impianti — Allegati');
    props.setProperty('FOLDER_ID', folder.getId());
    Logger.log('Creata cartella Drive: ' + folder.getUrl());
  }
  COLS.forEach(function (col) { sheetFor(col); });
  if (CLIENT_ID.indexOf('INSERISCI_QUI') === 0) {
    Logger.log('ATTENZIONE: sostituisci il valore di CLIENT_ID in cima al file con il tuo Client ID OAuth, poi salva.');
  } else {
    Logger.log('CLIENT_ID configurato. Ora pubblica il Web App (Esegui il deployment → Nuovo deployment).');
  }
}

/* ---------- entry point del Web App ---------- */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, info: 'Gestionale Impianti — backend attivo. Usa POST per le richieste dati.' })).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  var out;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = handle(body);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle(body) {
  var caller = verifyToken(body.idToken);
  var role = resolveRole(caller);
  if (!role) return { ok: true, unauthorized: true, email: caller.email, name: caller.name };
  switch (body.action) {
    /* driveFolderId: l'id vero (fisso) della cartella allegati, letto da PropertiesService
       (lo stesso usato da apiUpload). Lo legge anche il browser, cosi' quando carica un file
       direttamente su Drive usa SEMPRE questo id invece di cercare la cartella per nome — la
       ricerca per nome, se due caricamenti partivano quasi insieme (due schede aperte, due
       salvataggi ravvicinati), poteva non trovarla ancora in nessuno dei due e farne creare una
       a testa: da qui le cartelle "Allegati" doppie viste più volte. */
    case 'whoami': return { ok: true, email: caller.email, name: caller.name, role: role, driveFolderId: PropertiesService.getScriptProperties().getProperty('FOLDER_ID') || '' };
    case 'listAll': return apiListAll(caller, role);
    case 'get': return apiGet(body, caller, role);
    case 'set': return apiSet(body, caller, role);
    case 'update': return apiUpdate(body, caller, role);
    case 'delete': return apiDelete(body, caller, role);
    case 'upload': return apiUpload(body, caller, role);
    case 'deleteFile': return apiDeleteFile(body, caller, role);
    case 'archiveFile': return apiArchiveFile(body, caller, role);
    case 'resolveFolder': return apiResolveFolder(body, caller, role);
    default: return { ok: false, error: 'Azione sconosciuta: ' + body.action };
  }
}

/* ---------- identità (accesso Supabase) ---------- */
/* L'app ora fa l'accesso con Google tramite Supabase: qui si verifica il token di Supabase e si
   legge il ruolo dalla tabella "docs" con lo stesso token (così valgono le stesse regole del database).
   Il risultato per lo stesso token resta in cache pochi minuti. */
var SUPABASE_URL = 'https://uagjyykdrwvzhcemlcwk.supabase.co';
var SUPABASE_KEY = 'sb_publishable_soPbd6TCzDMDyyOvwPpXKQ_zYaFHlNh';
function verifyToken(idToken) {
  if (!idToken) throw new Error('Token mancante: accedi di nuovo.');
  var cache = CacheService.getScriptCache();
  var key = 'sbtok_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
  var cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  var res = UrlFetchApp.fetch(SUPABASE_URL + '/auth/v1/user', { muteHttpExceptions: true, headers: { apikey: SUPABASE_KEY, Authorization: 'Bearer ' + idToken } });
  if (res.getResponseCode() !== 200) throw new Error('Sessione scaduta: accedi di nuovo.');
  var u = JSON.parse(res.getContentText());
  var email = String(u.email || '').toLowerCase();
  var provider = u.app_metadata && u.app_metadata.provider;
  if (!email || provider !== 'google') throw new Error('Serve l\u2019accesso con un account Google.');
  var m = u.user_metadata || {};
  var out = { email: email, name: m.full_name || m.name || email, token: idToken };
  try { cache.put(key, JSON.stringify(out), 300); } catch (e) { /* cache piena: si continua senza */ }
  return out;
}
function sbHeaders(token) { return { apikey: SUPABASE_KEY, Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }; }
function resolveRole(caller) {
  var token = caller.token;
  var a = UrlFetchApp.fetch(SUPABASE_URL + '/rest/v1/rpc/is_admin', { method: 'post', muteHttpExceptions: true, headers: sbHeaders(token), payload: '{}' });
  if (a.getResponseCode() === 200 && JSON.parse(a.getContentText()) === true) return 'admin';
  var b = UrlFetchApp.fetch(SUPABASE_URL + '/rest/v1/docs?select=path&path=eq.' + encodeURIComponent('accessi/' + caller.email), { muteHttpExceptions: true, headers: sbHeaders(token) });
  if (b.getResponseCode() === 200 && JSON.parse(b.getContentText()).length) return 'dipendente';
  return null;
}

/* ---------- regole di accesso (identiche a quelle dell'artifact originale) ---------- */
function selfIdMatches(col, id, email) {
  if (col === 'accessi') return id === email;
  if (col === 'ore') return !!id && id.split('::')[0] === email;
  return false;
}
function canRead(col, id, role, email) {
  if (role === 'admin') return true;
  if (col === 'commesseElenco') return true;
  return selfIdMatches(col, id, email);
}
function canWrite(col, id, role, email) {
  if (role === 'admin') return true;
  if (col === 'ore') return selfIdMatches(col, id, email);
  return false;
}

/* ---------- operazioni dati ---------- */
function apiListAll(caller, role) {
  if (role === 'admin') return { ok: true, cols: readManyBatch(COLS) };
  var out = {};
  var own = readDoc('accessi', caller.email);
  out.accessi = own ? toIdMap(caller.email, own) : {};
  out.commesseElenco = readAll('commesseElenco');
  return { ok: true, cols: out };
}
function toIdMap(id, doc) { var m = {}; m[id] = doc; return m; }

function apiGet(body, caller, role) {
  var col = body.col, id = body.id;
  if (!canRead(col, id, role, caller.email)) return { ok: true, found: false };
  var doc = readDoc(col, id);
  return { ok: true, found: !!doc, doc: doc };
}
function apiSet(body, caller, role) {
  var col = body.col, id = body.id;
  if (!canWrite(col, id, role, caller.email)) return { ok: false, error: 'Non hai i permessi per modificare questi dati.' };
  writeDoc(col, id, body.data || {});
  return { ok: true };
}
function apiUpdate(body, caller, role) {
  var col = body.col, id = body.id;
  if (!canWrite(col, id, role, caller.email)) return { ok: false, error: 'Non hai i permessi per modificare questi dati.' };
  var merged = updateDoc(col, id, body.data || {});
  return { ok: true, doc: merged };
}
function apiDelete(body, caller, role) {
  var col = body.col, id = body.id;
  if (!canWrite(col, id, role, caller.email)) return { ok: false, error: 'Non hai i permessi per modificare questi dati.' };
  deleteDoc(col, id);
  return { ok: true };
}
/* risolve (creando se serve) la cartella annidata dentro quella radice che corrisponde a
   pagina/commessa/sezione, cosi' i documenti caricati da Apps Script finiscono nella stessa
   struttura di cartelle di quelli caricati direttamente dal browser, invece che tutti insieme
   in un'unica cartella piatta.
   Il lucchetto (withLock, lo stesso usato per i salvataggi dei dati) protegge SOLO il pezzo
   di percorso che deve ancora essere creato: i livelli già esistenti si leggono prima, senza
   lucchetto (leggere "esiste già?" è sempre sicuro, nessuna corsa possibile lì), così la
   stragrande maggioranza dei caricamenti — che vanno in cartelle già create da tempo — non
   aspetta mai il lucchetto e non rallenta gli altri salvataggi dell'app. Solo il PRIMO
   caricamento che crea per davvero una cartella nuova entra nel lucchetto, e solo per quel
   pezzo mancante. Senza nessuna protezione, due caricamenti quasi contemporanei nella stessa
   sottocartella NON ANCORA creata (due schede, due persone, o il tentativo diretto dal
   browser e quello di riserva via Apps Script in corsa fra loro) potevano trovarla entrambi
   assente e crearla due volte — la stessa corsa già risolta per la cartella radice (vedi
   whoami/driveFolderId in handle()), ma qui capitava per ogni sottocartella (bug reale,
   segnalato il 2026-10-19: cartelle doppie dentro gli allegati). */
function ensureFolderPath(rootFolder, segments) {
  var clean = (segments || []).slice(0, 6);
  var folder = rootFolder, mancaDa = -1;
  for (var i = 0; i < clean.length; i++) {
    var nome = String(clean[i] || '').trim().slice(0, 100);
    if (!nome) continue;
    var it = folder.getFoldersByName(nome);
    if (it.hasNext()) { folder = it.next(); } else { mancaDa = i; break; }
  }
  if (mancaDa === -1) return folder;
  return withLock(function () {
    var f = folder;
    for (var j = mancaDa; j < clean.length; j++) {
      var nm = String(clean[j] || '').trim().slice(0, 100);
      if (!nm) continue;
      var it2 = f.getFoldersByName(nm);
      f = it2.hasNext() ? it2.next() : f.createFolder(nm);
    }
    return f;
  });
}
/* permette al browser di risolvere (creandola se serve) una sottocartella SENZA passare dal
   caricamento: così anche il caricamento diretto su Drive (più veloce, usato quando c'è già
   un token Drive valido) ottiene l'id della cartella da qui invece di cercarla e crearla da
   sé — che era proprio la ricerca-poi-crea non protetta da nessun lucchetto, e quindi la
   causa delle cartelle doppie quando due caricamenti partivano quasi insieme.
   Come apiUpload e apiArchiveFile: solo l'amministratore può crearle. Nell'app di oggi solo
   l'amministratore carica allegati (i dipendenti vedono "I miei documenti" in sola lettura),
   ma il controllo va comunque messo qui — non basta che l'interfaccia non mostri il
   pulsante — perché chiunque abbia un token valido può altrimenti chiamare direttamente
   questa azione. */
function apiResolveFolder(body, caller, role) {
  if (role !== 'admin') return { ok: false, error: 'Non hai i permessi per caricare allegati.' };
  var folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!folderId) return { ok: false, error: 'Cartella allegati non configurata: esegui setup().' };
  var folder = ensureFolderPath(DriveApp.getFolderById(folderId), body.folderPath);
  return { ok: true, folderId: folder.getId() };
}
function apiUpload(body, caller, role) {
  if (role !== 'admin') return { ok: false, error: 'Non hai i permessi per caricare allegati.' };
  var folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!folderId) return { ok: false, error: 'Cartella allegati non configurata: esegui setup().' };
  var folder = ensureFolderPath(DriveApp.getFolderById(folderId), body.folderPath);
  var bytes = Utilities.base64Decode(body.base64 || '');
  var blob = Utilities.newBlob(bytes, body.mimeType || 'application/octet-stream', body.filename || 'file');
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, id: file.getId(), url: file.getUrl(), name: body.filename || file.getName() };
}
/* cestina (non elimina per sempre) la copia su Drive di un allegato sostituito o rimosso da una
   scheda: pulizia di riserva, usata solo quando il caricamento diretto dal browser non è
   disponibile. Non blocca mai nulla se fallisce (file già cestinato, permessi…): è una pulizia,
   non un'operazione essenziale. */
function apiDeleteFile(body, caller, role) {
  try {
    var file = DriveApp.getFileById(body.fileId);
    file.setTrashed(true);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: 'File non trovato o non eliminabile.' };
  }
}
/* sposta (non elimina) la copia su Drive di un allegato sostituito dentro una sottocartella
   "Scaduti" della stessa cartella in cui si trovava: pulizia di riserva, usata solo quando il
   caricamento diretto dal browser non è disponibile. Non blocca mai nulla se fallisce (file già
   spostato, permessi…): è un riordino, non un'operazione essenziale. */
function apiArchiveFile(body, caller, role) {
  if (role !== 'admin') return { ok: false, error: 'Non hai i permessi per modificare gli allegati.' };
  try {
    var folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
    if (!folderId) return { ok: false, error: 'Cartella allegati non configurata.' };
    var dest = ensureFolderPath(DriveApp.getFolderById(folderId), (body.folderPath || []).concat('Scaduti'));
    var file = DriveApp.getFileById(body.fileId);
    var parents = file.getParents();
    while (parents.hasNext()) {
      var p = parents.next();
      if (p.getId() !== dest.getId()) p.removeFile(file);
    }
    dest.addFile(file);
    return { ok: true };
  } catch (e) {
    /* il lucchetto occupato (ensureFolderPath, sopra) è un errore diverso da "file non
       trovato": tenerli distinti permette di capire dai log se il problema è stato un
       sovraccarico momentaneo oppure un file davvero sparito. Non cambia nulla per chi usa
       l'app (questa pulizia di riserva fallisce in silenzio anche lato browser), serve solo
       a non confondersi leggendo i log. */
    return { ok: false, error: (e && /momentaneamente occupato/.test(e.message)) ? e.message : 'File non trovato o non spostabile.' };
  }
}

/* ---------- storage: Google Sheet come database, una scheda per collezione ---------- */
/* Il foglio e le singole schede vengono aperti una sola volta per esecuzione
   (una chiamata al backend può leggere fino a 18 collezioni in "listAll":
   riaprirle ogni volta era la causa principale della lentezza). */
var _ssCache = null, _sheetCache = {};
function ss() {
  if (_ssCache) return _ssCache;
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Il backend non è configurato: esegui setup() dall\'editor di Apps Script.');
  _ssCache = SpreadsheetApp.openById(id);
  return _ssCache;
}
function sheetFor(col) {
  if (_sheetCache[col]) return _sheetCache[col];
  var s = ss(), sh = s.getSheetByName(col);
  if (!sh) { sh = s.insertSheet(col); sh.appendRow(['id', 'json', 'aggiornato']); sh.setFrozenRows(1); }
  _sheetCache[col] = sh;
  return sh;
}
/* Legge solo la colonna A (gli id), non l'intero foglio: prima leggeva anche la colonna con il
   JSON di ogni riga (a volte grande, es. una commessa con iter e checklist) solo per scartarlo
   subito dopo — un lavoro inutile ripetuto a ogni singola lettura/scrittura/eliminazione, che
   con molte righe si sentiva soprattutto nelle eliminazioni. */
function findRow(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function readAll(col) {
  var sh = sheetFor(col), vals = sh.getDataRange().getValues(), out = {};
  for (var i = 1; i < vals.length; i++) {
    var id = vals[i][0], json = vals[i][1];
    if (!id) continue;
    try { out[id] = JSON.parse(json); } catch (e) { /* riga corrotta: la saltiamo */ }
  }
  return out;
}
/* Legge più collezioni in UNA sola chiamata di rete (invece di una per collezione)
   usando il servizio avanzato "Sheets API" (Servizi → + → Google Sheets API, nell'editor
   di Apps Script). Se non è stato attivato, o per qualunque altro motivo la chiamata
   fallisce, si torna automaticamente al metodo più lento ma sempre funzionante: l'app
   non si rompe mai per questo, va solo più piano finché il servizio non è attivo. */
function readManyBatch(cols) {
  try {
    if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets || !Sheets.Spreadsheets.Values) throw new Error('adv-service-off');
    var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    var ranges = cols.map(function (c) { return "'" + c + "'!A2:C"; });
    var resp = Sheets.Spreadsheets.Values.batchGet(id, { ranges: ranges });
    var out = {};
    (resp.valueRanges || []).forEach(function (vr, i) {
      var col = cols[i], map = {};
      (vr.values || []).forEach(function (row) {
        var rid = row[0], json = row[1];
        if (!rid) return;
        try { map[rid] = JSON.parse(json); } catch (e2) { /* riga corrotta: la saltiamo */ }
      });
      out[col] = map;
    });
    return out;
  } catch (e) {
    var out2 = {};
    cols.forEach(function (c) { out2[c] = readAll(c); });
    return out2;
  }
}
function readDoc(col, id) {
  var sh = sheetFor(col), row = findRow(sh, id);
  if (row < 0) return null;
  try { return JSON.parse(sh.getRange(row, 2).getValue()); } catch (e) { return null; }
}
/* Ogni scrittura (su qualunque collezione, di chiunque) passa da qui: un unico "semaforo"
   condiviso, perché Apps Script non offre un lucchetto per singola collezione, solo uno per
   script. Due scritture sulla STESSA riga devono per forza aspettare il proprio turno (altrimenti
   una delle due si perderebbe), quindi il lucchetto resta necessario; ma prima, se era occupato,
   si aspettava fino a 30 secondi e poi uscivano il messaggio d'errore grezzo di Apps Script
   (in inglese, poco chiaro). Ora l'attesa massima è più breve e l'errore, se capita, è chiaro. */
function withLock(fn) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { throw new Error('Il server è momentaneamente occupato: riprova tra qualche secondo.'); }
  try { return fn(); } finally { lock.releaseLock(); }
}
function writeDoc(col, id, obj) {
  return withLock(function () {
    var sh = sheetFor(col), row = findRow(sh, id);
    var json = JSON.stringify(obj), now = new Date().toISOString();
    if (row < 0) sh.appendRow([id, json, now]); else sh.getRange(row, 1, 1, 3).setValues([[id, json, now]]);
  });
}
function updateDoc(col, id, part) {
  return withLock(function () {
    var cur = {};
    var sh = sheetFor(col), row = findRow(sh, id);
    if (row > 0) { try { cur = JSON.parse(sh.getRange(row, 2).getValue()) || {}; } catch (e) { cur = {}; } }
    var merged = Object.assign({}, cur, part);
    var json = JSON.stringify(merged), now = new Date().toISOString();
    if (row < 0) sh.appendRow([id, json, now]); else sh.getRange(row, 1, 1, 3).setValues([[id, json, now]]);
    return merged;
  });
}
function deleteDoc(col, id) {
  return withLock(function () {
    var sh = sheetFor(col), row = findRow(sh, id);
    if (row > 0) sh.deleteRow(row);
  });
}
function todayStr() {
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

/* ---------- migrazione una tantum: sposta gli allegati già caricati nelle sottocartelle ---------- */
/* Prima di questa modifica tutti i file finivano nell'unica cartella radice "Gestionale
   Impianti - Allegati". Questa funzione sposta quelli già presenti nelle stesse sottocartelle
   che userebbe un caricamento nuovo (Commesse/<commessa>/Sicurezza, Personale/<dipendente>/
   Formazione, ecc.), cosi' anche i documenti caricati prima si ritrovano organizzati.
   Si può rilanciare più volte senza problemi: ogni volta sposta solo i file ancora nella
   cartella radice (quelli già spostati in una sottocartella non vengono più toccati), e non
   tocca MAI un file semplicemente collegato con un link "Oppure link a Google Drive" (quei
   file non sono dentro questa cartella, restano dove l'utente li aveva messi su Drive).
   Uso, dall'editor di Apps Script: scegli "migraCartelleAnteprima" dal menu delle funzioni in
   alto ed esegui (▷): nei log (Visualizza → Log, o Ctrl+Cmd+H) si vede quanti file verrebbero
   spostati e in quali cartelle, senza spostare nulla. Quando il risultato convince, esegui
   "migraCartelle" per spostarli davvero. */
/* legge tutte le collezioni e restituisce, per ogni allegato ancora collegato a una scheda
   (campo "file" singolo, gallerie foto, campi personalizzati di tipo documento, righe dei
   registri...), il suo id su Drive e il percorso di cartelle che dovrebbe avere. Usata sia da
   migraCartelleEsegui (per sapere dove spostare) sia da trovaFileOrfani (per sapere quali id
   sono ancora referenziati, e quindi NON sono orfani): tenerla in un unico punto evita che le
   due funzioni, nel tempo, finiscano per "vedere" elenchi diversi di allegati. */
function raccogliRiferimentiFile() {
  var clienti = readAll('clienti'), commesse = readAll('commesse'), dipendenti = readAll('dipendenti'),
    mezzi = readAll('mezzi'), categorie = readAll('categorie'), documenti = readAll('documenti'),
    riservato = readAll('riservato'), verbali = readAll('verbali'), docAzienda = readAll('docAzienda'),
    assicurazioni = readAll('assicurazioni'), registri = readAll('registri'), impostazioni = readAll('impostazioni');

  var campiCfg = impostazioni.campi || {};
  var registriDefs = {}; ((impostazioni.registri || {}).items || []).forEach(function (r) { registriDefs[r.id] = r; });

  var clienteNome = function (id) { var c = clienti[id]; return c ? c.ragioneSociale : ''; };
  var commessaLabel = function (c) { return (c.codice || '') + ' · ' + clienteNome(c.clienteId); };
  var nomeD = function (d) { return (d.cognome || '') + ' ' + (d.nome || ''); };
  var mezzoLabel = function (m) { return m.descrizione + (m.targa ? ' ' + m.targa : ''); };
  var mezzoCartella = function (tipo) { return tipo === 'ATTREZZATURA' ? 'Attrezzature' : 'Mezzi'; };
  var TIPI_DOC_TEC = ['PROGETTO', 'VALUTAZIONE_ENERGETICA', 'BOLLETTA_CLIENTE', 'SCHEDA_TECNICA'];
  var docCommCartella = function (tipo) { return TIPI_DOC_TEC.indexOf(tipo) >= 0 ? 'Documentazione tecnica' : 'Documentazione amministrativa'; };
  var catPath = function (id) {
    var c = categorie[id]; if (!c) return '';
    var p = c.parentId && categorie[c.parentId];
    return p ? p.nome + ' / ' + c.nome : c.nome;
  };

  var tasks = [];
  var add = function (fileOrId, segments) {
    var id = fileOrId && typeof fileOrId === 'object' ? fileOrId.id : fileOrId;
    if (!id) return;
    tasks.push({ fileId: id, segments: segments });
  };

  Object.keys(commesse).forEach(function (id) {
    var c = commesse[id], lab = ['Commesse', commessaLabel(c)];
    (c.sicurezza || []).forEach(function (s) { if (s.file) add(s.file, lab.concat('Sicurezza')); });
    (c.documentiAmm || []).forEach(function (x) { if (x.file) add(x.file, lab.concat(docCommCartella(x.tipo))); });
    (c.fotoSopralluogo || []).forEach(function (f) { add(f, lab.concat('Foto sopralluogo')); });
    (c.fotoInstallazione || []).forEach(function (f) { add(f, lab.concat('Foto installazione')); });
    (campiCfg.commesse || []).forEach(function (cf) { if (cf.t === 'file' && c[cf.k]) add(c[cf.k], lab); });
  });

  Object.keys(dipendenti).forEach(function (id) {
    var d = dipendenti[id], lab = ['Personale', nomeD(d)];
    (d.abilitazioni || []).forEach(function (a) { if (a.file) add(a.file, lab.concat('Formazione')); });
    (d.documenti || []).forEach(function (x) { if (x.file) add(x.file, lab.concat('Documenti personali')); });
    (campiCfg.dipendenti || []).forEach(function (cf) { if (cf.t === 'file' && d[cf.k]) add(d[cf.k], lab); });
  });

  Object.keys(mezzi).forEach(function (id) {
    var m = mezzi[id], lab = ['Azienda', 'Mezzi e attrezzature', mezzoCartella(m.tipo), mezzoLabel(m)];
    (m.controlli || []).forEach(function (ct) { if (ct.file) add(ct.file, lab.concat('Controlli')); });
    (campiCfg.mezzi || []).forEach(function (cf) { if (cf.t === 'file' && m[cf.k]) add(m[cf.k], lab); });
  });

  Object.keys(clienti).forEach(function (id) {
    var cl = clienti[id];
    (campiCfg.clienti || []).forEach(function (cf) { if (cf.t === 'file' && cl[cf.k]) add(cl[cf.k], ['Clienti', cl.ragioneSociale]); });
  });

  var docLabel = function (x) {
    if (x.commessaId && commesse[x.commessaId]) return ['Commesse', commessaLabel(commesse[x.commessaId]), 'Documenti'];
    return ['Documentale', catPath(x.categoriaId) || 'Senza categoria'];
  };
  Object.keys(documenti).forEach(function (id) { var x = documenti[id]; if (x.file) add(x.file, docLabel(x)); });
  Object.keys(riservato).forEach(function (id) {
    var x = riservato[id];
    if (x.kind === 'doc' && x.file) add(x.file, docLabel(x));
    else if (x.kind === 'verbale' && x.file) add(x.file, ['Documentale', 'Verbali']);
  });
  Object.keys(verbali).forEach(function (id) { var x = verbali[id]; if (x.file) add(x.file, ['Documentale', 'Verbali']); });
  Object.keys(docAzienda).forEach(function (id) { var x = docAzienda[id]; if (x.file) add(x.file, ['Azienda', 'Documenti impresa']); });
  Object.keys(assicurazioni).forEach(function (id) { var x = assicurazioni[id]; if (x.file) add(x.file, ['Azienda', 'Assicurazioni']); });
  Object.keys(registri).forEach(function (id) {
    var x = registri[id], def = registriDefs[x.reg]; if (!def) return;
    (def.campi || []).forEach(function (cf) { if (cf.t === 'file' && x[cf.k]) add(x[cf.k], ['Registri', def.nome]); });
  });

  return tasks;
}

/* trova tutte le cartelle "Allegati" (quella principale collegata a FOLDER_ID, più eventuali
   doppioni): usata sia dalla migrazione sia dalla ricerca dei file orfani. Cerca per
   CONTENUTO del nome ("contains", non il nome esatto): un confronto col nome esatto, trattino
   compreso, si è dimostrato fragile — un copia-incolla manuale del codice può cambiare quel
   trattino senza che si noti, e da quel momento la ricerca non trova più i doppioni già
   esistenti (bug successo più volte: trattino diverso trovato il 2026-10-01, poi di nuovo con
   due cartelle create quasi insieme il 2026-10-06). Cercare "contiene" invece che "è uguale a"
   funziona qualunque sia il trattino usato. */
function trovaRadiciAllegati(root) {
  var extra = [];
  var it = DriveApp.searchFolders("title contains 'Gestionale Impianti' and title contains 'Allegati' and trashed = false");
  while (it.hasNext()) {
    var f = it.next();
    if (f.getId() !== root.getId()) extra.push(f);
  }
  return [root].concat(extra);
}

function migraCartelleAnteprima() { return migraCartelleEsegui(true); }
function migraCartelle() { return migraCartelleEsegui(false); }
/* cerca (senza creare nulla) la sottocartella che corrisponde a un percorso di pagina/sezione,
   partendo da una cartella radice: restituisce null se un pezzo del percorso non esiste ancora
   (usata per capire se un file è GIÀ al posto giusto, senza il rischio di crearlo per sbaglio
   solo per fare questo controllo). */
function trovaFolderIdEsistente(radiceFolder, segments) {
  var folder = radiceFolder;
  for (var i = 0; i < (segments || []).length && i < 6; i++) {
    var nome = String(segments[i] || '').trim().slice(0, 100);
    if (!nome) continue;
    var it = folder.getFoldersByName(nome);
    if (!it.hasNext()) return null;
    folder = it.next();
  }
  return folder.getId();
}
function migraCartelleEsegui(dryRun) {
  var folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!folderId) { Logger.log('Cartella allegati non configurata: esegui prima setup().'); return; }
  var root = DriveApp.getFolderById(folderId);

  var radici = trovaRadiciAllegati(root);
  if (radici.length > 1) Logger.log('Trovate ' + (radici.length - 1) + ' cartella/e "Allegati" in più oltre a quella principale: i loro file verranno spostati nella struttura dentro "' + root.getName() + '" e poi resteranno vuote (puoi eliminarle a mano da Drive).');
  var radiciIds = {}; radici.forEach(function (r) { radiciIds[r.getId()] = true; });

  var tasks = raccogliRiferimentiFile();

  var daSpostare = 0, nonInRadice = 0, nonTrovati = 0, errori = 0, cartelleUsate = {};
  tasks.forEach(function (t) {
    var file;
    try { file = DriveApp.getFileById(t.fileId); }
    catch (e) { nonTrovati++; return; }

    var parents = file.getParents(), direct = parents.hasNext() ? parents.next() : null;

    /* se il genitore diretto del file è già esattamente la cartella corretta (dentro la
       cartella radice principale), non c'è nulla da fare: era solo questo controllo che prima
       mancava — un file già organizzato in una sottocartella (anche dentro la cartella doppia
       sbagliata, con la STESSA struttura Commesse/Personale/ecc.) non veniva riconosciuto come
       "da spostare" perché il controllo guardava solo se il genitore diretto fosse una delle
       cartelle radice, non una loro sottocartella. */
    var destIdEsistente = trovaFolderIdEsistente(root, t.segments);
    if (direct && destIdEsistente && direct.getId() === destIdEsistente) return;

    /* risale la catena delle cartelle superiori (fino a 8 livelli, più che sufficienti) per
       capire se il file si trova comunque DENTRO l'albero di una delle cartelle "Allegati"
       conosciute (radice principale o doppioni), anche se non è nel punto esatto giusto —
       senza toccare MAI un file che sta altrove nel Drive dell'utente (es. un link scelto con
       "Sfoglia Drive" da una cartella sua, che non c'entra nulla con questo). */
    var dentro = false, cursore = direct, profondita = 0;
    while (cursore && profondita < 8) {
      if (radiciIds[cursore.getId()]) { dentro = true; break; }
      var pp = cursore.getParents();
      cursore = pp.hasNext() ? pp.next() : null;
      profondita++;
    }
    if (!dentro) { nonInRadice++; return; }

    var chiave = t.segments.join(' / ');
    cartelleUsate[chiave] = (cartelleUsate[chiave] || 0) + 1;
    if (dryRun) { daSpostare++; return; }
    try {
      var dest = ensureFolderPath(root, t.segments);
      dest.addFile(file);
      var curParents = file.getParents();
      while (curParents.hasNext()) {
        var cp = curParents.next();
        if (cp.getId() !== dest.getId()) cp.removeFile(file);
      }
      daSpostare++;
    } catch (e2) { errori++; }
  });

  Logger.log((dryRun ? '[ANTEPRIMA, nessun file è stato spostato] ' : '[FATTO] ') + 'File ' + (dryRun ? 'da spostare' : 'spostati') + ': ' + daSpostare + ' · non dentro a nessuna cartella "Allegati" (già organizzati in precedenza, oppure link a file altrove su Drive — non toccati): ' + nonInRadice + ' · non trovati (eliminati o senza permesso): ' + nonTrovati + (errori ? ' · errori: ' + errori : ''));
  Logger.log('Cartelle coinvolte:\n' + Object.keys(cartelleUsate).sort().map(function (k) { return '  ' + k + ' (' + cartelleUsate[k] + ')'; }).join('\n'));
  return { daSpostare: daSpostare, nonInRadice: nonInRadice, nonTrovati: nonTrovati, errori: errori };
}

/* ---------- una tantum: unisce le sottocartelle doppie (stesso nome, stesso genitore) ---------- */
/* Prima che ensureFolderPath venisse protetta da un lucchetto (vedi sopra), due caricamenti
   quasi contemporanei in una sottocartella non ancora creata (due schede aperte, due persone,
   o il tentativo diretto dal browser e quello di riserva via Apps Script in corsa fra loro)
   potevano crearne due con lo stesso nome, sotto lo stesso genitore (es. due cartelle
   "Sicurezza" dentro la stessa commessa, o due cartelle con il nome di uno stesso dipendente).
   Questa funzione le trova, a qualunque livello dentro "Gestionale Impianti — Allegati", e le
   unisce: tiene la più vecchia, sposta dentro di lei tutti i file e le sottocartelle delle
   altre con lo stesso nome, poi cestina quelle rimaste vuote (non c'è nulla da perdere: sono
   vuote solo perché tutto il loro contenuto è stato spostato nella cartella tenuta).
   Uso, dall'editor di Apps Script: esegui prima "unisciCartelleDoppieAnteprima" (▷) e guarda
   il log (Visualizza → Log, o Ctrl+Cmd+H); se il numero ti convince, esegui
   "unisciCartelleDoppie" per farlo davvero. */
function unisciCartelleDoppieAnteprima() { return unisciCartelleDoppieEsegui(true); }
function unisciCartelleDoppie() { return unisciCartelleDoppieEsegui(false); }
function unisciCartelleDoppieEsegui(dryRun) {
  var folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!folderId) { Logger.log('Cartella allegati non configurata: esegui prima setup().'); return; }
  var root = DriveApp.getFolderById(folderId);
  var gruppi = [], fileSpostati = 0, cartelleCestinate = 0, errori = 0;

  var elabora = function (folder, path) {
    var perNome = {};
    var sub = folder.getFolders();
    while (sub.hasNext()) {
      var f = sub.next(), nome = f.getName();
      (perNome[nome] = perNome[nome] || []).push(f);
    }
    Object.keys(perNome).forEach(function (nome) {
      var lista = perNome[nome];
      lista.sort(function (a, b) { return a.getDateCreated().getTime() - b.getDateCreated().getTime(); });
      var tieni = lista[0];
      if (lista.length > 1) {
        gruppi.push((path ? path + ' / ' : '') + nome + ' (' + lista.length + ' cartelle)');
        lista.slice(1).forEach(function (d) {
          if (!dryRun) {
            try {
              var fileDaSpostare = [];
              var fileIt = d.getFiles();
              while (fileIt.hasNext()) fileDaSpostare.push(fileIt.next());
              fileDaSpostare.forEach(function (fl) { tieni.addFile(fl); d.removeFile(fl); fileSpostati++; });

              var subDaSpostare = [];
              var subIt = d.getFolders();
              while (subIt.hasNext()) subDaSpostare.push(subIt.next());
              subDaSpostare.forEach(function (sf) { tieni.addFolder(sf); d.removeFolder(sf); });

              d.setTrashed(true);
              cartelleCestinate++;
            } catch (e) { errori++; }
          }
        });
      }
      elabora(tieni, (path ? path + ' / ' : '') + nome);
    });
  };
  elabora(root, '');

  Logger.log((dryRun ? '[ANTEPRIMA, nessuna modifica] ' : '[FATTO] ') + 'Gruppi di cartelle doppie trovati: ' + gruppi.length + (dryRun ? '' : ' · file spostati: ' + fileSpostati + ' · cartelle vuote cestinate: ' + cartelleCestinate + (errori ? ' · errori: ' + errori : '')));
  if (gruppi.length) Logger.log('Cartelle coinvolte:\n' + gruppi.map(function (g) { return '  ' + g; }).join('\n'));
  return { gruppi: gruppi.length, fileSpostati: fileSpostati, cartelleCestinate: cartelleCestinate, errori: errori };
}

/* ---------- trova (senza eliminare nulla) i file rimasti nelle cartelle "Allegati" che non
   sono più collegati a nessuna scheda ---------- */
/* Succedeva soprattutto con le sostituzioni fatte prima della correzione del 2026-10-01 (quando
   si sostituiva un allegato, la copia vecchia restava per sempre su Drive senza che nessuna
   scheda la referenziasse più). Questa funzione NON cancella nulla: scorre tutte le cartelle
   "Allegati" (quella principale e eventuali doppioni) e tutte le loro sottocartelle, confronta
   ogni file trovato con l'elenco di quelli ancora referenziati nei dati, e scrive l'elenco di
   quelli "orfani" (fisicamente su Drive ma non più collegati a nessuna scheda) in un nuovo
   foglio chiamato "FileOrfani" dentro lo stesso Google Sheet dei dati dell'app — con nome,
   cartella, link diretto e dimensione di ciascuno — cosi' li puoi controllare con calma, uno per
   uno, prima di decidere se eliminarli (da Drive, a mano: dal link nel foglio).
   Uso: dal menu delle funzioni in alto nell'editor scegli "trovaFileOrfani" ed esegui (▷), poi
   apri il Google Sheet dei dati (quello creato da setup(), non questo editor) e guarda il foglio
   "FileOrfani". */
function trovaFileOrfani() {
  var folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!folderId) { Logger.log('Cartella allegati non configurata: esegui prima setup().'); return; }
  var root = DriveApp.getFolderById(folderId);
  var radici = trovaRadiciAllegati(root);

  var trovati = [];
  var scorri = function (folder, path) {
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      trovati.push({ id: f.getId(), name: f.getName(), url: f.getUrl(), path: path || '(cartella radice)', size: f.getSize(), date: f.getLastUpdated() });
    }
    var subs = folder.getFolders();
    while (subs.hasNext()) {
      var sf = subs.next();
      scorri(sf, path ? path + ' / ' + sf.getName() : sf.getName());
    }
  };
  radici.forEach(function (r) { scorri(r, r.getName()); });

  var referenziati = {};
  raccogliRiferimentiFile().forEach(function (t) { referenziati[t.fileId] = true; });

  var orfani = trovati.filter(function (x) { return !referenziati[x.id]; });

  var sh = ss().getSheetByName('FileOrfani');
  if (sh) sh.clear(); else sh = ss().insertSheet('FileOrfani');
  sh.appendRow(['Nome file', 'Cartella', 'Link', 'Dimensione (KB)', 'Ultima modifica']);
  orfani.forEach(function (o) { sh.appendRow([o.name, o.path, o.url, Math.round(o.size / 1024), o.date]); });
  sh.setFrozenRows(1);
  if (orfani.length) sh.autoResizeColumns(1, 5);

  var totKB = Math.round(orfani.reduce(function (s, o) { return s + o.size; }, 0) / 1024);
  Logger.log('File trovati nelle cartelle "Allegati": ' + trovati.length + ' · non più collegati a nessuna scheda (orfani): ' + orfani.length + ' (circa ' + totKB + ' KB) · elenco scritto nel foglio "FileOrfani" del Google Sheet dei dati. Nessun file è stato eliminato: controllali con calma prima di cancellarli a mano da Drive.');
  return { trovati: trovati.length, orfani: orfani.length };
}

/* ---------- una tantum: rende visibili a chiunque abbia il link le foto già collegate con
   "Sfoglia Drive" ---------- */
/* Prima della correzione del 2026-10-01, le foto di sopralluogo/installazione collegate con
   "Sfoglia Drive" (non caricate come copia) restavano condivise solo con chi aveva già accesso
   al file: per questo non si vedeva l'anteprima nella scheda della commessa. Questa funzione
   scorre tutte le commesse e, per ogni foto nelle due gallerie, imposta la condivisione
   "chiunque abbia il link" (sola lettura) — lo stesso trattamento che l'app applica da sola
   alle nuove foto da ora in poi. Non sposta né elimina nulla, cambia solo chi può vederle.
   Se una foto è dentro un Drive condiviso su cui questo account non ha i permessi per
   cambiare la condivisione, viene saltata senza bloccare le altre.
   Uso: scegli "condividiFotoGiaCollegate" dal menu delle funzioni in alto ed esegui (▷), poi
   guarda il risultato nei log. */
function condividiFotoGiaCollegate() {
  var commesse = readAll('commesse');
  var fatte = 0, saltate = 0, totale = 0;
  Object.keys(commesse).forEach(function (id) {
    var c = commesse[id];
    ['fotoSopralluogo', 'fotoInstallazione'].forEach(function (key) {
      (c[key] || []).forEach(function (f) {
        if (!f || !f.id) return;
        totale++;
        try {
          DriveApp.getFileById(f.id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          fatte++;
        } catch (e) { saltate++; }
      });
    });
  });
  Logger.log('Foto trovate nelle gallerie delle commesse: ' + totale + ' · rese visibili con il link: ' + fatte + (saltate ? ' · saltate (permessi non modificabili, o file non più esistente): ' + saltate : ''));
  return { totale: totale, fatte: fatte, saltate: saltate };
}
