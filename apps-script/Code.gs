/**
 * Gestionale I-Energy — backend Google Apps Script (solo file su Google Drive).
 *
 * L'accesso e i dati stanno su Supabase. Questo script serve unicamente per i file:
 * crea le cartelle, carica gli allegati e sposta in "Scaduti" quelli sostituiti.
 * Chi lo chiama deve avere un accesso valido (token di Supabase); solo l'amministratore può
 * caricare o spostare file.
 */

/* ---------- setup iniziale (da eseguire una sola volta dall'editor di Apps Script) ---------- */
function setup() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FOLDER_ID')) {
    var folder = DriveApp.createFolder('Gestionale Impianti — Allegati');
    props.setProperty('FOLDER_ID', folder.getId());
    Logger.log('Creata cartella Drive: ' + folder.getUrl());
  }
}

/* ---------- entry point del Web App ---------- */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, info: 'Gestionale I-Energy — backend file attivo. Usa POST.' })).setMimeType(ContentService.MimeType.JSON);
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
    case 'whoami': return { ok: true, email: caller.email, name: caller.name, role: role, driveFolderId: PropertiesService.getScriptProperties().getProperty('FOLDER_ID') || '' };
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
  if (role !== 'admin') return { ok: false, error: 'Non hai i permessi per modificare gli allegati.' };
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

function withLock(fn) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { throw new Error('Il server è momentaneamente occupato: riprova tra qualche secondo.'); }
  try { return fn(); } finally { lock.releaseLock(); }
}
