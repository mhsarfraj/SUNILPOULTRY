/**
 * SUNIL POULTRY - Google Sheets backend (Google Apps Script web app)
 *
 * Every record the website saves is written into the Google Sheet this script is
 * attached to. Photos are saved in a Google Drive folder; the Sheet stores their IDs.
 *
 * FIRST-TIME SETUP: see DEPLOY.md  (short version: paste Code.gs + SeedData.gs,
 * run setupAndSeed once, Deploy > New deployment > Web app > "Anyone").
 */

// ====================== CONFIG (edit these) ======================
const CONFIG = {
  OWNER_ID: 'owner',                    // owner's login ID
  OWNER_INITIAL_PASSWORD: 'Sunil@2026', // used only the first time; change it inside the app afterwards
  SAMPLE_FARMER_PASSWORD: 'farmer@123', // password given to the sample farmers by seedSampleData()
  REQUIRE_PHOTOS: true,                 // feed + weight photos always, mortality photo when deaths > 0
  TZ: 'Asia/Kolkata',
  PHOTO_FOLDER_NAME: 'Sunil Poultry Photos',
  SESSION_DAYS: 30
};

// ====================== SHEET SCHEMA ======================
const SCHEMA = {
  Farmers:    ['id','code','name','phone','village','status','photo','passHash','salt','createdAt'],
  Batches:    ['id','farmerId','name','startDate','initial','breed','notes','status','completedAt'],
  Reports:    ['id','batchId','farmerId','date','deaths','feedBags','weight','notes','photoFeed','photoMortality','photoWeight','submittedAt'],
  Dispatches: ['id','batchId','farmerId','date','quantity','shift','time','truckNo','challan','driver','destination','notes','createdAt'],
  BagFlow:    ['id','farmerId','batchId','qty','origin','status','note','ownerNote','requestedAt','assignedAt','respondedAt'],
  Settings:   ['key','value']
};
const NUM_FIELDS  = {Batches:['initial'], Reports:['deaths','feedBags','weight'], Dispatches:['quantity'], BagFlow:['qty']};
const DATE_FIELDS = {Batches:['startDate'], Reports:['date'], Dispatches:['date']};
const TEXT_FIELDS = {Farmers:['phone','code','id']};

// ====================== ENTRY POINTS ======================
function doGet() { return json_({ ok: true, service: 'Sunil Poultry API', time: new Date().toISOString() }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    ensureSetup_();
    return json_(route_(req));
  } catch (err) {
    return json_({ ok: false, error: err && err.userMessage ? err.userMessage : 'Server error: ' + (err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function fail_(msg) { const e = new Error(msg); e.userMessage = msg; throw e; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function route_(req) {
  const a = req.action;
  if (a === 'login') return login_(req);
  if (a === 'resetPassword') return resetPassword_(req);

  const who = verifyToken_(req.token);
  const owner = who.r === 'owner';
  const needOwner = () => { if (!owner) fail_('Owner access required'); };
  const needFarmer = () => { if (owner) fail_('Farmer access required'); };

  switch (a) {
    case 'getData':           break;
    case 'saveFarmer':        needOwner(); saveFarmer_(req); break;
    case 'saveBatch':         needOwner(); saveBatch_(req); break;
    case 'setBatchStatus':    needOwner(); setBatchStatus_(req); break;
    case 'submitReport':      needFarmer(); submitReport_(req, who); break;
    case 'deleteReport':      needOwner(); deleteReport_(req); break;
    case 'saveDispatch':      saveDispatch_(req, who); break;
    case 'deleteDispatch':    needOwner(); remove_('Dispatches', req.id); break;
    case 'assignBags':        needOwner(); assignBags_(req); break;
    case 'requestBags':       needFarmer(); requestBags_(req, who); break;
    case 'respondRequest':    needOwner(); respondRequest_(req); break;
    case 'respondAssignment': needFarmer(); respondAssignment_(req, who); break;
    case 'cancelBags':        cancelBags_(req, who); break;
    case 'changeOwnerPassword': needOwner(); changeOwnerPassword_(req); break;
    default: fail_('Unknown action');
  }
  return { ok: true, data: getData_(who) };
}

// ====================== SETUP ======================
function ensureSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]).setFontWeight('bold').setBackground('#d9f0f3');
      sh.setFrozenRows(1);
      const textCols = (DATE_FIELDS[name] || []).concat(TEXT_FIELDS[name] || []);
      textCols.forEach(f => {
        const c = SCHEMA[name].indexOf(f) + 1;
        sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@');
      });
    }
  });
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) { try { ss.deleteSheet(def); } catch (e) {} }
  if (!getSetting_('ownerHash')) {
    const salt = Utilities.getUuid().slice(0, 8);
    setSetting_('ownerSalt', salt);
    setSetting_('ownerHash', hash_(CONFIG.OWNER_INITIAL_PASSWORD, salt));
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
}

/** Run this ONCE from the editor: creates the sheets and loads the sample data. */
function setupAndSeed() { ensureSetup_(); seedSampleData(); }

// ====================== TABLE HELPERS ======================
function sheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd');
  return String(v == null ? '' : v).slice(0, 10);
}

function readAll_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const head = SCHEMA[name];
  const vals = sh.getRange(2, 1, last - 1, head.length).getValues();
  const nums = NUM_FIELDS[name] || [], dates = DATE_FIELDS[name] || [];
  const out = [];
  vals.forEach(row => {
    if (row[0] === '' || row[0] == null) return;
    const o = {};
    head.forEach((k, i) => {
      let v = row[i];
      if (nums.indexOf(k) >= 0) v = Number(v) || 0;
      else if (dates.indexOf(k) >= 0) v = normDate_(v);
      else if (v instanceof Date) v = Utilities.formatDate(v, CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");
      else v = (v == null ? '' : String(v));
      o[k] = v;
    });
    out.push(o);
  });
  return out;
}

function rowFor_(name, o) { return SCHEMA[name].map(k => (o[k] === undefined || o[k] === null) ? '' : o[k]); }
function insert_(name, o) { sheet_(name).appendRow(rowFor_(name, o)); }

function findRow_(name, id) {
  const sh = sheet_(name), last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function update_(name, id, patch) {
  const r = findRow_(name, id);
  if (r < 0) fail_('Record not found');
  const sh = sheet_(name), head = SCHEMA[name];
  const cur = sh.getRange(r, 1, 1, head.length).getValues()[0];
  head.forEach((k, i) => { if (patch[k] !== undefined) cur[i] = patch[k]; });
  sh.getRange(r, 1, 1, head.length).setValues([cur]);
}

function remove_(name, id) {
  const r = findRow_(name, id);
  if (r < 0) fail_('Record not found');
  sheet_(name).deleteRow(r);
}

function bulkInsert_(name, rows) {
  if (!rows.length) return;
  const sh = sheet_(name), start = sh.getLastRow() + 1;
  if (start + rows.length > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), start + rows.length - sh.getMaxRows() + 50);
  sh.getRange(start, 1, rows.length, SCHEMA[name].length).setValues(rows);
}

function getSetting_(key) {
  const r = readAll_('Settings').filter(x => x.key === key)[0];
  return r ? r.value : '';
}
function setSetting_(key, value) {
  const sh = sheet_('Settings'), r = findRow_('Settings', key);
  if (r < 0) sh.appendRow([key, value]); else sh.getRange(r, 2).setValue(value);
}

function todayStr_() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'); }
function nowStr_() { return Utilities.formatDate(new Date(), CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function uid_(p) { return p + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

// ====================== AUTH ======================
function hash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + String(password), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function b64url_(bytes) { return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }

function makeToken_(payload) {
  const body = b64url_(Utilities.newBlob(JSON.stringify(payload)).getBytes());
  const sig = b64url_(Utilities.computeHmacSha256Signature(body, PropertiesService.getScriptProperties().getProperty('SECRET')));
  return body + '.' + sig;
}

function verifyToken_(token) {
  if (!token || String(token).indexOf('.') < 0) fail_('Please sign in again');
  const parts = String(token).split('.');
  const sig = b64url_(Utilities.computeHmacSha256Signature(parts[0], PropertiesService.getScriptProperties().getProperty('SECRET')));
  if (sig !== parts[1]) fail_('Please sign in again');
  let p;
  try {
    const padded = parts[0].replace(/-/g, '+').replace(/_/g, '/');
    p = JSON.parse(Utilities.newBlob(Utilities.base64Decode(padded + '===='.slice(0, (4 - padded.length % 4) % 4))).getDataAsString());
  } catch (e) { fail_('Please sign in again'); }
  if (!p || p.e < Date.now()) fail_('Session expired, please sign in again');
  if (p.r === 'farmer') {
    const f = readAll_('Farmers').filter(x => x.id === p.f)[0];
    if (!f || f.status === 'Inactive') fail_('This farmer account is inactive');
  }
  return p;
}

function throttle_(key, max, seconds) {
  const cache = CacheService.getScriptCache();
  const k = 'thr_' + key.toLowerCase().replace(/[^a-z0-9]/g, '');
  const n = Number(cache.get(k) || 0);
  if (n >= max) fail_('Too many attempts. Please wait a few minutes and try again.');
  return () => cache.put(k, String(n + 1), seconds);
}

function normPhone_(p) { return String(p || '').replace(/\D/g, '').slice(-10); }
function normText_(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }

function login_(req) {
  const role = req.role, id = String(req.id || '').trim(), pw = String(req.password || '');
  if (!id || !pw) fail_('Enter your ID and password');
  const bump = throttle_('login' + id, 8, 600);
  if (role === 'owner') {
    if (id.toLowerCase() !== CONFIG.OWNER_ID.toLowerCase() || hash_(pw, getSetting_('ownerSalt')) !== getSetting_('ownerHash')) { bump(); fail_('Wrong owner ID or password'); }
    const who = { r: 'owner', f: '', e: Date.now() + CONFIG.SESSION_DAYS * 864e5 };
    return { ok: true, token: makeToken_(who), data: getData_(who) };
  }
  const idl = id.toLowerCase(), ph = normPhone_(id);
  const f = readAll_('Farmers').filter(x => x.code.toLowerCase() === idl || (ph.length === 10 && normPhone_(x.phone) === ph))[0];
  if (!f || !f.passHash || hash_(pw, f.salt) !== f.passHash) { bump(); fail_('Wrong farmer code / phone or password'); }
  if (f.status === 'Inactive') fail_('This farmer account is inactive. Contact the owner.');
  const who = { r: 'farmer', f: f.id, e: Date.now() + CONFIG.SESSION_DAYS * 864e5 };
  return { ok: true, token: makeToken_(who), data: getData_(who) };
}

/** Farmer forgot / wants to change password: must match ALL personal details saved by the owner. */
function resetPassword_(req) {
  const code = String(req.code || '').trim().toLowerCase();
  const np = String(req.newPassword || '');
  if (!code) fail_('Enter your farmer code');
  if (np.length < 4) fail_('New password must be at least 4 characters');
  const bump = throttle_('reset' + code, 5, 900);
  const f = readAll_('Farmers').filter(x => x.code.toLowerCase() === code)[0];
  const ok = f && normText_(f.name) === normText_(req.name) && normPhone_(f.phone) === normPhone_(req.phone) &&
             normPhone_(f.phone).length === 10 && normText_(f.village) === normText_(req.village);
  if (!ok) { bump(); fail_('Details do not match what the owner saved. Check name, phone and village.'); }
  const salt = Utilities.getUuid().slice(0, 8);
  update_('Farmers', f.id, { salt: salt, passHash: hash_(np, salt) });
  return { ok: true, message: 'Password changed. You can sign in now.' };
}

function changeOwnerPassword_(req) {
  if (hash_(String(req.oldPassword || ''), getSetting_('ownerSalt')) !== getSetting_('ownerHash')) fail_('Current password is wrong');
  if (String(req.newPassword || '').length < 6) fail_('New password must be at least 6 characters');
  const salt = Utilities.getUuid().slice(0, 8);
  setSetting_('ownerSalt', salt);
  setSetting_('ownerHash', hash_(req.newPassword, salt));
}

// ====================== DATA ======================
function getData_(who) {
  const owner = who.r === 'owner';
  let farmers = readAll_('Farmers').map(f => { const c = Object.assign({}, f); c.hasPassword = !!f.passHash; delete c.passHash; delete c.salt; return c; });
  let batches = readAll_('Batches'), reports = readAll_('Reports'), dispatches = readAll_('Dispatches'), bagflow = readAll_('BagFlow');
  if (!owner) {
    const fid = who.f;
    farmers = farmers.filter(f => f.id === fid);
    batches = batches.filter(b => b.farmerId === fid);
    reports = reports.filter(r => r.farmerId === fid);
    dispatches = dispatches.filter(d => d.farmerId === fid);
    bagflow = bagflow.filter(x => x.farmerId === fid);
  }
  return {
    today: todayStr_(), now: nowStr_(),
    me: { role: who.r, farmerId: who.f || '' },
    requirePhotos: CONFIG.REQUIRE_PHOTOS,
    farmers: farmers, batches: batches, reports: reports, dispatches: dispatches, bagflow: bagflow
  };
}

function deathsOf_(batchId, reports) { return reports.filter(r => r.batchId === batchId).reduce((s, r) => s + r.deaths, 0); }
function loadedOf_(batchId, disp) { return disp.filter(d => d.batchId === batchId).reduce((s, d) => s + d.quantity, 0); }

// ====================== PHOTOS ======================
function photoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(CONFIG.PHOTO_FOLDER_NAME);
  props.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}

function saveImage_(dataUrl, prefix) {
  if (!dataUrl) return '';
  const m = /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+\/=]+)$/.exec(String(dataUrl));
  if (!m) fail_('Unsupported image. Use a JPG or PNG photo.');
  if (m[2].length > 3.5 * 1024 * 1024) fail_('Photo is too large. Try again with a smaller photo.');
  const ext = m[1] === 'image/png' ? 'png' : (m[1] === 'image/webp' ? 'webp' : 'jpg');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], prefix + '_' + Date.now() + '.' + ext);
  const file = photoFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

// ====================== FARMERS ======================
function saveFarmer_(req) {
  const f = req.farmer || {};
  const name = String(f.name || '').trim(), phone = normPhone_(f.phone), village = String(f.village || '').trim();
  if (!name) fail_('Enter the farmer name');
  if (phone.length !== 10) fail_('Enter a valid 10-digit phone number');
  if (!village) fail_('Enter the village / location (used to verify password changes)');
  const farmers = readAll_('Farmers');
  if (farmers.some(x => x.id !== f.id && normPhone_(x.phone) === phone)) fail_('Another farmer already uses this phone number');
  const pw = String(req.password || '');
  let photo = '';
  if (req.photo) photo = saveImage_(req.photo, 'farmer');

  if (f.id) {
    const cur = farmers.filter(x => x.id === f.id)[0];
    if (!cur) fail_('Farmer not found');
    const patch = { name: name, phone: phone, village: village, status: f.status === 'Inactive' ? 'Inactive' : 'Active' };
    if (photo) patch.photo = photo;
    if (pw) { if (pw.length < 4) fail_('Password must be at least 4 characters'); patch.salt = Utilities.getUuid().slice(0, 8); patch.passHash = hash_(pw, patch.salt); }
    update_('Farmers', f.id, patch);
  } else {
    if (pw.length < 4) fail_('Set a password of at least 4 characters for the farmer');
    const next = farmers.reduce((m, x) => Math.max(m, Number((/^F(\d+)$/i.exec(x.code) || [])[1] || 0)), 0) + 1;
    const salt = Utilities.getUuid().slice(0, 8);
    insert_('Farmers', { id: uid_('f'), code: 'F' + next, name: name, phone: phone, village: village, status: 'Active', photo: photo, passHash: hash_(pw, salt), salt: salt, createdAt: nowStr_() });
  }
}

// ====================== BATCHES ======================
function saveBatch_(req) {
  const b = req.batch || {};
  const initial = Math.round(Number(b.initial));
  if (!b.farmerId || !readAll_('Farmers').some(f => f.id === b.farmerId)) fail_('Choose a farmer');
  if (!(initial > 0)) fail_('Enter the number of chicks placed');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(b.startDate || ''))) fail_('Choose the placement date');
  const batches = readAll_('Batches');
  if (b.id) {
    update_('Batches', b.id, { farmerId: b.farmerId, name: String(b.name || '').trim(), startDate: b.startDate, initial: initial, breed: b.breed || 'Broiler', notes: String(b.notes || '').trim() });
  } else {
    const next = batches.reduce((m, x) => Math.max(m, Number((/^B-(\d+)$/.exec(x.id) || [])[1] || 0)), 0) + 1;
    const id = 'B-' + ('000' + next).slice(-3);
    insert_('Batches', { id: id, farmerId: b.farmerId, name: String(b.name || '').trim() || ('Batch ' + id), startDate: b.startDate, initial: initial, breed: b.breed || 'Broiler', notes: String(b.notes || '').trim(), status: 'Active', completedAt: '' });
  }
}

function setBatchStatus_(req) {
  if (req.status === 'Finished') update_('Batches', req.id, { status: 'Finished', completedAt: nowStr_() });
  else update_('Batches', req.id, { status: 'Active', completedAt: '' });
}

// ====================== DAILY REPORT (once per batch per day) ======================
function submitReport_(req, who) {
  const r = req.report || {};
  const batch = readAll_('Batches').filter(b => b.id === r.batchId && b.farmerId === who.f)[0];
  if (!batch) fail_('Batch not found');
  if (batch.status === 'Finished') fail_('This batch is finished. Ask the owner to reopen it.');
  const date = todayStr_();
  const reports = readAll_('Reports');
  if (reports.some(x => x.batchId === batch.id && x.date === date)) fail_('Today\'s report for this batch is already submitted.');

  const deaths = Math.round(Number(r.deaths)), feed = Number(r.feedBags), weight = Number(r.weight);
  if (!(deaths >= 0)) fail_('Enter the number of dead chicks (0 if none)');
  if (!(feed >= 0) || r.feedBags === '' || r.feedBags === null) fail_('Enter feed bags used (0 if none)');
  if (!(weight > 0 && weight < 10)) fail_('Enter the average weight of one chick in kg');
  const live = batch.initial - deathsOf_(batch.id, reports) - loadedOf_(batch.id, readAll_('Dispatches'));
  if (deaths > live) fail_('Deaths cannot be more than live chicks (' + live + ')');
  if (CONFIG.REQUIRE_PHOTOS) {
    if (!r.photoFeed) fail_('Add the feed bags photo');
    if (!r.photoWeight) fail_('Add the weight photo');
    if (deaths > 0 && !r.photoMortality) fail_('Add the mortality photo');
  }
  const rec = {
    id: uid_('r'), batchId: batch.id, farmerId: who.f, date: date, deaths: deaths, feedBags: feed, weight: weight,
    notes: String(r.notes || '').trim().slice(0, 500),
    photoFeed: saveImage_(r.photoFeed, 'feed_' + batch.id),
    photoMortality: saveImage_(r.photoMortality, 'mortality_' + batch.id),
    photoWeight: saveImage_(r.photoWeight, 'weight_' + batch.id),
    submittedAt: nowStr_()
  };
  insert_('Reports', rec);
}

function deleteReport_(req) { remove_('Reports', req.id); }

// ====================== LOADING / DISPATCH ======================
function saveDispatch_(req, who) {
  const d = req.dispatch || {};
  const owner = who.r === 'owner';
  const batch = readAll_('Batches').filter(b => b.id === d.batchId && (owner || b.farmerId === who.f))[0];
  if (!batch) fail_('Batch not found');
  const qty = Math.round(Number(d.quantity));
  if (!(qty > 0)) fail_('Enter the number of birds loaded');
  const live = batch.initial - deathsOf_(batch.id, readAll_('Reports')) - loadedOf_(batch.id, readAll_('Dispatches'));
  if (qty > live) fail_('Only ' + live + ' live birds are left in this batch');
  insert_('Dispatches', {
    id: uid_('D'), batchId: batch.id, farmerId: batch.farmerId, date: /^\d{4}-\d{2}-\d{2}$/.test(d.date || '') ? d.date : todayStr_(),
    quantity: qty, shift: d.shift || '', time: d.time || '', truckNo: String(d.truckNo || '').trim().toUpperCase(),
    challan: String(d.challan || '').trim(), driver: String(d.driver || '').trim(), destination: String(d.destination || '').trim(),
    notes: String(d.notes || '').trim(), createdAt: nowStr_()
  });
}

// ====================== FEED BAGS (assign / request / accept) ======================
function assignBags_(req) {
  const qty = Math.round(Number(req.qty));
  if (!readAll_('Farmers').some(f => f.id === req.farmerId)) fail_('Choose a farmer');
  if (!(qty > 0)) fail_('Enter the number of bags');
  insert_('BagFlow', { id: uid_('bg'), farmerId: req.farmerId, batchId: req.batchId || '', qty: qty, origin: 'owner', status: 'Pending',
    note: '', ownerNote: String(req.note || '').trim(), requestedAt: '', assignedAt: nowStr_(), respondedAt: '' });
}

function requestBags_(req, who) {
  const qty = Math.round(Number(req.qty));
  if (!(qty > 0)) fail_('Enter how many bags you need');
  if (req.batchId && !readAll_('Batches').some(b => b.id === req.batchId && b.farmerId === who.f)) fail_('Batch not found');
  insert_('BagFlow', { id: uid_('bg'), farmerId: who.f, batchId: req.batchId || '', qty: qty, origin: 'farmer', status: 'Requested',
    note: String(req.note || '').trim(), ownerNote: '', requestedAt: nowStr_(), assignedAt: '', respondedAt: '' });
}

function respondRequest_(req) {
  const x = readAll_('BagFlow').filter(b => b.id === req.id)[0];
  if (!x || x.status !== 'Requested') fail_('This request is no longer open');
  if (req.decision === 'approve') {
    const qty = Math.round(Number(req.qty || x.qty));
    if (!(qty > 0)) fail_('Enter the number of bags to send');
    update_('BagFlow', x.id, { qty: qty, status: 'Pending', ownerNote: String(req.note || '').trim(), assignedAt: nowStr_() });
  } else {
    update_('BagFlow', x.id, { status: 'Declined', ownerNote: String(req.note || '').trim(), respondedAt: nowStr_() });
  }
}

function respondAssignment_(req, who) {
  const x = readAll_('BagFlow').filter(b => b.id === req.id && b.farmerId === who.f)[0];
  if (!x || x.status !== 'Pending') fail_('Nothing to respond to');
  update_('BagFlow', x.id, { status: req.decision === 'accept' ? 'Accepted' : 'Rejected', note: String(req.note || x.note || '').trim(), respondedAt: nowStr_() });
}

function cancelBags_(req, who) {
  const x = readAll_('BagFlow').filter(b => b.id === req.id)[0];
  if (!x) fail_('Record not found');
  if (who.r === 'owner') { if (x.status !== 'Pending' && x.status !== 'Requested') fail_('Only open items can be cancelled'); }
  else if (x.farmerId !== who.f || x.status !== 'Requested') fail_('You can only cancel your own open request');
  update_('BagFlow', x.id, { status: 'Cancelled', respondedAt: nowStr_() });
}

// ====================== SAMPLE DATA -> SHEET ======================
/** Writes the sample farmers, batches, daily reports and feed-bag entries into the Sheet (only when empty). */
function seedSampleData(force) {
  ensureSetup_();
  if (readAll_('Farmers').length && force !== true) { Logger.log('Farmers sheet already has data - seed skipped. Run seedSampleDataForce() to add anyway.'); return; }
  const now = nowStr_();
  const farmers = SEED.farmers.map(f => {
    const salt = Utilities.getUuid().slice(0, 8);
    return rowFor_('Farmers', { id: f[0], code: f[1], name: f[2], phone: f[3], village: f[4], status: 'Active', photo: f[5] || '', passHash: hash_(CONFIG.SAMPLE_FARMER_PASSWORD, salt), salt: salt, createdAt: now });
  });
  bulkInsert_('Farmers', farmers);
  bulkInsert_('Batches', SEED.batches.map(b => rowFor_('Batches', { id: b[0], farmerId: b[1], name: b[2], startDate: b[3], initial: b[4], breed: b[5], notes: b[6], status: 'Active', completedAt: '' })));
  bulkInsert_('Reports', SEED.reports.map(r => rowFor_('Reports', { id: r[0], batchId: r[1], farmerId: r[2], date: r[3], deaths: r[4], feedBags: r[5], weight: r[6], notes: r[7], photoFeed: '', photoMortality: '', photoWeight: '', submittedAt: r[3] + 'T18:00:00' })));
  bulkInsert_('BagFlow', SEED.bagflow.map(x => rowFor_('BagFlow', { id: x[0], farmerId: x[1], batchId: x[2], qty: x[3], origin: x[4], status: x[5], note: x[6], ownerNote: x[7], requestedAt: x[8], assignedAt: x[9], respondedAt: x[10] })));
  Logger.log('Sample data written: ' + farmers.length + ' farmers, ' + SEED.batches.length + ' batches, ' + SEED.reports.length + ' reports, ' + SEED.bagflow.length + ' feed-bag entries.');
}
function seedSampleDataForce() { seedSampleData(true); }
