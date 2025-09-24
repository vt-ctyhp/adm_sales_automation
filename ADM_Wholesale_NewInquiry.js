/*** ADM_Wholesale_NewInquiry_DEBUG_v1.gs
 * New Inquiry / Create SO + upload to 05-3D + write product image to cell with deep logging.
 * Requires Script Properties:
 *   ADM_CUSTOMERS_ROOT_FOLDER_ID = <Drive Folder ID that contains per-customer folders>
 * Optional Script Properties:
 *   ADM_SO_TZ = America/Los_Angeles (default)
 *   ADM_DEBUG = "true"|"false"  (when true, verbose logs appear)
 */

/*** === CONFIG === ***/
const MASTER_SHEET = '00_Master Wholesale';
const DROPDOWN_TAB = 'Dropdown'; // (present for parity; not required by this module)

function ADM_PROP(k, d) {
  try { return PropertiesService.getScriptProperties().getProperty(k) || d || ''; }
  catch (_) { return d || ''; }
}
const ADM_TZ = ADM_PROP('ADM_SO_TZ','America/Los_Angeles');
const CUSTOMERS_ROOT_ID = ADM_PROP('ADM_CUSTOMERS_ROOT_FOLDER_ID','');

/*** === SHEET HELPERS === ***/
function ss_(){ return SpreadsheetApp.getActive(); }
function sh_(name){ const s=ss_().getSheetByName(name); if(!s) throw new Error('Missing sheet: '+name); return s; }
function headers_(name){
  const row=sh_(name).getRange(1,1,1,sh_(name).getLastColumn()).getDisplayValues()[0];
  const H={}; row.forEach((h,i)=>{ h=String(h||'').trim(); if(h) H[h]=i+1; }); return H;
}
function headerIndex1_(sh){
  const row = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const m = {}; row.forEach((h,i)=>{ const k=String(h||'').trim(); if(k) m[k]=i+1; }); return m;
}
function setCell_(row, colName, val){
  const H=headers_(MASTER_SHEET); const c=H[colName]; if (!c) return; sh_(MASTER_SHEET).getRange(row,c).setValue(val);
}
function setRich_(row, colName, val, url){
  const H=headers_(MASTER_SHEET); const c=H[colName]; if (!c) return;
  const rng = sh_(MASTER_SHEET).getRange(row,c);
  if (url) {
    const rt = SpreadsheetApp.newRichTextValue().setText(val).setLinkUrl(url).build();
    rng.setRichTextValue(rt);
  } else {
    rng.setValue(val);
  }
}

/*** === SO normalizers === ***/
function soKey_(raw){
  let s = String(raw==null?'':raw).trim().replace(/^'+/, '');
  s = s.replace(/^SO#?/i,'').trim();
  const m = s.match(/^(\d{2})\.(\d{4})$/);
  if (!m) return '';
  return m[1] + m[2];
}
function soPretty_(raw){ const k = soKey_(raw); return k ? (k.slice(0,2) + '.' + k.slice(2)) : ''; }
function soDisplay_(raw){ const p = soPretty_(raw); return p ? ('SO' + p) : ''; }

/*** === SAFE DRIVE HELPERS === ***/
function sanitizeFolderName_(s){
  return String(s||'').replace(/[\\/:*?"<>|]/g,' ').replace(/\s+/g,' ').trim();
}
function ensureFolderChild_(parent, name){
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function ensureCustomerFolder_(customerName){
  if (!CUSTOMERS_ROOT_ID) throw new Error('Script Property ADM_CUSTOMERS_ROOT_FOLDER_ID is not set.');
  const parent = DriveApp.getFolderById(CUSTOMERS_ROOT_ID);
  const name = sanitizeFolderName_(customerName || 'Unknown Customer');
  const it = parent.getFoldersByName(name);
  const f = it.hasNext() ? it.next() : parent.createFolder(name);
  DBG('[ensureCustomerFolder_] → %s (%s)', f.getName(), f.getId());
  return f;
}

/*** ORDER FOLDER  (SO12.3456 — [Product]) + 05-3D ***/
function ensureOrderFolder_(customerFolder, soPretty, product){
  const name = soDisplay_(soPretty) + ' — ' + sanitizeFolderName_(product||'Product');
  let orderFolder = (function(){
    const it = customerFolder.getFoldersByName(name);
    return it.hasNext() ? it.next() : customerFolder.createFolder(name);
  })();
  const need = ['04-Deposit','05-3D','09-ReadyForDelivery','10-Completed'];
  const child = {};
  need.forEach(n => child[n] = ensureFolderChild_(orderFolder, n));
  DBG('[ensureOrderFolder_] %s → 05-3D=%s', orderFolder.getName(), child['05-3D'].getId());
  return {
    orderFolder, orderFolderId: orderFolder.getId(), orderFolderUrl: orderFolder.getUrl(),
    threeDFolderId: child['05-3D'].getId(), threeDFolderUrl: child['05-3D'].getUrl()
  };
}

/** Robust resolver used by upload(): prefers customer/product when provided (no tree-scan). */
function resolve05_3DFolderForSO_(soPretty, customerName, product){
  const pretty = soPretty_(soPretty);
  DBG('[resolve05_3DFolderForSO_] in so=%s customer=%s product=%s', pretty, customerName||'', product||'');
  // If we have customer & product (we do from the client), ensure directly:
  if (customerName && product) {
    const c = ensureCustomerFolder_(customerName);
    const ord = ensureOrderFolder_(c, pretty, product);
    return DriveApp.getFolderById(ord.threeDFolderId);
  }

  // Fallback: scan customers root (rare path).
  const root = DriveApp.getFolderById(CUSTOMERS_ROOT_ID);
  const prefix = soDisplay_(pretty) + ' — ';
  const itCust = root.getFolders(); let found = null;
  while (itCust.hasNext() && !found){
    const cust = itCust.next();
    const itOrders = cust.getFolders();
    while (itOrders.hasNext()){
      const of = itOrders.next();
      if (String(of.getName()||'').indexOf(prefix)===0){
        found = ensureFolderChild_(of, '05-3D');
        break;
      }
    }
  }
  if (found) return found;

  // Last resort: create under a “_Unassigned” container to avoid losing uploads
  const unassigned = ensureFolderChild_(DriveApp.getFolderById(CUSTOMERS_ROOT_ID), '_Unassigned');
  const order = ensureFolderChild_(unassigned, soDisplay_(pretty) + ' — Uploads');
  return ensureFolderChild_(order, '05-3D');
}

/*** === DRIVE → Thumbnail URL (no public share required) === ***/
function driveThumbUrl_(fileId, size){
  try {
    if (!fileId) return '';
    var token = ScriptApp.getOAuthToken();
    var url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
              '?fields=thumbnailLink,mimeType,hasThumbnail&supportsAllDrives=true';
    var res = UrlFetchApp.fetch(url, {
      method:'get', headers:{ Authorization:'Bearer ' + token }, muteHttpExceptions:true
    });
    DBG('[driveThumbUrl_] GET %s → %s', url, res.getResponseCode());
    if (res.getResponseCode() !== 200) { DBG('[driveThumbUrl_] body=%s', res.getContentText()); return ''; }
    var j = {}; try { j = JSON.parse(res.getContentText()); } catch(e){ DBG('[driveThumbUrl_] JSON parse err %s', e); }
    var u = j && j.thumbnailLink || '';
    if (!u) return '';
    var s = Math.max(64, Math.min(1024, Number(size||512)));
    u = u.replace(/=s\d+(?=[-?&]|$)/i, '=s' + s);
    DBG('[driveThumbUrl_] thumb=%s', u);
    return u;
  } catch (e) {
    DBG('[driveThumbUrl_] EX %s', e && (e.stack||e));
    return '';
  }
}

/** Find the most recent image/* in the SO's 05-3D folder */
function findLatestImageIn05_3D_(customerName, soPretty, product){
  try {
    const f05 = resolve05_3DFolderForSO_(soPretty, customerName, product);
    if (!f05) { DBG('[findLatestImageIn05_3D_] no 05-3D resolved'); return null; }
    let it = f05.getFiles(), best=null, bestTime=0;
    while (it.hasNext()){
      const f = it.next();
      if (/^image\//i.test(String(f.getMimeType()||''))) {
        const t = f.getLastUpdated() ? f.getLastUpdated().getTime() : 0;
        if (t > bestTime) { best = f; bestTime = t; }
      }
    }
    if (!best) { DBG('[findLatestImageIn05_3D_] no image/* files'); return null; }
    const fileId = best.getId();
    DBG('[findLatestImageIn05_3D_] picked id=%s', fileId);
    return { fileId, thumbUrl: '' }; // no thumbnail; we embed bytes instead
  } catch (e) {
    DBG('[findLatestImageIn05_3D_] EX %s', e && (e.stack||e));
    return null;
  }
}



/*** === OVER-GRID IMAGE helpers === ***/
function _clearOverGridImagesInCell_(sheet, row, col){
  try {
    var imgs = sheet.getImages ? sheet.getImages() : [];
    for (var i = imgs.length - 1; i >= 0; i--) {
      var img = imgs[i], a = img.getAnchorCell && img.getAnchorCell();
      if (a && a.getRow() === row && a.getColumn() === col) { try { img.remove(); } catch(_){ } }
    }
  } catch(_){}
}
function _embedImageAtCell_(sheet, row, col, fileId, opts){
  opts = opts || {};
  var width  = Number(opts.width  || 150);
  var height = Number(opts.height || 150);
  _clearOverGridImagesInCell_(sheet, row, col);
  var blob = DriveApp.getFileById(fileId).getBlob();
  var img  = sheet.insertImage(blob, col, row);
  try { img.setAnchorCell(sheet.getRange(row, col)); } catch(_){}
  try { img.setWidth(width).setHeight(height); } catch(_){}
  try {
    if (sheet.getColumnWidth(col) < width + 16) sheet.setColumnWidth(col, width + 16);
    if (sheet.getRowHeight(row)   < height + 8) sheet.setRowHeight(row, height + 8);
  } catch(_){}
  DBG('[embedImage] placed over-grid at r=%s c=%s (w=%s h=%s)', row, col, width, height);
  return img;
}

/*** === In-cell image from URL (CellImage) === ***/
function setProductImageCellByUrl_(sheet, row, col, url, alt){
  if (!url) return false;
  try {
    _clearOverGridImagesInCell_(sheet, row, col);
    var cellImg = SpreadsheetApp.newCellImage()
      .setSourceUrl(url)
      .setAltTextTitle(alt || 'Product image')
      .setAltTextDescription('Product image')
      .build();
    sheet.getRange(row, col).setValue(cellImg);
    if (sheet.getRowHeight(row) < 130) sheet.setRowHeight(row, 130);
    if (sheet.getColumnWidth(col) < 130) sheet.setColumnWidth(col, 130);
    DBG('[setProductImageCellByUrl_] OK url=%s r=%s c=%s', url, row, col);
    return true;
  } catch (e) {
    DBG('[setProductImageCellByUrl_] FAIL url=%s → %s', url, e && (e.message||e));
    return false;
  }
}


/*** CellImage from a private Drive file (embed data URL) ***/
function cellImageFromDriveFileId_(fileId, alt){
  if (!fileId) throw new Error('cellImageFromDriveFileId_: missing fileId');
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var mime = String(blob.getContentType() || 'image/png');
  var dataUrl = 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
  return SpreadsheetApp.newCellImage()
    .setSourceUrl(dataUrl)
    .setAltTextTitle(alt || file.getName())
    .setAltTextDescription('Product image')
    .build();
}

// Safe wrapper for Debug logging
function DBG(...args) {
  if (ADM_PROP('ADM_DEBUG', 'false') === 'true') {
    console.log(...args);
  }
}

// Safe wrapper for JSON Debug logging  
function DBJ(label, obj) {
  if (ADM_PROP('ADM_DEBUG', 'false') === 'true') {
    console.log(label, JSON.stringify(obj, null, 2));
  }
}

// Timer helper
function T(label) {
  const start = new Date().getTime();
  return function() {
    const elapsed = new Date().getTime() - start;
    console.log(`[TIMER] ${label}: ${elapsed}ms`);
  };
}

function setProductImageCell_(sh, row, col, fileId, alt){
  var img = cellImageFromDriveFileId_(fileId, alt);
  var rng = sh.getRange(row, col);
  rng.setValue(img);
  try {
    if (sh.getRowHeight(row) < 120) sh.setRowHeight(row, 120);
    if (sh.getColumnWidth(col) < 120) sh.setColumnWidth(col, 120);
  } catch(_){}
  DBG('[setProductImageCell_] OK fileId=%s r=%s c=%s', fileId, row, col);
}


/*** === ODOO text builder (type-aware) === ***/
function admBuildOdooPaste(form){
  DBG('[admBuildOdooPaste] form.type=%s', form && form.productType);
  function linesRing(f){
    const notes = (String(f.notes||'').split(/\r?\n/).map(s=>s.trim()).filter(Boolean)).slice(0,3);
    return [
      '— 3D DESIGN REQUEST — START 3D / CREATE NEW SO —',
      'SETTING',
      '• Accent Diamond: ' + (f.accType||''),
      '• Ring Style    : ' + (f.ringStyle||''),
      '• Metal         : ' + (f.metal||''),
      '• US Size       : ' + (f.ringSize||''),
      '• Band Width    : ' + (f.bandWidth||''),
      '',
      'DESIGN NOTES',
      '• ' + (notes[0]||''),
      '• ' + (notes[1]||''),
      '• ' + (notes[2]||''),
      '',
      'CENTER STONE',
      '• Type          : ' + (f.centerType||''),
      '• Shape         : ' + (f.shape||''),
      '• Dimension     : ' + (f.dimensions||''),
      '',
      '(Mode: Start 3D Design / Create New SO)'
    ].join('\n');
  }
  function kv(label, val){ return '• ' + label.padEnd(12) + ': ' + (val||''); }
  function linesGeneric(title, f, fields){
    const body = fields.map(([lbl, key]) => kv(lbl, f[key]||''));
    return ['— 3D DESIGN REQUEST — START 3D / CREATE NEW SO —', title, ...body, '', '(Mode: Start 3D Design / Create New SO)'].join('\n');
  }

  const t = String(form.productType||'').toLowerCase();
  if (t === 'ring setting' || t === 'ring') return linesRing(form);
  if (t === 'pendant')    return linesGeneric('PENDANT', form, [['Metal','metal'],['Chain Style','chainStyle'],['Length','length'],['Bail Type','bailType'],['Notes','notes']]);
  if (t === 'chain')      return linesGeneric('CHAIN', form,   [['Metal','metal'],['Chain Style','chainStyle'],['Width (mm)','width'],['Length (in)','length'],['Notes','notes']]);
  if (t === 'earrings')   return linesGeneric('EARRINGS', form,[['Metal','metal'],['Type','earringType'],['Back Type','backType'],['Notes','notes']]);
  if (t === 'bracelet')   return linesGeneric('BRACELET', form,[['Metal','metal'],['Style','braceletType'],['Length','length'],['Notes','notes']]);
  return linesGeneric('CUSTOM ITEM', form, [['Metal','metal'],['Description','notes']]);
}

/*** === Upload endpoint (dialog → server) === ***/
function admUploadFile(payload) {
  try {
    console.log('[admUploadFile] Starting, keys:', Object.keys(payload || {}));
    
    if (!payload || !payload.bytesBase64) {
      console.log('[admUploadFile] No file data provided');
      return { ok: false, reason: 'NO_FILE' };
    }
    
    // Decode the base64 data
    const bytes = Utilities.base64Decode(payload.bytesBase64);
    const blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream', payload.filename || 'upload');
    
    // Handle product images specially
    if (payload.isProductImage) {
      console.log('[admUploadFile] Processing as product image');
      
      // Create file in Drive root first (we'll move it later)
      const file = DriveApp.createFile(blob);
      const fileId = file.getId();
      
      console.log('[admUploadFile] Product image created, ID:', fileId);
      
      // Try to move to 05-3D folder if we have the info
      if (payload.so && payload.customerName && payload.product) {
        try {
          const targetFolder = resolve05_3DFolderForSO_(payload.so, payload.customerName, payload.product);
          if (targetFolder) {
            file.moveTo(targetFolder);
            console.log('[admUploadFile] Moved to 05-3D folder');
          }
        } catch (moveErr) {
          console.log('[admUploadFile] Could not move to 05-3D (non-critical):', moveErr.toString());
        }
      }
      
      return {
        ok: true,
        fileId: fileId
      };
    }
    
    // Regular file - upload directly to 05-3D
    const targetFolder = resolve05_3DFolderForSO_(payload.so, payload.customerName, payload.product);
    if (!targetFolder) {
      console.log('[admUploadFile] Could not resolve 05-3D folder');
      return { ok: false, reason: 'NO_05_3D_FOLDER' };
    }
    
    const file = targetFolder.createFile(blob);
    console.log('[admUploadFile] File uploaded to 05-3D, ID:', file.getId());
    
    return {
      ok: true,
      fileId: file.getId()
    };
    
  } catch (e) {
    console.error('[admUploadFile] Error:', e.toString(), e.stack);
    return { ok: false, reason: e.toString() };
  }
}

function testUploadEndpoint() {
  // Test function to verify the upload endpoint is working
  const testPayload = {
    bytesBase64: "SGVsbG8gV29ybGQ=", // "Hello World" in base64
    filename: "test.txt",
    mimeType: "text/plain",
    isProductImage: false
  };
  
  const result = admUploadFile(testPayload);
  console.log('Test upload result:', result);
  return result;
}

function admUploadProductImage(bytesBase64, mimeType, filename) {
  try {
    console.log('[admUploadProductImage] Starting product image upload');
    
    // Create blob from base64 data
    var blob = Utilities.newBlob(
      Utilities.base64Decode(bytesBase64),
      mimeType || 'image/jpeg',
      filename || 'product-image.jpg'
    );
    
    // Create file in Drive root (temporary location)
    var file = DriveApp.createFile(blob);
    var fileId = file.getId();
    
    console.log('[admUploadProductImage] Image saved with ID:', fileId);
    
    // Don't try to get thumbnails or change sharing
    return {
      ok: true,
      fileId: fileId
    };
  } catch (e) {
    console.error('[admUploadProductImage] Error:', e);
    return {
      ok: false,
      error: e.toString()
    };
  }
}

/*** Insert a new top row at 3 and prime it from row 4 ***/
function insertTopRowFromRow4_(sh){
  const lc = sh.getLastColumn();
  sh.insertRowsBefore(3, 1);
  const src  = sh.getRange(4, 1, 1, lc);
  const dest = sh.getRange(3, 1, 1, lc);
  src.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT,          false);
  src.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  src.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMULA,         false);
  return 3;
}
function addDaysYMD_(ymd, days){
  if (!ymd) return '';
  var parts = String(ymd).split('-');
  if (parts.length !== 3) return '';
  var d = new Date(parts[0], parts[1]-1, parts[2], 9, 0, 0);
  d.setDate(d.getDate() + (Number(days)||0));
  return Utilities.formatDate(d, ADM_TZ, 'yyyy-MM-dd');
}

/*** === CRM (01_CRM) — New Customer === ***/
const CRM_SHEET = '01_CRM';
const CRM_HEADERS = [
  'Customer ID','Business Name','Contact Name','Contact Phone','Contact Email',
  'Preferred Contact Method','Street','City','State','ZIP','High Interest Products',
  'Additional Notes','Customer Folder URL','Customer Order Tracker URL','Added On'
];

function ensureCRMTab_(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CRM_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CRM_SHEET);
    sh.getRange(1,1,1,CRM_HEADERS.length).setValues([CRM_HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }
  const H = headerIndex1_(sh);
  let changed = false;
  CRM_HEADERS.forEach(h=>{
    if (!H[h]) { sh.getRange(1, sh.getLastColumn()+1).setValue(h); changed = true; }
  });
  if (changed) SpreadsheetApp.flush();
  return sh;
}

function listCRMCustomers_(){
  const sh = ensureCRMTab_();
  const H = headerIndex1_(sh);
  const col = H['Business Name'] || 0;
  if (!col) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, col, last-1, 1).getDisplayValues()
                 .map(r => String(r[0] || '').trim())
                 .filter(Boolean);
  const seen = Object.create(null), out = [];
  vals.forEach(s => { const k = s.toLowerCase(); if (!seen[k]) { seen[k] = true; out.push(s); } });
  const sorted = out.sort((a,b)=> a.localeCompare(b));
  DBG('[listCRMCustomers_] count=%s', sorted.length);
  return sorted;
}

function _linkUrlFromRichCell_(sh, row, headerLabel){
  const H = headerIndex1_(sh), c = H[headerLabel] || 0;
  if (!c) return '';
  try {
    const rt = sh.getRange(row, c).getRichTextValue();
    if (!rt) return '';
    const direct = (rt.getLinkUrl && rt.getLinkUrl()) || '';
    if (direct) return direct;
    if (rt.getRuns) {
      const runs = rt.getRuns();
      for (var i=0; i<runs.length; i++) {
        const u = runs[i].getLinkUrl && runs[i].getLinkUrl();
        if (u) return u;
      }
    }
  } catch(_){}
  return '';
}

function findOrCreateCustomerWorkbook_(customerFolder, customerName){
  var it = customerFolder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var mt = f.getMimeType();
    if (mt === MimeType.GOOGLE_SHEETS || mt === 'application/vnd.google-apps.spreadsheet') {
      return { id: f.getId(), url: f.getUrl(), existed: true };
    }
  }
  var file = SpreadsheetApp.create('ADM — ' + (customerName || 'Customer'));
  var id = file.getId();
  DriveApp.getFileById(id).moveTo(customerFolder);
  var sh = file.getActiveSheet();
  sh.setName('Quotes & Orders');
  sh.getRange(1, 1, 1, 6).setValues([['Timestamp','SO#','Type','Description','Amount','Status']]);
  return { id: id, url: file.getUrl(), existed: false };
}

function resolveCustomerWorkbookFor_(customerName){
  const crm = ensureCRMTab_();
  const H = headerIndex1_(crm);
  const iName = H['Business Name'] || 0;
  if (!iName) throw new Error('01_CRM missing "Business Name" header.');
  const last = crm.getLastRow();
  if (last < 2) throw new Error('01_CRM has no rows. Add customer first.');

  let row = 0;
  const vals = crm.getRange(2, iName, last-1, 1).getDisplayValues();
  for (let i=0; i<vals.length; i++){
    if (String(vals[i][0]||'').trim() === customerName) { row = i+2; break; }
  }
  if (!row) throw new Error('Customer not found in 01_CRM: ' + customerName);

  const urlRaw  = H['Customer Order Tracker URL'] ? String(crm.getRange(row, H['Customer Order Tracker URL']).getValue()||'').trim() : '';
  const urlRich = H['Customer Orders'] ? _linkUrlFromRichCell_(crm, row, 'Customer Orders') : '';
  let url = urlRaw || urlRich;

  let id = '', existed = true;
  if (!url) {
    const folder = ensureCustomerFolder_(customerName);
    const wb = findOrCreateCustomerWorkbook_(folder, customerName);
    id = wb.id; url = wb.url; existed = wb.existed;

    if (H['Customer Orders']) {
      const rt = SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(url).build();
      crm.getRange(row, H['Customer Orders']).setRichTextValue(rt);
    }
    if (H['Customer Order Tracker URL']) crm.getRange(row, H['Customer Order Tracker URL']).setValue(url);
  } else {
    const m = url.match(/[-\w]{25,}/);
    id = m ? m[0] : '';
  }
  return { id, url, existed };
}

/*** === NEW CUSTOMER DIALOG OPEN === ***/
function admOpenNewCustomerDialog(){
  const t = HtmlService.createTemplateFromFile('dlg_adm_new_customer');
  const html = t.evaluate().setWidth(620).setHeight(560).setTitle('Add New Customer');
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Customer');
}

/*** === NEW CUSTOMER SUBMIT === ***/
function formatPhoneUS_Server_(raw){
  var d = String(raw||'').replace(/\D+/g,'');
  if (d.length === 10) return '('+d.slice(0,3)+') '+d.slice(3,6)+'-'+d.slice(6);
  if (d.length === 11 && d[0]==='1') return '+1 ('+d.slice(1,4)+') '+d.slice(4,7)+'-'+d.slice(7);
  return raw || '';
}
function makeCustomerId_(businessName, stateCode){
  const safeName = String(businessName || '').replace(/[^A-Za-z]/g, '').toUpperCase();
  const prefix   = safeName.slice(0,3).padEnd(3,'X');
  const state    = (String(stateCode||'').replace(/[^A-Za-z]/g,'').toUpperCase() || 'ZZ');
  const randLetters = Array.from({length:2}, ()=> String.fromCharCode(65 + Math.floor(Math.random()*26))).join('');
  const randDigits  = ('0' + Math.floor(Math.random()*100)).slice(-2);
  return prefix + state + '-' + randLetters + randDigits;
}

function admSubmitNewCustomer(payload){
  const stop = T('admSubmitNewCustomer');
  function clean(s){ return String(s||'').trim(); }

  const businessName = clean(payload.businessName);
  const contactName  = clean(payload.contactName);
  const phone        = clean(payload.phone);
  const email        = clean(payload.email);

  if (!businessName) throw new Error('Business Name is required.');
  if (!contactName)  throw new Error('Contact Name is required.');
  if (!phone && !email) throw new Error('Provide at least one: Contact Phone or Contact Email.');

  const phoneStd = formatPhoneUS_Server_(phone);

  const street   = clean(payload.street);
  const city     = clean(payload.city);
  const state    = clean(payload.stateCode || payload.state);
  const zip      = clean(payload.zip);
  const notes    = clean(payload.notes || payload.additionalNotes);

  const pcmArr = (payload.pcm && payload.pcm.length) ? payload.pcm : [];
  const pcm    = pcmArr.join(', ');

  const rawArr  = (payload.interestList || payload.interests || payload.interest || []);
  const asArr   = Array.isArray(rawArr) ? rawArr.slice() : String(rawArr||'').split(',').map(s=>s.trim()).filter(Boolean);
  const otherText = clean(payload.interestOtherText || payload.other);
  if (otherText) asArr.push(otherText);
  const interestCsv = asArr.filter(Boolean).join(', ');

  const crm = ensureCRMTab_();
  const H   = headerIndex1_(crm);

  const customerFolder = ensureCustomerFolder_(businessName);
  const folderUrl = customerFolder.getUrl();

  const stateCodeForId = clean(payload.stateCode || payload.state);
  const customerId = makeCustomerId_(businessName, stateCodeForId);
  const tracker = findOrCreateCustomerWorkbook_(customerFolder, businessName);

  const row = crm.getLastRow() + 1;
  const addedOn = new Date();
  function put(label, val){ if (H[label]) crm.getRange(row, H[label]).setValue(val); }

  put('Customer ID', customerId);
  put('Business Name', businessName);
  put('Contact Name', contactName);
  put('Contact Phone', phoneStd || phone);
  put('Contact Email', email);
  put('Preferred Contact Method', pcm);
  put('Street', street);
  put('City',   city);
  put('State',  state);
  put('ZIP',    zip);
  put('High Interest Products', interestCsv);
  put('Additional Notes', notes);
  put('Added On', addedOn);

  if (H['Customer Folder URL']) {
    const rng = crm.getRange(row, H['Customer Folder URL']);
    const rt = SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(folderUrl).build();
    rng.setRichTextValue(rt);
  }
  if (H['Customer Order Tracker URL'] && tracker && tracker.url) {
    const rng = crm.getRange(row, H['Customer Order Tracker URL']);
    const rt = SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(tracker.url).build();
    rng.setRichTextValue(rt);
  }

  DBJ('[admSubmitNewCustomer] summary', { customerId, folderUrl, trackerUrl: tracker.url });
  if (stop) stop();

  return { ok: true, row, customerId, folderUrl, trackerUrl: (tracker && tracker.url) || '', summary: {
    customerId, businessName, contactName, phone: phoneStd||phone, email,
    address: [street, city, state, zip].filter(Boolean).join(', '), interests: interestCsv
  }};
}

/*** === NEW INQUIRY SUBMIT (creates row + folders + image) === ***/
function admSubmitNewInquiry(payload){
  const stop = T('admSubmitNewInquiry');
  DBJ('[admSubmitNewInquiry] payload', payload);

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MASTER_SHEET);
  if (!sh) throw new Error('Sheet "'+MASTER_SHEET+'" not found.');

  const row = insertTopRowFromRow4_(sh);
  const H   = headerIndex1_(sh);
  function put(label, val){ if (H[label]) sh.getRange(row, H[label]).setValue(val); }

  const customerName = String(payload.customerName||'').trim();
  if (!customerName) throw new Error('Customer Name is required.');

  const soPretty = String(payload.so||'').trim();
  const product  = String(payload.product||'').trim();

  const customerFolder = ensureCustomerFolder_(customerName);
  const order          = ensureOrderFolder_(customerFolder, soPretty, product);
  const workbook       = resolveCustomerWorkbookFor_(customerName);

  const inquiryDate = String(payload.inquiryDate||'').trim();
  if (!inquiryDate) throw new Error('Inquiry Date is required.');
  const deadline3D = addDaysYMD_(inquiryDate, 3);

  const v = (k)=> (payload && payload[k]) || '';
  const metal = String(payload.metal || payload.pd_metal || payload.ch_metal || payload.er_metal || payload.br_metal || '').trim();
  const soUrl = String(v('soUrl')||'').trim();

  if (H['SO#']) {
    const rt = soUrl
      ? SpreadsheetApp.newRichTextValue().setText(soPretty).setLinkUrl(soUrl).build()
      : SpreadsheetApp.newRichTextValue().setText(soPretty).build();
    sh.getRange(row, H['SO#']).setRichTextValue(rt);
  }

  put('Customer Name', customerName);
  put('Product', product);
  put('Inquiry Date', inquiryDate);
  put('Quotation Date', v('quotationDate'));

  put('Sales Stage', 'Lead');
  put('Conversion Status', 'Quotation Requested');
  put('Custom Order Status', '3D Requested');
  put('In Production Status', '');

  put('Product Details', v('odoo'));
  put('Design Details', v('odoo'));
  put('Quantity', v('quantity') || 1);
  put('Gold Type', metal);
  put('Priority Level', v('priorityLevel'));
  put('R&D Deadline', deadline3D);
  put('Customer Order Tracker URL', workbook.url);
  put('Customer Sheet URL', workbook.url);
  put('Customer Folder ID', customerFolder.getId());

  // === Product Image: prefer in-cell embedded image (no sharing required)
    if (H['Product Image']) {
      var col = H['Product Image'];
      var fid = String(v('productImageFileId') || '').trim();
      var ok = false;
      
      console.log('[Image] Starting image processing, fileId:', fid);
      
      // If user didn't upload a Product Image in the dialog, pick the latest from 05-3D
      if (!fid) {
        console.log('[Image] No upload provided, checking 05-3D folder');
        var pick = findLatestImageIn05_3D_(customerName, soPretty, product);
        if (pick && pick.fileId) {
          fid = pick.fileId;
          console.log('[Image] Found existing image in 05-3D:', fid);
        }
      }
      
      if (fid) {
        try {
          // Clear any existing images first
          _clearOverGridImagesInCell_(sh, row, col);
          
          // Try to embed the image directly in the cell
          console.log('[Image] Attempting to embed image in cell');
          var file = DriveApp.getFileById(fid);
          var blob = file.getBlob();
          
          // Create a data URL from the blob for in-cell image
          var dataUrl = 'data:' + blob.getContentType() + ';base64,' + 
                        Utilities.base64Encode(blob.getBytes());
          
          var cellImage = SpreadsheetApp.newCellImage()
            .setSourceUrl(dataUrl)
            .setAltTextTitle(product || 'Product Image')
            .setAltTextDescription('Product image for ' + (product || 'this item'))
            .build();
          
          sh.getRange(row, col).setValue(cellImage);
          
          // Adjust cell size if needed
          if (sh.getRowHeight(row) < 120) sh.setRowHeight(row, 120);
          if (sh.getColumnWidth(col) < 120) sh.setColumnWidth(col, 120);
          
          ok = true;
          console.log('[Image] Successfully embedded image in cell');
          
          // Move the image file to the 05-3D folder if it's still in root
          try {
            var fileObj = DriveApp.getFileById(fid);
            var parents = fileObj.getParents();
            var isInRoot = false;
            while (parents.hasNext()) {
              var parent = parents.next();
              if (parent.getId() === DriveApp.getRootFolder().getId()) {
                isInRoot = true;
                break;
              }
            }
            
            if (isInRoot) {
              console.log('[Image] Moving image from root to 05-3D folder');
              var targetFolder = resolve05_3DFolderForSO_(soPretty, customerName, product);
              if (targetFolder) {
                fileObj.moveTo(targetFolder);
              }
            }
          } catch (moveError) {
            console.log('[Image] Could not move file (non-critical):', moveError);
          }
          
        } catch (e) {
          console.error('[Image] Failed to embed image:', e);
          
          // Fallback: try over-grid image
          try {
            console.log('[Image] Falling back to over-grid image');
            _embedImageAtCell_(sh, row, col, fid, { width: 120, height: 120 });
            ok = true;
          } catch (e2) {
            console.error('[Image] Over-grid fallback also failed:', e2);
          }
        }
      }
      
      // If no image or all attempts failed, clear the cell
      if (!ok) {
        console.log('[Image] No image to display, clearing cell');
        sh.getRange(row, col).setValue('');
      }
      
      console.log('[Image] Processing complete, success:', ok);
    }

  if (stop) stop();
  return {
    ok: true,
    masterRow: row,
    orderFolderUrl: order.orderFolderUrl || '',
    threeDFolderUrl: order.threeDFolderUrl || '',
    customerSheetUrl: workbook.url || ''
  };
}

/*** === CLIENT STATUS UPDATE DIALOG === ***/
function admOpenClientStatusDialog(){
  const html = HtmlService.createHtmlOutputFromFile('dlg_wh_status_update')
    .setWidth(640)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Client Status');
}

function admOpenSalesReportDialog(){
  const html = HtmlService.createHtmlOutputFromFile('dlg_sales_report')
    .setWidth(680)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sales Report');
}

function collectColumnOptions_(label, defaults){
  const sh = sh_(MASTER_SHEET);
  const H = headerIndex1_(sh);
  const col = H[label] || 0;
  const seen = new Set();
  const out = [];
  (defaults || []).forEach(v => {
    const s = String(v || '').trim();
    if (s && !seen.has(s)) { seen.add(s); out.push(s); }
  });
  if (!col) return out;
  const last = sh.getLastRow();
  if (last < 2) return out;
  const vals = sh.getRange(2, col, last-1, 1).getDisplayValues();
  vals.forEach(r => {
    const s = String((r && r[0]) || '').trim();
    if (s && !seen.has(s)) { seen.add(s); out.push(s); }
  });
  return out;
}

function pickHeader_(H, aliases){
  const list = Array.isArray(aliases) ? aliases : [];
  for (let i = 0; i < list.length; i++) {
    const key = String(list[i] || '').trim();
    if (key && H[key]) return H[key];
  }
  return 0;
}

function matchesFilterValue_(cellValue, filterValue){
  const filter = String(filterValue || '').trim();
  if (!filter) return true;
  const cell = String(cellValue == null ? '' : cellValue).trim();
  if (filter === '__BLANK__') return cell === '';
  return cell.toLowerCase() === filter.toLowerCase();
}

function admSalesReportBootstrap(){
  const options = {
    salesStage: collectColumnOptions_('Sales Stage', ['Lead','Quotation Sent','Order Won','Order Lost','In Production']),
    conversionStatus: collectColumnOptions_('Conversion Status', ['Quotation Requested','Quotation Sent','Converted','Lost']),
    customOrderStatus: collectColumnOptions_('Custom Order Status', ['3D Requested','3D In Progress','3D Complete','Production','Shipped']),
    inProductionStatus: collectColumnOptions_('In Production Status', ['Not Started','CAD','Casting','Setting','QA','Completed'])
  };
  return { options };
}

function buildSalesReportRows_(filters){
  const sh = sh_(MASTER_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const H = headerIndex1_(sh);
  const columns = {
    companyId: pickHeader_(H, ['Company ID','Customer ID','Customer (Company) ID','CustomerID','ClientID','Account Code']),
    businessName: pickHeader_(H, ['Business Name','Customer Name']),
    soNumber: pickHeader_(H, ['SO#','SO','Sales Order','Sales Order #']),
    salesStage: H['Sales Stage'] || 0,
    conversionStatus: H['Conversion Status'] || 0,
    customOrderStatus: H['Custom Order Status'] || 0,
    inProductionStatus: H['In Production Status'] || 0,
    orderTotal: pickHeader_(H, ['Order Total','OrderTotal','Total']),
    paidToDate: pickHeader_(H, ['Paid-to-Date','Paid To Date','Paid-To-Date','Paid to Date','Paid']),
    lastUpdatedBy: pickHeader_(H, ['Last Updated By','Updated By','Modified By','Owner']),
    lastUpdatedOn: pickHeader_(H, ['Last Updated On','Updated On','Updated At','Modified At','Last Updated'])
  };

  const getCell = (row, col) => {
    if (!col) return '';
    const val = row[col - 1];
    return val == null ? '' : String(val);
  };

  if (lastRow < 2 || lastCol < 1) {
    return { rows: [], columns };
  }

  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const rows = [];
  const f = filters || {};
  values.forEach(row => {
    if (!matchesFilterValue_(getCell(row, columns.salesStage), f.salesStage)) return;
    if (!matchesFilterValue_(getCell(row, columns.conversionStatus), f.conversionStatus)) return;
    if (!matchesFilterValue_(getCell(row, columns.customOrderStatus), f.customOrderStatus)) return;
    if (!matchesFilterValue_(getCell(row, columns.inProductionStatus), f.inProductionStatus)) return;

    const soRaw = getCell(row, columns.soNumber);
    rows.push([
      getCell(row, columns.companyId),
      getCell(row, columns.businessName),
      soDisplay_(soRaw) || soRaw || '',
      getCell(row, columns.salesStage),
      getCell(row, columns.conversionStatus),
      getCell(row, columns.customOrderStatus),
      getCell(row, columns.inProductionStatus),
      getCell(row, columns.orderTotal),
      getCell(row, columns.paidToDate),
      getCell(row, columns.lastUpdatedBy),
      getCell(row, columns.lastUpdatedOn)
    ]);
  });
  return { rows, columns };
}

function describeFilter_(label, value){
  const v = String(value || '').trim();
  if (!v) return '';
  if (v === '__BLANK__') return label + ': (Blank)';
  return label + ': ' + v;
}

function admGenerateSalesReport(filters){
  const f = filters || {};
  try {
    const { rows } = buildSalesReportRows_(f);
    const headers = [
      'Company ID','Business Name','SO#','Sales Stage','Conversion Status',
      'Custom Order Status','In Production Status','Order Total','Paid-to-Date',
      'Last Updated By','Last Updated On'
    ];

    const tableRows = rows.slice();
    if (!tableRows.length) {
      const empty = new Array(headers.length).fill('');
      empty[0] = 'No rows matched the selected filters.';
      tableRows.push(empty);
    }

    const tz = ADM_TZ || Session.getScriptTimeZone() || 'America/Los_Angeles';
    const timestamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmm');
    const prettyStamp = Utilities.formatDate(new Date(), tz, 'MMMM d, yyyy h:mm a');
    const name = 'Sales Report ' + timestamp;

    const doc = DocumentApp.create(name);
    const body = doc.getBody();

    const POINTS_PER_INCH = 72;
    body.setPageWidth(11 * POINTS_PER_INCH);
    body.setPageHeight(8.5 * POINTS_PER_INCH);
    const margin = 0.5 * POINTS_PER_INCH;
    body.setMarginTop(margin);
    body.setMarginBottom(margin);
    body.setMarginLeft(margin);
    body.setMarginRight(margin);
    body.setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
      [DocumentApp.Attribute.FONT_SIZE]: 10,
      [DocumentApp.Attribute.LINE_SPACING]: 1.15
    });

    const heading = body.appendParagraph('Sales Report')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setSpacingAfter(6);
    heading.editAsText().setForegroundColor('#111827').setFontSize(20).setBold(true);

    body.appendParagraph('Generated: ' + prettyStamp)
      .setFontSize(10)
      .setFontFamily('Arial')
      .setForegroundColor('#4b5563')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    const summaryParts = [
      describeFilter_('Sales Stage', f.salesStage),
      describeFilter_('Conversion Status', f.conversionStatus),
      describeFilter_('Custom Order Status', f.customOrderStatus),
      describeFilter_('In Production Status', f.inProductionStatus)
    ].filter(Boolean);
    const summaryText = summaryParts.length ? summaryParts.join(' | ') : 'None (all records)';
    const metaStyle = {
      [DocumentApp.Attribute.FONT_SIZE]: 10,
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#4b5563'
    };
    body.appendParagraph('Filters: ' + summaryText)
      .setAttributes(metaStyle)
      .setSpacingAfter(2);
    body.appendParagraph('Rows: ' + (rows.length || 0))
      .setAttributes(metaStyle)
      .setSpacingAfter(12);

    const table = body.appendTable([headers].concat(tableRows));
    table.setBorderWidth(0.5);
    table.setBorderColor('#d1d5db');
    const headerRow = table.getRow(0);
    headerRow.setBackgroundColor('#1f2937');
    for (let i = 0; i < headerRow.getNumCells(); i++) {
      const cell = headerRow.getCell(i);
      cell.setBackgroundColor('#1f2937');
      const text = cell.getChild(0);
      if (text && text.editAsText) {
        const t = text.editAsText();
        t.setBold(true).setForegroundColor('#ffffff').setFontFamily('Arial');
      }
      cell.setPaddingTop(6);
      cell.setPaddingBottom(6);
      cell.setPaddingLeft(6);
      cell.setPaddingRight(6);
      if (text && text.getType && text.getType() === DocumentApp.ElementType.PARAGRAPH) {
        text.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    }
    const rightAlign = new Set([7, 8]);
    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (r % 2 === 1) {
        row.setBackgroundColor('#f9fafb');
      }
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        cell.setPaddingTop(6);
        cell.setPaddingBottom(6);
        cell.setPaddingLeft(6);
        cell.setPaddingRight(6);
        const child = cell.getChild(0);
        if (child && child.editAsText) {
          const text = child.editAsText();
          text.setFontSize(10).setForegroundColor('#1f2937').setFontFamily('Arial');
          if (rightAlign.has(c) && child.getType && child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            child.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
          }
        }
      }
    }

    doc.saveAndClose();

    const docId = doc.getId();
    const pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF);
    const folder = ensureFolderChild_(DriveApp.getRootFolder(), 'ADM Reports');
    const pdfFile = folder.createFile(pdfBlob).setName(name + '.pdf');
    DriveApp.getFileById(docId).setTrashed(true);

    return {
      ok: true,
      rowCount: rows.length,
      pdfUrl: pdfFile.getUrl(),
      pdfId: pdfFile.getId(),
      fileName: pdfFile.getName()
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function formatDateYMD_(val){
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, ADM_TZ, 'yyyy-MM-dd');
  }
  const s = String(val || '').trim();
  if (!s) return '';
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return iso[0];
  const d = new Date(s);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, ADM_TZ, 'yyyy-MM-dd');
}

function parseYMD_(ymd){
  if (!ymd) return null;
  const m = String(ymd || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function findMasterRowBySO_(soRaw){
  const key = soKey_(soRaw);
  if (!key) throw new Error('Enter a valid SO number (format 12.3456).');
  const sh = sh_(MASTER_SHEET);
  const H = headerIndex1_(sh);
  const col = H['SO#'] || 0;
  if (!col) throw new Error('00_Master Wholesale is missing an "SO#" column.');
  const last = sh.getLastRow();
  if (last < 2) throw new Error('00_Master Wholesale has no data rows.');
  const vals = sh.getRange(2, col, last-1, 1).getDisplayValues();
  for (let i=0; i<vals.length; i++){
    const cell = (vals[i] && vals[i][0]) || '';
    if (soKey_(cell) === key) {
      const pretty = soPretty_(cell || soRaw);
      const display = soDisplay_(cell || soRaw);
      return { row: i+2, key, soPretty: pretty, soDisplay: display };
    }
  }
  throw new Error('SO not found: ' + soPretty_(soRaw || key));
}

function requireActiveMasterRow_(){
  const ss = ss_();
  const sheet = ss.getActiveSheet();
  if (!sheet || sheet.getName() !== MASTER_SHEET) {
    throw new Error('Select a row in 00_Master Wholesale before opening this dialog.');
  }
  const range = ss.getActiveRange();
  if (!range) {
    throw new Error('Select a row in 00_Master Wholesale before opening this dialog.');
  }
  const row = range.getRow();
  if (row < 2) {
    throw new Error('Select a data row in 00_Master Wholesale before opening this dialog.');
  }
  const last = sheet.getLastRow();
  if (last < 2) {
    throw new Error('00_Master Wholesale has no data rows.');
  }
  if (row > last) {
    throw new Error('The selected row is empty.');
  }
  return row;
}

function buildStatusPayloadFromRow_(row){
  const sh = sh_(MASTER_SHEET);
  const last = sh.getLastRow();
  if (row < 2 || row > last) {
    throw new Error('Select a populated row in 00_Master Wholesale.');
  }
  const H = headerIndex1_(sh);
  const lastCol = sh.getLastColumn();
  const rowRange = sh.getRange(row, 1, 1, lastCol);
  const displayRow = (rowRange.getDisplayValues()[0]) || [];
  const valueRow = (rowRange.getValues()[0]) || [];
  function getDisplay(label){
    const col = H[label] || 0;
    if (!col) return '';
    const val = displayRow[col - 1];
    return val == null ? '' : String(val);
  }
  function getValue(label){
    const col = H[label] || 0;
    if (!col) return '';
    return valueRow[col - 1];
  }
  function getDisplayMulti(labels){
    const list = Array.isArray(labels) ? labels : [labels];
    for (let i = 0; i < list.length; i++) {
      const col = H[list[i]] || 0;
      if (!col) continue;
      const val = displayRow[col - 1];
      if (val != null && String(val).trim() !== '') {
        return String(val);
      }
    }
    return '';
  }
  const soCell = getDisplay('SO#');
  const statuses = {
    salesStage: getDisplay('Sales Stage'),
    conversionStatus: getDisplay('Conversion Status'),
    customOrderStatus: getDisplay('Custom Order Status'),
    inProductionStatus: getDisplay('In Production Status')
  };
  const dates = {
    orderDate: formatDateYMD_(getValue('Order Date')),
    threeDDeadline: formatDateYMD_(getValue('3D Deadline')),
    productionDeadline: formatDateYMD_(getValue('Production Deadline'))
  };
  const contact = {
    email: getDisplayMulti(['Email','Primary Email','Contact Email','Customer Email']),
    phone: getDisplayMulti(['Phone','Phone Number','Contact Phone','Customer Phone'])
  };
  const address = {
    line1: getDisplayMulti(['Address','Address Line 1','Street Address']),
    line2: getDisplayMulti(['Address Line 2','Street','Street 2','Suite/Unit']),
    cityStateZip: getDisplayMulti(['City, State ZIP','City/State/ZIP','City State Zip','City / State / Zip'])
  };
  const assignedRep = getDisplayMulti(['Assigned Rep','Assigned To','Sales Rep','Account Manager']);
  const nextSteps = getDisplayMulti(['Next Steps','Status Notes','Next Step','Order Next Steps']);
  let trackerUrl = '';
  if (H['Customer Order Tracker URL']) {
    trackerUrl = extractUrlFromCell_(rowRange.getCell(1, H['Customer Order Tracker URL']));
  } else if (H['Customer Sheet URL']) {
    trackerUrl = extractUrlFromCell_(rowRange.getCell(1, H['Customer Sheet URL']));
  }
  return {
    ok: true,
    row,
    soDisplay: soDisplay_(soCell || ''),
    soPretty: soPretty_(soCell || ''),
    soNumber: soCell || '',
    statuses,
    dates,
    customerId: getDisplayMulti(['Customer ID','Customer (Company) ID','CustomerID','ClientID','Account Code']),
    businessName: getDisplayMulti(['Business Name']),
    customerName: getDisplayMulti(['Customer Name','Business Name']),
    product: getDisplayMulti(['Product']),
    productDescription: getDisplayMulti(['Product Description','Prod Description','Product','Description','Short Description']),
    trackerUrl,
    contact,
    address,
    assignedRep,
    nextSteps
  };
}

function admStatusUpdateBootstrap(){
  const options = {
    salesStage: collectColumnOptions_('Sales Stage', ['Lead','Quotation Sent','Order Won','Order Lost','In Production']),
    conversionStatus: collectColumnOptions_('Conversion Status', ['Quotation Requested','Quotation Sent','Converted','Lost']),
    customOrderStatus: collectColumnOptions_('Custom Order Status', ['3D Requested','3D In Progress','3D Complete','Production','Shipped']),
    inProductionStatus: collectColumnOptions_('In Production Status', ['Not Started','CAD','Casting','Setting','QA','Completed'])
  };
  const res = {
    options,
    today: Utilities.formatDate(new Date(), ADM_TZ, 'yyyy-MM-dd')
  };
  try {
    const row = requireActiveMasterRow_();
    res.prefill = buildStatusPayloadFromRow_(row);
  } catch (err) {
    res.prefillError = err && err.message ? err.message : String(err);
  }
  return res;
}

function extractUrlFromCell_(range){
  try {
    const rt = range.getRichTextValue();
    if (rt) {
      const direct = rt.getLinkUrl && rt.getLinkUrl();
      if (direct) return direct;
      if (rt.getRuns) {
        const runs = rt.getRuns();
        for (let i=0; i<runs.length; i++){
          const url = runs[i].getLinkUrl && runs[i].getLinkUrl();
          if (url) return url;
        }
      }
    }
  } catch (_) {}
  return range.getDisplayValue();
}

function admFetchStatusForSO(soRaw){
  const found = findMasterRowBySO_(soRaw);
  return buildStatusPayloadFromRow_(found.row);
}

function trackerIdFromUrl_(url){
  const m = String(url || '').match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

function trackerTabName_(soDisplay){
  const base = soDisplay || 'SO';
  return base + ' 3D Tracker';
}

const TRACKER_SF_TZ = 'America/Los_Angeles';
const TRACKER_LOG_HEADERS = [
  'Log Date',
  'Sales Stage',
  'Conversion Status',
  'Custom Order Status',
  'In Production Status',
  'Next Steps',
  'Deadline Type',
  'Deadline Move',
  'Move Count',
  'Assigned Rep',
  'Updated By',
  'Updated At'
];
const TRACKER_LOG_START_COL = 2; // Column B
const TRACKER_TITLE_ROW = 2;
const TRACKER_LEFT_INFO_START_ROW = 4;
const TRACKER_RIGHT_INFO_START_COL = TRACKER_LOG_START_COL + 4; // Column F
const TRACKER_INFO_ROWS = 5;
const TRACKER_LOG_LABEL_ROW = 9;
const TRACKER_LOG_HEADER_ROW = 11;

function trackerFormatYMDForDisplay_(ymd){
  if (!ymd) return '';
  const dt = parseYMD_(ymd);
  if (!dt) return '';
  return Utilities.formatDate(dt, TRACKER_SF_TZ, 'MM/dd/yy');
}

function trackerFormatDeadlineMove_(beforeYMD, afterYMD){
  const beforeDisplay = trackerFormatYMDForDisplay_(beforeYMD);
  const afterDisplay = trackerFormatYMDForDisplay_(afterYMD);
  if (beforeDisplay && afterDisplay) {
    if (beforeDisplay === afterDisplay) return '';
    return beforeDisplay + ' → ' + afterDisplay;
  }
  if (!beforeDisplay && afterDisplay) {
    return 'Set to ' + afterDisplay;
  }
  if (beforeDisplay && !afterDisplay) {
    return 'Cleared (was ' + beforeDisplay + ')';
  }
  return '';
}

function trackerMaybeResetLegacyLayout_(sheet){
  if (!sheet) return;
  const firstCell = String(sheet.getRange(1, 1).getDisplayValue() || '').trim();
  if (firstCell === 'Field') {
    sheet.clear();
  }
}

function trackerEnsureLayout_(sheet, payload){
  trackerMaybeResetLegacyLayout_(sheet);
  const maxNeededCols = TRACKER_LOG_START_COL + TRACKER_LOG_HEADERS.length;
  if (sheet.getMaxColumns() < maxNeededCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxNeededCols - sheet.getMaxColumns());
  }

  const titleRange = sheet.getRange(TRACKER_TITLE_ROW, TRACKER_LOG_START_COL, 1, TRACKER_LOG_HEADERS.length);
  titleRange.breakApart();
  titleRange.merge();
  const businessName = (payload && (payload.businessName || payload.customerName)) || '';
  const soDisplay = (payload && payload.soDisplay) || (payload && payload.soPretty) || '';
  let title = '';
  if (businessName && soDisplay) {
    title = businessName + ' - (' + soDisplay + ' Tracker)';
  } else if (soDisplay) {
    title = soDisplay + ' Tracker';
  } else if (businessName) {
    title = businessName + ' Tracker';
  } else {
    title = 'Customer Tracker';
  }
  titleRange.setValue(title).setFontWeight('bold').setFontSize(16).setHorizontalAlignment('left');

  const leftInfo = [
    ['Company ID:', (payload && payload.customerId) || ''],
    ['Business Name:', businessName || ''],
    ['Address:', (payload && payload.address && payload.address.line1) || ''],
    ['Street:', (payload && payload.address && payload.address.line2) || ''],
    ['City, State ZIP:', (payload && payload.address && payload.address.cityStateZip) || '']
  ];
  sheet.getRange(TRACKER_LEFT_INFO_START_ROW, TRACKER_LOG_START_COL, TRACKER_INFO_ROWS, 2).setValues(leftInfo);
  sheet.getRange(TRACKER_LEFT_INFO_START_ROW, TRACKER_LOG_START_COL).offset(0, 0, TRACKER_INFO_ROWS, 1).setFontWeight('bold');

  const orderDate = (payload && payload.dates && payload.dates.orderDate) || '';
  const rightInfo = [
    ['Order Date:', trackerFormatYMDForDisplay_(orderDate)],
    ['Assigned Rep:', (payload && payload.assignedRep) || ''],
    ['Email:', (payload && payload.contact && payload.contact.email) || ''],
    ['Phone:', (payload && payload.contact && payload.contact.phone) || ''],
    ['SO Number:', (payload && payload.soDisplay) || (payload && payload.soNumber) || '']
  ];
  sheet.getRange(TRACKER_LEFT_INFO_START_ROW, TRACKER_RIGHT_INFO_START_COL, TRACKER_INFO_ROWS, 2).setValues(rightInfo);
  sheet.getRange(TRACKER_LEFT_INFO_START_ROW, TRACKER_RIGHT_INFO_START_COL).offset(0, 0, TRACKER_INFO_ROWS, 1).setFontWeight('bold');

  const logLabelRange = sheet.getRange(TRACKER_LOG_LABEL_ROW, TRACKER_LOG_START_COL, 1, TRACKER_LOG_HEADERS.length);
  logLabelRange.breakApart();
  logLabelRange.merge();
  logLabelRange.setValue('Status Update Log').setFontWeight('bold').setFontSize(12);

  const headerRange = sheet.getRange(TRACKER_LOG_HEADER_ROW, TRACKER_LOG_START_COL, 1, TRACKER_LOG_HEADERS.length);
  headerRange.setValues([TRACKER_LOG_HEADERS]);
  headerRange.setFontWeight('bold').setBackground('#e5e7eb');

  sheet.setFrozenRows(TRACKER_LOG_HEADER_ROW);
  for (let i = 0; i < TRACKER_LOG_HEADERS.length; i++) {
    const col = TRACKER_LOG_START_COL + i;
    try {
      sheet.setColumnWidth(col, i === 0 ? 110 : (i === TRACKER_LOG_HEADERS.length - 1 ? 160 : 150));
    } catch (_) {}
  }
}

function trackerExistingMoveCount_(sheet, type){
  if (!type) return 0;
  const last = sheet.getLastRow();
  if (last <= TRACKER_LOG_HEADER_ROW) return 0;
  const countRange = sheet.getRange(TRACKER_LOG_HEADER_ROW + 1, TRACKER_LOG_START_COL + 6, last - TRACKER_LOG_HEADER_ROW, 1).getDisplayValues();
  let count = 0;
  for (let i = 0; i < countRange.length; i++) {
    const cell = String((countRange[i] && countRange[i][0]) || '').trim();
    if (cell === type) count++;
  }
  return count;
}

function trackerBuildLogRows_(sheet, beforePayload, afterPayload, userEmail, now){
  const tz = TRACKER_SF_TZ;
  const rows = [];
  const afterStatuses = (afterPayload && afterPayload.statuses) || {};
  const afterDates = (afterPayload && afterPayload.dates) || {};
  const beforeDates = (beforePayload && beforePayload.dates) || {};
  const assignedRep = (afterPayload && afterPayload.assignedRep) || (beforePayload && beforePayload.assignedRep) || '';
  const nextSteps = (afterPayload && afterPayload.nextSteps) || (beforePayload && beforePayload.nextSteps) || '';
  const logDate = Utilities.formatDate(now, tz, 'MM/dd/yy');
  const updatedAt = Utilities.formatDate(now, tz, 'MM/dd/yy hh:mm a');
  const baseRow = [
    logDate,
    afterStatuses.salesStage || '',
    afterStatuses.conversionStatus || '',
    afterStatuses.customOrderStatus || '',
    afterStatuses.inProductionStatus || '',
    nextSteps || '',
    '',
    '',
    '',
    assignedRep || '',
    userEmail || '',
    updatedAt
  ];

  const deadlineChanges = [];
  if ((beforeDates.threeDDeadline || afterDates.threeDDeadline) && beforeDates.threeDDeadline !== afterDates.threeDDeadline) {
    deadlineChanges.push({
      type: '3D Deadline',
      before: beforeDates.threeDDeadline || '',
      after: afterDates.threeDDeadline || ''
    });
  }
  if ((beforeDates.productionDeadline || afterDates.productionDeadline) && beforeDates.productionDeadline !== afterDates.productionDeadline) {
    deadlineChanges.push({
      type: 'Production Deadline',
      before: beforeDates.productionDeadline || '',
      after: afterDates.productionDeadline || ''
    });
  }

  const existingCounts = {};
  const pendingCounts = {};
  function nextMoveCount(type){
    if (!type) return '';
    if (!(type in existingCounts)) {
      existingCounts[type] = trackerExistingMoveCount_(sheet, type);
    }
    const prior = existingCounts[type];
    const pending = pendingCounts[type] || 0;
    const next = prior + pending + 1;
    pendingCounts[type] = pending + 1;
    return next;
  }

  if (deadlineChanges.length === 0) {
    rows.push(baseRow);
  } else {
    deadlineChanges.forEach(change => {
      const row = baseRow.slice();
      row[6] = change.type;
      row[7] = trackerFormatDeadlineMove_(change.before, change.after);
      row[8] = row[7] ? nextMoveCount(change.type) : '';
      rows.push(row);
    });
  }

  return rows;
}

function updateCustomerTrackerSheet_(trackerUrl, beforePayload, afterPayload, userEmail){
  const id = trackerIdFromUrl_(trackerUrl);
  if (!id) return '';
  const ss = SpreadsheetApp.openById(id);
  const display = (afterPayload && afterPayload.soDisplay) || (beforePayload && beforePayload.soDisplay) || '';
  const preferredName = trackerTabName_(display);
  const legacyName = display || 'SO';
  let sheet = ss.getSheetByName(preferredName);
  if (!sheet && legacyName && legacyName !== preferredName) {
    sheet = ss.getSheetByName(legacyName);
    if (sheet) {
      try { sheet.setName(preferredName); }
      catch (_) {}
    }
  }
  if (!sheet) {
    sheet = ss.insertSheet(preferredName);
  }

  const infoPayload = afterPayload || beforePayload || {};
  trackerEnsureLayout_(sheet, infoPayload);

  const now = new Date();
  const rows = trackerBuildLogRows_(sheet, beforePayload, afterPayload, userEmail, now);
  if (rows.length) {
    const startRow = Math.max(sheet.getLastRow() + 1, TRACKER_LOG_HEADER_ROW + 1);
    sheet.getRange(startRow, TRACKER_LOG_START_COL, rows.length, TRACKER_LOG_HEADERS.length).setValues(rows);
    sheet.getRange(startRow, TRACKER_LOG_START_COL, rows.length, TRACKER_LOG_HEADERS.length).setBorder(true, true, true, true, false, false);
  }

  return sheet.getName();
}

function summarizeStatusChange_(arr, label, before, after){
  const prev = String(before == null ? '' : before).trim();
  const next = String(after == null ? '' : after).trim();
  if (prev === next) return;
  const display = (val) => {
    const s = String(val == null ? '' : val).trim();
    return s || '—';
  };
  arr.push({
    label,
    from: display(before),
    to: display(after)
  });
}

function buildStatusSummary_(before, after, row){
  const summary = {
    title: (after && after.soDisplay) || (before && before.soDisplay) || (row ? ('Row ' + row) : 'Selected Order'),
    details: [],
    changes: []
  };
  const detailSource = after || before || {};
  const details = [];
  if (detailSource.businessName) {
    details.push(detailSource.businessName);
  } else if (detailSource.customerName) {
    details.push(detailSource.customerName);
  }
  if (detailSource.customerId) {
    details.push('Customer ID: ' + detailSource.customerId);
  }
  if (detailSource.productDescription) {
    details.push(detailSource.productDescription);
  } else if (detailSource.product) {
    details.push(detailSource.product);
  }
  summary.details = details;

  const changes = [];
  if (before && before.statuses && after && after.statuses) {
    summarizeStatusChange_(changes, 'Sales Stage', before.statuses.salesStage, after.statuses.salesStage);
    summarizeStatusChange_(changes, 'Conversion Status', before.statuses.conversionStatus, after.statuses.conversionStatus);
    summarizeStatusChange_(changes, 'Custom Order Status', before.statuses.customOrderStatus, after.statuses.customOrderStatus);
    summarizeStatusChange_(changes, 'In Production Status', before.statuses.inProductionStatus, after.statuses.inProductionStatus);
  }
  if (before && before.dates && after && after.dates) {
    summarizeStatusChange_(changes, 'Order Date', before.dates.orderDate, after.dates.orderDate);
    summarizeStatusChange_(changes, '3D Deadline', before.dates.threeDDeadline, after.dates.threeDDeadline);
    summarizeStatusChange_(changes, 'Production Deadline', before.dates.productionDeadline, after.dates.productionDeadline);
  }
  summary.changes = changes;
  return summary;
}

function admSubmitStatusUpdate(payload){
  const stop = T('admSubmitStatusUpdate');
  if (!payload || !payload.row) throw new Error('Active row is required.');
  const row = Number(payload.row);
  if (!row || row < 2) throw new Error('Active row is required.');
  const sh = sh_(MASTER_SHEET);
  const last = sh.getLastRow();
  if (row > last) throw new Error('Select a populated row in 00_Master Wholesale.');
  const H = headerIndex1_(sh);

  const before = buildStatusPayloadFromRow_(row);
  const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'user';

  function put(label, value){
    if (H[label]) sh.getRange(row, H[label]).setValue(value || '');
  }
  function putDate(label, ymd){
    if (!H[label]) return;
    const rng = sh.getRange(row, H[label]);
    if (!ymd) {
      rng.clearContent();
      return;
    }
    const dt = parseYMD_(ymd);
    if (dt) {
      rng.setValue(dt);
    } else {
      rng.setValue(ymd);
    }
  }
  put('Sales Stage', payload.salesStage || '');
  put('Conversion Status', payload.conversionStatus || '');
  put('Custom Order Status', payload.customOrderStatus || '');
  put('In Production Status', payload.inProductionStatus || '');
  putDate('Order Date', payload.orderDate || '');
  putDate('3D Deadline', payload.threeDDeadline || '');
  putDate('Production Deadline', payload.productionDeadline || '');

  const after = buildStatusPayloadFromRow_(row);
  const trackerUrl = after.trackerUrl || '';
  let sheetName = '';

  if (trackerUrl) {
    try {
      sheetName = updateCustomerTrackerSheet_(
        trackerUrl,
        before,
        after,
        userEmail
      );
    } catch (err) {
      console.error('[admSubmitStatusUpdate] tracker update failed:', err);
    }
  }

  const summary = buildStatusSummary_(before, after, row);

  if (stop) stop();
  return { ok: true, trackerUrl, sheetName, summary };
}
