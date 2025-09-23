/*** ADM WHOLESALE — Record Payments (Multi‑SO + Templates + Filing + Credits + Supersedes)
 * Version: 2.1.0  (2025‑09‑23) — with ADM_DEBUG tracing
 */

const SP = PropertiesService.getScriptProperties();

// ============================= CONFIG =============================
const WH_LEDGER_TAB_NAME         = SP.getProperty('WH_LEDGER_TAB_NAME') || '400_payments ledger';
const WH_MASTER_TAB_NAME         = SP.getProperty('WH_MASTER_TAB_NAME') || '00_Master Wholesale';
const WH_DEFAULT_SHIP_PER_ORDER  = num_(SP.getProperty('WH_DEFAULT_SHIP_PER_ORDER'), 50);
const WH_SHIP_THRESHOLD_SUBTOTAL = num_(SP.getProperty('WH_SHIP_THRESHOLD_SUBTOTAL'), 2000);
const WH_DOC_PREFIX              = (SP.getProperty('WH_DOC_PREFIX') || 'ADM').replace(/\s+/g,'').toUpperCase();
const WH_DOCS_ENABLED            = String(SP.getProperty('WH_DOCS_ENABLED') || 'true').toLowerCase() === 'true';
const WH_DOCS_FOLDER_FALLBACK_ID = (SP.getProperty('WH_DOCS_FOLDER_FALLBACK_ID') || '').trim();

const WH_LEDGER_SCAN_START_ROW   = num_(SP.getProperty('WH_LEDGER_SCAN_START_ROW'), 0);
const WH_LEDGER_SCAN_WINDOW      = Math.max(200, num_(SP.getProperty('WH_LEDGER_SCAN_WINDOW'), 800));

// Which tabs hold wholesale orders (SO rows)
const WH_ORDERS_TAB_NAMES = (SP.getProperty('WH_ORDERS_TAB_NAMES_CSV') || '')
  .split(',').map(s=>s.trim()).filter(Boolean);

// Aliases for headers (include your exact labels)
const SO_ALIASES        = pickList_(SP.getProperty('WH_SO_COL_ALIASES'),        ['SO#','SO','Sales Order','Sales Order #']);
const SOFOLDER_ALIASES  = pickList_(SP.getProperty('WH_SOFOLDER_COL_ALIASES'),  ['SO Folder ID','SO Folder URL','Folder URL']);

const CUSTID_ALIASES    = pickList_(SP.getProperty('WH_CUSTID_COL_ALIASES'),
  ['Customer (Company) ID','Customer ID','CustomerID','ClientID','Account Code']);

const COMPANY_ALIASES   = pickList_(SP.getProperty('WH_COMPANY_COL_ALIASES'),
  ['Company Name','Business Name','Customer','Customer Name']);

const CONTACT_ALIASES   = pickList_(SP.getProperty('WH_CONTACT_COL_ALIASES'),
  ['Contact Name','Contact']);

const ADDRESS_ALIASES   = pickList_(SP.getProperty('WH_ADDRESS_ONE_LINE_ALIASES'),
  ['Business Address','Business Address (one line)','Address','Company Address']);

const SALES_STAGE_ALIASES = pickList_(SP.getProperty('WH_SALES_STAGE_COL_ALIASES'),
  ['Sales Stage','Stage','Status']);

const TRACKER_ALIASES   = pickList_(SP.getProperty('WH_TRACKER_URL_COL_ALIASES'),
  ['Customer Order Tracker URL','Order Tracker URL','Tracker URL']);

const MASTER_PTD_ALIASES = ['Paid-to-Date','Paid To Date','Paid-To-Date','Paid to Date','Paid'];
const MASTER_OT_ALIASES  = ['Order Total','OrderTotal','Total'];
const MASTER_RB_ALIASES  = ['Remaining Balance','Balance','RB'];

const CRM_ID_ALIASES       = ['Customer ID','CustomerID','Customer (Company) ID'];
const CRM_COMPANY_ALIASES  = ['Business Name','Company Name'];
const CRM_CONTACT_ALIASES  = ['Contact Name','Primary Contact'];
const CRM_PHONE_ALIASES    = ['Contact Phone','Phone'];
const CRM_EMAIL_ALIASES    = ['Contact Email','Email'];
const CRM_STREET_ALIASES   = ['Street','Address 1','Addr 1'];
const CRM_CITY_ALIASES     = ['City'];
const CRM_STATE_ALIASES    = ['State','ST'];
const CRM_ZIP_ALIASES      = ['ZIP','Postal Code','Zip'];
const CRM_TERMS_ALIASES    = ['Payment Terms','Terms'];
const CRM_TRACKER_ALIASES  = ['Customer Order Tracker URL','Tracker URL'];
const CRM_FOLDER_ALIASES   = ['Customer Folder URL','Folder URL'];

// For auto line descriptions
const PRODUCT_DESC_ALIASES = pickList_(SP.getProperty('WH_PRODUCT_DESC_ALIASES'),
  ['Product Description','Prod Description','Product','Description','Short Description']);

// (Optional) address parts fallback if one‑line address not present
const STREET_ALIASES = ['Street','Address 1','Addr 1','Address Line 1'];
const CITY_ALIASES   = ['City','Town'];
const STATE_ALIASES  = ['State','ST','Province'];
const ZIP_ALIASES    = ['Zip','ZIP','Postal','Postal Code'];

// Templates (Google Docs)
const TPL = {
  DI: (SP.getProperty('TEMPLATE_DEPOSIT_INVOICE_ADM') || '').trim(),
  DR: (SP.getProperty('TEMPLATE_DEPOSIT_RECEIPT_ADM') || '').trim(),
  SI: (SP.getProperty('TEMPLATE_SALES_INVOICE_ADM')   || '').trim(),
  SR: (SP.getProperty('TEMPLATE_SALES_RECEIPT_ADM')   || '').trim()
};

const LEDGER_HEADERS = [
  'PaymentID','TransactionID','InvoiceGroupID','DocNumber','DocType','DocStatus',
  'SupersedesDoc#','SupersedeAction','AppliesToDoc#',
  'CustomerID','CompanyName','ContactName','Address',
  'SO_List','SOsCSV','PrimarySO',
  'AllocationMode','AllocationsJSON',
  'LinesJSON','LinesSubtotal','ShippingJSON','ShippingTotal',
  'DOC_DATE','DueDate',
  'PaymentDateTime','AmountGross','Method','Reference','Notes',
  'FeePercent','FeeFlat','FeeAmount','AmountNet',
  'PDF_URL','DOC_URL','PrimarySO_FolderID','PrimarySO_FolderURL','ShortcutIDs_CSV',
  'CustomerOrderTrackerURL',
  'SubmittedBy','SubmittedAt'
];

let _feeRulesCache = null;

function feeMethodKey_(method){
  return String(method||'').trim().toLowerCase().replace(/[^a-z0-9]+/g,' ').replace(/\s+/g,' ').trim();
}

function loadFeeRules_(){
  if (_feeRulesCache) return _feeRulesCache;
  const props = SP.getProperties();
  function pickNumber(names, fallback){
    for (const name of names){
      if (props[name] !== undefined) {
        const val = num_(props[name], NaN);
        if (isFinite(val)) return val;
      }
    }
    return fallback;
  }
  const defaultPct  = pickNumber(['WH_FEE_PERCENT_DEFAULT','WH_FEE_DEFAULT_PERCENT','WH_FEE_PCT_DEFAULT'], 0);
  const defaultFlat = pickNumber(['WH_FEE_FLAT_DEFAULT','WH_FEE_DEFAULT_FLAT'], 0);
  const rules = { __default: { pct: defaultPct, flat: defaultFlat } };

  const rawJson = props.WH_FEE_RULES_JSON || props.WH_FEE_METHODS_JSON || '';
  if (rawJson) {
    try {
      const parsed = JSON.parse(rawJson);
      const assign = (method, obj)=>{
        const key = feeMethodKey_(method);
        if (!key) return;
        const pct = (obj && obj.pct !== undefined) ? num_(obj.pct, defaultPct) :
                    (obj && obj.percent !== undefined) ? num_(obj.percent, defaultPct) :
                    (obj && obj.rate !== undefined) ? num_(obj.rate, defaultPct) : defaultPct;
        const flat = (obj && obj.flat !== undefined) ? num_(obj.flat, defaultFlat) :
                     (obj && obj.fee !== undefined) ? num_(obj.fee, defaultFlat) : defaultFlat;
        rules[key] = { pct, flat };
      };
      if (Array.isArray(parsed)) {
        parsed.forEach(item => {
          if (!item) return;
          const method = item.method || item.name || item.type || (Array.isArray(item) ? item[0] : '');
          const pct = item.pct ?? item.percent ?? item.rate ?? (Array.isArray(item) ? item[1] : undefined);
          const flat = item.flat ?? item.fee ?? (Array.isArray(item) ? item[2] : undefined);
          assign(method, { pct, flat });
        });
      } else if (parsed && typeof parsed === 'object') {
        if (parsed.default && typeof parsed.default === 'object') {
          const d = parsed.default;
          rules.__default = {
            pct: num_(d.pct ?? d.percent ?? d.rate, defaultPct),
            flat: num_(d.flat ?? d.fee, defaultFlat)
          };
        }
        if (parsed.methods && typeof parsed.methods === 'object') {
          Object.keys(parsed.methods).forEach(method => assign(method, parsed.methods[method]));
        }
        Object.keys(parsed).forEach(method => {
          if (method === 'default' || method === 'methods') return;
          assign(method, parsed[method]);
        });
      }
    } catch(err) {
      dbg('loadFeeRules_: parse error', String(err));
    }
  }

  Object.keys(props).forEach(key => {
    if (key.startsWith('WH_FEE_PCT_')) {
      const method = key.slice('WH_FEE_PCT_'.length).replace(/_/g,' ');
      const rule = rules[feeMethodKey_(method)] || { pct: defaultPct, flat: defaultFlat };
      rule.pct = num_(props[key], defaultPct);
      rules[feeMethodKey_(method)] = rule;
    }
    if (key.startsWith('WH_FEE_FLAT_')) {
      const method = key.slice('WH_FEE_FLAT_'.length).replace(/_/g,' ');
      const rule = rules[feeMethodKey_(method)] || { pct: defaultPct, flat: defaultFlat };
      rule.flat = num_(props[key], defaultFlat);
      rules[feeMethodKey_(method)] = rule;
    }
  });

  Object.keys(rules).forEach(key => {
    const rule = rules[key];
    if (!rule) return;
    rule.pct = isFinite(rule.pct) ? rule.pct : defaultPct;
    rule.flat = isFinite(rule.flat) ? rule.flat : defaultFlat;
  });

  _feeRulesCache = rules;
  return _feeRulesCache;
}

function wh_calcFeeForMethod_(method, amount){
  const amt = num_(amount,0);
  const rules = loadFeeRules_();
  const key = feeMethodKey_(method);
  const rule = (key && rules[key]) || rules.__default || { pct:0, flat:0 };
  const pct = isFinite(rule.pct) ? rule.pct : 0;
  const flat = isFinite(rule.flat) ? rule.flat : 0;
  const fee = amt > 0 ? round2_( (amt * (pct/100)) + flat ) : 0;
  const net = round2_(Math.max(0, amt - fee));
  return { pct, flat, fee, net };
}

function wh_feeRulesForUi_(){
  const rules = loadFeeRules_();
  const out = {
    default: { pct: rules.__default ? rules.__default.pct || 0 : 0, flat: rules.__default ? rules.__default.flat || 0 : 0 },
    methods: {}
  };
  Object.keys(rules).forEach(key => {
    if (key === '__default') return;
    const rule = rules[key] || {};
    out.methods[key] = { pct: rule.pct || 0, flat: rule.flat || 0 };
  });
  return out;
}

// ============================= INIT / CONTEXT =============================
function wh_init(){
  const ctx = readActiveContext_();
  const debugEnabled = ADM_isDebug();
  const out = {
    nowIso: new Date().toISOString(),
    ctx,
    defaults: {
      shipPerOrder: WH_DEFAULT_SHIP_PER_ORDER,
      shipThresholdSubtotal: WH_SHIP_THRESHOLD_SUBTOTAL,
      docPrefix: WH_DOC_PREFIX
    },
    feeRules: wh_feeRulesForUi_(),
    debug: debugEnabled ? { enabled: true, props: { defaults: { shipPerOrder: WH_DEFAULT_SHIP_PER_ORDER, shipThresholdSubtotal: WH_SHIP_THRESHOLD_SUBTOTAL } } } : { enabled: false }
  };
  dbg('wh_init ->', out);
  return out;
}

/** Try active sheet first; if missing fields, fall back to finding the SO on orders tabs. */
function readActiveContext_(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const rng = sh.getActiveRange();
  if (!rng) { dbg('readActiveContext_: no active range'); return blankCtx_(); }
  const row = rng.getRow();
  if (row <= 1) { dbg('readActiveContext_: active row is header'); return blankCtx_(); }

  const hdr = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn()))
    .getDisplayValues()[0].map(s=>String(s||'').trim());
  const H = hIndex_(hdr);
  dbg('readActiveContext_: active sheet + headers', {sheet: sh.getName(), headers: hdr});

  const soCol   = pickH_(H, SO_ALIASES);
  const cidCol  = pickH_(H, CUSTID_ALIASES);
  const compCol = pickH_(H, COMPANY_ALIASES);
  const contact = pickH_(H, CONTACT_ALIASES);
  const addrCol = pickH_(H, ADDRESS_ALIASES);
  const trkCol  = pickH_(H, TRACKER_ALIASES);

  dbg('readActiveContext_: alias matches', {soCol, cidCol, compCol, contact, addrCol, trkCol});

  const rowVals = sh.getRange(row,1,1,sh.getLastColumn()).getDisplayValues()[0];

  // One-line address; else compose
  let address = addrCol ? String(rowVals[addrCol-1]||'').trim() : '';
  if (!address) {
    const cStreet = pickH_(H, STREET_ALIASES), cCity = pickH_(H, CITY_ALIASES),
          cState  = pickH_(H, STATE_ALIASES), cZip  = pickH_(H, ZIP_ALIASES);
    if (cStreet || cCity || cState || cZip) {
      const street = cStreet?String(rowVals[cStreet-1]||'').trim():'';
      const city   = cCity?String(rowVals[cCity-1]||'').trim():'';
      const state  = cState?String(rowVals[cState-1]||'').trim():'';
      const zip    = cZip?String(rowVals[cZip-1]||'').trim():'';
      address = [street, [city, state].filter(Boolean).join(', '), zip].filter(Boolean).join(', ');
    }
  }

  let ctx = {
    sheetName: sh.getName(),
    rowIndex: row,
    soNumber:   soCol ? String(rowVals[soCol-1]||'').trim() : '',
    customerId: cidCol ? String(rowVals[cidCol-1]||'').trim() : '',
    companyName:compCol? String(rowVals[compCol-1]||'').trim() : '',
    contactName:contact? String(rowVals[contact-1]||'').trim() : '',
    address:    address,
    trackerUrl: trkCol ? String(rowVals[trkCol-1]||'').trim() : '',
    contactPhone:'',
    contactEmail:'',
    paymentTerms:'',
    customerFolderUrl:''
  };
  dbg('readActiveContext_: initial ctx from active sheet', ctx);

  // Fallback via orders tabs if something is still blank but we have a primary SO
  if (!ctx.customerId || !ctx.companyName || !ctx.contactName || !ctx.address || !ctx.trackerUrl) {
    const extra = findOrdersRowBySO_(ctx.soNumber);
    if (extra) {
      ctx = {
        ...ctx,
        customerId: ctx.customerId || extra.customerId || '',
        companyName: ctx.companyName || extra.companyName || '',
        contactName: ctx.contactName || extra.contactName || '',
        address: ctx.address || extra.address || '',
        trackerUrl: ctx.trackerUrl || extra.trackerUrl || ''
      };
      dbg('readActiveContext_: filled blanks from orders tabs via SO', {so: ctx.soNumber, filled: ctx});
    } else {
      dbg('readActiveContext_: no orders‑tab fallback found for SO', ctx.soNumber);
    }
  }

  const crmDetails = loadCrmDetailsById_(ctx.customerId);
  if (crmDetails) {
    ctx = {
      ...ctx,
      companyName: crmDetails.businessName || ctx.companyName,
      contactName: crmDetails.contactName || ctx.contactName,
      address: crmDetails.addressOneLine || ctx.address,
      trackerUrl: crmDetails.trackerUrl || ctx.trackerUrl,
      contactPhone: crmDetails.contactPhone || ctx.contactPhone,
      contactEmail: crmDetails.contactEmail || ctx.contactEmail,
      paymentTerms: crmDetails.paymentTerms || ctx.paymentTerms,
      customerFolderUrl: crmDetails.customerFolderUrl || ctx.customerFolderUrl
    };
  }

  return ctx;
}

function blankCtx_(){ return { sheetName:'', rowIndex:0, soNumber:'', customerId:'', companyName:'', contactName:'', address:'', trackerUrl:'', contactPhone:'', contactEmail:'', paymentTerms:'', customerFolderUrl:'' }; }

function loadCrmDetailsById_(customerId){
  const id = String(customerId||'').trim();
  if (!id) return null;
  if (typeof ensureCRMTab_ !== 'function') return null;
  let sh;
  try { sh = ensureCRMTab_(); } catch(_) { return null; }
  if (!sh) return null;
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return null;
  const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
  const H = hIndex_(hdr);
  const cId = pickH_(H, CRM_ID_ALIASES);
  if (!cId) return null;

  const cCompany = pickH_(H, CRM_COMPANY_ALIASES);
  const cContact = pickH_(H, CRM_CONTACT_ALIASES);
  const cPhone   = pickH_(H, CRM_PHONE_ALIASES);
  const cEmail   = pickH_(H, CRM_EMAIL_ALIASES);
  const cStreet  = pickH_(H, CRM_STREET_ALIASES);
  const cCity    = pickH_(H, CRM_CITY_ALIASES);
  const cState   = pickH_(H, CRM_STATE_ALIASES);
  const cZip     = pickH_(H, CRM_ZIP_ALIASES);
  const cTerms   = pickH_(H, CRM_TERMS_ALIASES);
  const cTracker = pickH_(H, CRM_TRACKER_ALIASES);
  const cFolder  = pickH_(H, CRM_FOLDER_ALIASES);

  const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();
  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    if (String(row[cId-1]||'').trim() !== id) continue;
    const rowIndex = i+2;
    function get(col){ return col ? String(row[col-1]||'').trim() : ''; }

    const street = get(cStreet);
    const city   = get(cCity);
    const state  = get(cState);
    const zip    = get(cZip);

    let locality = '';
    if (city && state) locality = `${city}, ${state}`;
    else if (city) locality = city;
    else if (state) locality = state;
    if (zip) locality = locality ? `${locality} ${zip}` : zip;
    const addressOneLine = [street, locality].filter(Boolean).join(', ');

    const trackerUrl = getLinkFromCell_(sh, rowIndex, cTracker) || get(cTracker);
    const folderUrl  = getLinkFromCell_(sh, rowIndex, cFolder) || get(cFolder);

    return {
      row: rowIndex,
      businessName: get(cCompany),
      contactName: get(cContact),
      contactPhone: get(cPhone),
      contactEmail: get(cEmail),
      street, city, state, zip,
      addressOneLine,
      paymentTerms: get(cTerms),
      trackerUrl,
      customerFolderUrl: folderUrl
    };
  }
  return null;
}

function getLinkFromCell_(sh, row, col){
  if (!sh || !col) return '';
  try {
    const cell = sh.getRange(row, col);
    const rt = cell.getRichTextValue && cell.getRichTextValue();
    if (rt) {
      const direct = (rt.getLinkUrl && rt.getLinkUrl()) || '';
      if (direct) return direct;
      if (rt.getRuns) {
        const runs = rt.getRuns();
        for (let i=0;i<runs.length;i++) {
          const u = runs[i].getLinkUrl && runs[i].getLinkUrl();
          if (u) return u;
        }
      }
    }
    const raw = cell.getValue();
    const rawStr = String(raw||'').trim();
    if (rawStr) {
      const m = rawStr.match(/https?:\/\/\S+/);
      return m ? m[0] : rawStr;
    }
    const display = String(cell.getDisplayValue()||'').trim();
    if (display) {
      const m2 = display.match(/https?:\/\/\S+/);
      return m2 ? m2[0] : display;
    }
  } catch(_) {}
  return '';
}

/** Find a single row on the configured orders tabs by SO number and return key fields. */
function findOrdersRowBySO_(soNumber){
  const so = String(soNumber||'').trim(); if (!so) return null;
  const ss = SpreadsheetApp.getActive();
  const tabs = WH_ORDERS_TAB_NAMES.length ? WH_ORDERS_TAB_NAMES : ss.getSheets().map(s=>s.getName());
  dbg('findOrdersRowBySO_: scanning tabs', tabs);

  for (const name of tabs) {
    const sh = ss.getSheetByName(name); if (!sh) continue;
    const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr<2) continue;
    const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
    const H = hIndex_(hdr);

    const cSO  = pickH_(H, SO_ALIASES);
    const cCID = pickH_(H, CUSTID_ALIASES);
    const cCMP = pickH_(H, COMPANY_ALIASES);
    const cCON = pickH_(H, CONTACT_ALIASES);
    const cADR = pickH_(H, ADDRESS_ALIASES);
    const cTRK = pickH_(H, TRACKER_ALIASES);
    const cPD  = pickH_(H, PRODUCT_DESC_ALIASES);

    if (!cSO) continue;
    const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();
    for (let i=0;i<vals.length;i++){
      const r = vals[i];
      const s = String(r[cSO-1]||'').trim();
      if (!s) continue;
      if (soEq_(s, so)) {
        const out = {
          sheet: name, rowIndex: i+2,
          customerId: cCID ? String(r[cCID-1]||'').trim() : '',
          companyName: cCMP ? String(r[cCMP-1]||'').trim() : '',
          contactName: cCON ? String(r[cCON-1]||'').trim() : '',
          address:     cADR ? String(r[cADR-1]||'').trim() : '',
          trackerUrl:  cTRK ? String(r[cTRK-1]||'').trim() : '',
          productDesc: cPD  ? String(r[cPD-1]||'').trim() : ''
        };
        dbg('findOrdersRowBySO_: match', out);
        return out;
      }
    }
  }
  return null;
}

// ============================= LOOKUPS =============================
function wh_listSOsForCustomer(customerId, limit){
  customerId = String(customerId||'').trim();
  const ss = SpreadsheetApp.getActive();
  const tabNames = WH_ORDERS_TAB_NAMES.length ? WH_ORDERS_TAB_NAMES : ss.getSheets().map(s=>s.getName());
  dbg('wh_listSOsForCustomer: args', {customerId, limit, tabNames});

  const out = [];
  for (const name of tabNames) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;
    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr < 2) continue;
    const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
    const H = hIndex_(hdr);

    const cSO    = pickH_(H, SO_ALIASES);
    const cCID   = pickH_(H, CUSTID_ALIASES);
    const cPD    = pickH_(H, PRODUCT_DESC_ALIASES);
    const cStage = pickH_(H, SALES_STAGE_ALIASES);
    dbg('wh_listSOsForCustomer: header map', {tab:name, cSO, cCID, cPD, cStage});

    if (!cSO) continue;
    const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();

    for (let i=0;i<vals.length;i++){
      const r = vals[i];
      const so = String(r[cSO-1]||'').trim();
      if (!so) continue;

      if (customerId) {
        const id = cCID ? String(r[cCID-1]||'').trim() : '';
        if (id && id !== customerId) continue;
      }

      if (name === WH_MASTER_TAB_NAME && cStage) {
        const stage = String(r[cStage-1]||'').trim().toLowerCase();
        if (stage === 'won') continue;
      }

      out.push({
        soNumber: so,
        customerId: cCID ? String(r[cCID-1]||'').trim() : '',
        productDesc: cPD ? String(r[cPD-1]||'').trim() : '',
        sheet: name,
        rowIndex: i+2
      });

      if (limit && out.length >= limit) {
        dbg('wh_listSOsForCustomer: hit limit', out.length);
        return out;
      }
    }
  }
  dbg('wh_listSOsForCustomer: result size', out.length);
  return out;
}

/** Helper for UI: if customerId empty, still return at least the primary SO row (for desc prefill). */
function wh_getKnownSOs(customerId, primarySO){
  if (String(customerId||'').trim()) return wh_listSOsForCustomer(customerId, 250);
  const found = findOrdersRowBySO_(primarySO);
  return found ? [{ soNumber: String(primarySO||''), productDesc: found.productDesc||'', sheet: found.sheet, rowIndex: found.rowIndex, customerId: found.customerId||'' }] : [];
}

function wh_findSoFolderId(soNumber){
  soNumber = String(soNumber||'').trim(); if (!soNumber) return '';
  const ss = SpreadsheetApp.getActive();
  const tabNames = WH_ORDERS_TAB_NAMES.length ? WH_ORDERS_TAB_NAMES : ss.getSheets().map(s=>s.getName());
  for (const name of tabNames){
    const sh = ss.getSheetByName(name); if (!sh) continue;
    const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) continue;
    const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
    const H = hIndex_(hdr);
    const cSO = pickH_(H, SO_ALIASES); if (!cSO) continue;
    const cF  = pickH_(H, SOFOLDER_ALIASES); if (!cF) continue;
    const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();
    for (let i=0;i<vals.length;i++){
      const r = vals[i];
      const so = String(r[cSO-1]||'').trim();
      if (!so || !soEq_(so, soNumber)) continue;
      const raw = String(r[cF-1]||'').trim();
      const id  = fileIdFromUrl_(raw);
      if (id) return id;
    }
  }
  return '';
}

// ============================= SUBMIT =============================
function wh_submitPayment(payload){
  dbg('wh_submitPayment: payload', payload);
  const lock = LockService.getDocumentLock(); lock.waitLock(25*1000);
  try {
    ensureLedger_();
    require_(payload,'payload');

    const dt = String(payload.docType||'').toUpperCase().trim();
    let docKind = '', docFlavor = '';
    if (dt==='DEPOSIT_INVOICE')      { docKind='INVOICE'; docFlavor='DEPOSIT'; }
    else if (dt==='DEPOSIT_RECEIPT') { docKind='RECEIPT'; docFlavor='DEPOSIT'; }
    else if (dt==='SALES_INVOICE')   { docKind='INVOICE'; docFlavor='SALES';   }
    else if (dt==='SALES_RECEIPT')   { docKind='RECEIPT'; docFlavor='SALES';   }
    else if (dt==='CREDIT')          { docKind='CREDIT';  docFlavor='SALES';   }
    else throw new Error('Unsupported Document Type.');

    const docTypeLabelMap = {
      'DEPOSIT_INVOICE': 'Deposit Invoice',
      'DEPOSIT_RECEIPT': 'Deposit Receipt',
      'SALES_INVOICE': 'Sales Invoice',
      'SALES_RECEIPT': 'Sales Receipt',
      'CREDIT': 'Credit'
    };
    const docTypeLabel = docTypeLabelMap[dt] || dt;

    const customerId  = mustStr_(payload.customerId, 'customerId');
    const companyName = (payload.companyName||'').trim();
    const contactName = (payload.contactName||'').trim();
    const address     = (payload.address||'').trim();

    const primarySO = mustStr_(payload.primarySO, 'primarySO');
    const soCandidates = Array.isArray(payload.soList) ? payload.soList : [];
    const seenSOs = new Set();
    const soList = [];
    function pushSo(raw){
      const so = String(raw||'').trim();
      if (!so) return;
      const key = normSoKey_(so);
      if (!key || seenSOs.has(key)) return;
      seenSOs.add(key);
      soList.push(so);
    }
    pushSo(primarySO);
    soCandidates.forEach(pushSo);
    if (!soList.length) throw new Error('Select/type at least one SO.');

    const docDate = payload.docDateISO ? new Date(payload.docDateISO) : new Date();
    let dueDate = '';
    if (docKind === 'INVOICE' && payload.includeDueDate) {
      dueDate = payload.dueDateISO ? new Date(payload.dueDateISO) : addDays_(docDate, 2);
    }

    // Lines
    let lines = Array.isArray(payload.lines) ? payload.lines : [];
    if (!lines.length) {
      const catalog = wh_listSOsForCustomer(customerId, 999);
      const mapPD = {}; catalog.forEach(r => mapPD[String(r.soNumber)] = r.productDesc || '');
      lines = soList.map(so => ({ so, desc: mapPD[so]||'', qty: 1, amt: 0 }));
    } else {
      lines = lines.map(n => ({ so: String(n.so||primarySO), desc: String(n.desc||'').trim(), qty: num_(n.qty,0), amt: num_(n.amt,0) }))
                   .filter(l=> l.so );
    }
    const linesSubtotal = round2_(lines.reduce((s,l)=> s + (l.qty*l.amt), 0));
    dbg('wh_submitPayment: lines built', {lines, linesSubtotal});

    // Shipping (sales invoice only)
    let shipping = Array.isArray(payload.shipping) ? payload.shipping.map(s=>({ label: String(s.label||'Shipping').trim(), amount: num_(s.amount,0) })) : [];
    let shippingTotal = 0;
    if (docKind==='INVOICE' && docFlavor==='SALES' && payload.addShipping) {
      if (!shipping.length) {
        const defaultShip = (linesSubtotal < WH_SHIP_THRESHOLD_SUBTOTAL) ? (soList.length * WH_DEFAULT_SHIP_PER_ORDER) : 0;
        if (defaultShip>0) shipping = [{ label:'Shipping', amount: defaultShip }];
      }
      shippingTotal = round2_(shipping.reduce((s,x)=>s+num_(x.amount,0),0));
    } else { shipping=[]; shippingTotal=0; }
    dbg('wh_submitPayment: shipping', {shipping, shippingTotal});

    // Receipt allocation
    const pmt = (payload.pmt||{});
    pmt.amount = round2_(num_(pmt.amount, 0));
    pmt.method = String(pmt.method||'').trim();
    pmt.reference = String(pmt.reference||'').trim();
    pmt.notes = String(pmt.notes||'').trim();
    let allocationMode = '', allocMap = {};
    if (docKind==='RECEIPT') {
      if (!(pmt.amount>0)) throw new Error('Payment amount is required for receipts.');
      const evenSplit = (payload.evenSplit!==false);
      allocationMode = evenSplit ? 'EVEN' : 'MANUAL';
      if (evenSplit) {
        const per = round2_(pmt.amount / soList.length);
        soList.forEach(so => allocMap[so] = per);
        const sumEven = Object.values(allocMap).reduce((a,b)=>a+b,0);
        allocMap[soList[soList.length-1]] = round2_(allocMap[soList[soList.length-1]] + (pmt.amount - sumEven));
      } else {
        (payload.allocations||[]).forEach(a => {
          const so = String(a.so||'').trim();
          if (!so) return;
          if (!seenSOs.has(normSoKey_(so))) return;
          allocMap[so] = round2_(num_(a.amount,0));
        });
        const manualSum = Object.values(allocMap).reduce((sum,val)=>sum + num_(val,0),0);
        const diff = Math.abs(round2_(manualSum) - pmt.amount);
        if (diff > 0.01) {
          throw new Error('Manual allocations must total the payment amount. Off by $' + diff.toFixed(2));
        }
      }
      soList.forEach(so => { if (!(so in allocMap)) allocMap[so] = 0; });
    }
    dbg('wh_submitPayment: allocations', {allocationMode, allocMap, pmtAmount: pmt.amount});

    // IDs
    const transactionId = newTransactionId_(customerId, new Date());
    const invoiceGroupId = (soList.length>1) ? newInvoiceGroupId_(customerId, new Date()) : '';
    const docNumber = (String(payload.docNumberOverride||'').trim() || newDocNumber_());
    dbg('wh_submitPayment: ids', {transactionId, invoiceGroupId, docNumber});

    // Template selection
    const tplId = (docKind==='INVOICE' && docFlavor==='DEPOSIT') ? TPL.DI :
                  (docKind==='RECEIPT' && docFlavor==='DEPOSIT') ? TPL.DR :
                  (docKind==='INVOICE' && docFlavor==='SALES')   ? TPL.SI :
                  (docKind==='RECEIPT' && docFlavor==='SALES')   ? TPL.SR : '';
    if (WH_DOCS_ENABLED && !tplId) throw new Error('Missing ADM template id for this document type.');
    dbg('wh_submitPayment: template chosen', {docKind, docFlavor, tplId});

    // Build document
    let docUrl='', pdfUrl='', primaryFolderId='', primaryFolderUrl='', shortcutIds=[];
    if (WH_DOCS_ENABLED) {
      primaryFolderId = wh_findSoFolderId(primarySO) || WH_DOCS_FOLDER_FALLBACK_ID || '';
      if (primaryFolderId) {
        try {
          primaryFolderUrl = DriveApp.getFolderById(primaryFolderId).getUrl();
        } catch(e) {
          dbg('wh_submitPayment: primary folder lookup failed', { primarySO, error: String(e) });
          primaryFolderUrl = '';
        }
      }
      const model = {
        docKind, docFlavor, customerId, companyName, contactName, address,
        transactionId, invoiceGroupId, docNumber,
        soList, primarySO, lines, linesSubtotal, shipping, shippingTotal,
        docDate, dueDate, pmt
      };
      const out = wh_buildDocFromTemplate_(tplId, model, primaryFolderId);
      docUrl = out.docUrl; pdfUrl = out.pdfUrl;
      dbg('wh_submitPayment: doc built', out);

      const others = soList.filter(so => so !== primarySO);
      if (others.length && out.pdfId && Drive && Drive.Files) {
        others.forEach(so=>{
          const fid = wh_findSoFolderId(so);
          if (fid) {
            try {
              const shortcut = Drive.Files.insert({
                title: DriveApp.getFileById(out.pdfId).getName(),
                mimeType: MimeType.SHORTCUT,
                parents: [{ id: fid }],
                shortcutDetails: { targetId: out.pdfId }
              }, undefined, { supportsAllDrives:true });
              if (shortcut && shortcut.id) shortcutIds.push(shortcut.id);
            } catch(e) { dbg('Drive shortcut insert failed', {so, error: String(e)}); }
          }
        });
      }
    }

    // Ledger write
    const sh = ensureLedger_();
    const H  = headerMap_(sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getDisplayValues()[0]);
    const row = new Array(sh.getLastColumn()).fill('');
    setIf_(row,H,'PaymentID', 'PAY-' + Utilities.getUuid());
    setIf_(row,H,'TransactionID', transactionId);
    setIf_(row,H,'InvoiceGroupID', invoiceGroupId);
    setIf_(row,H,'DocNumber', docNumber);
    setIf_(row,H,'DocType', dt);
    setIf_(row,H,'DocStatus', 'ISSUED');

    const supersedesDocNumber = String(payload.supersedesDocNumber||'').trim();
    const supersedeAction     = String(payload.supersedeAction||'').trim().toUpperCase();
    if (supersedesDocNumber) {
      setIf_(row,H,'SupersedesDoc#', supersedesDocNumber);
      setIf_(row,H,'SupersedeAction', supersedeAction || 'REPLACE');
      wh_markSuperseded_(supersedesDocNumber, supersedeAction || 'REPLACE');
    }

    setIf_(row,H,'CustomerID', customerId);
    setIf_(row,H,'CompanyName', companyName);
    setIf_(row,H,'ContactName', contactName);
    setIf_(row,H,'Address', address);

    setIf_(row,H,'SO_List', JSON.stringify(soList));
    setIf_(row,H,'SOsCSV', soList.join(', '));
    setIf_(row,H,'PrimarySO', primarySO);
    if (docKind==='RECEIPT') setIf_(row,H,'AllocationMode', allocationMode);
    if (docKind==='RECEIPT' && Object.keys(allocMap).length) setIf_(row,H,'AllocationsJSON', JSON.stringify(allocMap));

    setIf_(row,H,'LinesJSON', JSON.stringify(lines));
    setIf_(row,H,'LinesSubtotal', linesSubtotal);
    setIf_(row,H,'ShippingJSON', JSON.stringify(shipping));
    setIf_(row,H,'ShippingTotal', shippingTotal);

    setIf_(row,H,'DOC_DATE', docDate);
    if (dueDate) setIf_(row,H,'DueDate', dueDate);

    const feeInfo = wh_calcFeeForMethod_(pmt.method, pmt.amount||0);
    const feePct = feeInfo.pct || 0;
    const feeFlat = feeInfo.flat || 0;
    const feeAmt = feeInfo.fee || 0;
    const amtNet = feeInfo.net || round2_(pmt.amount||0);
    setIf_(row,H,'PaymentDateTime', (docKind==='RECEIPT') ? (pmt.dateTimeISO ? new Date(pmt.dateTimeISO) : new Date()) : '');
    setIf_(row,H,'AmountGross', pmt.amount||'');
    setIf_(row,H,'Method', pmt.method||'');
    setIf_(row,H,'Reference', pmt.reference||'');
    setIf_(row,H,'Notes', (pmt.notes||payload.notes||''));

    setIf_(row,H,'FeePercent', feePct);
    setIf_(row,H,'FeeFlat', feeFlat);
    setIf_(row,H,'FeeAmount', feeAmt);
    setIf_(row,H,'AmountNet', amtNet);

    setIf_(row,H,'PDF_URL', pdfUrl);
    setIf_(row,H,'DOC_URL', docUrl);
    setIf_(row,H,'PrimarySO_FolderID', primaryFolderId);
    setIf_(row,H,'PrimarySO_FolderURL', primaryFolderUrl);
    setIf_(row,H,'ShortcutIDs_CSV', (shortcutIds||[]).join(','));
    setIf_(row,H,'CustomerOrderTrackerURL', (payload.trackerUrl||''));

    setIf_(row,H,'SubmittedBy', Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'user');
    setIf_(row,H,'SubmittedAt', new Date());
    sh.appendRow(row);
    dbg('wh_submitPayment: appended ledger row');

    try { if (payload.trackerUrl) mirrorTrackerUrl_(customerId, payload.trackerUrl); } catch(e){ dbg('mirrorTrackerUrl error', String(e)); }

    if (docKind==='RECEIPT') {
      wh_applyReceiptToOrders_(allocMap, soList);
    }

    let overage = 0;
    if (docKind==='RECEIPT') {
      const allocSum = Object.values(allocMap).reduce((a,b)=>a+b,0);
      if (pmt.amount > allocSum) {
        overage = round2_(pmt.amount - allocSum);
        writeCreditRow_(customerId, companyName, overage, pmt.method, transactionId, invoiceGroupId);
      }
    }

    return {
      ok: true,
      docType: dt,
      docTypeLabel,
      transactionId,
      invoiceGroupId,
      docNumber,
      pdfUrl,
      docUrl,
      folderUrl: primaryFolderUrl,
      primaryFolderId,
      shortcutIds,
      overageCredit: overage
    };
  } finally { try{ lock.releaseLock(); }catch(_){} }
}

function wh_markSuperseded_(docNumber, action){
  const sh = ensureLedger_(); const lr=sh.getLastRow(), lc=sh.getLastColumn(); if (lr<2) return;
  const H = headerMap_(sh.getRange(1,1,1,lc).getDisplayValues()[0]);
  const vals = sh.getRange(2,1,lr-1,lc).getValues();
  let changed=false;
  for (let i=0;i<vals.length;i++){
    const r = vals[i];
    if (String(r[H['DocNumber']-1]||'').trim() === docNumber) {
      r[H['DocStatus']-1] = (action==='VOID' ? 'VOID' : 'REPLACED');
      vals[i] = r; changed=true;
    }
  }
  if (changed) sh.getRange(2,1,lr-1,lc).setValues(vals);
}

// ============================= SUMMARY =============================
function wh_getSummary(params){
  const scope = String(params.scope||'SO').toUpperCase();
  const sh = ensureLedger_(); const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { scope, items:[], totals:{} };

  const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
  const H = headerMap_(hdr);
  const vals = sh.getRange(2,1,lr-1,lc).getValues();

  const items = [];
  for (let i=0;i<vals.length;i++){
    const r = vals[i]; const o={}; Object.keys(H).forEach(k=>o[k]=r[H[k]-1]);
    const match = scope==='SO'       ? (String(o.SOsCSV||'').split(',').map(s=>s.trim()).includes(String(params.soNumber||'').trim()))
                : scope==='CUSTOMER' ? (String(o.CustomerID||'').trim() === String(params.customerId||'').trim())
                : scope==='GROUP'    ? (String(o.InvoiceGroupID||'').trim() === String(params.invoiceGroupId||'').trim())
                : false;
    if (match) items.push(o);
  }
  const receipts = items.filter(it => String(it.DocType||'').toUpperCase()==='RECEIPT')
                        .reduce((s,it)=> s + num_(it.AmountGross,0),0);
  const credits  = items.filter(it => String(it.DocType||'').toUpperCase()==='CREDIT')
                        .reduce((s,it)=> s + num_(it.AmountGross,0),0);
  const applied  = items.filter(it => String(it.DocType||'').toUpperCase()==='CREDIT-APPLIED')
                        .reduce((s,it)=> s + num_(it.AmountGross,0),0);
  return {
    scope, items,
    totals: {
      receipts: round2_(receipts),
      creditsIssued: round2_(credits),
      creditsApplied: round2_(applied),
      creditUnappliedEstimate: round2_(credits - applied)
    }
  };
}

// ============================= CREDIT APPLY =============================
function wh_applyCreditNow(payload){
  const customerId = mustStr_(payload.customerId, 'customerId');
  const applyList = (payload.apply||[]).map(a=>({so:String(a.so||'').trim(), amount:num_(a.amount,0)}))
                                       .filter(a=>a.so && a.amount>0);
  if (!applyList.length) throw new Error('Nothing to apply.');
  const avail = getUnappliedCredit_(customerId);
  if (avail < applyList.reduce((s,a)=>s+a.amount,0)) throw new Error('Not enough credit.');

  const sh = ensureLedger_(); const H = headerMap_(sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0]);
  const row = new Array(sh.getLastColumn()).fill('');
  setIf_(row,H,'PaymentID','CRED-APPLY-'+Utilities.getUuid().slice(0,8));
  setIf_(row,H,'DocType','CREDIT-APPLIED');
  setIf_(row,H,'CustomerID', customerId);
  setIf_(row,H,'SOsCSV', applyList.map(x=>x.so+':'+x.amount).join(', '));
  setIf_(row,H,'AmountGross', applyList.reduce((s,a)=>s+a.amount,0));
  setIf_(row,H,'SubmittedAt', new Date());
  sh.appendRow(row);
  return { ok:true };
}

function getUnappliedCredit_(customerId){
  const sh = ensureLedger_(); const lr=sh.getLastRow(), lc=sh.getLastColumn();
  if (lr<2) return 0; const H = headerMap_(sh.getRange(1,1,1,lc).getDisplayValues()[0]);
  const vals = sh.getRange(2,1,lr-1,lc).getValues();
  let credit=0, applied=0;
  for (let i=0;i<vals.length;i++){
    const r=vals[i];
    if (String(r[H['CustomerID']-1]||'').trim()!==customerId) continue;
    const t = String(r[H['DocType']-1]||'').toUpperCase();
    if (t==='CREDIT') credit += num_(r[H['AmountGross']-1],0);
    if (t==='CREDIT-APPLIED') applied += num_(r[H['AmountGross']-1],0);
  }
  return Math.max(0, round2_(credit - applied));
}

function writeCreditRow_(customerId, companyName, amount, method, transactionId, invoiceGroupId){
  const sh = ensureLedger_(); const H = headerMap_(sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0]);
  const row = new Array(sh.getLastColumn()).fill('');
  setIf_(row,H,'PaymentID', 'CRED-' + Utilities.getUuid().slice(0,8));
  setIf_(row,H,'DocType','CREDIT');
  setIf_(row,H,'CustomerID', customerId);
  setIf_(row,H,'CompanyName', companyName);
  setIf_(row,H,'AmountGross', amount);
  setIf_(row,H,'Method', method||'');
  setIf_(row,H,'TransactionID', transactionId||'');
  setIf_(row,H,'InvoiceGroupID', invoiceGroupId||'');
  setIf_(row,H,'FeePercent', 0);
  setIf_(row,H,'FeeFlat', 0);
  setIf_(row,H,'FeeAmount', 0);
  setIf_(row,H,'AmountNet', amount);
  setIf_(row,H,'SubmittedAt', new Date());
  sh.appendRow(row);
}

// ============================= RECEIPT SUMS & MIRROR =============================
function ledgerScanValues_(sh, lc){
  const last = sh.getLastRow();
  if (last < 2) return [];
  let start = 2;
  if (WH_LEDGER_SCAN_START_ROW && WH_LEDGER_SCAN_START_ROW >= 2 && WH_LEDGER_SCAN_START_ROW <= last) {
    start = WH_LEDGER_SCAN_START_ROW;
  } else if (WH_LEDGER_SCAN_WINDOW && WH_LEDGER_SCAN_WINDOW > 0) {
    start = Math.max(2, last - WH_LEDGER_SCAN_WINDOW + 1);
  }
  return sh.getRange(start, 1, last - start + 1, lc).getValues();
}

function parseAllocationsJSON_(raw){
  if (!raw && raw !== 0) return {};
  let parsed = raw;
  if (typeof raw === 'string') {
    const txt = raw.trim();
    if (!txt) return {};
    try { parsed = JSON.parse(txt); }
    catch(_) { return {}; }
  }
  const out = {};
  if (Array.isArray(parsed)) {
    parsed.forEach(item => {
      if (!item) return;
      if (Array.isArray(item)) {
        const so = item[0];
        const amount = item[1];
        if (so !== undefined) out[String(so)] = num_(amount,0);
      } else if (typeof item === 'object') {
        const so = item.so || item.SO || item.id || item.name;
        const amount = item.amount ?? item.value ?? item.net ?? item.amt;
        if (so !== undefined) out[String(so)] = num_(amount,0);
      }
    });
  } else if (parsed && typeof parsed === 'object') {
    Object.keys(parsed).forEach(key => {
      const value = parsed[key];
      if (value === null || value === undefined) return;
      if (typeof value === 'object' && !Array.isArray(value)) {
        const so = value.so || value.SO || key;
        const amount = value.amount ?? value.value ?? value.net ?? value.amt;
        out[String(so)] = num_(amount,0);
      } else {
        out[String(key)] = num_(value,0);
      }
    });
  }
  return out;
}

function extractSoEntries_(row, H){
  const entries = [];
  const seen = new Set();
  const pushVal = (val)=>{
    const so = String(val||'').trim();
    if (!so) return;
    const key = normSoKey_(so);
    if (!key || seen.has(key)) return;
    seen.add(key);
    entries.push({ key, value: so });
  };

  if (H['SO_List']) {
    const raw = row[H['SO_List']-1];
    if (Array.isArray(raw)) {
      raw.forEach(pushVal);
    } else if (typeof raw === 'string') {
      const txt = raw.trim();
      if (txt) {
        if (txt.startsWith('[')) {
          try {
            const arr = JSON.parse(txt);
            if (Array.isArray(arr)) arr.forEach(pushVal);
          } catch(_) {}
        } else {
          txt.split(/[,\n]/).forEach(pushVal);
        }
      }
    }
  }

  if (!entries.length && H['SOsCSV']) {
    const csv = String(row[H['SOsCSV']-1]||'');
    csv.split(',').forEach(pushVal);
  }

  if (!entries.length && H['PrimarySO']) {
    pushVal(row[H['PrimarySO']-1]);
  }

  return entries;
}

function computeAllocationsForRow_(row, H, targetKeys){
  const allowed = targetKeys ? new Set(Array.from(targetKeys).map(k=>normSoKey_(k))) : null;
  const entries = extractSoEntries_(row, H);
  if (!entries.length) return {};
  const relevant = allowed ? entries.filter(e=>allowed.has(e.key)) : entries;
  if (!relevant.length) return {};

  const rawAlloc = H['AllocationsJSON'] ? row[H['AllocationsJSON']-1] : '';
  const allocParsed = parseAllocationsJSON_(rawAlloc);
  const out = {};
  Object.keys(allocParsed).forEach(so => {
    const key = normSoKey_(so);
    if (!key) return;
    if (allowed && !allowed.has(key)) return;
    out[key] = round2_(num_(allocParsed[so],0));
  });

  if (!Object.keys(out).length) {
    const amountNet = H['AmountNet'] ? num_(row[H['AmountNet']-1],0) : 0;
    const amountGross = num_(row[H['AmountGross']-1],0);
    const base = amountNet || amountGross;
    if (!(base>0)) return {};
    const shareTargets = relevant;
    let distributed = 0;
    const per = shareTargets.length ? round2_(base / shareTargets.length) : 0;
    shareTargets.forEach((entry, idx)=>{
      let val = per;
      if (idx === shareTargets.length-1) {
        val = round2_(base - distributed);
      }
      distributed = round2_(distributed + val);
      out[entry.key] = round2_(val);
    });
  }

  return out;
}

function wh_collectReceiptSums_(targetKeys){
  const sh = ensureLedger_();
  const lc = sh.getLastColumn();
  if (lc < 1) return {};
  const rows = ledgerScanValues_(sh, lc);
  if (!rows.length) return {};
  const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
  const H = headerMap_(hdr);
  const allowed = targetKeys ? new Set(Array.from(targetKeys).map(k=>normSoKey_(k))) : null;
  const sums = {};
  rows.forEach(row => {
    const docType = String(row[H['DocType']-1]||'').toUpperCase();
    if (!docType || docType.indexOf('RECEIPT') === -1) return;
    const status = H['DocStatus'] ? String(row[H['DocStatus']-1]||'').toUpperCase() : 'ISSUED';
    if (status === 'VOID' || status === 'REPLACED') return;
    const allocs = computeAllocationsForRow_(row, H, allowed);
    Object.keys(allocs).forEach(key => {
      if (allowed && !allowed.has(key)) return;
      sums[key] = round2_((sums[key]||0) + num_(allocs[key],0));
    });
  });
  return sums;
}

function wh_sumReceiptsForSO_(soNumber){
  const key = normSoKey_(soNumber);
  if (!key) return 0;
  const sums = wh_collectReceiptSums_(new Set([key]));
  return round2_(sums[key] || 0);
}

function wh_setPaidToDateOnMaster_(soList){
  const arr = Array.isArray(soList) ? soList : [soList];
  const keys = arr.map(normSoKey_).filter(Boolean);
  if (!keys.length) return;
  const keySet = new Set(keys);
  const sums = wh_collectReceiptSums_(keySet);

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(WH_MASTER_TAB_NAME);
  if (!sh) { dbg('wh_setPaidToDateOnMaster_: master sheet missing', WH_MASTER_TAB_NAME); return; }
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return;
  const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
  const H = hIndex_(hdr);
  const cSO = pickH_(H, SO_ALIASES);
  if (!cSO) return;
  const cPTD = pickH_(H, MASTER_PTD_ALIASES);
  const cOT  = pickH_(H, MASTER_OT_ALIASES);
  const cRB  = pickH_(H, MASTER_RB_ALIASES);
  const values = sh.getRange(2,1,lr-1,lc).getValues();
  const updates = {};

  for (let i=0;i<values.length;i++){
    const row = values[i];
    const so = String(row[cSO-1]||'').trim();
    const key = normSoKey_(so);
    if (!key || !keySet.has(key)) continue;
    const newPaid = round2_(sums[key] || 0);
    const info = updates[i+2] || {};
    if (cPTD) {
      const cur = num_(row[cPTD-1],0);
      if (Math.abs(cur - newPaid) > 0.009) {
        info.ptd = newPaid;
      }
    }
    if (cRB && cOT) {
      const ot = num_(row[cOT-1],0);
      const newBal = Math.max(0, round2_(ot - newPaid));
      const curBal = num_(row[cRB-1],0);
      if (Math.abs(curBal - newBal) > 0.009) {
        info.rb = newBal;
      }
    }
    if (Object.keys(info).length) updates[i+2] = info;
  }

  Object.keys(updates).forEach(rowIndex => {
    const idx = Number(rowIndex);
    const info = updates[rowIndex];
    if (cPTD && info.ptd !== undefined) sh.getRange(idx, cPTD).setValue(info.ptd);
    if (cRB && info.rb !== undefined) sh.getRange(idx, cRB).setValue(info.rb);
  });
}

// ============================= ORDERS WRITEBACK =============================
function wh_applyReceiptToOrders_(allocMap, soList){
  const sos = Array.isArray(soList) && soList.length ? soList : Object.keys(allocMap||{});
  if (!sos.length) return;
  wh_setPaidToDateOnMaster_(sos);
}

// ============================= DOC BUILDER =============================
function wh_buildDocFromTemplate_(templateId, model, primaryFolderId){
  const folder = primaryFolderId ? DriveApp.getFolderById(primaryFolderId) : DriveApp.getRootFolder();
  const name = [
    (model.docFlavor==='DEPOSIT' ? 'DEPOSIT ' : 'SALES '),
    (model.docKind==='INVOICE' ? 'INVOICE' : (model.docKind==='RECEIPT' ? 'RECEIPT' : 'DOC')),
    ' — ', (model.companyName || model.customerId || 'Client'),
    ' — ', model.docNumber
  ].join('');

  const docFile = DriveApp.getFileById(templateId).makeCopy(name, folder);
  const docId = docFile.getId(); const docUrl = docFile.getUrl();

  const d = DocumentApp.openById(docId);
  const body = d.getBody();

  const map = {
    '{{CompanyName}}': model.companyName || '',
    '{{CustomerID}}':  model.customerId || '',
    '{{ContactName}}': model.contactName || '',
    '{{Address}}':     model.address || '',
    '{{DOC_DATE}}':    fmtDate_(model.docDate||new Date()),
    '{{InvoiceGroupID}}': model.invoiceGroupId || '',
    '{{DOC_NUMBER}}':  model.docNumber || '',
    '{{TransactionID}}': model.transactionId || '',
    '{{ORDER_TOTAL}}': money_( (model.linesSubtotal||0) + (model.shippingTotal||0) ),
    '{{ShippingFee}}': money_( model.shippingTotal || 0 ),
    '{{PaidToDate}}':  money_( computePaidToDate_(model.customerId, model.soList) ),
    '{{BALANCE}}':     money_( computeBalance_(model.customerId, model.soList, (model.linesSubtotal||0) + (model.shippingTotal||0)) ),
    '{{PaymentMethod}}':  (model.pmt && model.pmt.method) || '',
    '{{PaymentReference}}': (model.pmt && model.pmt.reference) || '',
    '{{Notes}}': (model.pmt && model.pmt.notes) || ''
  };
  if (model.dueDate) map['{{DueDate}}'] = fmtDate_(model.dueDate);

  replaceAll_(body, map);
  if (!model.invoiceGroupId) replaceRegex_(body, /Invoice\s*#:\s*-\s*/g, 'Invoice #: ');
  injectItemsTable_(body, '{{ItemsTable}}', buildItemRows_(model.lines, model.shipping));
  injectPaymentsList_(body, '{{PaymentsList}}', model.customerId, model.soList);

  if (!model.dueDate) {
    const pars = body.getParagraphs();
    for (let i=0;i<pars.length;i++){
      const t=pars[i].getText();
      if (/Due Date\s*:/.test(t)) { body.removeChild(pars[i]); break; }
    }
  }

  d.saveAndClose();
  const pdf = DriveApp.getFileById(docId).getAs(MimeType.PDF);
  const pdfFile = folder.createFile(pdf).setName(name + '.pdf');
  return { docId, docUrl, pdfId: pdfFile.getId(), pdfUrl: pdfFile.getUrl() };
}

function buildItemRows_(lines, shipping){
  const rows = [];
  (lines||[]).forEach(ln=>{
    rows.push([
      String(ln.so||''),
      String(ln.desc||''),
      String(ln.qty||0),
      money_( (ln.qty||0)*(ln.amt||0) )
    ]);
  });
  (shipping||[]).forEach(s=>{
    rows.push(['', String(s.label||'Shipping'), '', money_( num_(s.amount,0) )]);
  });
  if (!rows.length) rows.push(['','','','']);
  return rows;
}

function injectItemsTable_(body, placeholder, rows){
  const range = body.findText(escapeForFind_(placeholder));
  if (range) {
    const p = range.getElement().getParent().asParagraph();
    const idx = body.getChildIndex(p);
    p.removeFromParent();
    const tbl = body.insertTable(idx, makeTable_(['ITEM/SO','DESCRIPTION','QTY','TOTAL'], rows));
    tbl.setBorderWidth(0.5);
  } else {
    body.appendTable(makeTable_(['ITEM/SO','DESCRIPTION','QTY','TOTAL'], rows)).setBorderWidth(0.5);
  }
}

function injectPaymentsList_(body, placeholder, customerId, soList){
  const txns = getPriorPayments_(customerId, soList);
  const text = txns.length
    ? txns.map(t => `${fmtDate_(t.date)}  —  ${money_(t.amount)}${t.method?(' ('+t.method+')'):''}`).join('\n')
    : '';
  replaceAll_(body, { [placeholder]: text });
}

function getPriorPayments_(customerId, soList){
  const sh = ensureLedger_();
  const lc = sh.getLastColumn();
  if (lc < 1) return [];
  const rows = ledgerScanValues_(sh, lc);
  if (!rows.length) return [];
  const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
  const H = headerMap_(hdr);
  const keys = new Set((soList||[]).map(normSoKey_).filter(Boolean));
  if (!keys.size) return [];
  const out = [];
  rows.forEach(row => {
    const docType = String(row[H['DocType']-1]||'').toUpperCase();
    if (!docType || docType.indexOf('RECEIPT') === -1) return;
    if (customerId) {
      const rowCust = String(row[H['CustomerID']-1]||'').trim();
      if (rowCust && rowCust !== customerId) return;
    }
    const status = H['DocStatus'] ? String(row[H['DocStatus']-1]||'').toUpperCase() : 'ISSUED';
    if (status === 'VOID' || status === 'REPLACED') return;
    const allocs = computeAllocationsForRow_(row, H, keys);
    let amount = 0;
    Object.keys(allocs).forEach(key => { if (keys.has(key)) amount += num_(allocs[key],0); });
    amount = round2_(amount);
    if (!(amount>0)) return;
    const dateVal = row[H['PaymentDateTime']-1] || row[H['DOC_DATE']-1] || new Date();
    out.push({ date: dateVal, amount, method: String(row[H['Method']-1]||'') });
  });
  out.sort((a,b)=> (new Date(a.date)) - (new Date(b.date)));
  return out;
}

function computePaidToDate_(customerId, soList){
  const keys = (soList||[]).map(normSoKey_).filter(Boolean);
  if (!keys.length) return 0;
  const sums = wh_collectReceiptSums_(new Set(keys));
  let total = 0;
  keys.forEach(key => { total += num_(sums[key],0); });
  return round2_(total);
}

function computeBalance_(customerId, soList, currentOrderTotal){
  const ptd = computePaidToDate_(customerId, soList);
  return Math.max(0, round2_( (currentOrderTotal||0) - ptd ));
}

// ============================= LEDGER / MIRROR =============================
function ensureLedger_(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(WH_LEDGER_TAB_NAME);
  if (!sh) {
    dbg('ensureLedger_: creating new ledger tab', WH_LEDGER_TAB_NAME);
    sh = ss.insertSheet(WH_LEDGER_TAB_NAME);
    sh.getRange(1,1,1,LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]);
    sh.setFrozenRows(1);
  } else {
    const have = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getDisplayValues()[0].map(s=>String(s||'').trim());
    const map = headerMap_(have);
    let col = have.length;
    LEDGER_HEADERS.forEach(h=>{ if (!map[h]) { col++; sh.getRange(1,col).setValue(h); } });
  }
  return sh;
}

function mirrorTrackerUrl_(customerId, trackerUrl){
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh=>{
    const lr=sh.getLastRow(), lc=sh.getLastColumn(); if (lr<2) return;
    const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
    const H = hIndex_(hdr);
    const cCID = pickH_(H, CUSTID_ALIASES); const cTRK = pickH_(H, TRACKER_ALIASES);
    if (!cCID || !cTRK) return;
    const colCID = sh.getRange(2,cCID,lr-1,1).getDisplayValues().map(a=>String(a[0]||'').trim());
    for (let i=0;i<colCID.length;i++){
      if (colCID[i] === customerId) {
        const cell = sh.getRange(i+2, cTRK);
        const cur = String(cell.getDisplayValue()||'').trim();
        if (!cur) cell.setValue(trackerUrl);
      }
    }
  });
}

// ============================= UTILS =============================
function pickList_(csv, d){ const a=(csv||'').split(',').map(s=>s.trim()).filter(Boolean); return a.length?a:d; }
function hIndex_(hdr){ const H={}; (hdr||[]).forEach((h,i)=>{ const k=String(h||'').trim(); if (k) H[k]=i+1; }); return H; }
function pickH_(H, names){ for (const n of (names||[])) if (H[n]) return H[n]; return 0; }
function headerMap_(hdrRow){ const m={}; hdrRow.forEach((h,i)=>{ m[String(h||'').trim()] = i+1; }); return m; }
function setIf_(row,H,key,val){ if (H[key]) row[H[key]-1] = val; }
function soEq_(a,b){ const sa=String(a||'').trim(), sb=String(b||'').trim(); if (sa===sb) return true; const na=Number(sa.replace(/[^\d.]/g,'')), nb=Number(sb.replace(/[^\d.]/g,'')); return (isFinite(na)&&isFinite(nb)) ? (Math.abs(na-nb)<1e-9) : false; }
function normSoKey_(so){ return String(so||'').trim().toUpperCase().replace(/\s+/g,''); }
function num_(v, d){ const n=parseFloat(String(v||'').replace(/[^\d.\-]/g,'')); return isFinite(n)?n:(d||0); }
function round2_(n){ return Math.round(num_(n,0)*100)/100; }
function money_(n){ n=num_(n,0); const s=n.toFixed(2); return '$'+s.replace(/\B(?=(\d{3})+(?!\d))/g,','); }
function mustStr_(v, name){ const s=String(v||'').trim(); if(!s) throw new Error('Missing '+name); return s; }
function require_(v, name){ if(!v) throw new Error('Missing '+name); }
function fileIdFromUrl_(s){ const m=String(s||'').match(/[-\w]{25,}/); return m?m[0]:''; }
function addDays_(d, days){ const x=new Date(d); x.setDate(x.getDate()+Number(days||0)); return x; }
function fmtDate_(d){ return Utilities.formatDate(d, Session.getScriptTimeZone()||'America/Los_Angeles', 'MMM d, yyyy'); }
function newTransactionId_(customerId, when){ return `TXN-${String(customerId||'NA').replace(/\s+/g,'')}-${Utilities.formatDate(when, Session.getScriptTimeZone()||'America/Los_Angeles', 'yyyyMMdd-HHmmss')}`; }
function newInvoiceGroupId_(customerId, when){ return `IG-${String(customerId||'NA').replace(/\s+/g,'')}-${Utilities.formatDate(when, Session.getScriptTimeZone()||'America/Los_Angeles', 'yyyyMMdd')}`; }
function newDocNumber_(){
  const tz = Session.getScriptTimeZone()||'America/Los_Angeles';
  const key = 'WH_DOC_SEQ_' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    const n = Number(SP.getProperty(key)||'0') + 1;
    SP.setProperty(key, String(n));
    return `${WH_DOC_PREFIX}-${Utilities.formatDate(new Date(), tz, 'yyyyMMdd')}-${('0000'+n).slice(-4)}`;
  } finally { try{lock.releaseLock();}catch(_){} }
}
function escapeForFind_(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function replaceAll_(body, map){
  Object.keys(map).forEach(k=>{
    let r; while ( (r = body.findText(escapeForFind_(k))) ) {
      r.getElement().asText().replaceText(escapeForFind_(k), String(map[k]));
    }
  });
}
function replaceRegex_(body, regex, repl){
  body.getParagraphs().forEach(p=>{
    const t=p.getText();
    if (regex.test(t)) p.setText(t.replace(regex, repl));
  });
}
function makeTable_(headers, rows){
  const temp = DocumentApp.create('tmp-tbl');
  const tb = temp.getBody().appendTable([headers]);
  rows.forEach(r=>tb.appendTableRow([String(r[0]||''), String(r[1]||''), String(r[2]||''), String(r[3]||'')]));
  const copy = tb.copy(); const id=temp.getId(); temp.saveAndClose(); DriveApp.getFileById(id).setTrashed(true);
  return copy;
}
