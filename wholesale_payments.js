/*** ADM WHOLESALE — Record Payments (Multi‑SO + Templates + Filing + Credits + Supersedes)
 * Version: 2.1.0  (2025‑09‑23) — with ADM_DEBUG tracing
 */

const SP = PropertiesService.getScriptProperties();

// ============================= CONFIG =============================
const WH_LEDGER_TAB_NAME         = SP.getProperty('WH_LEDGER_TAB_NAME') || '400_Payments Ledger';
const WH_CRM_TAB_NAME            = SP.getProperty('WH_CRM_TAB_NAME') || '01_CRM';
const WH_DEFAULT_SHIP_PER_ORDER  = num_(SP.getProperty('WH_DEFAULT_SHIP_PER_ORDER'), 50);
const WH_SHIP_THRESHOLD_SUBTOTAL = num_(SP.getProperty('WH_SHIP_THRESHOLD_SUBTOTAL'), 2000);
const WH_DOC_PREFIX              = (SP.getProperty('WH_DOC_PREFIX') || 'ADM').replace(/\s+/g,'').toUpperCase();
const WH_DOCS_ENABLED            = String(SP.getProperty('WH_DOCS_ENABLED') || 'true').toLowerCase() === 'true';
const WH_DOCS_FOLDER_FALLBACK_ID = (SP.getProperty('WH_DOCS_FOLDER_FALLBACK_ID') || '').trim();

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
  ['Contact Name','Contact','Primary Contact','Main Contact','Contact Person','Attn','Attention']);
const CONTACT_FIRST_ALIASES = pickList_(SP.getProperty('WH_CONTACT_FIRST_COL_ALIASES'),
  ['Contact First Name','Contact First','Primary Contact First Name','Primary Contact First','Contact FirstName']);
const CONTACT_LAST_ALIASES  = pickList_(SP.getProperty('WH_CONTACT_LAST_COL_ALIASES'),
  ['Contact Last Name','Contact Last','Primary Contact Last Name','Primary Contact Last','Contact LastName']);

const ADDRESS_ALIASES   = pickList_(SP.getProperty('WH_ADDRESS_ONE_LINE_ALIASES'),
  ['Business Address','Business Address (one line)','Business Address (One Line)','Address','Address (one line)','Company Address']);

const TRACKER_ALIASES   = pickList_(SP.getProperty('WH_TRACKER_URL_COL_ALIASES'),
  ['Customer Order Tracker URL','Order Tracker URL','Tracker URL']);

// For auto line descriptions
const PRODUCT_DESC_ALIASES = pickList_(SP.getProperty('WH_PRODUCT_DESC_ALIASES'),
  ['Product Description','Prod Description','Product','Description','Short Description']);

// (Optional) address parts fallback if one‑line address not present
const STREET_ALIASES = ['Street','Address 1','Addr 1','Address Line 1','Address Line1','Street Address','Business Street','Business Address Street'];
const CITY_ALIASES   = ['City','Town','City/Town','Business City'];
const STATE_ALIASES  = ['State','ST','Province','State/Province','Business State'];
const ZIP_ALIASES    = ['Zip','ZIP','Zip Code','Postal','Postal Code','Postcode','Business Zip'];

function joinAddressParts_(street, city, state, zip) {
  const streetPart = String(street||'').trim();
  const cityPart   = String(city||'').trim();
  const statePart  = String(state||'').trim();
  const zipPart    = String(zip||'').trim();
  const cityState  = [cityPart, statePart].filter(Boolean).join(', ').trim();
  const tail       = [cityState, zipPart].filter(Boolean).join(' ').trim();
  return [streetPart, tail].filter(Boolean).join(', ').trim();
}

// Templates (Google Docs)
const TPL = {
  DI: (SP.getProperty('TEMPLATE_DEPOSIT_INVOICE_ADM') || '').trim(),
  DR: (SP.getProperty('TEMPLATE_DEPOSIT_RECEIPT_ADM') || '').trim(),
  SI: (SP.getProperty('TEMPLATE_SALES_INVOICE_ADM')   || '').trim(),
  SR: (SP.getProperty('TEMPLATE_SALES_RECEIPT_ADM')   || '').trim()
};

const LEDGER_HEADERS = [
  'PaymentID','TransactionID','InvoiceGroupID','DocNumber','DocType','DocFlavor','DocStatus',
  'SupersedesDoc#','SupersedeAction','AppliesToDoc#',
  'CustomerID','CompanyName','ContactName','Address',
  'SOsCSV','PrimarySO',
  'AllocationMode','AllocationsJSON',
  'LinesJSON','LinesSubtotal','ShippingJSON','ShippingTotal',
  'DOC_DATE','DueDate',
  'PaymentDateTime','AmountGross','Method','Reference','Notes',
  'FeePercent','FeeAmount','AmountNet',
  'PDF_URL','DOC_URL','PrimarySO_FolderID','ShortcutIDs_CSV',
  'CustomerOrderTrackerURL',
  'SubmittedBy','SubmittedAt'
];

// ============================= INIT / CONTEXT =============================
function wh_init(){
  const ctx = readActiveContext_();
  const out = {
    nowIso: new Date().toISOString(),
    ctx,
    defaults: {
      shipPerOrder: WH_DEFAULT_SHIP_PER_ORDER,
      shipThresholdSubtotal: WH_SHIP_THRESHOLD_SUBTOTAL,
      docPrefix: WH_DOC_PREFIX
    },
    enabled: ADM_isDebug(),
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
      address = joinAddressParts_(street, city, state, zip);
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
    trackerUrl: trkCol ? String(rowVals[trkCol-1]||'').trim() : ''
  };

  if (!ctx.contactName) {
    const cFirst = pickH_(H, CONTACT_FIRST_ALIASES);
    const cLast  = pickH_(H, CONTACT_LAST_ALIASES);
    const first = cFirst ? String(rowVals[cFirst-1]||'').trim() : '';
    const last  = cLast  ? String(rowVals[cLast-1]||'').trim() : '';
    const combined = [first, last].filter(Boolean).join(' ').trim();
    if (combined) ctx.contactName = combined;
  }
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

  if (ctx.customerId || ctx.companyName) {
    const crm = findCrmRow_(ctx.customerId, ctx.companyName);
    if (crm) {
      ctx = {
        ...ctx,
        customerId: crm.customerId || ctx.customerId || '',
        companyName: crm.companyName || ctx.companyName || '',
        contactName: crm.contactName || ctx.contactName || '',
        address: crm.address || ctx.address || '',
        trackerUrl: crm.trackerUrl || ctx.trackerUrl || ''
      };
      dbg('readActiveContext_: enriched from CRM', { sheet: WH_CRM_TAB_NAME, merged: ctx });
    } else {
      dbg('readActiveContext_: CRM lookup returned no match', { sheet: WH_CRM_TAB_NAME, customerId: ctx.customerId, companyName: ctx.companyName });
    }
  }

  return ctx;
}

function blankCtx_(){ return { sheetName:'', rowIndex:0, soNumber:'', customerId:'', companyName:'', contactName:'', address:'', trackerUrl:'' }; }

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
    const cFirst = pickH_(H, CONTACT_FIRST_ALIASES);
    const cLast  = pickH_(H, CONTACT_LAST_ALIASES);
    const cStreet = pickH_(H, STREET_ALIASES);
    const cCity   = pickH_(H, CITY_ALIASES);
    const cState  = pickH_(H, STATE_ALIASES);
    const cZip    = pickH_(H, ZIP_ALIASES);

    if (!cSO) continue;
    const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();
    for (let i=0;i<vals.length;i++){
      const r = vals[i];
      const s = String(r[cSO-1]||'').trim();
      if (!s) continue;
      if (soEq_(s, so)) {
        const out = {
          sheet: name, rowIndex: i+2,
          soNumber: s,
          customerId: cCID ? String(r[cCID-1]||'').trim() : '',
          companyName: cCMP ? String(r[cCMP-1]||'').trim() : '',
          contactName: cCON ? String(r[cCON-1]||'').trim() : '',
          address:     cADR ? String(r[cADR-1]||'').trim() : '',
          trackerUrl:  cTRK ? String(r[cTRK-1]||'').trim() : '',
          productDesc: cPD  ? String(r[cPD-1]||'').trim() : ''
        };
        if (!out.contactName && (cFirst || cLast)) {
          const first = cFirst ? String(r[cFirst-1]||'').trim() : '';
          const last  = cLast  ? String(r[cLast-1]||'').trim() : '';
          out.contactName = [first, last].filter(Boolean).join(' ').trim();
        }
        if (!out.address && (cStreet || cCity || cState || cZip)) {
          const street = cStreet?String(r[cStreet-1]||'').trim():'';
          const city   = cCity?String(r[cCity-1]||'').trim():'';
          const state  = cState?String(r[cState-1]||'').trim():'';
          const zip    = cZip?String(r[cZip-1]||'').trim():'';
          out.address = joinAddressParts_(street, city, state, zip);
        }
        dbg('findOrdersRowBySO_: match', out);
        return out;
      }
    }
  }
  return null;
}

function findCrmRow_(customerId, companyName){
  const tab = String(WH_CRM_TAB_NAME||'').trim();
  if (!tab) return null;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(tab);
  if (!sh) { dbg('findCrmRow_: sheet not found', tab); return null; }
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) { dbg('findCrmRow_: sheet empty', {tab, lr}); return null; }

  const headers = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
  const H = hIndex_(headers);
  const cCID = pickH_(H, CUSTID_ALIASES);
  const cCMP = pickH_(H, COMPANY_ALIASES);
  const cCON = pickH_(H, CONTACT_ALIASES);
  const cFirst = pickH_(H, CONTACT_FIRST_ALIASES);
  const cLast  = pickH_(H, CONTACT_LAST_ALIASES);
  const cAddr  = pickH_(H, ADDRESS_ALIASES);
  const cStreet= pickH_(H, STREET_ALIASES);
  const cCity  = pickH_(H, CITY_ALIASES);
  const cState = pickH_(H, STATE_ALIASES);
  const cZip   = pickH_(H, ZIP_ALIASES);
  const cTracker = pickH_(H, TRACKER_ALIASES);

  const wantId = String(customerId||'').trim().toLowerCase();
  const wantName = String(companyName||'').trim().toLowerCase();
  const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();

  let fallback = null;
  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    const id = cCID ? String(row[cCID-1]||'').trim() : '';
    const name = cCMP ? String(row[cCMP-1]||'').trim() : '';
    const idLc = id.toLowerCase();
    const nameLc = name.toLowerCase();
    const matchId = wantId && idLc && idLc === wantId;
    const matchName = !matchId && wantName && nameLc && nameLc === wantName;
    if (!matchId && !matchName) {
      if (!fallback && wantName && nameLc.includes(wantName)) fallback = {row, id, name};
      continue;
    }
    return buildCrmPayload_(row, {id, name, cCON, cFirst, cLast, cAddr, cStreet, cCity, cState, cZip, cTracker});
  }

  if (fallback) {
    dbg('findCrmRow_: using partial match', {tab, name: fallback.name});
    return buildCrmPayload_(fallback.row, {
      id: fallback.id,
      name: fallback.name,
      cCON, cFirst, cLast, cAddr, cStreet, cCity, cState, cZip, cTracker
    });
  }
  return null;
}

function buildCrmPayload_(row, cols){
  const {
    id, name, cCON, cFirst, cLast, cAddr, cStreet, cCity, cState, cZip, cTracker
  } = cols;
  const out = {
    customerId: String(id||'').trim(),
    companyName: String(name||'').trim(),
    contactName: cCON ? String(row[cCON-1]||'').trim() : ''
  };
  if (!out.contactName) {
    const first = cFirst ? String(row[cFirst-1]||'').trim() : '';
    const last  = cLast  ? String(row[cLast-1]||'').trim() : '';
    out.contactName = [first, last].filter(Boolean).join(' ').trim();
  }
  let address = cAddr ? String(row[cAddr-1]||'').trim() : '';
  if (!address) {
    const street = cStreet?String(row[cStreet-1]||'').trim():'';
    const city   = cCity?String(row[cCity-1]||'').trim():'';
    const state  = cState?String(row[cState-1]||'').trim():'';
    const zip    = cZip?String(row[cZip-1]||'').trim():'';
    address = joinAddressParts_(street, city, state, zip);
  }
  out.address = address;
  out.trackerUrl = cTracker ? String(row[cTracker-1]||'').trim() : '';
  return out;
}

// ============================= LOOKUPS =============================
function wh_listSOsForCustomer(customerId, limit){
  customerId = String(customerId||'').trim();
  if (!customerId) return [];
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

    const cSO  = pickH_(H, SO_ALIASES);
    const cCID = pickH_(H, CUSTID_ALIASES);
    const cPD  = pickH_(H, PRODUCT_DESC_ALIASES);
    dbg('wh_listSOsForCustomer: header map', {tab:name, cSO, cCID, cPD});

    if (!cSO) continue;
    const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();

    for (let i=0;i<vals.length;i++){
      const r = vals[i];
      const so = String(r[cSO-1]||'').trim();
      if (!so) continue;

      const id = cCID ? String(r[cCID-1]||'').trim() : '';
      if (!id) {
        dbg('wh_listSOsForCustomer: skipping row (no customer id)', {tab:name, rowIndex:i+2, so});
        continue;
      }

      if (!customerIdsEqual_(id, customerId)) {
        dbg('wh_listSOsForCustomer: skipping row (id mismatch)', {
          tab: name,
          rowIndex: i+2,
          so,
          rowCustomerId: id,
          targetCustomerId: customerId
        });
        continue;
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
  const id = String(customerId||'').trim();
  if (id) return wh_listSOsForCustomer(id, 250);
  const found = findOrdersRowBySO_(primarySO);
  if (!found) return [];
  if (found.customerId) {
    const list = wh_listSOsForCustomer(found.customerId, 250);
    if (list.length) return list;
  }
  const soNum = String(found.soNumber || primarySO || '').trim();
  return [{
    soNumber: soNum,
    productDesc: found.productDesc || '',
    sheet: found.sheet,
    rowIndex: found.rowIndex,
    customerId: found.customerId || ''
  }];
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

    const customerId  = mustStr_(payload.customerId, 'customerId');
    const companyName = (payload.companyName||'').trim();
    const contactName = (payload.contactName||'').trim();
    const address     = (payload.address||'').trim();

    const primarySO = mustStr_(payload.primarySO, 'primarySO');
    const soList = (payload.soList||[]).map(s=>String(s||'').trim()).filter(Boolean);
    if (!soList.length) throw new Error('Select/type at least one SO.');
    if (!soList.includes(primarySO)) soList.unshift(primarySO);

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
    pmt.amount = num_(pmt.amount, 0);
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
        (payload.allocations||[]).forEach(a => { const so=String(a.so||'').trim(); if (so) allocMap[so]=num_(a.amount,0); });
      }
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
    let docUrl='', pdfUrl='', primaryFolderId='', shortcutIds=[];
    if (WH_DOCS_ENABLED) {
      primaryFolderId = wh_findSoFolderId(primarySO) || WH_DOCS_FOLDER_FALLBACK_ID || '';
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
    setIf_(row,H,'DocType', docKind);
    setIf_(row,H,'DocFlavor', docFlavor);
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

    const feePct = 0, feeAmt = round2_((pmt.amount||0)*feePct), amtNet = round2_((pmt.amount||0)-feeAmt);
    setIf_(row,H,'PaymentDateTime', (docKind==='RECEIPT') ? (pmt.dateTimeISO ? new Date(pmt.dateTimeISO) : new Date()) : '');
    setIf_(row,H,'AmountGross', pmt.amount||'');
    setIf_(row,H,'Method', pmt.method||'');
    setIf_(row,H,'Reference', pmt.reference||'');
    setIf_(row,H,'Notes', (pmt.notes||payload.notes||''));

    setIf_(row,H,'FeePercent', feePct);
    setIf_(row,H,'FeeAmount', feeAmt);
    setIf_(row,H,'AmountNet', amtNet);

    setIf_(row,H,'PDF_URL', pdfUrl);
    setIf_(row,H,'DOC_URL', docUrl);
    setIf_(row,H,'PrimarySO_FolderID', primaryFolderId);
    setIf_(row,H,'ShortcutIDs_CSV', (shortcutIds||[]).join(','));
    setIf_(row,H,'CustomerOrderTrackerURL', (payload.trackerUrl||''));

    setIf_(row,H,'SubmittedBy', Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'user');
    setIf_(row,H,'SubmittedAt', new Date());
    sh.appendRow(row);
    dbg('wh_submitPayment: appended ledger row');

    try { if (payload.trackerUrl) mirrorTrackerUrl_(customerId, payload.trackerUrl); } catch(e){ dbg('mirrorTrackerUrl error', String(e)); }

    if (docKind==='RECEIPT' && Object.keys(allocMap).length) {
      wh_applyReceiptToOrders_(allocMap);
    }

    let overage = 0;
    if (docKind==='RECEIPT') {
      const allocSum = Object.values(allocMap).reduce((a,b)=>a+b,0);
      if (pmt.amount > allocSum) {
        overage = round2_(pmt.amount - allocSum);
        writeCreditRow_(customerId, companyName, overage, pmt.method, transactionId, invoiceGroupId);
      }
    }

    return { ok:true, docType: dt, transactionId, invoiceGroupId, docNumber, pdfUrl, docUrl, overageCredit: overage };
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
  params = params || {};
  const debugging = ADM_isDebug();
  let scope = String(params.scope||'').trim().toUpperCase();
  let soNumber = String(params.soNumber||'').trim();
  let customerId = String(params.customerId||'').trim();
  let invoiceGroupId = String(params.invoiceGroupId||'').trim();

  if (debugging) {
    dbg('wh_getSummary: start', { params, scope, soNumber, customerId, invoiceGroupId });
  }

  if (!soNumber || !customerId) {
    const ctx = readActiveContext_();
    if (!soNumber && ctx.soNumber) soNumber = String(ctx.soNumber||'').trim();
    if (!customerId && ctx.customerId) customerId = String(ctx.customerId||'').trim();
  }

  if (!scope) {
    if (invoiceGroupId) {
      scope = 'GROUP';
    } else if (soNumber) {
      scope = 'SO';
    } else if (customerId) {
      scope = 'CUSTOMER';
    } else {
      scope = 'SO';
    }
  }

  if (scope === 'SO' && !soNumber && customerId) {
    scope = 'CUSTOMER';
  }

  if (scope === 'CUSTOMER' && !customerId && soNumber) {
    scope = 'SO';
  }

  if (debugging) {
    dbg('wh_getSummary: resolved scope', { scope, soNumber, customerId, invoiceGroupId });
  }

  const sh = ensureLedger_();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) {
    if (debugging) {
      dbg('wh_getSummary: ledger empty', { lastRow: lr, lastColumn: lc });
    }
    return { scope, items:[], totals:{}, groups:[], warnings:[] };
  }

  const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
  const H = headerMap_(hdr);
  const vals = sh.getRange(2,1,lr-1,lc).getValues();

  if (debugging) {
    dbg('wh_getSummary: ledger headers loaded', { headerCount: hdr.length, headers: hdr });
    dbg('wh_getSummary: scanning rows', { rowCount: vals.length });
  }

  const items = [];
  const tz = Session.getScriptTimeZone() || 'UTC';
  const now = new Date();
  const nowMs = now.getTime();
  const groupsByKey = new Map();
  const activeInvoicesBySO = new Map();
  const globalTotals = {
    invoiced: 0,
    receipts: 0,
    creditsIssued: 0,
    creditsApplied: 0
  };

  for (let i=0;i<vals.length;i++){
    const r = vals[i];
    const raw = {};
    Object.keys(H).forEach(k=>raw[k]=r[H[k]-1]);

    const soQuery = soNumber;
    const ledgerSOs = parseSummarySOList_(raw);
    const ledgerCustomerIds = parseSummaryCustomerIds_(raw);
    const ledgerGroupId = String(raw.InvoiceGroupID||'').trim();
    let match = false;
    let matchReason = '';

    if (scope === 'SO') {
      if (!soQuery) {
        matchReason = 'SO scope but soNumber missing';
      } else if (ledgerSOs.some(so => soEq_(so, soQuery))) {
        match = true;
        matchReason = 'SO matched';
      } else {
        matchReason = 'SO mismatch';
      }
    } else if (scope === 'CUSTOMER') {
      if (!customerId) {
        matchReason = 'CUSTOMER scope but customerId missing';
      } else if (ledgerCustomerIds.some(id => idEq_(id, customerId))) {
        match = true;
        matchReason = 'Customer matched';
      } else {
        matchReason = 'Customer mismatch';
      }
    } else if (scope === 'GROUP') {
      if (!invoiceGroupId) {
        matchReason = 'GROUP scope but invoiceGroupId missing';
      } else if (idEq_(ledgerGroupId, invoiceGroupId)) {
        match = true;
        matchReason = 'Group matched';
      } else {
        matchReason = 'Group mismatch';
      }
    }

    const docTypeRaw = String(raw.DocType||'');
    const docTypeUpper = docTypeRaw.toUpperCase();
    const amountInfo = resolveSummaryAmount_(raw, docTypeUpper);

    if (debugging) {
      const dbgRow = {
        rowIndex: i + 2,
        paymentId: raw.PaymentID || '',
        docNumber: raw.DocNumber || '',
        docType: raw.DocType || '',
        amount: amountInfo.amount,
        amountSource: amountInfo.source,
        ledgerSOs,
        ledgerCustomerIds,
        ledgerGroupId,
        scope,
        soQuery,
        customerId,
        invoiceGroupId,
        match,
        matchReason
      };
      dbg('wh_getSummary: ledger row inspected', dbgRow);
    }

    if (!match) continue;

    items.push(raw);

    const docFlavor = String(raw.DocFlavor||'');
    const docStatus = String(raw.DocStatus||'');
    const amount = amountInfo.amount;

    const docDate = coerceDate_(raw.DOC_DATE || raw.PaymentDateTime || raw.SubmittedAt);
    const paymentDate = coerceDate_(raw.PaymentDateTime);
    const dueDate = coerceDate_(raw.DueDate);

    const docDateMs = docDate ? docDate.getTime() : (paymentDate ? paymentDate.getTime() : null);
    const dueDateMs = dueDate ? dueDate.getTime() : null;

    const soNumbers = parseSummarySOList_(raw);

    const docOut = {
      docType: docTypeRaw,
      docFlavor,
      docLabel: buildSummaryDocLabel_(docTypeRaw, docFlavor),
      docNumber: String(raw.DocNumber||'').trim(),
      docStatus,
      displayDate: formatDateForSummary_(docDate || paymentDate || raw.SubmittedAt, tz),
      docDateDisplay: formatDateForSummary_(docDate, tz),
      paymentDateDisplay: formatDateForSummary_(paymentDate, tz),
      dueDateDisplay: formatDateForSummary_(dueDate, tz),
      docDateIso: docDate ? docDate.toISOString() : '',
      dueDateIso: dueDate ? dueDate.toISOString() : '',
      amount,
      method: String(raw.Method||'').trim(),
      pdfUrl: String(raw.PDF_URL||'').trim(),
      invoiceGroupId: String(raw.InvoiceGroupID||'').trim(),
      soNumbers,
      isInvoice: isInvoiceDoc_(docTypeUpper),
      isReceipt: isReceiptDoc_(docTypeUpper),
      isCredit: docTypeUpper === 'CREDIT',
      isCreditApplied: docTypeUpper === 'CREDIT-APPLIED',
      isActiveInvoice: false,
      dueDateMs,
      activityMs: docDateMs,
      notes: String(raw.Notes||'').trim(),
      amountSource: amountInfo.source
    };

    docOut.isActiveInvoice = docOut.isInvoice && isActiveInvoiceStatus_(docStatus);

    const groupKey = docOut.invoiceGroupId || '__' + (docOut.docNumber || raw.PaymentID || ('ROW'+(i+2)));
    if (!groupsByKey.has(groupKey)) {
      groupsByKey.set(groupKey, {
        key: groupKey,
        invoiceGroupId: docOut.invoiceGroupId,
        docs: [],
        soNumbers: new Set(),
        totals: { invoiced: 0, receipts: 0, creditsIssued: 0, creditApplied: 0 },
        dueDates: [],
        latestActivityMs: docDateMs || 0,
        warnings: []
      });
    }
    const group = groupsByKey.get(groupKey);
    group.docs.push(docOut);
    if (docOut.activityMs && docOut.activityMs > (group.latestActivityMs||0)) {
      group.latestActivityMs = docOut.activityMs;
    }
    soNumbers.forEach(so => { if (so) group.soNumbers.add(so); });

    if (docOut.isInvoice) {
      group.totals.invoiced = round2_(group.totals.invoiced + amount);
      globalTotals.invoiced = round2_(globalTotals.invoiced + amount);
      if (dueDate) group.dueDates.push(dueDate);
    } else if (docOut.isReceipt) {
      group.totals.receipts = round2_(group.totals.receipts + amount);
      globalTotals.receipts = round2_(globalTotals.receipts + amount);
    } else if (docOut.isCreditApplied) {
      group.totals.creditApplied = round2_(group.totals.creditApplied + amount);
      globalTotals.creditsApplied = round2_(globalTotals.creditsApplied + amount);
    } else if (docOut.isCredit) {
      group.totals.creditsIssued = round2_(group.totals.creditsIssued + amount);
      globalTotals.creditsIssued = round2_(globalTotals.creditsIssued + amount);
    }

    if (docOut.isActiveInvoice) {
      soNumbers.forEach(so => {
        if (!so) return;
        if (!activeInvoicesBySO.has(so)) activeInvoicesBySO.set(so, []);
        activeInvoicesBySO.get(so).push({
          docNumber: docOut.docNumber,
          invoiceGroupId: docOut.invoiceGroupId,
          status: docStatus
        });
      });
    }
  }

  const groups = [];
  const generalWarnings = [];

  groupsByKey.forEach(group => {
    const soNumbers = Array.from(group.soNumbers).filter(Boolean).sort();
    const balance = round2_(group.totals.invoiced - (group.totals.receipts + group.totals.creditApplied));
    const dueDateMsList = group.dueDates.map(d=>d.getTime()).sort((a,b)=>a-b);
    const overdueDocs = group.docs.filter(doc => doc.isInvoice && doc.isActiveInvoice && doc.dueDateMs && doc.dueDateMs < nowMs);
    const hasOverdue = balance > 0 && overdueDocs.length > 0;
    const oldestOverdueMs = overdueDocs.length ? Math.min.apply(null, overdueDocs.map(doc=>doc.dueDateMs)) : null;
    const oldestOverdueDisplay = oldestOverdueMs ? formatDateForSummary_(new Date(oldestOverdueMs), tz) : '';

    if (hasOverdue && oldestOverdueDisplay) {
      group.warnings.push(`Outstanding balance is past due (oldest due ${oldestOverdueDisplay}).`);
    }

    group.docs.forEach(doc => {
      doc.isOverdue = hasOverdue && doc.isInvoice && doc.isActiveInvoice && doc.dueDateMs && doc.dueDateMs < nowMs;
    });

    const groupLabel = group.invoiceGroupId || (soNumbers.length ? `SO ${soNumbers.join(', ')}` : 'Ungrouped');

    groups.push({
      invoiceGroupId: group.invoiceGroupId,
      label: groupLabel,
      soNumbers,
      totals: {
        invoiced: round2_(group.totals.invoiced),
        receipts: round2_(group.totals.receipts),
        creditApplied: round2_(group.totals.creditApplied),
        creditsIssued: round2_(group.totals.creditsIssued),
        balance
      },
      hasOverdue,
      oldestOverdueDisplay,
      nextDueDateDisplay: dueDateMsList.length ? formatDateForSummary_(new Date(dueDateMsList[0]), tz) : '',
      latestActivityMs: group.latestActivityMs || 0,
      warnings: group.warnings.slice(),
      docs: group.docs.map(doc => ({
        docType: doc.docType,
        docFlavor: doc.docFlavor,
        docLabel: doc.docLabel,
        docNumber: doc.docNumber,
        docStatus: doc.docStatus,
        displayDate: doc.displayDate,
        docDateDisplay: doc.docDateDisplay,
        paymentDateDisplay: doc.paymentDateDisplay,
        dueDateDisplay: doc.dueDateDisplay,
        docDateIso: doc.docDateIso,
        dueDateIso: doc.dueDateIso,
        amount: doc.amount,
        method: doc.method,
        pdfUrl: doc.pdfUrl,
        isInvoice: doc.isInvoice,
        isReceipt: doc.isReceipt,
        isCredit: doc.isCredit,
        isCreditApplied: doc.isCreditApplied,
        isOverdue: doc.isOverdue,
        isActiveInvoice: doc.isActiveInvoice,
        soNumbers: doc.soNumbers,
        amountFormatted: formatCurrency_(doc.amount),
        amountSource: doc.amountSource
      }))
    });

    if (hasOverdue) {
      const balanceText = formatCurrency_(balance);
      generalWarnings.push(`Invoice group ${groupLabel} has an outstanding balance of ${balanceText} that is past due${oldestOverdueDisplay ? ' (oldest due '+oldestOverdueDisplay+')' : ''}.`);
    }
  });

  groups.sort((a,b)=> (b.latestActivityMs||0) - (a.latestActivityMs||0));

  activeInvoicesBySO.forEach((docs, so) => {
    if (docs.length <= 1) return;
    const docList = docs.map(d => d.docNumber || d.invoiceGroupId || 'invoice').join(', ');
    generalWarnings.push(`SO ${so} has multiple active invoices (${docList}).`);
  });

  const totals = {
    invoiced: round2_(globalTotals.invoiced),
    receipts: round2_(globalTotals.receipts),
    creditsIssued: round2_(globalTotals.creditsIssued),
    creditsApplied: round2_(globalTotals.creditsApplied),
    balance: round2_(globalTotals.invoiced - (globalTotals.receipts + globalTotals.creditsApplied))
  };

  const warnings = dedupeSummaryWarnings_(generalWarnings);

  if (debugging) {
    dbg('wh_getSummary: summary built', {
      scope,
      itemCount: items.length,
      groupCount: groups.length,
      totals,
      warnings
    });
  }

  return {
    scope,
    items,
    totals,
    groups,
    warnings
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

function parseSummarySOList_(row){
  if (!row || typeof row !== 'object') return [];
  const out = [];
  const seen = new Set();
  const add = so => {
    const val = String(so||'').trim();
    const norm = normalizeSo_(val);
    if (!norm || seen.has(norm)) return;
    seen.add(norm);
    out.push(val);
  };

  const addFromCsv = value => {
    String(value||'')
      .replace(/[|;]/g, ',')
      .split(/[\n,]/)
      .map(part => String(part||'').split(':')[0].trim())
      .forEach(part => add(part));
  };

  const addFromJson = value => {
    if (!value) return;
    try {
      const parsed = (typeof value === 'string') ? JSON.parse(value) : value;
      if (Array.isArray(parsed)) {
        parsed.forEach(add);
      } else if (parsed && typeof parsed === 'object') {
        Object.keys(parsed).forEach(add);
      }
    } catch(_) {
      addFromCsv(value);
    }
  };

  const keys = Object.keys(row);
  keys.forEach(key => {
    const lower = String(key||'').toLowerCase();
    const keyNorm = lower.replace(/[^a-z0-9]/g, '');
    if (keyNorm === 'primaryso' || keyNorm === 'primaryso#') {
      add(row[key]);
    } else if (keyNorm === 'soscsv' || keyNorm === 'solist' || keyNorm === 'sos') {
      addFromCsv(row[key]);
    } else if (keyNorm === 'allocationsjson' || keyNorm === 'sojson') {
      addFromJson(row[key]);
    }
  });

  // Explicit known columns for clarity
  add(row.PrimarySO);
  addFromCsv(row.SOsCSV);
  addFromCsv(row.SO_List);
  addFromJson(row.AllocationsJSON);

  return out;
}

function parseSummaryCustomerIds_(row){
  if (!row || typeof row !== 'object') return [];
  const out = [];
  const seen = new Set();
  const add = id => {
    const val = String(id||'').trim();
    const norm = normalizeId_(val);
    if (!norm || seen.has(norm)) return;
    seen.add(norm);
    out.push(val);
  };

  const keys = Object.keys(row);
  keys.forEach(key => {
    const lower = String(key||'').toLowerCase();
    const keyNorm = lower.replace(/[^a-z0-9]/g, '');
    if (keyNorm === 'customerid' || keyNorm === 'companyid' || keyNorm === 'customercompanyid') {
      add(row[key]);
    }
  });

  add(row.CustomerID);
  add(row.CompanyID);

  return out;
}

function buildSummaryDocLabel_(docType, docFlavor){
  const parts = [String(docType||'').trim(), String(docFlavor||'').trim()].filter(Boolean);
  return parts.join(' — ');
}

function resolveSummaryAmount_(row, docTypeUpper){
  const docType = String(docTypeUpper||'').toUpperCase();
  const gross = numFromLedger_(row ? row.AmountGross : null);
  const net = numFromLedger_(row ? row.AmountNet : null);
  const subtotal = numFromLedger_(row ? row.LinesSubtotal : null);
  const shipping = numFromLedger_(row ? row.ShippingTotal : null);

  const finish = (amount, source) => ({ amount: round2_(amount||0), source });

  if (isInvoiceDoc_(docType)) {
    if (gross !== null && gross !== 0) return finish(gross, 'AmountGross');
    if (net !== null && net !== 0) return finish(net, 'AmountNet');

    const hasSubtotalOrShipping = (subtotal !== null) || (shipping !== null);

    if (hasSubtotalOrShipping) {
      const sum = (subtotal||0) + (shipping||0);
      if (sum !== 0) return finish(sum, 'LinesSubtotal/ShippingTotal');
    }

    const linesTotal = sumLedgerLines_(row ? row.LinesJSON : null);
    if (linesTotal !== null) {
      const combined = linesTotal + (shipping||0);
      if (combined !== 0) {
        return finish(combined, shipping !== null ? 'LinesJSON+ShippingTotal' : 'LinesJSON');
      }
    }

    if (hasSubtotalOrShipping) {
      const sum = (subtotal||0) + (shipping||0);
      return finish(sum, 'LinesSubtotal/ShippingTotal');
    }

    if (gross !== null) return finish(gross, 'AmountGross');
    if (net !== null) return finish(net, 'AmountNet');
  } else {
    if (gross !== null && gross !== 0) return finish(gross, 'AmountGross');
    if (net !== null && net !== 0) return finish(net, 'AmountNet');
    if (gross !== null) return finish(gross, 'AmountGross');
    if (net !== null) return finish(net, 'AmountNet');
  }

  return finish(0, 'Default0');
}

function isInvoiceDoc_(docTypeUpper){
  return String(docTypeUpper||'').toUpperCase().indexOf('INVOICE') >= 0;
}

function isReceiptDoc_(docTypeUpper){
  return String(docTypeUpper||'').toUpperCase().indexOf('RECEIPT') >= 0;
}

function isActiveInvoiceStatus_(status){
  const s = String(status||'').toUpperCase();
  if (!s) return true;
  const inactiveFlags = ['VOID','VOIDED','SUPERSEDED','SUPERSEDE','CANCELLED','CANCELED','SUPERCEDED'];
  return !inactiveFlags.some(flag => s.indexOf(flag) >= 0);
}

function coerceDate_(value){
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  if (value === null || value === undefined || value === '') return null;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function formatDateForSummary_(value, tz){
  const d = coerceDate_(value);
  if (!d) return '';
  return Utilities.formatDate(d, tz || Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd');
}

function formatCurrency_(amount){
  const n = num_(amount, NaN);
  if (!isFinite(n)) return '';
  const fixed = round2_(n).toFixed(2);
  return '$' + fixed.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function dedupeSummaryWarnings_(warnings){
  const seen = new Set();
  const out = [];
  (warnings||[]).forEach(msg => {
    const text = String(msg||'').trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    out.push(text);
  });
  return out;
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
  setIf_(row,H,'SubmittedAt', new Date());
  sh.appendRow(row);
}

// ============================= ORDERS WRITEBACK =============================
function wh_applyReceiptToOrders_(allocMap){
  const ss = SpreadsheetApp.getActive();
  const tabNames = WH_ORDERS_TAB_NAMES.length ? WH_ORDERS_TAB_NAMES : ss.getSheets().map(s=>s.getName());
  const debugging = ADM_isDebug();

  const PTD_ALIASES  = ['Paid-to-Date','Paid To Date','Paid-To-Date','Paid to Date','Paid'];
  const OT_ALIASES   = ['Order Total','OrderTotal','Total'];
  const RB_ALIASES   = ['Remaining Balance','Balance','RB'];

  for (const tab of tabNames){
    const sh = ss.getSheetByName(tab); if (!sh) continue;
    const lr = sh.getLastRow(), lc=sh.getLastColumn(); if (lr<2) continue;
    const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
    const H = hIndex_(hdr);
    const cSO  = pickH_(H, SO_ALIASES); if (!cSO) continue;
    const cPTD = pickH_(H, PTD_ALIASES);
    const cOT  = pickH_(H, OT_ALIASES);
    const cRB  = pickH_(H, RB_ALIASES);

    if (debugging) {
      dbg('wh_applyReceiptToOrders_: inspecting sheet', {
        sheet: tab,
        lastRow: lr,
        lastColumn: lc,
        headers: hdr,
        soCol: cSO,
        paidToDateCol: cPTD,
        orderTotalCol: cOT,
        remainingBalanceCol: cRB
      });
    }

    const vals = sh.getRange(2,1,lr-1,lc).getValues();
    let touched = false;
    const touchLog = debugging ? [] : null;
    const writesByRow = new Map();

    const cellWithinRange_ = (range, rowIndex, columnIndex) => {
      if (!range || !rowIndex || !columnIndex) return false;
      try {
        const startRow = range.getRow();
        const startCol = range.getColumn();
        const endRow = startRow + range.getNumRows() - 1;
        const endCol = startCol + range.getNumColumns() - 1;
        return rowIndex >= startRow && rowIndex <= endRow &&
               columnIndex >= startCol && columnIndex <= endCol;
      } catch (err) {
        return false;
      }
    };

    const intersectingProtections = (rowIndex, columnIndex) => {
      const out = { isProtected: false, hits: [] };
      if (!rowIndex || !columnIndex) return out;

      let errorMessage = '';
      try {
        const { ProtectionType } = SpreadsheetApp;
        const cell = sh.getRange(rowIndex, columnIndex);

        if (cell && typeof cell.getProtections === 'function') {
          const rangeProtections = cell.getProtections(ProtectionType.RANGE) || [];
          rangeProtections.forEach(p => {
            let canEdit = null;
            try {
              canEdit = typeof p.canEdit === 'function' ? !!p.canEdit() : null;
            } catch (_) {
              canEdit = null;
            }
            if (canEdit) return;

            const range = typeof p.getRange === 'function' ? p.getRange() : null;
            out.hits.push({
              type: 'RANGE',
              description: typeof p.getDescription === 'function' ? (p.getDescription() || '') : '',
              a1: range && typeof range.getA1Notation === 'function' ? range.getA1Notation() : null,
              canEdit
            });
          });
        }

        const sheetProtections = (typeof sh.getProtections === 'function')
          ? (sh.getProtections(ProtectionType.SHEET) || [])
          : [];
        const blockingSheetProtections = Array.isArray(sheetProtections) ? sheetProtections : [];

        blockingSheetProtections.forEach(p => {
          let canEdit = null;
          try {
            canEdit = typeof p.canEdit === 'function' ? !!p.canEdit() : null;
          } catch (_) {
            canEdit = null;
          }
          if (canEdit) return;

          let unprotectedRanges = [];
          try {
            unprotectedRanges = typeof p.getUnprotectedRanges === 'function'
              ? (p.getUnprotectedRanges() || [])
              : [];
          } catch (_) {
            unprotectedRanges = [];
          }

          const withinUnprotected = (unprotectedRanges || []).some(range => cellWithinRange_(range, rowIndex, columnIndex));
          if (withinUnprotected) return;

          out.hits.push({
            type: 'SHEET',
            description: typeof p.getDescription === 'function' ? (p.getDescription() || '') : '',
            canEdit,
            unprotectedRanges: Array.isArray(unprotectedRanges) ? unprotectedRanges.length : 0
          });
        });

        out.isProtected = out.hits.length > 0;
      } catch (innerErr) {
        errorMessage = String(innerErr && innerErr.message || innerErr || '');
      }

      if (errorMessage) {
        return {
          isProtected: null,
          hits: [],
          error: errorMessage
        };
      }

      return out;
    };

    for (let i=0;i<vals.length;i++){
      const r = vals[i];
      const so = String(r[cSO-1]||'').trim(); if (!so || !(so in allocMap)) continue;
      const add = num_(allocMap[so],0); if (add<=0) continue;

      let rowTouched = false;
      let rowLog;
      if (debugging) {
        rowLog = {
          sheet: tab,
          rowIndex: i + 2,
          so,
          updates: []
        };
      }

      if (cPTD){
        const cur = num_(r[cPTD-1],0);
        const next = cur + add;
        r[cPTD-1] = next;
        touched = true;
        rowTouched = true;
        const sheetRowIndex = i + 2;
        if (!writesByRow.has(sheetRowIndex)) writesByRow.set(sheetRowIndex, new Map());
        writesByRow.get(sheetRowIndex).set(cPTD, next);
        if (debugging) {
          rowLog.updates.push({
            columnIndex: cPTD,
            columnName: hdr[cPTD-1] || ('Col'+cPTD),
            before: cur,
            after: next
          });
        }
      }
      if (cRB && cOT && cPTD){
        const ot=num_(r[cOT-1],0);
        const ptd=num_(r[cPTD-1],0);
        const prev=num_(r[cRB-1],0);
        const next=Math.max(0, round2_(ot-ptd));
        r[cRB-1] = next;
        touched=true;
        rowTouched = true;
        const sheetRowIndex = i + 2;
        if (!writesByRow.has(sheetRowIndex)) writesByRow.set(sheetRowIndex, new Map());
        writesByRow.get(sheetRowIndex).set(cRB, next);
        if (debugging) {
          rowLog.updates.push({
            columnIndex: cRB,
            columnName: hdr[cRB-1] || ('Col'+cRB),
            before: prev,
            after: next,
            basis: { orderTotal: ot, paidToDate: ptd }
          });
        }
      }
      vals[i] = r;
      if (debugging && rowTouched) touchLog.push(rowLog);
    }

    if (!touched){
      if (debugging) dbg('wh_applyReceiptToOrders_: no matches on sheet', { sheet: tab, allocKeys: Object.keys(allocMap||{}) });
      continue;
    }

    const range = sh.getRange(2,1,lr-1,lc);
    const plannedSegments = [];
    const rowIndices = Array.from(writesByRow.keys()).sort((a,b)=>a-b);
    rowIndices.forEach(rowIndex => {
      const colEntries = Array.from(writesByRow.get(rowIndex).entries()).sort((a,b)=>a[0]-b[0]);
      let segmentStart = null;
      let segmentValues = [];
      let segmentColumns = [];
      let prevCol = null;
      const commitSegment = () => {
        if (!segmentValues.length) return;
        plannedSegments.push({
          rowIndex,
          startColumn: segmentStart,
          columnCount: segmentValues.length,
          columnIndices: segmentColumns.slice(),
          values: segmentValues.slice()
        });
        segmentStart = null;
        segmentValues = [];
        segmentColumns = [];
      };
      colEntries.forEach(([colIndex, value]) => {
        if (segmentStart === null) {
          segmentStart = colIndex;
          segmentValues = [value];
          segmentColumns = [colIndex];
        } else if (prevCol !== null && colIndex === prevCol + 1) {
          segmentValues.push(value);
          segmentColumns.push(colIndex);
        } else {
          commitSegment();
          segmentStart = colIndex;
          segmentValues = [value];
          segmentColumns = [colIndex];
        }
        prevCol = colIndex;
      });
      commitSegment();
    });

    if (debugging){
      dbg('wh_applyReceiptToOrders_: attempting write', {
        sheet: tab,
        range: range.getA1Notation(),
        rows: lr-1,
        columns: lc,
        headers: hdr,
        touches: touchLog,
        segments: plannedSegments
      });
    }

    try {
      plannedSegments.forEach(segment => {
        const targetRange = sh.getRange(segment.rowIndex, segment.startColumn, 1, segment.columnCount);
        targetRange.setValues([segment.values]);
      });
    } catch (err) {
      if (debugging) {
        const cellDiagnostics = [];
        try {
          const seen = {};
          (touchLog||[]).forEach(rowInfo => {
            (rowInfo.updates||[]).forEach(update => {
              const key = rowInfo.rowIndex + ':' + update.columnIndex;
              if (seen[key]) return;
              seen[key] = true;
              const cell = sh.getRange(rowInfo.rowIndex, update.columnIndex);
              const formula = cell.getFormula();
              const protectionDetails = intersectingProtections(rowInfo.rowIndex, update.columnIndex);
              const diag = {
                rowIndex: rowInfo.rowIndex,
                columnIndex: update.columnIndex,
                columnName: update.columnName,
                formula: formula || null,
                hasFormula: !!formula,
                isBlank: cell.isBlank(),
                note: cell.getNote() || null
              };
              if (protectionDetails) {
                diag.isProtected = protectionDetails.isProtected;
                if (protectionDetails.error) diag.protectionError = protectionDetails.error;
                if (protectionDetails.hits && protectionDetails.hits.length) {
                  diag.protections = protectionDetails.hits;
                }
              }
              cellDiagnostics.push(diag);
            });
          });
        } catch (inner){
          cellDiagnostics.push({ error: 'cell diagnostics failed', message: inner && inner.message });
        }
        dbg('wh_applyReceiptToOrders_: write failed', {
          sheet: tab,
          range: range.getA1Notation(),
          headers: hdr,
          touches: touchLog,
          cells: cellDiagnostics,
          error: String(err && err.message || err)
        });
      }
      throw err;
    }
  }
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
  dbg('buildItemRows_: input', {
    linesCount: Array.isArray(lines) ? lines.length : 0,
    shippingCount: Array.isArray(shipping) ? shipping.length : 0
  });
  const rows = [];
  (lines||[]).forEach(ln=>{
    rows.push([
      String(ln.so||''),
      String(ln.desc||''),
      String(ln.qty||0),
      money_( (ln.qty||0)*(ln.amt||0) )
    ]);
  });
  if (!rows.length) rows.push(['','','','']);
  dbg('buildItemRows_: output', { rowCount: rows.length });
  return rows;
}

function injectItemsTable_(body, placeholder, rows){
  const headers = ['ITEM/SO','DESCRIPTION','QTY','TOTAL'];
  const range = body.findText(escapeForFind_(placeholder));
  dbg('injectItemsTable_: start', {
    placeholder,
    rowCount: Array.isArray(rows) ? rows.length : 0,
    hasRange: !!range
  });
  const table = makeTable_(headers, rows, { includeHeader: false });

  if (!range) {
    dbg('injectItemsTable_: placeholder not found, appending to body');
    body.appendTable(table).setBorderWidth(0.5);
    return;
  }

  let el = range.getElement();
  while (el && el.getParent && el.getType &&
         el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
         el.getType() !== DocumentApp.ElementType.LIST_ITEM) {
    el = el.getParent();
  }

  if (!el || !el.getParent) {
    dbg('injectItemsTable_: element missing parent, appending to body');
    body.appendTable(table).setBorderWidth(0.5);
    return;
  }

  const paragraph = (el.getType && el.getType() === DocumentApp.ElementType.LIST_ITEM)
    ? el.asListItem()
    : el.asParagraph();
  const container = paragraph.getParent();
  const paragraphType = paragraph && paragraph.getType ? String(paragraph.getType()) : 'unknown';
  const containerType = container && container.getType ? String(container.getType()) : 'unknown';
  dbg('injectItemsTable_: resolved container', { containerType, paragraphType });

  if (container && typeof container.getChildIndex === 'function' && typeof container.insertTable === 'function') {
    const idx = container.getChildIndex(paragraph);
    const inserted = container.insertTable(idx, table);
    inserted.setBorderWidth(0.5);
    paragraph.removeFromParent();
    dbg('injectItemsTable_: inserted via container.insertTable', { containerType, childIndex: idx });
    return;
  }

  if (container && container.getType && container.getType() === DocumentApp.ElementType.BODY_SECTION) {
    const idx = body.getChildIndex(paragraph);
    const inserted = body.insertTable(idx, table);
    inserted.setBorderWidth(0.5);
    paragraph.removeFromParent();
    dbg('injectItemsTable_: inserted via body.insertTable', { childIndex: idx });
  } else {
    paragraph.removeFromParent();
    dbg('injectItemsTable_: fallback append to body (container unsupported)', { containerType });
    body.appendTable(table).setBorderWidth(0.5);
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
  const sh = ensureLedger_(); const lr=sh.getLastRow(), lc=sh.getLastColumn(); if (lr<2) return [];
  const H = headerMap_(sh.getRange(1,1,1,lc).getDisplayValues()[0]);
  const vals = sh.getRange(2,1,lr-1,lc).getValues();
  const set = new Set(soList.map(s=>String(s).trim()));
  const out = [];
  for (let i=0;i<vals.length;i++){
    const r = vals[i];
    if (String(r[H['CustomerID']-1]||'').trim() !== customerId) continue;
    if (String(r[H['DocType']-1]||'').toUpperCase() !== 'RECEIPT') continue;
    const csv = String(r[H['SOsCSV']-1]||'');
    const any = csv.split(',').map(s=>s.trim()).some(s => set.has(s));
    if (!any) continue;
    out.push({ date: r[H['PaymentDateTime']-1] || r[H['DOC_DATE']-1] || new Date(),
               amount: num_(r[H['AmountGross']-1],0),
               method: String(r[H['Method']-1]||'') });
  }
  out.sort((a,b)=> (new Date(a.date)) - (new Date(b.date)));
  return out;
}

function computePaidToDate_(customerId, soList){
  const sh = ensureLedger_(); const lr=sh.getLastRow(), lc=sh.getLastColumn(); if (lr<2) return 0;
  const H = headerMap_(sh.getRange(1,1,1,lc).getDisplayValues()[0]);
  const vals = sh.getRange(2,1,lr-1,lc).getValues();
  const set = new Set(soList.map(s=>String(s).trim()));
  let sum=0;
  for (let i=0;i<vals.length;i++){
    const r = vals[i];
    if (String(r[H['CustomerID']-1]||'').trim() !== customerId) continue;
    if (String(r[H['DocType']-1]||'').toUpperCase() !== 'RECEIPT') continue;
    const csv = String(r[H['SOsCSV']-1]||'');
    const any = csv.split(',').map(s=>s.trim()).some(s => set.has(s));
    if (!any) continue;
    sum += num_(r[H['AmountGross']-1],0);
  }
  return round2_(sum);
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
function pickList_(csv, d){
  const defaults = Array.isArray(d) ? d : [];
  const custom = String(csv||'')
    .split(',')
    .map(s=>s.trim())
    .filter(Boolean);
  const seen = new Set();
  const list = [];
  custom.concat(defaults).forEach(label=>{
    if (!seen.has(label)) {
      seen.add(label);
      list.push(label);
    }
  });
  return list.length ? list : defaults;
}
function hIndex_(hdr){ const H={}; (hdr||[]).forEach((h,i)=>{ const k=String(h||'').trim(); if (k) H[k]=i+1; }); return H; }
function pickH_(H, names){ for (const n of (names||[])) if (H[n]) return H[n]; return 0; }
function headerMap_(hdrRow){ const m={}; hdrRow.forEach((h,i)=>{ m[String(h||'').trim()] = i+1; }); return m; }
function setIf_(row,H,key,val){ if (H[key]) row[H[key]-1] = val; }

function numFromLedger_(value){
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return isFinite(value) ? value : null;
  if (value instanceof Date) return null;
  const text = String(value).trim();
  if (!text) return null;
  const cleaned = text.replace(/[^0-9.\-]/g, '');
  if (!cleaned) return null;
  const parsed = parseFloat(cleaned);
  return isFinite(parsed) ? parsed : null;
}

function sumLedgerLines_(linesValue){
  if (linesValue === null || linesValue === undefined) return null;
  let parsed = linesValue;
  if (typeof parsed === 'string') {
    const trimmed = parsed.trim();
    if (!trimmed) return null;
    try {
      parsed = JSON.parse(trimmed);
    } catch (_) {
      return null;
    }
  }

  if (!Array.isArray(parsed)) return null;

  let total = 0;
  let hasValue = false;

  parsed.forEach(line => {
    if (!line || typeof line !== 'object') return;

    const explicit = numFromLedger_(line.total || line.lineTotal || line.extended || line.extendedAmount || line.amountTotal);
    if (explicit !== null) {
      total += explicit;
      hasValue = true;
      return;
    }

    const qty = numFromLedger_(line.qty ?? line.quantity ?? line.qtyOrdered ?? line.qtyBilled ?? line.units);
    const amt = numFromLedger_(line.amt ?? line.amount ?? line.rate ?? line.price ?? line.unitPrice);

    let lineTotal = null;
    if (qty !== null && amt !== null) {
      lineTotal = qty * amt;
    } else if (amt !== null) {
      lineTotal = amt;
    }

    if (lineTotal !== null) {
      total += lineTotal;
      hasValue = true;
    }
  });

  return hasValue ? total : null;
}
function normalizeSo_(value){
  const raw = String(value||'').trim().toUpperCase();
  if (!raw) return '';
  return raw.replace(/[^A-Z0-9]/g, '');
}

function normalizeId_(value){
  const raw = String(value||'').trim().toUpperCase();
  if (!raw) return '';
  return raw.replace(/[^A-Z0-9]/g, '');
}

function soEq_(a,b){
  const na = normalizeSo_(a);
  const nb = normalizeSo_(b);
  if (!na || !nb) return false;
  return na === nb;
}

function idEq_(a,b){
  const na = normalizeId_(a);
  const nb = normalizeId_(b);
  if (!na || !nb) return false;
  return na === nb;
}
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
function makeTable_(headers, rows, opts){
  const includeHeader = !(opts && opts.includeHeader === false);
  const cols = (Array.isArray(headers) && headers.length) ? headers.length : ((rows && rows[0] && rows[0].length) || 1);
  const temp = DocumentApp.create('tmp-tbl');
  const body = temp.getBody();
  const tb = body.appendTable([new Array(cols).fill('')]);

  const data = Array.isArray(rows) ? rows : [];
  let dataIndex = 0;

  if (includeHeader) {
    const headerRow = tb.getRow(0);
    for (let i=0;i<cols;i++) {
      const val = (Array.isArray(headers) && headers[i] != null) ? String(headers[i]) : '';
      headerRow.getChild(i).asTableCell().setText(val);
    }
  } else if (data.length) {
    const first = tb.getRow(0);
    for (let i=0;i<cols;i++) {
      const val = (data[0] && data[0][i] != null) ? String(data[0][i]) : '';
      first.getChild(i).asTableCell().setText(val);
    }
    dataIndex = 1;
  }

  for (let rIndex = dataIndex; rIndex < data.length; rIndex++) {
    const tr = tb.appendTableRow();
    for (let i=0;i<cols;i++) {
      const cellVal = (data[rIndex] && data[rIndex][i] != null) ? data[rIndex][i] : '';
      tr.appendTableCell(String(cellVal));
    }
  }

  const copy = tb.copy();
  const id = temp.getId();
  temp.saveAndClose();
  DriveApp.getFileById(id).setTrashed(true);
  return copy;
}

function normalizeCustomerId_(id){
  return String(id||'')
    .replace(/\u00a0/g, ' ')
    .trim()
    .toUpperCase();
}

function customerIdsEqual_(a, b){
  return normalizeCustomerId_(a) === normalizeCustomerId_(b);
}
