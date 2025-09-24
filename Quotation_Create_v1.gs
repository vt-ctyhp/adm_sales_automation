/**
 * Create Quotation flow — dialog + document writer
 */

const QC_SP = (typeof SP !== 'undefined' && SP) ? SP : PropertiesService.getScriptProperties();

const QC_MASTER_TAB_NAMES = (function () {
  if (typeof WH_ORDERS_TAB_NAMES !== 'undefined' && Array.isArray(WH_ORDERS_TAB_NAMES) && WH_ORDERS_TAB_NAMES.length) {
    return WH_ORDERS_TAB_NAMES.slice();
  }
  const csv = QC_SP && QC_SP.getProperty ? QC_SP.getProperty('WH_ORDERS_TAB_NAMES_CSV') : '';
  const list = String(csv || '').split(',').map(s => s.trim()).filter(Boolean);
  return list.length ? list : ['00_Master Wholesale'];
})();

const QC_CRM_TAB_NAME = (typeof WH_CRM_TAB_NAME !== 'undefined' && WH_CRM_TAB_NAME)
  ? WH_CRM_TAB_NAME
  : ((QC_SP && QC_SP.getProperty ? QC_SP.getProperty('WH_CRM_TAB_NAME') : '') || '01_CRM');

const QC_SO_ALIASES = pickList_(QC_SP && QC_SP.getProperty ? QC_SP.getProperty('WH_SO_COL_ALIASES') : '',
  ['SO#', 'SO', 'SO Number', 'Sales Order', 'Sales Order #']);
const QC_CUSTOMER_ID_ALIASES = pickList_(QC_SP && QC_SP.getProperty ? QC_SP.getProperty('WH_CUSTID_COL_ALIASES') : '',
  ['Customer (Company) ID', 'Customer ID', 'CustomerID', 'ClientID', 'Account Code']);
const QC_BUSINESS_NAME_ALIASES = pickList_(QC_SP && QC_SP.getProperty ? QC_SP.getProperty('WH_COMPANY_COL_ALIASES') : '',
  ['Business Name', 'Company Name', 'Customer', 'Customer Name']);
const QC_CONTACT_NAME_ALIASES = pickList_(QC_SP && QC_SP.getProperty ? QC_SP.getProperty('WH_CONTACT_COL_ALIASES') : '',
  ['Contact Name', 'Primary Contact', 'Main Contact', 'Contact', 'Attn', 'Attention']);
const QC_CONTACT_FIRST_ALIASES = ['Contact First Name', 'Contact First', 'Primary Contact First Name', 'Primary Contact First', 'Contact FirstName'];
const QC_CONTACT_LAST_ALIASES = ['Contact Last Name', 'Contact Last', 'Primary Contact Last Name', 'Primary Contact Last', 'Contact LastName'];
const QC_EMAIL_ALIASES = ['Contact Email', 'Email', 'EmailLower'];
const QC_PHONE_ALIASES = ['Contact Phone', 'Phone', 'PhoneNorm'];
const QC_PRODUCT_DESC_ALIASES = ['Product Description', 'Product', 'Product Name', 'Setting Description', 'Design Request', '3D Design Request'];
const QC_PRODUCT_DETAILS_ALIASES = ['Product Details', 'Design Notes', 'Ring Style', 'Metal', 'US Size', 'Center Type', 'Diamond Dimension'];
const QC_QUANTITY_ALIASES = ['Quantity', 'Qty'];
const QC_TRACKER_ALIASES = ['Customer Order Tracker URL', 'Order Tracker URL', 'Tracker URL'];
const QC_QUOTATION_URL_ALIASES = ['Quotation URL'];
const QC_CRM_EMAIL_ALIASES = ['Contact Email', 'Email', 'EmailLower'];
const QC_CRM_PHONE_ALIASES = ['Contact Phone', 'Phone', 'PhoneNorm'];

const QC_MONEY_HEADERS = ['V1 Quotation', 'V2 Quotation', 'Approved Price'];

function qc_openCreateQuotation() {
  const html = HtmlService.createHtmlOutputFromFile('dlg_create_quotation_v1')
    .setWidth(1100)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Quotation');
}

function qc_init() {
  const t = ADM_time && ADM_time('qc_init');
  try {
    const ss = SpreadsheetApp.getActive();
    const activeSheet = ss.getActiveSheet();
    const allowedNames = QC_MASTER_TAB_NAMES && QC_MASTER_TAB_NAMES.length ? QC_MASTER_TAB_NAMES : [];
    if (!activeSheet || (allowedNames.length && allowedNames.indexOf(activeSheet.getName()) === -1)) {
      const label = allowedNames.length ? allowedNames.join(', ') : 'a wholesale orders tab';
      throw new Error('Please activate one of the wholesale order tabs (' + label + ') and select a data row.');
    }
    const master = activeSheet;
    const rng = master.getActiveRange();
    if (!rng) throw new Error('Select a row first.');
    const rowIndex = rng.getRow();
    if (rowIndex <= 1) throw new Error('Select a data row (below the header).');

    const lastCol = master.getLastColumn();
    const headerRow = master.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(s => String(s || '').trim());
    const H = hIndex_(headerRow);

    const rowRange = master.getRange(rowIndex, 1, 1, lastCol);
    const rowValues = rowRange.getValues()[0];
    const rowDisplay = rowRange.getDisplayValues()[0];
    const rowRich = rowRange.getRichTextValues()[0];

    const ctx = qc_buildContext_(ss, master, master.getName(), H, rowIndex, rowDisplay, rowValues);
    const product = qc_buildProductSnapshot_(H, rowDisplay, rowValues, ctx.trackerUrl);
    const money = qc_buildMoneyPrefill_(H, rowDisplay, rowValues);
    const links = qc_buildLinks_(H, rowDisplay, rowRich);
    const known = qc_collectKnownSOs_(ss, ctx, product);

    const payload = {
      context: {
        rowIndex,
        sheetId: master.getSheetId(),
        sheetName: master.getName(),
        masterUrl: ss.getUrl(),
        SO: ctx.SO,
        customerName: ctx.customerName,
        businessName: ctx.businessName,
        customerId: ctx.customerId,
        contactName: ctx.contactName,
        emailLower: ctx.emailLower,
        emailDisplay: ctx.emailDisplay,
        phoneNorm: ctx.phoneNorm,
        phoneDisplay: ctx.phoneDisplay,
        trackerUrl: ctx.trackerUrl
      },
      productSnapshot: product,
      knownSOs: known,
      money,
      links,
      ui: {
        currencyHint: '$1,234.00',
        knownSoCount: known.length
      }
    };

    ADM_dbg && ADM_dbg('qc_init payload', payload);
    return payload;
  } finally {
    try { t && t(); } catch (ignored) {}
  }
}

function qc_submit(payload) {
  const timer = ADM_time && ADM_time('qc_submit');
  try {
    if (!payload || typeof payload !== 'object') throw new Error('Missing payload.');
    const ctx = payload.context || {};
    const rowIndex = Number(ctx.rowIndex || payload.rowIndex || 0);
    if (!rowIndex) throw new Error('Row index missing.');

    const items = Array.isArray(payload.items) ? payload.items : [];
    if (!items.length) throw new Error('Add at least one line item.');

    const cleanItems = [];
    const selectedSOs = [];
    items.forEach((item, idx) => {
      const qty = Math.max(1, Math.floor(Number(item.qty || item.quantity || 0)));
      if (!qty || !isFinite(qty)) throw new Error('Line ' + (idx + 1) + ' has an invalid quantity.');
      const so = String(item.so || '').trim();
      if (so && selectedSOs.indexOf(so) === -1) selectedSOs.push(so);
      cleanItems.push({
        so,
        productDescription: String(item.productDescription || '').trim(),
        productDetails: String(item.productDetails || '').trim(),
        qty
      });
    });

    const pricing = payload.pricing || {};
    const v1 = qc_parseMoney_(pricing.v1);
    if (v1 === '' || v1 < 0) throw new Error('“V1 Quotation” must be a number greater than or equal to 0.');
    const v2 = qc_parseMoney_(pricing.v2);
    const approved = qc_parseMoney_(pricing.approved);

    const ss = SpreadsheetApp.getActive();
    const master = qc_resolveOrdersSheet_(ss, ctx.sheetId, ctx.sheetName);
    const lastCol = master.getLastColumn();
    const headerRow = master.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(s => String(s || '').trim());
    const H = hIndex_(headerRow);

    const now = new Date();
    const tz = Session.getScriptTimeZone() || 'America/Los_Angeles';
    const todayIso = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const todayPretty = Utilities.formatDate(now, tz, 'MMMM d, yyyy');

    let docUrl = '';
    if (!payload.saveOnly) {
      const templateId = quotationTemplateIdForBrand_('');
      if (!templateId) {
        throw new Error('No quotation template configured. Set QUOTATION_TEMPLATE_ID_DEFAULT in Script Properties.');
      }

      const existingLink = qc_extractLinkFromSheet_(master, H, rowIndex);
      const filename = qc_buildQuotationFilename_(ctx.businessName || ctx.customerName, selectedSOs, ctx.customerId);
      const selectedLabel = selectedSOs.length ? selectedSOs.join(', ') : (ctx.SO || '');
      const placeholders = {
        BRAND: '',
        BUSINESS_NAME: ctx.businessName || ctx.customerName || '',
        CUSTOMER_BUSINESS_NAME: ctx.businessName || ctx.customerName || '',
        CUSTOMER_NAME: ctx.businessName || ctx.customerName || '',
        CUSTOMER_ID: ctx.customerId || '',
        CONTACT_NAME: ctx.contactName || '',
        CUSTOMER_EMAIL: ctx.emailLower || '',
        CUSTOMER_EMAIL_DISPLAY: ctx.emailDisplay || ctx.emailLower || '',
        CUSTOMER_PHONE: ctx.phoneNorm || '',
        CUSTOMER_PHONE_DISPLAY: ctx.phoneDisplay || ctx.phoneNorm || '',
        ROOT_APPT_ID: '',
        TRACKER_URL: ctx.trackerUrl || '',
        SELECTED_SOS: selectedLabel,
        DATE_TODAY: todayPretty,
        V1_QUOTATION: money_(v1),
        V2_QUOTATION: v2 === '' ? '' : money_(v2),
        APPROVED_PRICE: approved === '' ? '' : money_(approved)
      };

      const doc = ensureAndFillQuotationForRow_({
        rowIndex,
        templateId,
        filename,
        placeholders,
        items: cleanItems,
        existingUrl: existingLink && existingLink.url
      });
      docUrl = doc && doc.url ? doc.url : '';
    }

    const updates = {
      'V1 Quotation': v1,
      'V2 Quotation': v2 === '' ? '' : v2,
      'Approved Price': approved === '' ? '' : approved
    };
    if (docUrl) {
      updates['Quotation URL'] = docUrl;
      updates['Quotation Date'] = todayIso;
    }

    qc_writeBack_(master, H, rowIndex, updates);

    const result = { ok: true, url: docUrl, rowIndex };
    ADM_dbg && ADM_dbg('qc_submit result', result);
    return result;
  } catch (err) {
    ADM_dbg && ADM_dbg('qc_submit error', String(err && err.stack || err));
    throw err;
  } finally {
    try { timer && timer(); } catch (ignored) {}
  }
}

function qc_buildContext_(ss, sheet, sheetName, H, rowIndex, rowDisplay, rowValues) {
  const get = (aliases) => qc_pickFirst_(H, aliases, rowDisplay, rowValues);
  const so = get(QC_SO_ALIASES);
  const businessName = get(QC_BUSINESS_NAME_ALIASES);
  const customerId = get(QC_CUSTOMER_ID_ALIASES);
  let contactName = get(QC_CONTACT_NAME_ALIASES);
  if (!contactName) {
    const first = get(QC_CONTACT_FIRST_ALIASES);
    const last = get(QC_CONTACT_LAST_ALIASES);
    contactName = [first, last].filter(Boolean).join(' ').trim();
  }
  const trackerFromRow = get(QC_TRACKER_ALIASES);
  const crm = qc_lookupCrmRow_(ss, customerId, businessName);

  let emailDisplay = '';
  let emailLower = '';
  let phoneDisplay = '';
  let phoneNorm = '';
  let trackerUrl = trackerFromRow || '';

  if (crm) {
    emailDisplay = crm.emailDisplay || crm.emailLower || '';
    emailLower = (crm.emailLower || emailDisplay).toLowerCase();
    phoneDisplay = crm.phoneDisplay || crm.phoneNorm || '';
    phoneNorm = qc_normPhone_(crm.phoneNorm || crm.phoneDisplay || '');
    trackerUrl = trackerUrl || crm.trackerUrl || '';
    contactName = crm.contactName || contactName;
  }

  if (!emailLower) {
    const rawEmail = get(QC_EMAIL_ALIASES);
    emailDisplay = emailDisplay || rawEmail;
    emailLower = (rawEmail || '').toLowerCase();
  }
  if (!phoneNorm) {
    const rawPhone = get(QC_PHONE_ALIASES);
    phoneDisplay = phoneDisplay || rawPhone;
    phoneNorm = qc_normPhone_(rawPhone);
  }

  return {
    sheetId: sheet.getSheetId(),
    sheetName,
    rowIndex,
    SO: so,
    customerName: businessName || '',
    businessName: businessName || '',
    customerId: customerId || '',
    contactName: contactName || '',
    emailLower,
    emailDisplay: emailDisplay || emailLower || '',
    phoneNorm,
    phoneDisplay: phoneDisplay || phoneNorm || '',
    trackerUrl
  };
}

function qc_buildProductSnapshot_(H, rowDisplay, rowValues, trackerUrl) {
  const get = (aliases) => qc_pickFirst_(H, aliases, rowDisplay, rowValues);
  let productDescription = get(QC_PRODUCT_DESC_ALIASES);
  let productDetails = get(QC_PRODUCT_DETAILS_ALIASES);
  let quantityRaw = get(QC_QUANTITY_ALIASES);
  let quantity = Math.max(1, Math.round(num_(quantityRaw || 1, 1)) || 1);

  if (!productDescription || !productDetails) {
    const fallback = qc_fetchProductSnapshotFromTracker_(trackerUrl);
    if (fallback) {
      if (!productDescription && fallback.productDescription) productDescription = fallback.productDescription;
      if (!productDetails && fallback.productDetails) productDetails = fallback.productDetails;
      if ((!quantity || quantity === 1) && fallback.quantity) quantity = fallback.quantity;
    }
  }

  return {
    so: get(QC_SO_ALIASES),
    productDescription: productDescription || '',
    productDetails: productDetails || '',
    quantity
  };
}

function qc_buildMoneyPrefill_(H, rowDisplay, rowValues) {
  const out = {};
  QC_MONEY_HEADERS.forEach(label => {
    const col = pickH_(H, [label]);
    const raw = col ? rowValues[col - 1] : '';
    const display = col ? rowDisplay[col - 1] : '';
    const num = qc_parseMoney_(raw);
    out[label] = {
      number: num === '' ? '' : num,
      display: display || (num === '' ? '' : money_(num))
    };
  });
  return out;
}

function qc_buildLinks_(H, rowDisplay, rowRich) {
  const res = {};
  const col = pickH_(H, QC_QUOTATION_URL_ALIASES);
  if (col) {
    res.quotationUrl = qc_extractLink_(rowRich[col - 1], rowDisplay[col - 1]);
  }
  return res;
}

function qc_collectKnownSOs_(ss, ctx, productSnapshot) {
  const tabNames = QC_MASTER_TAB_NAMES && QC_MASTER_TAB_NAMES.length ? QC_MASTER_TAB_NAMES : ss.getSheets().map(s => s.getName());
  const targetId = String(ctx.customerId || '').trim().toLowerCase();
  const targetEmail = String(ctx.emailLower || '').trim();
  const targetPhone = String(ctx.phoneNorm || '').trim();
  const list = [];
  const seenSo = new Set();

  tabNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const lr = sheet.getLastRow();
    const lc = sheet.getLastColumn();
    if (lr <= 1 || lc < 1) return;
    const headers = sheet.getRange(1, 1, 1, lc).getDisplayValues()[0].map(s => String(s || '').trim());
    const H = hIndex_(headers);
    const soCol = pickH_(H, QC_SO_ALIASES);
    if (!soCol) return;

    const idCol = pickH_(H, QC_CUSTOMER_ID_ALIASES);
    const emailCol = pickH_(H, QC_EMAIL_ALIASES);
    const phoneCol = pickH_(H, QC_PHONE_ALIASES);
    const descCol = pickH_(H, QC_PRODUCT_DESC_ALIASES);
    const detailsCol = pickH_(H, QC_PRODUCT_DETAILS_ALIASES);
    const qtyCol = pickH_(H, QC_QUANTITY_ALIASES);

    const values = sheet.getRange(2, 1, lr - 1, lc).getValues();
    const display = sheet.getRange(2, 1, lr - 1, lc).getDisplayValues();

    for (let i = 0; i < values.length; i++) {
      const rowVals = values[i];
      const rowDisp = display[i];
      const so = String(rowDisp[soCol - 1] || rowVals[soCol - 1] || '').trim();
      if (!so) continue;
      const soKey = so.toLowerCase();
      const rowId = idCol ? String(rowVals[idCol - 1] || rowDisp[idCol - 1] || '').trim().toLowerCase() : '';
      const rowEmail = emailCol ? String(rowVals[emailCol - 1] || rowDisp[emailCol - 1] || '').trim().toLowerCase() : '';
      const rowPhone = phoneCol ? qc_normPhone_(rowVals[phoneCol - 1] || rowDisp[phoneCol - 1] || '') : '';

      let match = false;
      if (targetId && rowId && rowId === targetId) {
        match = true;
      } else if (targetEmail && rowEmail && rowEmail === targetEmail) {
        match = true;
      } else if (targetPhone && rowPhone && rowPhone === targetPhone) {
        match = true;
      }
      if (!match) continue;
      if (seenSo.has(soKey)) continue;
      seenSo.add(soKey);

      const productDescription = descCol ? String(rowDisp[descCol - 1] || rowVals[descCol - 1] || '') : '';
      const productDetails = detailsCol ? String(rowDisp[detailsCol - 1] || rowVals[detailsCol - 1] || '') : '';
      const qtyRaw = qtyCol ? (rowVals[qtyCol - 1] || rowDisp[qtyCol - 1] || '') : '';
      const qty = Math.max(1, Math.round(num_(qtyRaw || 1, 1)) || 1);
      const metaParts = [];
      if (productDescription) metaParts.push(productDescription);
      if (name !== ctx.sheetName) metaParts.push(name);
      list.push({
        key: name + '|' + (i + 2) + '|' + so,
        so,
        label: so,
        sheetName: name,
        rowIndex: i + 2,
        productDescription: productDescription || '',
        productDetails: productDetails || '',
        quantity: qty,
        meta: metaParts.join(' • ')
      });
    }
  });

  const activeSo = String(ctx.SO || '').trim();
  if (activeSo) {
    const key = activeSo.toLowerCase();
    if (!seenSo.has(key)) {
      seenSo.add(key);
      const metaParts = [];
      if (productSnapshot && productSnapshot.productDescription) {
        metaParts.push(productSnapshot.productDescription);
      }
      list.unshift({
        key: 'active|' + (ctx.sheetName || '') + '|' + (ctx.rowIndex || '') + '|' + activeSo,
        so: activeSo,
        label: activeSo,
        sheetName: ctx.sheetName || '',
        rowIndex: ctx.rowIndex || 0,
        productDescription: productSnapshot ? (productSnapshot.productDescription || '') : '',
        productDetails: productSnapshot ? (productSnapshot.productDetails || '') : '',
        quantity: productSnapshot ? (productSnapshot.quantity || 1) : 1,
        meta: metaParts.join(' • ')
      });
    } else {
      const existing = list.find(entry => entry.so === activeSo);
      if (existing && productSnapshot) {
        existing.productDescription = existing.productDescription || productSnapshot.productDescription || '';
        existing.productDetails = existing.productDetails || productSnapshot.productDetails || '';
        existing.quantity = existing.quantity || productSnapshot.quantity || 1;
        if (!existing.meta && productSnapshot.productDescription) {
          existing.meta = productSnapshot.productDescription;
        }
      }
    }
  }

  return list;
}

function qc_lookupCrmRow_(ss, customerId, businessName) {
  const tab = QC_CRM_TAB_NAME;
  if (!tab) return null;
  const sheet = ss.getSheetByName(tab);
  if (!sheet) return null;
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr <= 1 || lc < 1) return null;

  const headers = sheet.getRange(1, 1, 1, lc).getDisplayValues()[0].map(s => String(s || '').trim());
  const H = hIndex_(headers);
  const idCol = pickH_(H, QC_CUSTOMER_ID_ALIASES);
  const nameCol = pickH_(H, QC_BUSINESS_NAME_ALIASES);
  const emailCol = pickH_(H, QC_CRM_EMAIL_ALIASES);
  const phoneCol = pickH_(H, QC_CRM_PHONE_ALIASES);
  const contactCol = pickH_(H, QC_CONTACT_NAME_ALIASES);
  const contactFirstCol = pickH_(H, QC_CONTACT_FIRST_ALIASES);
  const contactLastCol = pickH_(H, QC_CONTACT_LAST_ALIASES);
  const trackerCol = pickH_(H, QC_TRACKER_ALIASES);

  if (!idCol && !nameCol) return null;

  const values = sheet.getRange(2, 1, lr - 1, lc).getValues();
  const display = sheet.getRange(2, 1, lr - 1, lc).getDisplayValues();

  const wantId = String(customerId || '').trim().toLowerCase();
  const wantName = String(businessName || '').trim().toLowerCase();

  let fallback = null;
  for (let i = 0; i < values.length; i++) {
    const rowVals = values[i];
    const rowDisp = display[i];
    const rowIdRaw = idCol ? String(rowVals[idCol - 1] || rowDisp[idCol - 1] || '').trim() : '';
    const rowNameRaw = nameCol ? String(rowVals[nameCol - 1] || rowDisp[nameCol - 1] || '').trim() : '';
    const rowId = rowIdRaw.toLowerCase();
    const rowName = rowNameRaw.toLowerCase();
    const matchId = wantId && rowId && rowId === wantId;
    const matchName = !matchId && wantName && rowName && rowName === wantName;
    if (matchId || matchName) {
      return qc_buildCrmLookupResult_(rowVals, rowDisp, {
        id: rowIdRaw,
        name: rowNameRaw,
        emailCol,
        phoneCol,
        contactCol,
        contactFirstCol,
        contactLastCol,
        trackerCol
      });
    }
    if (!fallback && wantName && rowName && rowName.includes(wantName)) {
      fallback = { index: i, id: rowIdRaw, name: rowNameRaw };
    }
  }

  if (fallback) {
    const rowVals = values[fallback.index];
    const rowDisp = display[fallback.index];
    return qc_buildCrmLookupResult_(rowVals, rowDisp, {
      id: fallback.id,
      name: fallback.name,
      emailCol,
      phoneCol,
      contactCol,
      contactFirstCol,
      contactLastCol,
      trackerCol
    });
  }
  return null;
}

function qc_buildCrmLookupResult_(rowVals, rowDisp, cols) {
  const emailDisplay = cols.emailCol ? String(rowDisp[cols.emailCol - 1] || rowVals[cols.emailCol - 1] || '').trim() : '';
  const emailLower = emailDisplay.toLowerCase();
  const phoneDisplay = cols.phoneCol ? String(rowDisp[cols.phoneCol - 1] || rowVals[cols.phoneCol - 1] || '').trim() : '';
  const phoneNorm = qc_normPhone_(phoneDisplay);
  let contactName = cols.contactCol ? String(rowDisp[cols.contactCol - 1] || rowVals[cols.contactCol - 1] || '').trim() : '';
  if (!contactName) {
    const first = cols.contactFirstCol ? String(rowDisp[cols.contactFirstCol - 1] || rowVals[cols.contactFirstCol - 1] || '').trim() : '';
    const last = cols.contactLastCol ? String(rowDisp[cols.contactLastCol - 1] || rowVals[cols.contactLastCol - 1] || '').trim() : '';
    contactName = [first, last].filter(Boolean).join(' ').trim();
  }
  const trackerUrl = cols.trackerCol ? String(rowDisp[cols.trackerCol - 1] || rowVals[cols.trackerCol - 1] || '').trim() : '';

  return {
    customerId: String(cols.id || '').trim(),
    businessName: String(cols.name || '').trim(),
    contactName,
    emailDisplay,
    emailLower,
    phoneDisplay,
    phoneNorm,
    trackerUrl
  };
}

function qc_fetchProductSnapshotFromTracker_(trackerUrl) {
  try {
    if (!trackerUrl) return null;
    const fileId = fileIdFromUrl_(trackerUrl);
    if (!fileId) return null;
    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheets()[0];
    if (!sheet) return null;
    const lr = sheet.getLastRow();
    const lc = sheet.getLastColumn();
    if (lr < 2 || lc < 1) return null;
    const hdr = sheet.getRange(1, 1, 1, lc).getDisplayValues()[0].map(s => String(s || '').trim());
    const H = hIndex_(hdr);
    const row = sheet.getRange(lr, 1, 1, lc).getDisplayValues()[0];
    const get = (aliases) => {
      for (let i = 0; i < aliases.length; i++) {
        const col = pickH_(H, [aliases[i]]);
        if (col) return row[col - 1];
      }
      return '';
    };
    const quantityRaw = get(QC_QUANTITY_ALIASES);
    return {
      productDescription: get(QC_PRODUCT_DESC_ALIASES) || '',
      productDetails: get(QC_PRODUCT_DETAILS_ALIASES) || '',
      quantity: Math.max(1, Math.round(num_(quantityRaw || 1, 1)) || 1)
    };
  } catch (err) {
    ADM_dbg && ADM_dbg('qc_fetchProductSnapshotFromTracker_ error', String(err && err.stack || err));
    return null;
  }
}

function qc_extractLink_(rich, display) {
  if (rich) {
    const runs = rich.getRuns();
    for (let i = 0; i < runs.length; i++) {
      const run = runs[i];
      const url = run.getLinkUrl();
      if (url) {
        return { text: run.getText(), url };
      }
    }
    const url = rich.getLinkUrl && rich.getLinkUrl();
    if (url) return { text: display || '', url };
  }
  const url = (display && /https?:\/\//i.test(display)) ? display : '';
  return { text: display || '', url };
}

function qc_extractLinkFromSheet_(sheet, H, rowIndex) {
  const col = pickH_(H, QC_QUOTATION_URL_ALIASES);
  if (!col) return null;
  const cell = sheet.getRange(rowIndex, col);
  const rich = cell.getRichTextValue();
  const display = cell.getDisplayValue();
  return qc_extractLink_(rich, display);
}

function qc_resolveOrdersSheet_(ss, sheetId, sheetName) {
  const id = Number(sheetId || 0);
  if (id) {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const sh = sheets[i];
      if (sh && sh.getSheetId && sh.getSheetId() === id) return sh;
    }
  }
  if (sheetName) {
    const byName = ss.getSheetByName(String(sheetName));
    if (byName) return byName;
  }
  const names = QC_MASTER_TAB_NAMES && QC_MASTER_TAB_NAMES.length ? QC_MASTER_TAB_NAMES : [];
  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  throw new Error('Unable to locate the wholesale orders sheet for the selected row.');
}

function qc_writeBack_(sheet, H, rowIndex, updates) {
  if (!updates) return;
  let maxCol = sheet.getLastColumn();
  Object.keys(updates).forEach(header => {
    if (!header) return;
    const col = pickH_(H, [header]);
    if (!col) {
      maxCol += 1;
      sheet.getRange(1, maxCol).setValue(header);
      H[header] = maxCol;
    }
  });
  const lc = sheet.getLastColumn();
  const rowRange = sheet.getRange(rowIndex, 1, 1, lc);
  const rowValues = rowRange.getValues()[0];
  Object.keys(updates).forEach(header => {
    if (header === 'Quotation URL') return;
    const col = pickH_(H, [header]);
    if (col) rowValues[col - 1] = updates[header];
  });
  rowRange.setValues([rowValues]);

  if (updates['Quotation URL']) {
    const col = pickH_(H, ['Quotation URL']);
    if (col) {
      const cell = sheet.getRange(rowIndex, col);
      const url = String(updates['Quotation URL']);
      if (url) {
        const rt = SpreadsheetApp.newRichTextValue().setText('Open').setLinkUrl(url).build();
        cell.setRichTextValue(rt);
      } else {
        cell.clearContent();
      }
    }
  }
}

function qc_normPhone_(phone) {
  const digits = String(phone || '').replace(/\D+/g, '');
  return digits || '';
}

function qc_pickFirst_(H, aliases, displayRow, valueRow) {
  return qc_pickFirstRow_(H, aliases, displayRow, valueRow) || '';
}

function qc_pickFirstRow_(H, aliases, displayRow, valueRow) {
  if (!aliases) return '';
  for (let i = 0; i < aliases.length; i++) {
    const col = pickH_(H, [aliases[i]]);
    if (col) {
      const val = valueRow[col - 1];
      const display = displayRow[col - 1];
      if (val != null && val !== '') return String(val);
      if (display != null && display !== '') return String(display);
    }
  }
  return '';
}

function quotationTemplateIdForBrand_(brand) {
  const sp = PropertiesService.getScriptProperties();
  const key = 'QUOTATION_TEMPLATE_ID_' + String(brand || '').trim().toUpperCase();
  const fallbackKey = 'QUOTATION_TEMPLATE_ID_DEFAULT';
  return (sp.getProperty(key) || sp.getProperty(fallbackKey) || '').trim();
}

function ensureAndFillQuotationForRow_(opts) {
  if (!opts || !opts.templateId) throw new Error('Missing quotation template ID.');
  const templateFile = DriveApp.getFileById(opts.templateId);
  const filename = opts.filename || ('Quotation ' + new Date().toISOString());

  let docFile;
  if (opts.existingUrl) {
    try {
      const existingId = fileIdFromUrl_(opts.existingUrl);
      if (existingId) docFile = DriveApp.getFileById(existingId);
    } catch (err) {
      ADM_dbg && ADM_dbg('ensureAndFillQuotationForRow_ existing file error', String(err));
    }
  }
  if (!docFile) {
    docFile = templateFile.makeCopy(filename);
  } else {
    docFile.setName(filename);
  }

  const doc = DocumentApp.openById(docFile.getId());
  const body = doc.getBody();
  const placeholders = opts.placeholders || {};
  Object.keys(placeholders).forEach(key => {
    const needle = '{{' + key + '}}';
    body.replaceText(qc_escapeForRegex_(needle), String(placeholders[key] == null ? '' : placeholders[key]));
  });

  if (Array.isArray(opts.items) && opts.items.length) {
    const tableHeader = ['SO#', 'Product Description', 'Product Details', 'Qty'];
    const tableData = [tableHeader.slice()];
    opts.items.forEach(item => {
      tableData.push([
        item.so || '',
        item.productDescription || '',
        item.productDetails || '',
        String(item.qty || '')
      ]);
    });
    const placeholder = body.findText('{{ITEMS_TABLE}}');
    if (placeholder) {
      const el = placeholder.getElement();
      const parent = el.getParent();
      const index = body.getChildIndex(parent);
      body.insertTable(index, tableData);
      parent.removeFromParent();
    } else {
      const existing = qc_findQuotationItemsTable_(body, tableHeader);
      if (existing) {
        const index = body.getChildIndex(existing);
        body.removeChild(existing);
        body.insertTable(index, tableData);
      } else {
        body.appendTable(tableData);
      }
    }
  }

  doc.saveAndClose();
  return { id: doc.getId(), url: doc.getUrl() };
}

function qc_escapeForRegex_(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function qc_buildQuotationFilename_(businessName, selectedSOs, customerId) {
  const parts = [];
  if (businessName) parts.push(businessName);
  const firstSo = selectedSOs && selectedSOs.length ? selectedSOs[0] : '';
  if (firstSo) {
    parts.push(firstSo);
  } else if (customerId) {
    parts.push(customerId);
  }
  parts.push('Quotation');
  return parts.filter(Boolean).join(' – ');
}

function qc_findQuotationItemsTable_(body, headerRow) {
  const childCount = body.getNumChildren();
  for (let i = 0; i < childCount; i++) {
    const child = body.getChild(i);
    if (child.getType && child.getType() === DocumentApp.ElementType.TABLE) {
      const table = child.asTable();
      if (table.getNumRows() === 0) continue;
      const firstRow = table.getRow(0);
      if (firstRow.getNumCells() < headerRow.length) continue;
      let matches = true;
      for (let c = 0; c < headerRow.length; c++) {
        const text = firstRow.getCell(c).getText();
        if (String(text || '').trim() !== String(headerRow[c] || '').trim()) {
          matches = false;
          break;
        }
      }
      if (matches) return table;
    }
  }
  return null;
}

function qc_parseMoney_(value) {
  if (value === '' || value == null) return '';
  const n = parseFloat(String(value).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? n : '';
}
