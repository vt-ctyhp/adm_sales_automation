/**
 * Create Quotation flow — dialog + document writer
 */

const QC_MASTER_TAB_NAME = '00_Master Appointments';

const QC_SO_ALIASES = ['SO#', 'SO', 'SO Number'];
const QC_CUSTOMER_NAME_ALIASES = ['Customer Name', 'Customer', 'Client Name'];
const QC_EMAIL_ALIASES = ['EmailLower', 'Email'];
const QC_PHONE_ALIASES = ['PhoneNorm', 'Phone'];
const QC_BRAND_ALIASES = ['Brand', 'Company'];
const QC_ROOT_APPT_ALIASES = ['RootApptID', 'Root Appt ID', 'Root Appointment ID', 'Appt ID', 'ApptID'];
const QC_PRODUCT_DESC_ALIASES = ['Product Description', 'Product', 'Product Name', 'Setting Description', 'Design Request', '3D Design Request'];
const QC_PRODUCT_DETAILS_ALIASES = ['Product Details', 'Design Notes', 'Ring Style', 'Metal', 'US Size', 'Center Type', 'Diamond Dimension'];
const QC_QUANTITY_ALIASES = ['Quantity', 'Qty'];
const QC_TRACKER_ALIASES = ['Customer Order Tracker URL', 'Order Tracker URL', 'Tracker URL'];
const QC_QUOTATION_URL_ALIASES = ['Quotation URL'];

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
    const master = ss.getSheetByName(QC_MASTER_TAB_NAME);
    if (!master) {
      throw new Error('Sheet "' + QC_MASTER_TAB_NAME + '" not found.');
    }
    const activeSheet = ss.getActiveSheet();
    if (!activeSheet || activeSheet.getSheetId() !== master.getSheetId()) {
      throw new Error('Please activate the "' + QC_MASTER_TAB_NAME + '" sheet and select a data row.');
    }
    const rng = activeSheet.getActiveRange();
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

    const ctx = qc_buildContext_(ss, master, H, rowIndex, rowDisplay, rowValues);
    const product = qc_buildProductSnapshot_(H, rowDisplay, rowValues, ctx.trackerUrl);
    const money = qc_buildMoneyPrefill_(H, rowDisplay, rowValues);
    const links = qc_buildLinks_(H, rowDisplay, rowRich);
    const known = qc_collectKnownSOs_(master, H, ctx.emailLower, ctx.phoneNorm);

    const payload = {
      context: {
        rowIndex,
        sheetId: master.getSheetId(),
        masterUrl: ss.getUrl(),
        brand: ctx.brand,
        RootApptID: ctx.RootApptID,
        SO: ctx.SO,
        customerName: ctx.customerName,
        emailLower: ctx.emailLower,
        phoneNorm: ctx.phoneNorm
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
      if (so) selectedSOs.push(so);
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
    const master = ss.getSheetByName(QC_MASTER_TAB_NAME);
    if (!master) throw new Error('Sheet "' + QC_MASTER_TAB_NAME + '" not found.');
    const lastCol = master.getLastColumn();
    const headerRow = master.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(s => String(s || '').trim());
    const H = hIndex_(headerRow);

    const now = new Date();
    const tz = Session.getScriptTimeZone() || 'America/Los_Angeles';
    const todayIso = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const todayPretty = Utilities.formatDate(now, tz, 'MMMM d, yyyy');

    let docUrl = '';
    if (!payload.saveOnly) {
      const templateId = quotationTemplateIdForBrand_(ctx.brand);
      if (!templateId) throw new Error('No quotation template configured for brand "' + (ctx.brand || 'Unknown') + '".');

      const existingLink = qc_extractLinkFromSheet_(master, H, rowIndex);
      const filename = qc_buildQuotationFilename_(ctx.brand, selectedSOs, ctx.RootApptID);
      const placeholders = {
        BRAND: ctx.brand || '',
        CUSTOMER_NAME: ctx.customerName || '',
        CUSTOMER_EMAIL: ctx.emailLower || '',
        CUSTOMER_PHONE: ctx.phoneNorm || '',
        ROOT_APPT_ID: ctx.RootApptID || '',
        SELECTED_SOS: selectedSOs.join(', '),
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

function qc_buildContext_(ss, sheet, H, rowIndex, rowDisplay, rowValues) {
  const get = (aliases) => qc_pickFirst_(H, aliases, rowDisplay, rowValues);
  const brand = get(QC_BRAND_ALIASES);
  const so = get(QC_SO_ALIASES);
  const customerName = get(QC_CUSTOMER_NAME_ALIASES);
  const emailLower = (get(QC_EMAIL_ALIASES) || '').toLowerCase();
  const phoneNorm = qc_normPhone_(get(QC_PHONE_ALIASES));
  const rootApptId = get(QC_ROOT_APPT_ALIASES);
  const trackerUrl = get(QC_TRACKER_ALIASES);

  return {
    sheetId: sheet.getSheetId(),
    rowIndex,
    brand,
    SO: so,
    customerName,
    emailLower,
    phoneNorm,
    RootApptID: rootApptId,
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

function qc_collectKnownSOs_(sheet, H, emailLower, phoneNorm) {
  const lr = sheet.getLastRow();
  if (lr <= 1) return [];
  const lc = sheet.getLastColumn();
  const range = sheet.getRange(2, 1, lr - 1, lc);
  const vals = range.getValues();
  const display = range.getDisplayValues();
  const emailCol = pickH_(H, QC_EMAIL_ALIASES);
  const phoneCol = pickH_(H, QC_PHONE_ALIASES);
  const brandCol = pickH_(H, QC_BRAND_ALIASES);
  const soCol = pickH_(H, QC_SO_ALIASES);

  const seen = new Set();
  const list = [];
  for (let i = 0; i < vals.length; i++) {
    const rowVals = vals[i];
    const rowDisp = display[i];
    const email = emailCol ? String(rowVals[emailCol - 1] || rowDisp[emailCol - 1] || '').toLowerCase() : '';
    const phone = phoneCol ? qc_normPhone_(rowVals[phoneCol - 1] || rowDisp[phoneCol - 1] || '') : '';
    if (emailLower && email && emailLower === email) {
      // pass
    } else if (phoneNorm && phone && phoneNorm === phone) {
      // pass
    } else {
      continue;
    }
    const so = soCol ? String(rowDisp[soCol - 1] || '').trim() : '';
    if (!so) continue;
    const brand = brandCol ? String(rowDisp[brandCol - 1] || '').trim() : '';
    const key = (brand || '') + '|' + so;
    if (seen.has(key)) continue;
    seen.add(key);
    const productDescription = qc_pickFirstRow_(H, QC_PRODUCT_DESC_ALIASES, rowDisp, rowVals);
    const productDetails = qc_pickFirstRow_(H, QC_PRODUCT_DETAILS_ALIASES, rowDisp, rowVals);
    const qtyRaw = qc_pickFirstRow_(H, QC_QUANTITY_ALIASES, rowDisp, rowVals);
    const qty = Math.max(1, Math.round(num_(qtyRaw || 1, 1)) || 1);
    list.push({
      key,
      brand,
      so,
      label: [brand, so].filter(Boolean).join(' • '),
      productDescription: productDescription || '',
      productDetails: productDetails || '',
      quantity: qty
    });
  }
  return list;
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

function qc_buildQuotationFilename_(brand, selectedSOs, rootApptId) {
  const label = selectedSOs && selectedSOs.length ? selectedSOs[0] : (rootApptId || '');
  return [brand || 'Brand', label || 'Row', 'Quotation'].filter(Boolean).join(' – ');
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
