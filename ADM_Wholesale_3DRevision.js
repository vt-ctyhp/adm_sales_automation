/*** ADM Wholesale — 3D Revision Request ***/

const REV3D_ACTION = '3D Revision Requested';
const REV3D_MODE = '3D Revision Request';
const REV3D_COPY_HEADER = '— 3D REVISION REQUEST — START 3D / UPDATE SO —';

function open3DRevision(){
  const t = HtmlService.createTemplateFromFile('dlg_revision3d_v1');
  const today = Utilities.formatDate(new Date(), ADM_TZ, 'yyyy-MM-dd');
  t.BOOTSTRAP = {
    today,
    productTypes: ['Ring','Pendant','Chain','Earrings','Bracelet'],
    priorityLevels: ['', 'P1', 'P2'],
    metals: ['14K YG','18K YG','14K WG','18K WG','14K RG','18K RG','Pt950','Silver'],
    ringStyles: ['Solitaire','Pavé','Cathedral','Halo','Bezel','Knife-edge','Vintage'],
    accentTypes: ['Natural','Lab-grown','None'],
    ringSizes: ['4.5','5','5.5','6','6.5','7','7.5','8'],
    shapes: ['Round','Oval','Cushion','Emerald','Radiant','Pear','Marquise','Princess'],
    chainStyles: ['Curb','Cable','Franco','Rope','Box'],
    earringTypes: ['Studs','Hoops','Halo','Drop'],
    braceletTypes: ['Tennis','Curb','Bangle'],
    uploadTypes: ['Product Image','Reference','3D File','Design Sketch','Other']
  };
  const html = t.evaluate().setWidth(700).setHeight(640).setTitle('3D Revision Request');
  SpreadsheetApp.getUi().showModalDialog(html, '3D Revision Request');
}

function getCellLinkOrText_(sheet, row, colIndex){
  if (!sheet || !colIndex) return '';
  try {
    const rng = sheet.getRange(row, colIndex);
    const rich = rng.getRichTextValue();
    if (rich) {
      const direct = rich.getLinkUrl && rich.getLinkUrl();
      if (direct) return direct;
      if (rich.getRuns) {
        const runs = rich.getRuns();
        for (let i = 0; i < runs.length; i++) {
          const u = runs[i].getLinkUrl && runs[i].getLinkUrl();
          if (u) return u;
        }
      }
      return rich.getText();
    }
    const val = rng.getDisplayValue();
    if (val != null && val !== '') return val;
    const raw = rng.getValue();
    return raw == null ? '' : raw;
  } catch (e) {
    DBG && DBG('[getCellLinkOrText_] ex', e);
    return '';
  }
}

function getActiveMasterRow_(){
  const ss = ss_();
  const sheet = sh_(MASTER_SHEET);
  const range = ss.getActiveRange();
  if (!range || range.getSheet().getName() !== MASTER_SHEET) {
    throw new Error('Select a valid row on "00_Master Wholesale" and try again.');
  }
  const row = range.getRow();
  if (row < 3) throw new Error('Select a valid row on "00_Master Wholesale" and try again.');
  const headers = headerIndex1_(sheet);
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const display = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
  const rich = sheet.getRange(row, 1, 1, lastCol).getRichTextValues()[0];
  return { sheet, row, headers, values, display, rich };
}

function normalizeSO_(raw){
  const pretty = soPretty_(raw || '');
  if (!pretty) {
    return { pretty: '', display: '', key: '' };
  }
  const key = soKey_(pretty);
  return { pretty, display: soDisplay_(pretty), key };
}

function soSheetName_(soPretty){
  return 'SO' + soPretty;
}

const REVISION_HEADERS = [
  'Timestamp','User','Action','Revision #','Mode','Accent Type','Ring Style','Metal','US Size','Band Width (mm)',
  'Center Type','Shape','Diamond Dimension','Design Notes','Short Tag','SO#','Brand','Odoo SO URL','Master Link'
];

function ensureRevisionSheetInCustomerTracker_(trackerId, soPretty){
  const ss = SpreadsheetApp.openById(trackerId);
  const sheetName = soSheetName_(soPretty);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, REVISION_HEADERS.length).setValues([REVISION_HEADERS]);
    sh.setFrozenRows(1);
  } else {
    const H = headerIndex1_(sh);
    let last = sh.getLastColumn();
    REVISION_HEADERS.forEach(h => {
      if (!H[h]) {
        last += 1;
        sh.getRange(1, last).setValue(h);
      }
    });
  }
  return { sheet: sh, sheetName };
}

function computeNextRevisionNumber_(sheet){
  const H = headerIndex1_(sheet);
  const idx = H['Revision #'] || 0;
  if (!idx) return 1;
  const last = sheet.getLastRow();
  if (last < 2) return 1;
  const vals = sheet.getRange(2, idx, last - 1, 1).getValues();
  let max = 0;
  vals.forEach(r => {
    const n = Number(r[0]);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

function appendRevisionRowToCustomerTracker_(trackerId, sheetName, rowObj){
  const ss = SpreadsheetApp.openById(trackerId);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Couldn\'t write the revision row. Please try again; if it persists, contact admin.');

  // Ensure headers exist and extend if needed
  const baseHeaders = REVISION_HEADERS.slice();
  const extraHeaders = Object.keys(rowObj).filter(k => baseHeaders.indexOf(k) === -1 && k !== '__rich');
  const H0 = headerIndex1_(sh);
  let last = sh.getLastColumn();
  [...baseHeaders, ...extraHeaders].forEach(h => {
    if (!H0[h]) {
      last += 1;
      sh.getRange(1, last).setValue(h);
    }
  });
  const H = headerIndex1_(sh);

  const arr = new Array(sh.getLastColumn()).fill('');
  Object.keys(rowObj).forEach(k => {
    if (k === '__rich') return;
    const idx = H[k];
    if (idx) arr[idx - 1] = rowObj[k];
  });
  sh.appendRow(arr);

  const newRow = sh.getLastRow();
  const rich = rowObj.__rich || {};
  Object.keys(rich).forEach(header => {
    const idx = H[header];
    if (idx) {
      try {
        const cfg = rich[header];
        const rt = SpreadsheetApp.newRichTextValue().setText(cfg.text || '').setLinkUrl(cfg.url || '').build();
        sh.getRange(newRow, idx).setRichTextValue(rt);
      } catch (e) {
        try {
          sh.getRange(newRow, idx).setValue(cfg.url || cfg.text || '');
        } catch(_){}
      }
    }
  });

  return { row: newRow };
}

function buildOdooRevisionText_(form){
  const f = form || {};
  const t = String(f.productType || '').toLowerCase();
  const notes = String(f.DesignNotes || '')
    .split(/\r?\n/)
    .map(s => s.replace(/^[-•\s]+/, '').trim())
    .filter(Boolean)
    .slice(0, 3);

  function block(title, pairs){
    const body = pairs
      .filter(([_, key]) => (f[key] || '') !== '')
      .map(([label, key]) => `• ${label.padEnd(15)}: ${f[key] || ''}`);
    return [REV3D_COPY_HEADER, title, ...body, '', '(Mode: ' + REV3D_MODE + ')'].join('\n');
  }

  if (t === 'ring' || t === 'ring setting') {
    const lines = [
      REV3D_COPY_HEADER,
      'SETTING',
      `• Accent Diamond : ${f.AccentDiamondType || ''}`,
      `• Ring Style     : ${f.RingStyle || ''}`,
      `• Metal          : ${f.Metal || ''}`,
      `• US Size        : ${f.USSize || ''}`,
      `• Band Width     : ${f.BandWidthMM || ''}`,
      '',
      'DESIGN NOTES',
      `• ${notes[0] || ''}`,
      `• ${notes[1] || ''}`,
      `• ${notes[2] || ''}`,
      '',
      'CENTER STONE',
      `• Type          : ${f.CenterDiamondType || ''}`,
      `• Shape         : ${f.Shape || ''}`,
      `• Dimension     : ${f.DiamondDimension || ''}`,
      '',
      '(Mode: ' + REV3D_MODE + ')'
    ];
    return lines.join('\n');
  }

  if (t === 'pendant') {
    return block('PENDANT', [
      ['Metal', 'Metal'],
      ['Chain Style', 'ChainStyle'],
      ['Length (in)', 'LengthIn'],
      ['Bail Type', 'BailType'],
      ['Notes', 'DesignNotes']
    ]);
  }
  if (t === 'chain') {
    return block('CHAIN', [
      ['Metal', 'Metal'],
      ['Chain Style', 'ChainStyle'],
      ['Width (mm)', 'WidthMM'],
      ['Length (in)', 'LengthIn'],
      ['Notes', 'DesignNotes']
    ]);
  }
  if (t === 'earrings') {
    return block('EARRINGS', [
      ['Metal', 'Metal'],
      ['Type', 'EarringType'],
      ['Back Type', 'BackType'],
      ['Notes', 'DesignNotes']
    ]);
  }
  if (t === 'bracelet') {
    return block('BRACELET', [
      ['Metal', 'Metal'],
      ['Bracelet Type', 'BraceletType'],
      ['Length (in)', 'LengthIn'],
      ['Notes', 'DesignNotes']
    ]);
  }
  return block('CUSTOM ITEM', [
    ['Metal', 'Metal'],
    ['Notes', 'DesignNotes']
  ]);
}

function rev3d_init(){
  const ctx = getActiveMasterRow_();
  const H = ctx.headers;
  const sheet = ctx.sheet;
  const row = ctx.row;

  if (!H['SO#']) throw new Error('SO# is required (format 12.3456).');
  const soCell = getCellLinkOrText_(sheet, row, H['SO#']);
  const soRaw = ctx.display[H['SO#'] - 1] || ctx.values[H['SO#'] - 1] || soCell;
  const soNorm = normalizeSO_(soRaw);
  if (!soNorm.pretty) throw new Error('SO# is required (format 12.3456).');

  const customerName = H['Customer Name'] ? String(ctx.values[H['Customer Name'] - 1] || '').trim() : '';
  if (!customerName) throw new Error('Customer Name is required on the selected row.');

  let odooUrl = '';
  if (H['Odoo SO URL']) {
    const val = getCellLinkOrText_(sheet, row, H['Odoo SO URL']);
    if (/^https?:/i.test(String(val || '').trim())) odooUrl = val;
  }
  if (!odooUrl && soCell && /^https?:/i.test(String(soCell || '').trim())) odooUrl = soCell;

  const threeDFolderUrl = H['05-3D Folder'] ? getCellLinkOrText_(sheet, row, H['05-3D Folder']) : '';

  const tracker = resolveCustomerWorkbookFor_(customerName, { master: ctx });
  if (!tracker || !tracker.id) throw new Error('Couldn\'t resolve the Customer Order Tracker for this customer.');

  const today = Utilities.formatDate(new Date(), ADM_TZ, 'yyyy-MM-dd');
  const prefill = {
    productType: 'Ring',
    priorityLevel: '',
    quantity: '1',
    DesignNotes: '',
    AccentDiamondType: '',
    RingStyle: '',
    Metal: '',
    USSize: '',
    BandWidthMM: '',
    CenterDiamondType: '',
    Shape: '',
    DiamondDimension: '',
    ChainStyle: '',
    LengthIn: '',
    BailType: '',
    WidthMM: '',
    EarringType: '',
    BackType: '',
    BraceletType: ''
  };

  const trackerSs = SpreadsheetApp.openById(tracker.id);
  const soSheet = trackerSs.getSheetByName(soSheetName_(soNorm.pretty));
  if (soSheet) {
    const last = soSheet.getLastRow();
    if (last >= 2) {
      const data = soSheet.getRange(2, 1, last - 1, soSheet.getLastColumn()).getValues();
      const Hrev = headerIndex1_(soSheet);
      for (let i = data.length - 1; i >= 0; i--) {
        const rowVals = data[i];
        if (!rowVals.some(v => v !== '' && v != null)) continue;
        const byHeader = {};
        Object.keys(Hrev).forEach(name => {
          const idx = Hrev[name];
          if (idx) byHeader[name] = rowVals[idx - 1];
        });
        if (byHeader['Design Notes']) prefill.DesignNotes = byHeader['Design Notes'];
        if (byHeader['Accent Type']) prefill.AccentDiamondType = byHeader['Accent Type'];
        if (byHeader['Ring Style']) prefill.RingStyle = byHeader['Ring Style'];
        if (byHeader['Metal']) prefill.Metal = byHeader['Metal'];
        if (byHeader['US Size']) prefill.USSize = byHeader['US Size'];
        if (byHeader['Band Width (mm)']) prefill.BandWidthMM = byHeader['Band Width (mm)'];
        if (byHeader['Center Type']) prefill.CenterDiamondType = byHeader['Center Type'];
        if (byHeader['Shape']) prefill.Shape = byHeader['Shape'];
        if (byHeader['Diamond Dimension']) prefill.DiamondDimension = byHeader['Diamond Dimension'];
        if (byHeader['Chain Style']) prefill.ChainStyle = byHeader['Chain Style'];
        if (byHeader['Length (in)']) prefill.LengthIn = byHeader['Length (in)'];
        if (byHeader['Bail Type']) prefill.BailType = byHeader['Bail Type'];
        if (byHeader['Width (mm)']) prefill.WidthMM = byHeader['Width (mm)'];
        if (byHeader['Earring Type']) prefill.EarringType = byHeader['Earring Type'];
        if (byHeader['Back Type']) prefill.BackType = byHeader['Back Type'];
        if (byHeader['Bracelet Type']) prefill.BraceletType = byHeader['Bracelet Type'];
        if (byHeader['Priority']) prefill.priorityLevel = String(byHeader['Priority']);
        if (byHeader['Quantity']) prefill.quantity = String(byHeader['Quantity']);

        const typeGuess = (function(){
          if (prefill.BraceletType) return 'Bracelet';
          if (prefill.EarringType || prefill.BackType) return 'Earrings';
          if (prefill.WidthMM && !prefill.BailType) return 'Chain';
          if (prefill.BailType) return 'Pendant';
          if (prefill.RingStyle || prefill.USSize || prefill.AccentDiamondType) return 'Ring';
          return prefill.productType;
        })();
        prefill.productType = typeGuess;
        break;
      }
    }
  }

  function toYMD(val){
    if (val instanceof Date) {
      return Utilities.formatDate(val, ADM_TZ, 'yyyy-MM-dd');
    }
    const str = String(val || '').trim();
    if (!str) return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, ADM_TZ, 'yyyy-MM-dd');
    }
    return '';
  }

  let inquiryDateDefault = today;
  if (H['Inquiry Date']) {
    const raw = ctx.values[H['Inquiry Date'] - 1];
    const ymd = toYMD(raw);
    if (ymd) inquiryDateDefault = ymd;
  }
  let quotationDateDefault = '';
  if (H['Quotation Date']) {
    const raw = ctx.values[H['Quotation Date'] - 1];
    const ymd = toYMD(raw);
    if (ymd) quotationDateDefault = ymd;
  }

  return {
    ok: true,
    today,
    masterRow: row,
    soPretty: soNorm.pretty,
    soDisplay: soNorm.display,
    customerName,
    odooUrl,
    trackerId: tracker.id,
    trackerUrl: tracker.url,
    threeDFolderUrl,
    prefill,
    defaultInquiryDate: inquiryDateDefault,
    defaultQuotationDate: quotationDateDefault
  };
}

function previewRevOdooPaste(form){
  return buildOdooRevisionText_(form || {});
}

function submit3DRevision(payload){
  payload = payload || {};
  const ctx = getActiveMasterRow_();
  const H = ctx.headers;

  const soInput = payload.so || payload.soDisplay || '';
  const soNorm = normalizeSO_(soInput);
  if (!soNorm.pretty) throw new Error('SO# is required (format 12.3456).');

  const inquiryDate = String(payload.inquiryDate || '').trim();
  if (!inquiryDate) throw new Error('Inquiry Date is required.');

  const customerName = H['Customer Name'] ? String(ctx.values[H['Customer Name'] - 1] || '').trim() : '';
  if (!customerName) throw new Error('Customer Name is required on the selected row.');

  const tracker = resolveCustomerWorkbookFor_(customerName, { master: ctx });
  if (!tracker || !tracker.id) throw new Error('Couldn\'t resolve the Customer Order Tracker for this customer.');

  const { sheet } = ensureRevisionSheetInCustomerTracker_(tracker.id, soNorm.pretty);
  const nextRev = computeNextRevisionNumber_(sheet);

  const form = payload.form || {};
  const odooText = buildOdooRevisionText_(form);
  const masterSheet = ctx.sheet;
  const masterLinkUrl = ss_().getUrl() + '#gid=' + masterSheet.getSheetId() + '&range=' + ctx.row + ':' + ctx.row;
  let existingOdooUrl = '';
  if (H['Odoo SO URL']) {
    const val = getCellLinkOrText_(masterSheet, ctx.row, H['Odoo SO URL']);
    if (/^https?:/i.test(String(val || '').trim())) existingOdooUrl = val;
  }
  if (!existingOdooUrl && H['SO#']) {
    const link = getCellLinkOrText_(masterSheet, ctx.row, H['SO#']);
    if (/^https?:/i.test(String(link || '').trim())) existingOdooUrl = link;
  }
  const odooUrl = String(payload.soUrl || payload.odooUrl || existingOdooUrl || '').trim();
  const masterBrand = H['Brand'] ? ctx.values[H['Brand'] - 1] : '';

  function shortTagFrom(form){
    const parts = [form.Shape, form.RingStyle].map(v => String(v || '').trim()).filter(Boolean);
    if (!parts.length) return '';
    const joined = parts.join(' ');
    const title = joined.split(/\s+/).map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
    return title.length > 24 ? title.slice(0, 24).trim() : title;
  }

  const rowMap = {
    'Timestamp': new Date(),
    'User': (function(){
      try { return Session.getActiveUser().getEmail() || ''; }
      catch (_) { return ''; }
    })(),
    'Action': REV3D_ACTION,
    'Revision #': nextRev,
    'Mode': REV3D_MODE,
    'Accent Type': form.AccentDiamondType || '',
    'Ring Style': form.RingStyle || '',
    'Metal': form.Metal || '',
    'US Size': form.USSize || '',
    'Band Width (mm)': form.BandWidthMM || '',
    'Center Type': form.CenterDiamondType || '',
    'Shape': form.Shape || '',
    'Diamond Dimension': form.DiamondDimension || '',
    'Design Notes': form.DesignNotes || '',
    'Short Tag': shortTagFrom(form),
    'SO#': soNorm.display,
    'Brand': masterBrand || '',
    'Odoo SO URL': odooUrl || '',
    'Master Link': masterLinkUrl,
    'Chain Style': form.ChainStyle || '',
    'Length (in)': form.LengthIn || '',
    'Bail Type': form.BailType || '',
    'Width (mm)': form.WidthMM || '',
    'Earring Type': form.EarringType || '',
    'Back Type': form.BackType || '',
    'Bracelet Type': form.BraceletType || '',
    'Priority': form.priorityLevel || '',
    'Quantity': form.quantity || '',
    'Inquiry Date': inquiryDate,
    'Quotation Date': String(payload.quotationDate || '').trim()
  };

  rowMap.__rich = {
    'Master Link': { text: 'Master Row ' + ctx.row, url: masterLinkUrl },
    'Odoo SO URL': odooUrl ? { text: 'Odoo SO', url: odooUrl } : null
  };
  if (!rowMap.__rich['Odoo SO URL']) delete rowMap.__rich['Odoo SO URL'];

  appendRevisionRowToCustomerTracker_(tracker.id, sheet.getName(), rowMap);

  if (H['Custom Order Status']) {
    masterSheet.getRange(ctx.row, H['Custom Order Status']).setValue(REV3D_ACTION);
  }

  const summary = {
    ok: true,
    trackerUrl: tracker.url,
    sheetName: sheet.getName(),
    soPretty: soNorm.pretty,
    masterRow: ctx.row,
    revision: nextRev,
    copy: odooText
  };

  return summary;
}

