
/*** PUBLIC — menu & open dialog ***/
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('ADM Wholesale')
    .addItem('New Inquiry / Create SO','admOpenNewInquiryDialog')
    .addItem('➕ Add New Customer','admOpenNewCustomerDialog')
    .addSeparator()
    .addItem('Record Payment (wholesale)…', 'openWholesaleRecordPayment')
    .addItem('Payment Summary (selected SO)…', 'openWholesalePaymentSummary')
    .addSeparator()
    .addItem('Refresh Config Cache', 'wh_refreshCaches')
    .addToUi();
}

function admOpenNewInquiryDialog(){
  const t = HtmlService.createTemplateFromFile('dlg_adm_new_inquiry'); // no trailing underscore

  // Bootstrap lists to client
  t.BOOTSTRAP = {
    tz: ADM_TZ,
    today: Utilities.formatDate(new Date(), ADM_TZ, 'yyyy-MM-dd'), // NEW: for default Inquiry Date
    customers: listCRMCustomers_(),                                 // NEW: CRM → Business Name list
    productTypes: ['Ring Setting','Pendant','Chain','Earrings','Bracelet','Other'],
    metals: ['14K YG','18K YG','14K WG','18K WG','14K RG','18K RG','Pt950','Silver'],
    ringStyles: ['Solitaire','Pavé','Cathedral','Halo','Bezel','Knife-edge','Vintage'],
    accentTypes: ['Natural','Lab-grown','None'],
    ringSizes: ['4.5','5','5.5','6','6.5','7','7.5','8'],
    shapes: ['Round','Oval','Cushion','Emerald','Radiant','Pear','Marquise','Princess'],
    chainStyles: ['Curb','Cable','Franco','Rope','Box'],
    earringTypes: ['Studs','Hoops','Halo','Drop'],
    braceletTypes: ['Tennis','Curb','Bangle'],
    uploadTypes: ['Product Image','Design Sketch','3D File','PO Acknowledgment','Other']
  };

  const html = t.evaluate().setWidth(760).setHeight(720).setTitle('New Inquiry / Create SO');
  SpreadsheetApp.getUi().showModalDialog(html,'New Inquiry / Create SO');
}


function openWholesaleRecordPayment() {
  ensureLedger_();
  const html = HtmlService.createHtmlOutputFromFile('dlg_wh_record_payment')
    .setWidth(1000).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Record Wholesale Payment');
}

function openWholesalePaymentSummary() {
  ensureLedger_();
  const html = HtmlService.createHtmlOutputFromFile('dlg_wh_summary')
    .setWidth(900).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Payment Summary (Wholesale)');
}

function wh_refreshCaches() {
  // kept for parity with future caching; currently a no-op
  SpreadsheetApp.getUi().alert('Config refreshed.');
}
