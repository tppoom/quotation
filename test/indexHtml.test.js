const assert = require('assert');
const fs = require('fs');
const path = require('path');

async function runTests() {
  console.log('Testing index.html layout, dependencies, bindings, and print styling...');

  const htmlPath = path.resolve(__dirname, '../index.html');
  assert(fs.existsSync(htmlPath), 'index.html must exist at project root');

  const html = fs.readFileSync(htmlPath, 'utf8');
  assert(html.length > 500, 'index.html must not be empty');

  // 1. CDN Dependencies
  console.log('1. Checking CDN dependencies...');
  assert(html.includes('cdn.tailwindcss.com'), 'Must load Tailwind CSS CDN');
  assert(html.includes('fonts.googleapis.com') && (html.includes('Sarabun') || html.includes('Prompt')), 'Must link Google Fonts Sarabun & Prompt');
  assert(html.includes('alpinejs@3'), 'Must load Alpine.js 3.x CDN');
  assert(html.includes('exceljs@4.4.0') || html.includes('exceljs'), 'Must load ExcelJS CDN');
  assert(html.includes('html2canvas@1.4.1') || html.includes('html2canvas'), 'Must load html2canvas CDN');
  assert(html.includes('jspdf@2.5.1') || html.includes('jspdf'), 'Must load jsPDF CDN');

  // 2. Modular script tags and order
  console.log('2. Checking modular application scripts and loading order...');
  const bahtIdx = html.indexOf('js/bahtText.js');
  const storageIdx = html.indexOf('js/storage.js');
  const excelIdx = html.indexOf('js/excelExport.js');
  const pdfIdx = html.indexOf('js/pdfExport.js');
  const appIdx = html.indexOf('js/app.js');

  assert(bahtIdx !== -1, 'Must include js/bahtText.js');
  assert(storageIdx !== -1, 'Must include js/storage.js');
  assert(excelIdx !== -1, 'Must include js/excelExport.js');
  assert(pdfIdx !== -1, 'Must include js/pdfExport.js');
  assert(appIdx !== -1, 'Must include js/app.js');

  assert(bahtIdx < storageIdx, 'bahtText.js must load before storage.js');
  assert(storageIdx < excelIdx, 'storage.js must load before excelExport.js');
  assert(excelIdx < pdfIdx, 'excelExport.js must load before pdfExport.js');
  assert(pdfIdx < appIdx, 'pdfExport.js must load before app.js');

  // 3. Root Alpine registration and preview element ID
  console.log('3. Checking Alpine bindings and quotation-preview ID...');
  assert(html.includes('x-data="quotationApp') || html.includes('x-data="quotationApp()"'), 'Root element must have x-data="quotationApp"');
  assert(html.includes('id="quotation-preview"'), 'Must contain id="quotation-preview" for PDF/PNG/Print export');

  // 4. Form inputs and data models
  console.log('4. Checking form input bindings...');
  assert(html.includes('quotation.docNo'), 'Form must bind quotation.docNo');
  assert(html.includes('quotation.docDate'), 'Form must bind quotation.docDate');
  assert(html.includes('quotation.jobTitle'), 'Form must bind quotation.jobTitle');
  assert(html.includes('quotation.customer.name'), 'Form must bind quotation.customer.name');
  assert(html.includes('quotation.customer.address'), 'Form must bind quotation.customer.address');

  // 5. Dynamic item actions
  console.log('5. Checking dynamic row actions...');
  assert(html.includes('addItemRow()'), 'Must have button to call addItemRow()');
  assert(html.includes('addSubItemRow()'), 'Must have button to call addSubItemRow()');
  assert(html.includes('addPaymentTermsPreset()'), 'Must have button to call addPaymentTermsPreset()');
  assert(html.includes('removeItemRow('), 'Must have button to call removeItemRow()');

  // 6. Action triggers (Save, New, History, Settings, Exports)
  console.log('6. Checking action triggers...');
  assert(html.includes('saveQuotation()'), 'Must have saveQuotation() trigger');
  assert(html.includes('startNewQuotation()'), 'Must have startNewQuotation() trigger');
  assert(html.includes('showHistoryModal'), 'Must handle showHistoryModal');
  assert(html.includes('showSellerModal'), 'Must handle showSellerModal');
  assert(html.includes('exportExcel()'), 'Must have exportExcel() trigger');
  assert(html.includes('exportPDF()'), 'Must have exportPDF() trigger');
  assert(html.includes('exportPNG()'), 'Must have exportPNG() trigger');
  assert(html.includes('printDocument()'), 'Must have printDocument() trigger');

  // 7. Preview layout content matching Excel
  console.log('7. Checking live preview content matching Excel template...');
  assert(html.includes('ใบเสนอราคา'), 'Preview must have ใบเสนอราคา header');
  assert(html.includes('QUOTATION'), 'Preview must have QUOTATION header');
  assert(html.includes('เรียน/Attention'), 'Preview must have เรียน/Attention');
  assert(html.includes('เลขที่/ No.'), 'Preview must have เลขที่/ No.');
  assert(html.includes('ลำดับที่'), 'Preview must have ลำดับที่ table column');
  assert(html.includes('ITEM'), 'Preview must have ITEM table column');
  assert(html.includes('DESCRIPTION'), 'Preview must have DESCRIPTION table column');
  assert(html.includes('รวมเงิน'), 'Preview must have รวมเงิน summary');
  assert(html.includes('ผู้อนุมัติ'), 'Preview must have ผู้อนุมัติ signature');
  assert(html.includes('ผู้เสนอ'), 'Preview must have ผู้เสนอ signature');

  // 8. Print stylesheet
  console.log('8. Checking print stylesheet...');
  assert(html.includes('@media print'), 'Must include @media print CSS');
  assert(html.includes('size: A4 portrait') || html.includes('size: a4 portrait'), 'Print stylesheet must specify A4 portrait');
  assert(html.includes('no-print'), 'Must have no-print CSS class');

  // 9. Zero private credentials
  console.log('9. Checking for zero private credentials...');
  assert(!html.includes('SKT'), 'Should not contain private SKT names in initial source');
  assert(!html.includes('081-888-8888'), 'Should not contain hardcoded private phone numbers');

  console.log('✅ All index.html validation checks passed successfully!');
}

runTests().catch(err => {
  console.error('❌ Test failed:', err);
  process.exit(1);
});
