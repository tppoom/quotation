const assert = require('assert');

// Mock localStorage for node environment
global.localStorage = (function() {
  let store = {};
  return {
    getItem: (key) => (Object.prototype.hasOwnProperty.call(store, key) ? store[key] : null),
    setItem: (key, val) => { store[key] = String(val); },
    removeItem: (key) => { delete store[key]; },
    clear: () => { store = {}; }
  };
})();

const StorageManager = require('../js/storage.js');
const ExcelExporter = require('../js/excelExport.js');
const DocExporter = require('../js/pdfExport.js');
const { bahtText } = require('../js/bahtText.js');
const quotationApp = require('../js/app.js');

async function runTests() {
  console.log('Testing Alpine.js application controller (quotationApp)...');

  // =========================================================================
  // 1. Export structure & UMD loading
  // =========================================================================
  console.log('1. Testing module exports and factory...');
  assert.strictEqual(typeof quotationApp, 'function', 'quotationApp should be a factory function');
  assert.strictEqual(typeof quotationApp.quotationApp, 'function', 'quotationApp.quotationApp should be exported');

  // =========================================================================
  // 2. Initial state & startNewQuotation (Zero private credentials check)
  // =========================================================================
  console.log('2. Testing initialization without draft & privacy compliance...');
  localStorage.clear();

  const app = quotationApp();
  assert.strictEqual(typeof app, 'object', 'app instance should be an object');

  // Verify zero hardcoded private credentials in initial state
  assert.strictEqual(app.quotation.customer.name, '', 'Initial customer name must be empty');
  assert.strictEqual(app.quotation.customer.address, '', 'Initial customer address must be empty');
  assert.strictEqual(app.quotation.jobTitle, '', 'Initial job title must be empty');

  // Call init()
  app.init();
  assert(app.quotation.docNo.startsWith('QT-'), 'docNo should start with QT-');
  assert.strictEqual(app.quotation.items.length, 1, 'Initial quotation should have 1 item row');
  assert.strictEqual(app.quotation.items[0].type, 'item', 'First row should be of type item');
  assert.strictEqual(app.quotation.items[0].itemNo, '1', 'First row should have itemNo 1');
  assert.strictEqual(app.quotation.totalAmount, 0, 'Initial total should be 0');
  assert.strictEqual(app.quotation.totalTextThai, '(ศูนย์บาทถ้วน)', 'Initial total Thai text should be (ศูนย์บาทถ้วน)');
  assert.strictEqual(app.quotation.customer.name, '', 'New quotation customer name must be empty');

  // Check seller profile loaded from StorageManager
  const defaultSeller = StorageManager.getSellerProfile();
  assert.strictEqual(app.quotation.seller.name, defaultSeller.name, 'Seller profile name should match storage');

  // Check timezone-safe local date format
  const todayStr = StorageManager.getLocalDateStr(new Date());
  assert.strictEqual(app.quotation.docDate, todayStr, 'docDate should match local date YYYY-MM-DD');

  // =========================================================================
  // 3. Draft restoration and $watch autosave
  // =========================================================================
  console.log('3. Testing draft restoration and autosave...');
  const testDraft = {
    id: 'qt_draft_test',
    docNo: 'QT-202611-9999',
    docDate: '2026-11-03',
    jobTitle: 'งานปรับปรุงระบบน้ำ',
    customer: { name: 'ลูกค้าทดสอบ ก', address: 'กรุงเทพฯ' },
    seller: { ...defaultSeller },
    items: [
      { id: 'i1', type: 'item', itemNo: '1', description: 'ปั๊มน้ำ', unit: 'ตัว', qty: 2, unitPrice: 3500, amount: 7000 }
    ],
    totalAmount: 7000,
    totalTextThai: '(เจ็ดพันบาทถ้วน)'
  };
  StorageManager.saveCurrentDraft(testDraft);

  const draftApp = quotationApp();
  let watchRegistered = false;
  let watchHandler = null;
  draftApp.$watch = function(prop, handler) {
    if (prop === 'quotation') {
      watchRegistered = true;
      watchHandler = handler;
    }
  };

  draftApp.init();
  assert.strictEqual(draftApp.quotation.docNo, 'QT-202611-9999', 'Draft docNo should be restored');
  assert.strictEqual(draftApp.quotation.customer.name, 'ลูกค้าทดสอบ ก', 'Draft customer name should be restored');
  assert.strictEqual(draftApp.quotation.totalAmount, 7000, 'Draft total should be 7000');
  assert.strictEqual(watchRegistered, true, '$watch should be registered on quotation');

  // Trigger watch handler to verify autosave
  draftApp.quotation.items[0].qty = 3;
  if (watchHandler) watchHandler();
  assert.strictEqual(draftApp.quotation.totalAmount, 10500, 'Autosave watch should recalculate totals');
  assert.strictEqual(draftApp.quotation.totalTextThai, '(หนึ่งหมื่นห้าร้อยบาทถ้วน)', 'Thai baht text should update');
  const savedDraft = StorageManager.loadCurrentDraft();
  assert.strictEqual(savedDraft.totalAmount, 10500, 'Draft in storage should be updated to 10500');

  // =========================================================================
  // 4. Dynamic row management (items, bullets, payment terms, removal & re-indexing)
  // =========================================================================
  console.log('4. Testing dynamic row management & re-indexing...');
  const rowApp = quotationApp();
  rowApp.startNewQuotation();
  assert.strictEqual(rowApp.quotation.items.length, 1);
  assert.strictEqual(rowApp.quotation.items[0].itemNo, '1');

  // Add 2nd item row
  rowApp.addItemRow();
  assert.strictEqual(rowApp.quotation.items.length, 2);
  assert.strictEqual(rowApp.quotation.items[1].type, 'item');
  assert.strictEqual(rowApp.quotation.items[1].itemNo, '2');

  // Add bullet sub-item row
  rowApp.addSubItemRow();
  assert.strictEqual(rowApp.quotation.items.length, 3);
  assert.strictEqual(rowApp.quotation.items[2].type, 'bullet');
  assert.strictEqual(rowApp.quotation.items[2].description, ' - ');
  assert.strictEqual(rowApp.quotation.items[2].itemNo, undefined, 'Bullet rows should not have itemNo');

  // Add 3rd item row
  rowApp.addItemRow();
  assert.strictEqual(rowApp.quotation.items.length, 4);
  assert.strictEqual(rowApp.quotation.items[3].type, 'item');
  assert.strictEqual(rowApp.quotation.items[3].itemNo, '3', 'Next item row should be 3 despite intermediate bullet');

  // Add payment terms preset
  rowApp.addPaymentTermsPreset();
  assert.strictEqual(rowApp.quotation.items.length, 6);
  assert.strictEqual(rowApp.quotation.items[4].type, 'note');
  assert.strictEqual(rowApp.quotation.items[4].description, 'การเบิกค่างวดงาน');
  assert.strictEqual(rowApp.quotation.items[5].type, 'bullet');
  assert.strictEqual(rowApp.quotation.items[5].description, 'เบิกมัดจำ 70% และ 30% เมื่องานเสร็จสิ้น');

  // Remove first item row (index 0) and verify automatic re-indexing
  rowApp.removeItemRow(0);
  assert.strictEqual(rowApp.quotation.items.length, 5);
  // Remaining items should now be re-indexed as 1 and 2
  const reindexedItems = rowApp.quotation.items.filter(i => i.type === 'item');
  assert.strictEqual(reindexedItems.length, 2);
  assert.strictEqual(reindexedItems[0].itemNo, '1', 'Remaining first item should re-index to 1');
  assert.strictEqual(reindexedItems[1].itemNo, '2', 'Remaining second item should re-index to 2');

  // =========================================================================
  // 5. Total Calculations and Thai Baht Text
  // =========================================================================
  console.log('5. Testing total calculations and Baht text...');
  const calcApp = quotationApp();
  calcApp.startNewQuotation();
  calcApp.quotation.items = [
    { type: 'item', itemNo: '1', qty: 2, unitPrice: 1500, amount: 0 },
    { type: 'bullet', description: ' - ข้อต่อ PVC' },
    { type: 'item', itemNo: '2', qty: '3', unitPrice: '250.50', amount: 0 },
    { type: 'note', description: 'หมายเหตุ' },
    { type: 'item', itemNo: '3', qty: '1,000', unitPrice: '1.25', amount: 0 } // formatted strings with comma
  ];

  calcApp.calculateTotals();
  // Item 1: 2 * 1500 = 3000
  // Item 2: 3 * 250.50 = 751.50
  // Item 3: 1000 * 1.25 = 1250
  // Total: 3000 + 751.50 + 1250 = 5001.50
  assert.strictEqual(calcApp.quotation.items[0].amount, 3000);
  assert.strictEqual(calcApp.quotation.items[2].amount, 751.5);
  assert.strictEqual(calcApp.quotation.items[4].amount, 1250);
  assert.strictEqual(calcApp.quotation.totalAmount, 5001.5);
  assert.strictEqual(calcApp.quotation.totalTextThai, '(ห้าพันหนึ่งบาทห้าสิบสตางค์)');

  // Test zero and empty quantities
  calcApp.quotation.items = [
    { type: 'item', itemNo: '1', qty: 0, unitPrice: 1000, amount: 0 },
    { type: 'item', itemNo: '2', qty: '', unitPrice: '', amount: 0 }
  ];
  calcApp.calculateTotals();
  assert.strictEqual(calcApp.quotation.totalAmount, 0);
  assert.strictEqual(calcApp.quotation.totalTextThai, '(ศูนย์บาทถ้วน)');

  // =========================================================================
  // 6. History Management (save, search, filter, load, duplicate, delete)
  // =========================================================================
  console.log('6. Testing history management...');
  localStorage.clear();
  const histApp = quotationApp();
  histApp.startNewQuotation();
  histApp.quotation.customer.name = 'บริษัท ก ทดสอบ จำกัด';
  histApp.quotation.jobTitle = 'ติดตั้งระบบน้ำสนามหญ้า';
  histApp.quotation.items = [
    { type: 'item', itemNo: '1', qty: 1, unitPrice: 20000, amount: 20000 }
  ];
  histApp.saveQuotation();

  assert.strictEqual(histApp.historyList.length, 1, 'History list should have 1 item');
  assert.strictEqual(histApp.historyList[0].customer.name, 'บริษัท ก ทดสอบ จำกัด');
  assert(histApp.saveNotice.includes('บันทึกเอกสาร'), 'saveNotice should be populated');

  // Search & Filter History
  histApp.historySearch = 'สนามหญ้า';
  assert.strictEqual(histApp.filteredHistory().length, 1, 'Filter should find match on jobTitle');
  histApp.historySearch = 'บริษัท ก';
  assert.strictEqual(histApp.filteredHistory().length, 1, 'Filter should find match on customer name');
  histApp.historySearch = 'ไม่พบข้อมูล';
  assert.strictEqual(histApp.filteredHistory().length, 0, 'Filter should return empty for unmatched query');
  histApp.historySearch = '';
  assert.strictEqual(histApp.filteredHistory().length, 1, 'Filter should return all for empty search');

  // Duplicate from history
  const origId = histApp.historyList[0].id;
  histApp.duplicateFromHistory(origId);
  assert.strictEqual(histApp.historyList.length, 2, 'History list should have 2 items after duplicate');
  assert.notStrictEqual(histApp.quotation.id, origId, 'Duplicated item should have new ID');
  assert.strictEqual(histApp.quotation.customer.name, 'บริษัท ก ทดสอบ จำกัด');
  assert.strictEqual(histApp.showHistoryModal, false, 'History modal should close after duplicate');

  // Load from history
  const firstItem = histApp.historyList[1];
  histApp.loadFromHistory(firstItem);
  assert.strictEqual(histApp.quotation.id, firstItem.id, 'Loaded item ID should match');
  assert.strictEqual(histApp.showHistoryModal, false, 'History modal should close after load');

  // Delete from history
  const deleteId = histApp.historyList[0].id;
  histApp.deleteFromHistory(deleteId);
  assert.strictEqual(histApp.historyList.length, 1, 'History list should have 1 item after delete');

  // Test confirm cancellation
  const originalConfirm = global.confirm;
  global.confirm = () => false;
  histApp.deleteFromHistory(histApp.historyList[0].id);
  assert.strictEqual(histApp.historyList.length, 1, 'Item should not be deleted when confirmation cancelled');
  global.confirm = originalConfirm;

  // =========================================================================
  // 7. Seller Profile Settings Modal
  // =========================================================================
  console.log('7. Testing seller profile settings modal...');
  const sellerApp = quotationApp();
  sellerApp.startNewQuotation();
  sellerApp.openSellerSettings();
  assert.strictEqual(sellerApp.showSellerModal, true, 'showSellerModal should be true');

  sellerApp.sellerProfileForm = {
    name: 'ร้านช่างทอง การประปา',
    address: '99/1 ถ.มิตรภาพ นครราชสีมา',
    tel1: '081-111-2222',
    tel2: '082-333-4444',
    bank: 'ไทยพาณิชย์ 123-4-56789-0'
  };
  sellerApp.saveSellerSettings();
  assert.strictEqual(sellerApp.showSellerModal, false, 'showSellerModal should be false after save');
  assert.strictEqual(sellerApp.quotation.seller.name, 'ร้านช่างทอง การประปา');

  // Verify stored in StorageManager
  const storedSeller = StorageManager.getSellerProfile();
  assert.strictEqual(storedSeller.name, 'ร้านช่างทอง การประปา');

  // =========================================================================
  // 8. Formatting and Export Helpers
  // =========================================================================
  console.log('8. Testing formatting and export triggers...');
  const exportApp = quotationApp();
  exportApp.startNewQuotation();
  exportApp.quotation.docNo = 'QT-202611-0005';
  exportApp.quotation.customer.name = 'บริษัท ตัวอย่าง จำกัด';

  // Currency formatting
  assert.strictEqual(exportApp.formatCurrency(1234.5), '1,234.50');
  assert.strictEqual(exportApp.formatCurrency('1234567.89'), '1,234,567.89');
  assert.strictEqual(exportApp.formatCurrency(0), '0.00');
  assert.strictEqual(exportApp.formatCurrency(''), '');
  assert.strictEqual(exportApp.formatCurrency(null), '');
  assert.strictEqual(exportApp.formatCurrency(undefined), '');

  // Date formatting
  assert.strictEqual(exportApp.formatDateShort('2026-11-03'), '3 พ.ย. 69');
  assert.strictEqual(exportApp.formatDateLong('2026-11-03'), '3 พฤศจิกายน 2569');

  // Filename generation
  const xlsxFilename = exportApp.getFilename('xlsx');
  assert.strictEqual(xlsxFilename, 'ใบเสนอราคา_บริษัท ตัวอย่าง จำกัด_QT-202611-0005.xlsx');
  const pngFilename = exportApp.getFilename('png');
  assert.strictEqual(pngFilename, 'ใบเสนอราคา_บริษัท ตัวอย่าง จำกัด_QT-202611-0005.png');

  // Flash notice timeout
  exportApp.flashNotice('ข้อความทดสอบ');
  assert.strictEqual(exportApp.saveNotice, 'ข้อความทดสอบ');

  // Export triggers (gracefully handled when canvas/DOM not present)
  let excelCalled = false;
  ExcelExporter.downloadExcelFile = async (data) => { excelCalled = true; };
  await exportApp.exportExcel();
  assert.strictEqual(excelCalled, true, 'exportExcel should call ExcelExporter.downloadExcelFile');

  let pngCalled = false;
  DocExporter.exportToPNG = async (elId, fname) => { pngCalled = true; };
  await exportApp.exportPNG();
  assert.strictEqual(pngCalled, true, 'exportPNG should call DocExporter.exportToPNG');

  let pdfCalled = false;
  DocExporter.exportToPDF = async (elId, fname) => { pdfCalled = true; };
  await exportApp.exportPDF();
  assert.strictEqual(pdfCalled, true, 'exportPDF should call DocExporter.exportToPDF');

  let printCalled = false;
  DocExporter.printDocument = () => { printCalled = true; };
  exportApp.printDocument();
  assert.strictEqual(printCalled, true, 'printDocument should call DocExporter.printDocument');

  // =========================================================================
  // 9. Alpine.data auto-registration
  // =========================================================================
  console.log('9. Testing Alpine registration mock...');
  let registeredName = '';
  let registeredFactory = null;
  global.Alpine = {
    data: (name, factory) => {
      registeredName = name;
      registeredFactory = factory;
    }
  };
  global.document = {
    addEventListener: (event, handler) => {
      if (event === 'alpine:init') handler();
    }
  };

  // Re-require app.js with mocked Alpine & document
  delete require.cache[require.resolve('../js/app.js')];
  const reloadedApp = require('../js/app.js');
  assert.strictEqual(registeredName, 'quotationApp', 'Alpine.data should register quotationApp');
  assert.strictEqual(typeof registeredFactory, 'function', 'Registered Alpine factory should be a function');
  const instanceFromAlpine = registeredFactory();
  assert.strictEqual(typeof instanceFromAlpine.startNewQuotation, 'function');

  console.log('✅ All quotationApp tests passed successfully!');
}

runTests().catch(err => {
  console.error('❌ Test failed:', err);
  process.exit(1);
});
