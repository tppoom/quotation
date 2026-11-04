const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const ExcelJS = require('exceljs');

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

async function runE2ETests() {
  console.log('================================================================');
  console.log('🚀 Running End-to-End Functional Verification Test Suite');
  console.log('================================================================');

  // ===========================================================================
  // 1. Init Fresh App State (Privacy compliance & initial defaults)
  // ===========================================================================
  console.log('\n[1/9] Initializing fresh app state and verifying privacy defaults...');
  localStorage.clear();

  const app = quotationApp();
  assert.strictEqual(typeof app, 'object', 'Application controller should initialize');

  // Verify initial data state is completely clean with zero hardcoded credentials
  assert.strictEqual(app.quotation.customer.name, '', 'Initial customer name must be empty');
  assert.strictEqual(app.quotation.customer.address, '', 'Initial customer address must be empty');
  assert.strictEqual(app.quotation.jobTitle, '', 'Initial job title must be empty');

  app.init();
  assert(app.quotation.docNo.startsWith('QT-'), 'Document number must start with QT-');
  assert.strictEqual(app.quotation.items.length, 1, 'Initial quotation must have 1 item row');
  assert.strictEqual(app.quotation.items[0].type, 'item', 'Initial row type must be "item"');
  assert.strictEqual(app.quotation.items[0].itemNo, '1', 'Initial row must be indexed as 1');
  assert.strictEqual(app.quotation.totalAmount, 0, 'Initial total must be 0');
  assert.strictEqual(app.quotation.totalTextThai, '(ศูนย์บาทถ้วน)', 'Initial Thai text must be (ศูนย์บาทถ้วน)');
  console.log('  ✓ Fresh state initialized with zero private data.');

  // ===========================================================================
  // 2. Set Custom Seller Profile in StorageManager
  // ===========================================================================
  console.log('\n[2/9] Configuring custom seller profile in StorageManager...');
  const customSeller = {
    name: 'นายสมเกียรติ สวนสวรรค์',
    address: '99/88 หมู่บ้านธรรมชาติ ต.คลองหลวง อ.คลองหลวง จ.ปทุมธานี 12120',
    tel1: '081-999-8877',
    tel2: '089-111-2233',
    bank: 'ธ.กสิกรไทย 123-4-56789-0'
  };

  app.sellerProfileForm = { ...customSeller };
  app.saveSellerSettings();

  const loadedSeller = StorageManager.getSellerProfile();
  assert.strictEqual(loadedSeller.name, customSeller.name, 'Seller name must match saved profile');
  assert.strictEqual(loadedSeller.address, customSeller.address, 'Seller address must match saved profile');
  assert.strictEqual(loadedSeller.tel1, customSeller.tel1, 'Seller tel1 must match saved profile');
  assert.strictEqual(loadedSeller.tel2, customSeller.tel2, 'Seller tel2 must match saved profile');
  assert.strictEqual(loadedSeller.bank, customSeller.bank, 'Seller bank must match saved profile');
  assert.strictEqual(app.quotation.seller.name, customSeller.name, 'Quotation seller must reflect profile changes');
  console.log('  ✓ Custom seller profile persisted and applied.');

  // ===========================================================================
  // 3. Create Quotation: Customer, Main Items, Sub-item Bullets, Payment Terms
  // ===========================================================================
  console.log('\n[3/9] Building quotation with customer, main items, bullets, and payment preset...');
  app.quotation.customer.name = 'บริษัท กรีนเทค การ์เดน จำกัด';
  app.quotation.customer.address = '82 ถนนสุขุมวิท 71 แขวงพระโขนงเหนือ เขตวัฒนา กรุงเทพฯ 10110';
  app.quotation.jobTitle = 'งานติดตั้งระบบน้ำอัตโนมัติ 4 โซน (ด้านหน้า)';
  app.quotation.docDate = '2026-11-03';

  // Configure first item
  app.quotation.items[0].description = 'ระบบรดน้ำอัตโนมัติ 4 โซน (ด้านหน้า)';
  app.quotation.items[0].unit = 'ชุด';
  app.quotation.items[0].qty = 1;
  app.quotation.items[0].unitPrice = 57552.90;

  // Add sub-item bullets under main item 1
  app.addSubItemRow();
  app.quotation.items[app.quotation.items.length - 1].description = ' - หัวป็อบอัพ 4 หุนจำนวน 5 หัว';

  app.addSubItemRow();
  app.quotation.items[app.quotation.items.length - 1].description = ' - ท่อ PE 20 มิลลิเมตร พร้อมอุปกรณ์ข้อต่อ';

  app.addSubItemRow();
  app.quotation.items[app.quotation.items.length - 1].description = ' - ตู้ควบคุมพร้อมไทม์เมอร์อัตโนมัติ 1 ชุด';

  // Add payment terms preset
  app.addPaymentTermsPreset();

  // Verify structure of items array
  assert.strictEqual(app.quotation.items.length, 6, 'Should have 1 main item + 3 bullets + 1 note + 1 bullet = 6 rows');
  assert.strictEqual(app.quotation.items[0].type, 'item', 'Row 0 is main item');
  assert.strictEqual(app.quotation.items[1].type, 'bullet', 'Row 1 is sub-bullet');
  assert.strictEqual(app.quotation.items[2].type, 'bullet', 'Row 2 is sub-bullet');
  assert.strictEqual(app.quotation.items[3].type, 'bullet', 'Row 3 is sub-bullet');
  assert.strictEqual(app.quotation.items[4].type, 'note', 'Row 4 is note header');
  assert.strictEqual(app.quotation.items[4].description, 'การเบิกค่างวดงาน');
  assert.strictEqual(app.quotation.items[5].type, 'bullet', 'Row 5 is payment bullet');
  assert.strictEqual(app.quotation.items[5].description, 'เบิกมัดจำ 70% และ 30% เมื่องานเสร็จสิ้น');
  console.log('  ✓ Quotation structure built successfully.');

  // ===========================================================================
  // 4. Verify Dynamic Total Calculation & Thai Baht Text Conversion
  // ===========================================================================
  console.log('\n[4/9] Verifying dynamic total calculation and Thai Baht text conversion...');
  app.calculateTotals();
  assert.strictEqual(app.quotation.totalAmount, 57552.90, 'Total amount must equal 57,552.90');
  assert.strictEqual(
    app.quotation.totalTextThai,
    '(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)',
    'Thai Baht text must correctly format floating satang'
  );

  // Add another item and check real-time recalculation
  app.addItemRow();
  const secondItem = app.quotation.items[app.quotation.items.length - 1];
  assert.strictEqual(secondItem.itemNo, '2', 'Second main item should be numbered 2');
  secondItem.description = 'ค่าบริการติดตั้งและทดสอบระบบ';
  secondItem.unit = 'จุด';
  secondItem.qty = 2;
  secondItem.unitPrice = 1200;
  app.calculateTotals();

  assert.strictEqual(app.quotation.totalAmount, 59952.90, 'Total should be 57552.90 + 2400 = 59952.90');
  assert.strictEqual(
    app.quotation.totalTextThai,
    '(ห้าหมื่นเก้าพันเก้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)',
    'Thai Baht text must update dynamically'
  );

  // Remove the second item and verify re-calculation and indexing
  app.removeItemRow(app.quotation.items.length - 1);
  assert.strictEqual(app.quotation.totalAmount, 57552.90, 'Total should return to 57,552.90');
  assert.strictEqual(
    app.quotation.totalTextThai,
    '(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)',
    'Thai Baht text should return to original'
  );
  console.log('  ✓ Dynamic calculations and Baht text validated.');

  // ===========================================================================
  // 5. Verify Running Quotation Number Auto-Incrementation & Collision Prevention
  // ===========================================================================
  console.log('\n[5/9] Verifying running quotation number generation and non-collision...');
  localStorage.clear();
  // Set custom seller again after clearing
  StorageManager.saveSellerProfile(customSeller);

  const docNo1 = StorageManager.generateDocNo('2026-11-03');
  assert.strictEqual(docNo1, 'QT-202611-0001', 'First quotation in 2026-11 should be QT-202611-0001');

  // Saving doc1 to history should advance sequence
  app.quotation.docNo = docNo1;
  app.saveQuotation();

  const docNo2 = StorageManager.generateDocNo('2026-11-03');
  assert.strictEqual(docNo2, 'QT-202611-0002', 'Second quotation in 2026-11 should be QT-202611-0002');

  // Create quotation 2 in history
  const q2 = { ...app.quotation, id: 'qt_test_2', docNo: docNo2 };
  StorageManager.saveToHistory(q2);

  // Deleting quotation 2 must NOT cause sequence regression or collision
  StorageManager.deleteFromHistory('qt_test_2');
  const docNo3 = StorageManager.generateDocNo('2026-11-03');
  assert.strictEqual(docNo3, 'QT-202611-0003', 'Sequence must monotonically increment even if records deleted');

  // Next month resets counter to 0001
  const docNoNextMonth = StorageManager.generateDocNo('2026-12-01');
  assert.strictEqual(docNoNextMonth, 'QT-202612-0001', 'New month must start at 0001');
  console.log('  ✓ Quotation number sequence and anti-collision logic verified.');

  // ===========================================================================
  // 6. Excel Workbook Export and Comparison with Reference Template
  // ===========================================================================
  console.log('\n[6/9] Exporting Excel workbook and comparing with reference template structure...');
  const refPath = path.resolve(__dirname, '../ใบเสนอราคา บ.เพิ่ม เอสเคท  จำกัด.xlsx');
  assert(fs.existsSync(refPath), 'Reference template file must exist');

  const refWb = new ExcelJS.Workbook();
  await refWb.xlsx.readFile(refPath);
  const refWs = refWb.worksheets[0];

  const genWb = await ExcelExporter.generateExcelWorkbook(app.quotation, ExcelJS);
  const genWs = genWb.getWorksheet('Sheet1');
  assert(genWs, 'Generated workbook must contain Sheet1');

  // 6.1 Compare Column Widths (100% fidelity)
  for (let c = 1; c <= 6; c++) {
    const colName = String.fromCharCode(64 + c);
    const refWidth = Math.round(refWs.getColumn(c).width * 100) / 100;
    const genWidth = Math.round(genWs.getColumn(colName).width * 100) / 100;
    assert.strictEqual(genWidth, refWidth, `Column ${colName} width should match template: ref=${refWidth}, gen=${genWidth}`);
  }
  console.log('  ✓ Column widths match reference template 100%.');

  // 6.2 Compare Row Heights
  const expectedRowHeights = {
    1: 30.75,
    2: 23.25,
    3: 25.5,
    4: 21.75,
    5: 26.25,
    9: 24,
    10: 18.75,
    11: 24,
    28: 21.95,
    29: 20.25,
    30: 11.25,
    31: 24,
    32: 21.95,
    33: 24
  };
  Object.entries(expectedRowHeights).forEach(([rowNum, expectedHeight]) => {
    const actualHeight = genWs.getRow(parseInt(rowNum, 10)).height;
    assert.strictEqual(actualHeight, expectedHeight, `Row ${rowNum} height must match template: expected ${expectedHeight}, got ${actualHeight}`);
  });
  console.log('  ✓ Row heights match reference template 100%.');

  // 6.3 Compare Header Texts and Labels
  assert.strictEqual(genWs.getCell('A1').value, '   ใบเสนอราคา', 'Row 1 title must match');
  assert.strictEqual(genWs.getCell('A2').value, '   QUOTATION', 'Row 2 subtitle must match');
  assert.strictEqual(genWs.getCell('D3').value, 'เลขที่/ No.');
  assert.strictEqual(genWs.getCell('A4').value, 'ที่อยู่  :');
  assert.strictEqual(genWs.getCell('D4').value, 'วันที่ Date ');
  assert.strictEqual(genWs.getCell('A9').value, 'ลำดับที่');
  assert.strictEqual(genWs.getCell('A10').value, 'ITEM');
  assert.strictEqual(genWs.getCell('B9').value, 'รายการ');
  assert.strictEqual(genWs.getCell('B10').value, 'DESCRIPTION');
  assert.strictEqual(genWs.getCell('C9').value, 'หน่วย');
  assert.strictEqual(genWs.getCell('D9').value, 'จำนวน');
  assert.strictEqual(genWs.getCell('E9').value, 'ราคา/หน่วย');
  assert.strictEqual(genWs.getCell('F9').value, 'จำนวนเงิน');
  assert.strictEqual(genWs.getCell('D28').value, 'รวมเงิน');
  assert.strictEqual(genWs.getCell('A31').value, 'ผู้อนุมัติ :');
  assert.strictEqual(genWs.getCell('D31').value, '   ผู้เสนอ :');
  assert.strictEqual(genWs.getCell('A32').value, 'ตำแหน่ง ');
  assert.strictEqual(genWs.getCell('B33').value, '   (                                                                                )');
  console.log('  ✓ Header texts and labels match template 100%.');

  // 6.4 Verify Merged Cell Ranges
  const genMerges = genWs.model.merges;
  const criticalMerges = [
    'A1:F1', 'A2:F2', 'A3:C3', 'D3:E3', 'B4:C4', 'D4:E4', 'D5:F5',
    'A6:C6', 'A7:C7', 'D7:E7', 'D8:F8', 'A28:C28', 'D28:E28', 'E31:F31', 'E32:F32',
    'C9:C10', 'D9:D10', 'E9:E10', 'F9:F10'
  ];
  criticalMerges.forEach(mergeRange => {
    assert(genMerges.includes(mergeRange), `Merged range ${mergeRange} must be present in exported worksheet`);
  });
  console.log('  ✓ Merged cells match template layout.');

  // 6.5 Verify SUM Formula & Total Results
  assert.deepStrictEqual(genWs.getCell('F28').value, {
    formula: 'SUM(F11:F27)',
    result: 57552.90
  }, 'Cell F28 must contain SUM(F11:F27) formula and match total');
  assert.strictEqual(genWs.getCell('F28').numFmt, '#,##0.00', 'Cell F28 format must be #,##0.00');
  assert.strictEqual(genWs.getCell('A28').value, '(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)');
  console.log('  ✓ SUM formula and total formatting verified.');

  // 6.6 Verify Export Filename Sanitization
  const exportFilename = ExcelExporter.getExportFilename(app.quotation, 'xlsx');
  assert.strictEqual(
    exportFilename,
    `ใบเสนอราคา_บริษัท กรีนเทค การ์เดน จำกัด_${app.quotation.docNo}.xlsx`,
    'Exported filename should be properly formatted and sanitized'
  );
  console.log('  ✓ Export filename formatting confirmed.');

  // ===========================================================================
  // 7. Export PDF / PNG Payload and Fallback Validation
  // ===========================================================================
  console.log('\n[7/9] Validating PNG and PDF export routines (2x Retina, 1-page A4 fit)...');

  // Mock DOM and canvas environment
  let h2cOptionsPassed = null;
  let jsPdfConstructorArgs = null;
  let jsPdfAddImageCalls = [];
  let jsPdfSaveCalls = [];
  let clickedDownload = false;
  let downloadFilename = '';

  const mockCanvas = {
    width: 1588,
    height: 2246,
    toDataURL: (type) => `data:${type};base64,mockImageData`
  };

  global.html2canvas = async function(element, options) {
    h2cOptionsPassed = options;
    return mockCanvas;
  };

  class MockJsPDF {
    constructor(options) {
      jsPdfConstructorArgs = options;
      this.internal = {
        pageSize: {
          getWidth: () => 210,
          getHeight: () => 297
        }
      };
    }
    addImage(imgData, format, x, y, width, height) {
      jsPdfAddImageCalls.push({ imgData, format, x, y, width, height });
    }
    save(name) {
      jsPdfSaveCalls.push(name);
    }
  }
  global.jspdf = { jsPDF: MockJsPDF };

  const mockPreviewEl = { id: 'quotation-preview' };
  global.document = {
    getElementById: (id) => (id === 'quotation-preview' ? mockPreviewEl : null),
    createElement: (tag) => {
      if (tag === 'a') {
        return {
          set download(val) { downloadFilename = val; },
          get download() { return downloadFilename; },
          href: '',
          click: () => { clickedDownload = true; }
        };
      }
      return {};
    },
    body: {
      appendChild: () => {},
      removeChild: () => {}
    }
  };

  // Test PNG Export
  await DocExporter.exportToPNG('quotation-preview', 'ใบเสนอราคา_ทดสอบ.png');
  assert(h2cOptionsPassed, 'html2canvas options should be passed');
  assert.strictEqual(h2cOptionsPassed.scale, 2, 'PNG export must use 2x scale for Retina clarity');
  assert.strictEqual(clickedDownload, true, 'PNG export must trigger anchor download click');
  assert.strictEqual(downloadFilename, 'ใบเสนอราคา_ทดสอบ.png', 'PNG download filename should match');
  console.log('  ✓ PNG export uses 2x retina scale and triggers browser download.');

  // Test PDF Export
  await DocExporter.exportToPDF('quotation-preview', 'ใบเสนอราคา_ทดสอบ.pdf');
  assert.deepStrictEqual(jsPdfConstructorArgs, {
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  }, 'PDF must be created in A4 portrait format');
  assert.strictEqual(jsPdfAddImageCalls.length, 1, 'PDF must add rendered preview image');
  assert.strictEqual(jsPdfAddImageCalls[0].width, 210, 'PDF image width must span 210mm');
  assert.strictEqual(jsPdfSaveCalls[0], 'ใบเสนอราคา_ทดสอบ.pdf', 'PDF save must use sanitized filename');
  console.log('  ✓ PDF export configured for 1-page A4 print fit.');

  // Test PDF Fallback to window.print() if jsPDF missing
  let printCalled = false;
  global.jspdf = undefined;
  global.window = { print: () => { printCalled = true; } };
  await DocExporter.exportToPDF('quotation-preview', 'fallback.pdf');
  assert.strictEqual(printCalled, true, 'PDF export should fallback to window.print() when library missing');
  console.log('  ✓ PDF export fallback to window.print() confirmed.');

  // ===========================================================================
  // 8. History Management: Save, Update, Duplicate, Delete, Restore
  // ===========================================================================
  console.log('\n[8/9] Verifying history lifecycle: save, update, duplicate, delete, restore...');
  // Clear history to start with known state
  localStorage.removeItem(StorageManager.KEYS.HISTORY);

  // 8.1 Save current quotation
  app.saveQuotation();
  assert.strictEqual(app.historyList.length, 1, 'History list should contain 1 quotation');
  assert.strictEqual(app.historyList[0].docNo, app.quotation.docNo, 'Saved quotation docNo should match');

  // 8.2 Update quotation in place
  app.quotation.jobTitle = 'งานติดตั้งระบบน้ำอัตโนมัติ (แก้ไขใหม่)';
  app.saveQuotation();
  assert.strictEqual(app.historyList.length, 1, 'Updating existing quotation must not create duplicate entry');
  assert.strictEqual(app.historyList[0].jobTitle, 'งานติดตั้งระบบน้ำอัตโนมัติ (แก้ไขใหม่)', 'Updated title saved');

  // 8.3 Duplicate quotation
  app.duplicateFromHistory(app.quotation.id);
  assert.strictEqual(app.historyList.length, 2, 'History must contain 2 quotations after duplication');
  assert.notStrictEqual(app.quotation.id, app.historyList[1].id, 'Cloned quotation must have unique ID');
  assert(app.quotation.docNo.endsWith('0002') || app.quotation.docNo.endsWith('0003'), 'Cloned quotation must receive new sequence number');
  assert.strictEqual(app.quotation.customer.name, 'บริษัท กรีนเทค การ์เดน จำกัด', 'Cloned quotation retains customer');
  assert.strictEqual(app.quotation.totalAmount, 57552.90, 'Cloned quotation retains calculated total');

  // 8.4 Delete original quotation
  global.confirm = () => true;
  const originalId = app.historyList[1].id;
  app.deleteFromHistory(originalId);
  assert.strictEqual(app.historyList.length, 1, 'History count should decrease to 1 after deletion');

  // 8.5 Restore remaining quotation from history
  const remaining = app.historyList[0];
  app.loadFromHistory(remaining);
  assert.strictEqual(app.quotation.id, remaining.id, 'Loaded quotation ID should match history record');
  assert.strictEqual(app.quotation.totalAmount, 57552.90, 'Restored quotation has accurate total');
  console.log('  ✓ History lifecycle operations (save, update, duplicate, delete, restore) passed.');

  // ===========================================================================
  // 9. Git Cleanliness & Zero Sensitive Data Leakage
  // ===========================================================================
  console.log('\n[9/9] Verifying Git status and privacy compliance...');
  const repoRoot = path.resolve(__dirname, '..');

  // 9.1 Check git status for any untracked private files
  const statusOut = execSync('git status --porcelain', { cwd: repoRoot }).toString();
  const untrackedFiles = statusOut
    .split('\n')
    .filter(line => line.startsWith('??'))
    .map(line => line.slice(3).trim());

  // Confirm no untracked .xlsx or private sensitive files exist
  const sensitiveUntracked = untrackedFiles.filter(f => f.endsWith('.xlsx') || f.endsWith('.xls') || f.startsWith('.env'));
  assert.strictEqual(sensitiveUntracked.length, 0, `No private data files should be untracked: ${sensitiveUntracked.join(', ')}`);

  // 9.2 Check gitignore ignores .xlsx files
  let isIgnored = false;
  try {
    execSync('git check-ignore -q "ใบเสนอราคา บ.เพิ่ม เอสเคท  จำกัด.xlsx"', { cwd: repoRoot });
    isIgnored = true;
  } catch (e) {
    isIgnored = false;
  }
  assert.strictEqual(isIgnored, true, 'Reference xlsx file must be ignored by .gitignore');

  // 9.3 Check source code for zero hardcoded real private credentials
  const sourceFiles = [
    path.join(repoRoot, 'index.html'),
    path.join(repoRoot, 'js/app.js'),
    path.join(repoRoot, 'js/storage.js'),
    path.join(repoRoot, 'js/excelExport.js'),
    path.join(repoRoot, 'js/pdfExport.js'),
    path.join(repoRoot, 'js/bahtText.js')
  ];

  const forbiddenStrings = [
    'ชัยพจน์', // Real person name from reference template
    '09 0864',  // Real phone from reference template
    '9330144529' // Real bank account number from reference template
  ];

  sourceFiles.forEach(file => {
    if (fs.existsSync(file)) {
      const content = fs.readFileSync(file, 'utf8');
      forbiddenStrings.forEach(str => {
        assert(!content.includes(str), `Source file ${path.basename(file)} must not contain private credential "${str}"`);
      });
    }
  });

  console.log('  ✓ Git repository is clean; zero private credentials in source code.');

  console.log('\n================================================================');
  console.log('🎉 ALL 9 END-TO-END FUNCTIONAL VERIFICATION CHECKS PASSED 100%!');
  console.log('================================================================\n');
}

runE2ETests().catch(err => {
  console.error('\n❌ E2E Verification failed:', err);
  process.exit(1);
});
