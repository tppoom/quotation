const assert = require('assert');
const ExcelJS = require('exceljs');
const fs = require('fs');

// Note: Test file will require ../js/excelExport.js once implemented
let ExcelExporter;
try {
  ExcelExporter = require('../js/excelExport.js');
} catch (e) {
  // Expected to fail before implementation
  ExcelExporter = null;
}

const { generateExcelWorkbook, getExportFilename, formatThaiDateShort, formatThaiDateLong, sanitizeFilename } = ExcelExporter || {};

async function runTests() {
  if (!generateExcelWorkbook) {
    throw new Error('Cannot find module ../js/excelExport.js or generateExcelWorkbook is not exported');
  }

  // 1. Placeholder sample data with generic details (NO private credentials)
  const sampleData = {
    docNo: 'QT-202611-0001',
    docDate: '2026-11-03',
    jobTitle: 'งานระบบน้ำอัตโนมัติ',
    customer: {
      name: 'บริษัท ตัวอย่าง จำกัด',
      address: '82 ถนนสุขุมวิท แขวงคลองเตย เขตคลองเตย กรุงเทพฯ'
    },
    seller: {
      name: 'นายสมชาย ใจดี',
      address: '88/99 ถนนสุขใจ ตำบลสุขใจ อำเภอสุขใจ จังหวัดปทุมธานี 12150',
      tel1: '08X-XXX-XXXX',
      tel2: '08Y-YYY-YYYY',
      bank: 'ธ.กรุงเทพ XXX-X-XXXXX-X'
    },
    items: [
      { id: '1', type: 'item', itemNo: '1', description: 'ระบบรดน้ำอัตโนมัติ 4 โซน (ด้านหน้า)', unit: 'ชุด', qty: 1, unitPrice: 57552.9, amount: 57552.9 },
      { id: '2', type: 'bullet', description: ' - หัวป็อบอัพ 4 หุนจำนวน 5 หัว' },
      { id: '3', type: 'bullet', description: ' - ท่อ PE 20 มิลลิเมตร' },
      { id: '4', type: 'note', description: 'การเบิกค่างวดงาน' },
      { id: '5', type: 'bullet', description: 'เบิกมัดจำ 70% และ 30% เมื่องานเสร็จสิ้น' }
    ],
    totalAmount: 57552.9,
    totalTextThai: '(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)'
  };

  console.log('Testing generateExcelWorkbook...');
  const wb = await generateExcelWorkbook(sampleData, ExcelJS);
  assert(wb, 'Workbook must be returned');

  const ws = wb.getWorksheet('Sheet1');
  assert(ws, 'Worksheet Sheet1 must exist');

  // Verify Page Setup
  assert.strictEqual(ws.pageSetup.orientation, 'portrait');
  assert.strictEqual(ws.pageSetup.fitToPage, true);
  assert.strictEqual(ws.pageSetup.fitToWidth, 1);
  assert.strictEqual(ws.pageSetup.fitToHeight, 1);

  // Verify Column Widths (matching original template)
  assert.strictEqual(ws.getColumn('A').width, 5.75, 'Column A width should be 5.75');
  assert.strictEqual(ws.getColumn('B').width, 54.88, 'Column B width should be 54.88');
  assert.strictEqual(ws.getColumn('C').width, 6.88, 'Column C width should be 6.88');
  assert.strictEqual(ws.getColumn('D').width, 6.75, 'Column D width should be 6.75');
  assert.strictEqual(ws.getColumn('E').width, 9.63, 'Column E width should be 9.63');
  assert.strictEqual(ws.getColumn('F').width, 10.25, 'Column F width should be 10.25');

  // Verify Row Heights (matching original template with Row 3 removed)
  assert.strictEqual(ws.getRow(1).height, 30.75, 'Row 1 height should match template');
  assert.strictEqual(ws.getRow(2).height, 23.25, 'Row 2 height should match template');
  assert.strictEqual(ws.getRow(3).height, 25.5, 'Row 3 height should match template');
  assert.strictEqual(ws.getRow(4).height, 21.75, 'Row 4 height should match template');
  assert.strictEqual(ws.getRow(5).height, 26.25, 'Row 5 height should match template');
  assert.strictEqual(ws.getRow(9).height, 24, 'Row 9 height should match template');
  assert.strictEqual(ws.getRow(10).height, 18.75, 'Row 10 height should match template');
  assert.strictEqual(ws.getRow(11).height, 24, 'Row 11 height should match template');
  assert.strictEqual(ws.getRow(28).height, 21.95, 'Row 28 height should match template');
  assert.strictEqual(ws.getRow(29).height, 20.25, 'Row 29 height should match template');
  assert.strictEqual(ws.getRow(30).height, 11.25, 'Row 30 height should match template');
  assert.strictEqual(ws.getRow(31).height, 24, 'Row 31 height should match template');
  assert.strictEqual(ws.getRow(32).height, 21.95, 'Row 32 height should match template');
  assert.strictEqual(ws.getRow(33).height, 24, 'Row 33 height should match template');

  // Verify Headers
  assert.strictEqual(ws.getCell('A1').value, '   ใบเสนอราคา');
  assert.strictEqual(ws.getCell('A2').value, '   QUOTATION');
  assert(ws.getCell('A3').value.includes('บริษัท ตัวอย่าง จำกัด'), 'A3 must include customer name');
  assert.strictEqual(ws.getCell('D3').value, 'เลขที่/ No.');
  assert.strictEqual(ws.getCell('F3').value, 'QT-202611-0001');

  // Verify Customer & Date Info
  assert.strictEqual(ws.getCell('A4').value, 'ที่อยู่  :');
  assert.strictEqual(ws.getCell('B4').value, '82 ถนนสุขุมวิท แขวงคลองเตย เขตคลองเตย กรุงเทพฯ');
  assert.strictEqual(ws.getCell('D4').value, 'วันที่ Date ');
  assert.strictEqual(ws.getCell('F4').value, '3 พ.ย. 69');
  assert.strictEqual(ws.getCell('D5').value, 'งานระบบน้ำอัตโนมัติ');

  // Verify Seller Info
  assert.strictEqual(ws.getCell('A6').value, 'ชื่อ : นายสมชาย ใจดี');
  assert.strictEqual(ws.getCell('A7').value, 'ที่อยู่ : 88/99 ถนนสุขใจ ตำบลสุขใจ อำเภอสุขใจ จังหวัดปทุมธานี 12150');
  assert.strictEqual(ws.getCell('D7').value, 'TEL. 08X-XXX-XXXX');
  assert.strictEqual(ws.getCell('F7').value, '08Y-YYY-YYYY');
  assert.strictEqual(ws.getCell('D8').value, 'ธ.กรุงเทพ XXX-X-XXXXX-X');

  // Verify Table Headers
  assert.strictEqual(ws.getCell('A9').value, 'ลำดับที่');
  assert.strictEqual(ws.getCell('A10').value, 'ITEM');
  assert.strictEqual(ws.getCell('B9').value, 'รายการ');
  assert.strictEqual(ws.getCell('B10').value, 'DESCRIPTION');
  assert.strictEqual(ws.getCell('C9').value, 'หน่วย');
  assert.strictEqual(ws.getCell('D9').value, 'จำนวน');
  assert.strictEqual(ws.getCell('E9').value, 'ราคา/หน่วย');
  assert.strictEqual(ws.getCell('F9').value, 'จำนวนเงิน');

  // Verify Data Rows (11-27)
  assert.strictEqual(ws.getCell('A11').value, '1');
  assert.strictEqual(ws.getCell('B11').value, 'ระบบรดน้ำอัตโนมัติ 4 โซน (ด้านหน้า)');
  assert.strictEqual(ws.getCell('C11').value, 'ชุด');
  assert.strictEqual(ws.getCell('D11').value, 1);
  assert.strictEqual(ws.getCell('E11').value, 57552.9);
  assert.strictEqual(ws.getCell('F11').value, 57552.9);
  assert.strictEqual(ws.getCell('B12').value, ' - หัวป็อบอัพ 4 หุนจำนวน 5 หัว');
  assert.strictEqual(ws.getCell('B13').value, ' - ท่อ PE 20 มิลลิเมตร');
  assert.strictEqual(ws.getCell('B14').value, 'การเบิกค่างวดงาน');
  assert.strictEqual(ws.getCell('B15').value, 'เบิกมัดจำ 70% และ 30% เมื่องานเสร็จสิ้น');

  // Verify row 27 has bottom border
  assert.strictEqual(ws.getCell('A27').border.bottom.style, 'thin');
  assert.strictEqual(ws.getCell('F27').border.bottom.style, 'thin');

  // Verify Row 28 Total & SUM Formula
  assert.strictEqual(ws.getCell('A28').value, '(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)');
  assert.strictEqual(ws.getCell('D28').value, 'รวมเงิน');
  assert.deepStrictEqual(ws.getCell('F28').value, {
    formula: 'SUM(F11:F27)',
    result: 57552.9
  });
  assert.strictEqual(ws.getCell('F28').numFmt, '#,##0.00');

  // Verify Signatures
  assert.strictEqual(ws.getCell('A31').value, 'ผู้อนุมัติ :');
  assert.strictEqual(ws.getCell('D31').value, '   ผู้เสนอ :');
  assert.strictEqual(ws.getCell('E31').value, 'นายสมชาย ใจดี');
  assert.strictEqual(ws.getCell('A32').value, 'ตำแหน่ง ');
  assert.strictEqual(ws.getCell('E32').value, '3 พฤศจิกายน 2569');
  assert.strictEqual(ws.getCell('B33').value, '   (                                                                                )');

  // Verify Merges
  const merges = ws.model.merges;
  const expectedMerges = [
    'A1:F1', 'A2:F2', 'A3:C3', 'D3:E3', 'B4:C4', 'D4:E4', 'D5:F5',
    'A6:C6', 'A7:C7', 'D7:E7', 'D8:F8', 'A28:C28', 'D28:E28', 'E31:F31', 'E32:F32',
    'C9:C10', 'D9:D10', 'E9:E10', 'F9:F10'
  ];
  expectedMerges.forEach(mergeRange => {
    assert(merges.includes(mergeRange), `Merged range ${mergeRange} must be present in worksheet`);
  });

  // 2. Test Filename generation and sanitization
  console.log('Testing filename generation...');
  const filename = getExportFilename(sampleData, 'xlsx');
  assert.strictEqual(filename, 'ใบเสนอราคา_บริษัท ตัวอย่าง จำกัด_QT-202611-0001.xlsx');

  const unsafeData = {
    customer: { name: 'บ.เอ/บี:ซี*?ดี<1>|2' },
    docNo: 'QT/2026/01'
  };
  const safeFilename = getExportFilename(unsafeData, 'xlsx');
  assert.strictEqual(safeFilename, 'ใบเสนอราคา_บ.เอ_บี_ซี__ดี_1__2_QT_2026_01.xlsx');

  // 3. Test Thai Date formatting helpers
  console.log('Testing Thai date helpers...');
  assert.strictEqual(formatThaiDateShort('2026-11-03'), '3 พ.ย. 69');
  assert.strictEqual(formatThaiDateLong('2026-11-03'), '3 พฤศจิกายน 2569');
  assert.strictEqual(formatThaiDateShort(''), '');
  assert.strictEqual(formatThaiDateLong(''), '');

  // 4. Test Browser UMD export
  console.log('Testing browser UMD export...');
  const code = fs.readFileSync(require.resolve('../js/excelExport.js'), 'utf8');
  const fakeWindow = {};
  const fn = new Function('self', 'module', 'exports', code);
  fn(fakeWindow, undefined, undefined);
  assert(fakeWindow.ExcelExporter, 'window.ExcelExporter must be defined');
  assert.strictEqual(typeof fakeWindow.ExcelExporter.generateExcelWorkbook, 'function');
  assert.strictEqual(typeof fakeWindow.ExcelExporter.getExportFilename, 'function');

  console.log('✅ Excel export test passed!');
}

runTests().catch(err => {
  console.error(err);
  process.exit(1);
});
