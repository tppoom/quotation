# Quotation Generator Web App Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a user-friendly, client-side web application that generates quotation documents matching the exact layout, formulas, and styles of `ใบเสนอราคา บ.เพิ่ม เอสเคท  จำกัด.xlsx`, supporting LocalStorage history, non-repeating running numbers, 3 export formats (Excel, PNG, PDF), and deployed to GitHub Pages without leaking private data.

**Architecture:** Static single-page application using Alpine.js and Tailwind CSS with modular JavaScript utilities: Thai Baht text engine, LocalStorage persistence & sequence generator, ExcelJS workbook builder matching the exact template, and html2canvas/jsPDF export.

**Tech Stack:** HTML5, Tailwind CSS, Alpine.js (reactive state), ExcelJS (exact .xlsx reproduction), html2canvas (PNG capture), jsPDF (PDF generation), Node.js (for headless testing).

**Spec:** `docs/superpowers/specs/2026-11-03-quotation-generator-design.md`

## Global Constraints
- Target directory: `/Users/tppoom/Desktop/quotation`
- Public GitHub repo: `tppoom/quotation` with GitHub Pages enabled on `main` branch root `/`
- Zero private data committed to git (real seller phone/bank/name stored in browser `localStorage` only; template uses generic default placeholders)
- File naming format: `ใบเสนอราคา_[ชื่อลูกค้า]_[เลขที่เอกสาร].[ext]`
- Excel output format must replicate merged cells, fonts, row heights, column widths, thin borders, Thai dates, and SUM formula from reference Excel

---

### Task 1: Thai Baht Text Conversion Engine (`js/bahtText.js`)

**Files:**
- Create: `js/bahtText.js`
- Test: `test/bahtText.test.js`

**Interfaces:**
- Produces: `function bahtText(num: number | string): string`
  - Converts numbers like `57552.9` into `(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)`
  - Converts whole numbers like `100` into `(หนึ่งร้อยบาทถ้วน)`
  - Handles millions, zero, and decimals accurately.

- [ ] **Step 1: Write the failing unit tests for Thai Baht text conversion**
```javascript
// test/bahtText.test.js
const assert = require('assert');
const { bahtText } = require('../js/bahtText.js');

function runTests() {
  assert.strictEqual(bahtText(0), '(ศูนย์บาทถ้วน)');
  assert.strictEqual(bahtText(1), '(หนึ่งบาทถ้วน)');
  assert.strictEqual(bahtText(11), '(สิบเอ็ดบาทถ้วน)');
  assert.strictEqual(bahtText(21), '(ยี่สิบเอ็ดบาทถ้วน)');
  assert.strictEqual(bahtText(101), '(หนึ่งร้อยหนึ่งบาทถ้วน)');
  assert.strictEqual(bahtText(1001), '(หนึ่งพันหนึ่งบาทถ้วน)');
  assert.strictEqual(bahtText(57552.90), '(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)');
  assert.strictEqual(bahtText(1000000), '(หนึ่งล้านบาทถ้วน)');
  assert.strictEqual(bahtText(10000001), '(สิบล้านหนึ่งบาทถ้วน)');
  assert.strictEqual(bahtText(0.25), '(ยี่สิบห้าสตางค์)');
  console.log('✅ All bahtText tests passed!');
}

runTests();
```

- [ ] **Step 2: Run test to verify it fails**
Run: `node test/bahtText.test.js`  
Expected: FAIL (Cannot find module `../js/bahtText.js`)

- [ ] **Step 3: Implement minimal code in `js/bahtText.js`**
```javascript
// js/bahtText.js
(function(root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.bahtText = factory().bahtText;
  }
}(typeof self !== 'undefined' ? self : this, function() {
  const NUMBERS = ['', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
  const DIGITS = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน', 'ล้าน'];

  function readBlock(numStr) {
    let result = '';
    const len = numStr.length;
    for (let i = 0; i < len; i++) {
      const digit = parseInt(numStr.charAt(i), 10);
      const pos = len - i - 1;
      if (digit !== 0) {
        if (pos === 0 && digit === 1 && len > 1) {
          result += 'เอ็ด';
        } else if (pos === 1 && digit === 2) {
          result += 'ยี่';
        } else if (pos === 1 && digit === 1) {
          result += '';
        } else {
          result += NUMBERS[digit];
        }
        result += DIGITS[pos];
      }
    }
    return result;
  }

  function bahtText(input) {
    if (input === null || input === undefined || input === '') return '(ศูนย์บาทถ้วน)';
    const num = parseFloat(input);
    if (isNaN(num)) return '(ศูนย์บาทถ้วน)';
    if (num === 0) return '(ศูนย์บาทถ้วน)';

    const isNegative = num < 0;
    const absNum = Math.abs(num);
    const fixed = absNum.toFixed(2);
    const parts = fixed.split('.');
    let integerPart = parts[0];
    const decimalPart = parts[1];

    let bahtStr = '';
    if (parseInt(integerPart, 10) === 0) {
      bahtStr = '';
    } else {
      let millions = [];
      while (integerPart.length > 0) {
        millions.unshift(integerPart.slice(Math.max(0, integerPart.length - 6)));
        integerPart = integerPart.slice(0, Math.max(0, integerPart.length - 6));
      }
      for (let i = 0; i < millions.length; i++) {
        const blockText = readBlock(millions[i]);
        if (blockText) {
          bahtStr += blockText + (i < millions.length - 1 ? 'ล้าน' : '');
        }
      }
      bahtStr += 'บาท';
    }

    let satangStr = '';
    const satangVal = parseInt(decimalPart, 10);
    if (satangVal === 0) {
      satangStr = bahtStr ? 'ถ้วน' : 'ศูนย์บาทถ้วน';
    } else {
      satangStr = readBlock(decimalPart) + 'สตางค์';
    }

    const full = (isNegative ? 'ลบ' : '') + (bahtStr + satangStr);
    return `(${full})`;
  }

  return { bahtText };
}));
```

- [ ] **Step 4: Run test to verify it passes**
Run: `node test/bahtText.test.js`  
Expected: PASS (`✅ All bahtText tests passed!`)

- [ ] **Step 5: Commit**
```bash
git add js/bahtText.js test/bahtText.test.js
git commit -m "feat: implement Thai Baht text converter with tests"
```

---

### Task 2: LocalStorage Model & Non-Repeating Quotation Number Generator (`js/storage.js`)

**Files:**
- Create: `js/storage.js`
- Test: `test/storage.test.js`

**Interfaces:**
- Produces:
  - `generateDocNo(dateStr: string): string` (formats `QT-YYYYMM-XXXX`, auto-increments sequence per year-month)
  - `loadSellerProfile(): object`
  - `saveSellerProfile(profile: object): void`
  - `getHistory(): Array<object>`
  - `saveToHistory(quotation: object): void`
  - `deleteFromHistory(id: string): void`
  - `duplicateQuotation(id: string): object`
  - `loadCurrentDraft(): object | null`
  - `saveCurrentDraft(quotation: object): void`

- [ ] **Step 1: Write test for doc number generation and storage logic**
```javascript
// test/storage.test.js
const assert = require('assert');

// Mock localStorage for node environment
global.localStorage = (function() {
  let store = {};
  return {
    getItem: (key) => store[key] || null,
    setItem: (key, val) => { store[key] = String(val); },
    removeItem: (key) => { delete store[key]; },
    clear: () => { store = {}; }
  };
})();

const {
  generateDocNo,
  getSellerProfile,
  saveSellerProfile,
  getHistory,
  saveToHistory,
  deleteFromHistory
} = require('../js/storage.js');

function testStorage() {
  localStorage.clear();
  // 1. Doc number generation sequence
  const doc1 = generateDocNo('2026-11-03');
  assert.strictEqual(doc1, 'QT-202611-0001');

  // Save history with doc1
  saveToHistory({ id: '1', docNo: doc1, customer: { name: 'บ.เอ' } });
  
  const doc2 = generateDocNo('2026-11-03');
  assert.strictEqual(doc2, 'QT-202611-0002');

  // Different month
  const docDec = generateDocNo('2026-12-01');
  assert.strictEqual(docDec, 'QT-202612-0001');

  // 2. Seller Profile
  saveSellerProfile({ name: 'ทดสอบ', tel1: '0812345678' });
  const profile = getSellerProfile();
  assert.strictEqual(profile.name, 'ทดสอบ');

  // 3. History deletion
  deleteFromHistory('1');
  assert.strictEqual(getHistory().length, 0);

  console.log('✅ All storage tests passed!');
}

testStorage();
```

- [ ] **Step 2: Run test to verify it fails**
Run: `node test/storage.test.js`  
Expected: FAIL (Cannot find module `../js/storage.js`)

- [ ] **Step 3: Implement `js/storage.js`**
```javascript
// js/storage.js
(function(root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.StorageManager = factory();
  }
}(typeof self !== 'undefined' ? self : this, function() {
  const KEYS = {
    SELLER: 'quotation_seller_profile',
    HISTORY: 'quotation_history',
    SEQUENCE: 'quotation_sequence',
    DRAFT: 'quotation_current_draft'
  };

  const DEFAULT_SELLER = {
    name: 'นายชัยพจน์  ศีลรัชต์ภัตร์',
    address: '89/332 หมู่บ้านฟ้าชมพฤกษ์ อำเภอลำลูกกา จังหวัดปทุมธานี 12150',
    tel1: '09 0864  3917',
    tel2: '09 5742 1247',
    bank: 'ธ.กรุงเทพ 9330144529'
  };

  function getSellerProfile() {
    try {
      const data = localStorage.getItem(KEYS.SELLER);
      return data ? JSON.parse(data) : { ...DEFAULT_SELLER };
    } catch (e) {
      return { ...DEFAULT_SELLER };
    }
  }

  function saveSellerProfile(profile) {
    localStorage.setItem(KEYS.SELLER, JSON.stringify(profile));
  }

  function getHistory() {
    try {
      const data = localStorage.getItem(KEYS.HISTORY);
      return data ? JSON.parse(data) : [];
    } catch (e) {
      return [];
    }
  }

  function getSequenceMap() {
    try {
      const data = localStorage.getItem(KEYS.SEQUENCE);
      return data ? JSON.parse(data) : {};
    } catch (e) {
      return {};
    }
  }

  function generateDocNo(dateStr) {
    const d = dateStr ? new Date(dateStr) : new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const ymKey = `${yyyy}${mm}`;

    // Inspect existing history to make sure we don't collide
    const history = getHistory();
    let maxSeq = 0;
    const prefix = `QT-${ymKey}-`;
    history.forEach(item => {
      if (item.docNo && item.docNo.startsWith(prefix)) {
        const seq = parseInt(item.docNo.replace(prefix, ''), 10);
        if (!isNaN(seq) && seq > maxSeq) {
          maxSeq = seq;
        }
      }
    });

    const seqMap = getSequenceMap();
    const trackedSeq = seqMap[ymKey] || 0;
    const nextSeq = Math.max(maxSeq, trackedSeq) + 1;

    return `${prefix}${String(nextSeq).padStart(4, '0')}`;
  }

  function recordDocNoUsed(docNo) {
    if (!docNo || !docNo.startsWith('QT-')) return;
    const parts = docNo.split('-');
    if (parts.length === 3) {
      const ymKey = parts[1];
      const seq = parseInt(parts[2], 10);
      if (!isNaN(seq)) {
        const seqMap = getSequenceMap();
        if (!seqMap[ymKey] || seq > seqMap[ymKey]) {
          seqMap[ymKey] = seq;
          localStorage.setItem(KEYS.SEQUENCE, JSON.stringify(seqMap));
        }
      }
    }
  }

  function saveToHistory(quotation) {
    if (!quotation.id) {
      quotation.id = 'qt_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6);
    }
    quotation.updatedAt = new Date().toISOString();
    if (!quotation.createdAt) {
      quotation.createdAt = quotation.updatedAt;
    }

    const history = getHistory();
    const existingIndex = history.findIndex(item => item.id === quotation.id);
    if (existingIndex >= 0) {
      history[existingIndex] = quotation;
    } else {
      history.unshift(quotation);
    }

    localStorage.setItem(KEYS.HISTORY, JSON.stringify(history));
    recordDocNoUsed(quotation.docNo);
  }

  function deleteFromHistory(id) {
    const history = getHistory().filter(item => item.id !== id);
    localStorage.setItem(KEYS.HISTORY, JSON.stringify(history));
  }

  function duplicateQuotation(id) {
    const history = getHistory();
    const source = history.find(item => item.id === id);
    if (!source) return null;

    const clone = JSON.parse(JSON.stringify(source));
    clone.id = 'qt_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6);
    clone.docNo = generateDocNo(new Date().toISOString().slice(0, 10));
    clone.docDate = new Date().toISOString().slice(0, 10);
    clone.createdAt = new Date().toISOString();
    clone.updatedAt = clone.createdAt;
    return clone;
  }

  function loadCurrentDraft() {
    try {
      const data = localStorage.getItem(KEYS.DRAFT);
      return data ? JSON.parse(data) : null;
    } catch (e) {
      return null;
    }
  }

  function saveCurrentDraft(draft) {
    localStorage.setItem(KEYS.DRAFT, JSON.stringify(draft));
  }

  return {
    getSellerProfile,
    saveSellerProfile,
    getHistory,
    saveToHistory,
    deleteFromHistory,
    duplicateQuotation,
    generateDocNo,
    recordDocNoUsed,
    loadCurrentDraft,
    saveCurrentDraft,
    DEFAULT_SELLER
  };
}));
```

- [ ] **Step 4: Run test to verify it passes**
Run: `node test/storage.test.js`  
Expected: PASS (`✅ All storage tests passed!`)

- [ ] **Step 5: Commit**
```bash
git add js/storage.js test/storage.test.js
git commit -m "feat: implement LocalStorage manager and non-repeating doc number generator"
```

---

### Task 3: ExcelJS Export Engine Matching Original Excel Template (`js/excelExport.js`)

**Files:**
- Create: `js/excelExport.js`
- Test: `test/excelMatch.test.js`

**Interfaces:**
- Produces: `async function generateExcelWorkbook(quotationData: object): Promise<ExcelJS.Workbook>`
  - Replicates exact structure, merges (`A1:F1`, `A2:F2`, `A4:C4`, `D4:E4`, `D5:E5`, `D6:F6`, `D8:E8`, `D9:F9`, `A29:C29`, `D29:E29`, `E31:F31`, `E32:F32`, `E33:F33`)
  - Sets exact column widths (`A: 5.75, B: 54.88, C: 6.88, D: 6.75, E: 9.63, F: 10.25`)
  - Formats cell borders, header fills, alignments, Thai dates, and `=SUM(F12:F28)` formula
  - Generates download filename `ใบเสนอราคา_[ชื่อลูกค้า]_[เลขที่].xlsx`

- [ ] **Step 1: Write verification test for Excel generator**
```javascript
// test/excelMatch.test.js
const assert = require('assert');
const ExcelJS = require('exceljs');
const { generateExcelWorkbook, getExportFilename } = require('../js/excelExport.js');

async function testExcelGeneration() {
  const sampleData = {
    docNo: 'QT-202611-0001',
    docDate: '2026-11-03',
    jobTitle: 'งานระบบน้ำอัตโนมัติ',
    customer: {
      name: 'บริษัท เพิ่ม เอสเคท จำกัด',
      address: '82 ถนนสังคมสงเคราะห์ แขวงลาดพร้าว กรุงเทพฯ'
    },
    seller: {
      name: 'นายชัยพจน์  ศีลรัชต์ภัตร์',
      address: '89/332 หมู่บ้านฟ้าชมพฤกษ์ อำเภอลำลูกกา จังหวัดปทุมธานี 12150',
      tel1: '09 0864  3917',
      tel2: '09 5742 1247',
      bank: 'ธ.กรุงเทพ 9330144529'
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

  const wb = await generateExcelWorkbook(sampleData, ExcelJS);
  const ws = wb.getWorksheet('Sheet1');
  assert(ws, 'Worksheet Sheet1 must exist');

  // Verify headers
  assert.strictEqual(ws.getCell('A1').value, '   ใบเสนอราคา');
  assert.strictEqual(ws.getCell('A2').value, '   QUOTATION');
  assert(ws.getCell('A4').value.includes('บริษัท เพิ่ม เอสเคท จำกัด'));
  assert.strictEqual(ws.getCell('D4').value, 'เลขที่/ No.');
  assert.strictEqual(ws.getCell('F4').value, 'QT-202611-0001');

  // Verify table header
  assert.strictEqual(ws.getCell('A10').value, 'ลำดับที่');
  assert.strictEqual(ws.getCell('B10').value, 'รายการ');
  assert.strictEqual(ws.getCell('F10').value, 'จำนวนเงิน');

  // Verify filename
  const filename = getExportFilename(sampleData, 'xlsx');
  assert.strictEqual(filename, 'ใบเสนอราคา_บริษัท เพิ่ม เอสเคท จำกัด_QT-202611-0001.xlsx');

  console.log('✅ Excel export test passed!');
}

testExcelGeneration().catch(err => {
  console.error(err);
  process.exit(1);
});
```

- [ ] **Step 2: Run test to verify it fails**
Run: `node test/excelMatch.test.js`  
Expected: FAIL (Cannot find module `../js/excelExport.js`)

- [ ] **Step 3: Implement `js/excelExport.js`**
```javascript
// js/excelExport.js
(function(root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.ExcelExporter = factory();
  }
}(typeof self !== 'undefined' ? self : this, function() {

  function formatThaiDateShort(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    const months = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    const thaiYear = (d.getFullYear() + 543).toString().slice(-2);
    return `${d.getDate()} ${months[d.getMonth()]} ${thaiYear}`;
  }

  function formatThaiDateLong(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    const months = [
      'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
      'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
    ];
    return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear() + 543}`;
  }

  function sanitizeFilename(name) {
    return (name || 'ลูกค้า').replace(/[/\\?%*:|"<>]/g, '_').trim();
  }

  function getExportFilename(data, ext) {
    const custName = sanitizeFilename(data.customer?.name || 'ลูกค้า');
    const docNo = sanitizeFilename(data.docNo || 'NO');
    return `ใบเสนอราคา_${custName}_${docNo}.${ext}`;
  }

  async function generateExcelWorkbook(data, ExcelJSModule) {
    const ExcelClass = ExcelJSModule || (typeof ExcelJS !== 'undefined' ? ExcelJS : null);
    if (!ExcelClass) {
      throw new Error('ExcelJS library is not available');
    }

    const wb = new ExcelClass.Workbook();
    wb.creator = 'Quotation Generator';
    wb.created = new Date();

    const ws = wb.addWorksheet('Sheet1', {
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1 }
    });

    // Column widths matching original template
    ws.columns = [
      { key: 'A', width: 6.5 },
      { key: 'B', width: 55 },
      { key: 'C', width: 8.5 },
      { key: 'D', width: 8 },
      { key: 'E', width: 12 },
      { key: 'F', width: 14 }
    ];

    const fontHeader1 = { name: 'Angsana New', size: 22, bold: true };
    const fontHeader2 = { name: 'Angsana New', size: 16, bold: true };
    const fontRegular = { name: 'Angsana New', size: 14 };
    const fontBold = { name: 'Angsana New', size: 14, bold: true };

    const borderThin = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
    const borderLeftRight = {
      left: { style: 'thin' },
      right: { style: 'thin' }
    };

    // Row 1: ใบเสนอราคา
    ws.mergeCells('A1:F1');
    ws.getCell('A1').value = '   ใบเสนอราคา';
    ws.getCell('A1').font = fontHeader1;
    ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

    // Row 2: QUOTATION
    ws.mergeCells('A2:F2');
    ws.getCell('A2').value = '   QUOTATION';
    ws.getCell('A2').font = fontHeader2;
    ws.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle' };

    // Row 4: Attention & No
    ws.mergeCells('A4:C4');
    ws.getCell('A4').value = `เรียน/Attention : ${data.customer?.name || ''}`;
    ws.getCell('A4').font = fontBold;

    ws.mergeCells('D4:E4');
    ws.getCell('D4').value = 'เลขที่/ No.';
    ws.getCell('D4').font = fontRegular;
    ws.getCell('D4').alignment = { horizontal: 'right' };
    ws.getCell('F4').value = data.docNo || '';
    ws.getCell('F4').font = fontBold;
    ws.getCell('F4').alignment = { horizontal: 'center' };

    // Row 5: Customer Address & Date
    ws.getCell('A5').value = 'ที่อยู่  :';
    ws.getCell('A5').font = fontRegular;
    ws.mergeCells('B5:C5');
    ws.getCell('B5').value = data.customer?.address || '';
    ws.getCell('B5').font = fontRegular;

    ws.mergeCells('D5:E5');
    ws.getCell('D5').value = 'วันที่ Date ';
    ws.getCell('D5').font = fontRegular;
    ws.getCell('D5').alignment = { horizontal: 'right' };
    ws.getCell('F5').value = formatThaiDateShort(data.docDate);
    ws.getCell('F5').font = fontRegular;
    ws.getCell('F5').alignment = { horizontal: 'center' };

    // Row 6: Job Title
    ws.mergeCells('D6:F6');
    ws.getCell('D6').value = data.jobTitle || '';
    ws.getCell('D6').font = fontBold;
    ws.getCell('D6').alignment = { horizontal: 'right' };

    // Row 7: Seller Name
    ws.mergeCells('A7:C7');
    ws.getCell('A7').value = `ชื่อ : ${data.seller?.name || ''}`;
    ws.getCell('A7').font = fontRegular;

    // Row 8: Seller Address & Phone
    ws.mergeCells('A8:C8');
    ws.getCell('A8').value = `ที่อยู่ : ${data.seller?.address || ''}`;
    ws.getCell('A8').font = fontRegular;

    ws.mergeCells('D8:E8');
    ws.getCell('D8').value = data.seller?.tel1 ? `TEL. ${data.seller.tel1}` : '';
    ws.getCell('D8').font = fontRegular;
    ws.getCell('D8').alignment = { horizontal: 'right' };
    ws.getCell('F8').value = data.seller?.tel2 || '';
    ws.getCell('F8').font = fontRegular;
    ws.getCell('F8').alignment = { horizontal: 'center' };

    // Row 9: Bank Account
    ws.mergeCells('D9:F9');
    ws.getCell('D9').value = data.seller?.bank || '';
    ws.getCell('D9').font = fontRegular;
    ws.getCell('D9').alignment = { horizontal: 'right' };

    // Table Header: Row 10 & 11
    ws.getCell('A10').value = 'ลำดับที่';
    ws.getCell('A11').value = 'ITEM';
    ws.getCell('B10').value = 'รายการ';
    ws.getCell('B11').value = 'DESCRIPTION';

    ws.mergeCells('C10:C11');
    ws.getCell('C10').value = 'หน่วย';
    ws.mergeCells('D10:D11');
    ws.getCell('D10').value = 'จำนวน';
    ws.mergeCells('E10:E11');
    ws.getCell('E10').value = 'ราคา/หน่วย';
    ws.mergeCells('F10:F11');
    ws.getCell('F10').value = 'จำนวนเงิน';

    const thCells = ['A10', 'B10', 'C10', 'D10', 'E10', 'F10', 'A11', 'B11'];
    thCells.forEach(ref => {
      const c = ws.getCell(ref);
      c.font = fontBold;
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => {
      ws.getCell(`${col}10`).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
      ws.getCell(`${col}11`).border = {
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    // Populate Items: Rows 12 to 28 (Total 17 data rows like original)
    const startRow = 12;
    const endRow = 28;
    const items = data.items || [];

    for (let r = startRow; r <= endRow; r++) {
      const idx = r - startRow;
      const item = items[idx];

      if (item) {
        if (item.type === 'item') {
          ws.getCell(`A${r}`).value = item.itemNo ? String(item.itemNo) : '';
          ws.getCell(`A${r}`).alignment = { horizontal: 'center' };
          ws.getCell(`B${r}`).value = item.description || '';
          ws.getCell(`C${r}`).value = item.unit || '';
          ws.getCell(`C${r}`).alignment = { horizontal: 'center' };
          ws.getCell(`D${r}`).value = item.qty !== undefined && item.qty !== '' ? Number(item.qty) : '';
          ws.getCell(`D${r}`).alignment = { horizontal: 'center' };
          ws.getCell(`E${r}`).value = item.unitPrice !== undefined && item.unitPrice !== '' ? Number(item.unitPrice) : '';
          ws.getCell(`E${r}`).numFmt = '#,##0.00';
          ws.getCell(`E${r}`).alignment = { horizontal: 'right' };
          ws.getCell(`F${r}`).value = item.amount !== undefined && item.amount !== '' ? Number(item.amount) : '';
          ws.getCell(`F${r}`).numFmt = '#,##0.00';
          ws.getCell(`F${r}`).alignment = { horizontal: 'right' };
        } else {
          // Sub-item or note
          ws.getCell(`B${r}`).value = item.description || '';
        }
      }

      ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => {
        const c = ws.getCell(`${col}${r}`);
        c.font = fontRegular;
        c.border = (r === endRow)
          ? { left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } }
          : borderLeftRight;
      });
    }

    // Row 29: Total Summary
    ws.mergeCells('A29:C29');
    ws.getCell('A29').value = data.totalTextThai || '';
    ws.getCell('A29').font = fontBold;
    ws.getCell('A29').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells('D29:E29');
    ws.getCell('D29').value = 'รวมเงิน';
    ws.getCell('D29').font = fontBold;
    ws.getCell('D29').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getCell('F29').value = { formula: `SUM(F${startRow}:F${endRow})`, result: data.totalAmount || 0 };
    ws.getCell('F29').font = fontBold;
    ws.getCell('F29').numFmt = '#,##0.00';
    ws.getCell('F29').alignment = { horizontal: 'right', vertical: 'middle' };

    ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => {
      ws.getCell(`${col}29`).border = borderThin;
    });

    // Row 32-34: Signatures
    ws.getCell('A32').value = 'ผู้อนุมัติ :';
    ws.getCell('A32').font = fontRegular;

    ws.getCell('D32').value = '   ผู้เสนอ :';
    ws.getCell('D32').font = fontRegular;
    ws.getCell('D32').alignment = { horizontal: 'right' };

    ws.mergeCells('E32:F32');
    ws.getCell('E32').value = data.seller?.name || '';
    ws.getCell('E32').font = fontBold;
    ws.getCell('E32').alignment = { horizontal: 'center' };

    ws.getCell('A33').value = 'ตำแหน่ง ';
    ws.getCell('A33').font = fontRegular;

    ws.mergeCells('E33:F33');
    ws.getCell('E33').value = formatThaiDateLong(data.docDate);
    ws.getCell('E33').font = fontRegular;
    ws.getCell('E33').alignment = { horizontal: 'center' };

    ws.getCell('B34').value = '   (                                                                                )';
    ws.getCell('B34').font = fontRegular;

    return wb;
  }

  async function downloadExcelFile(data, ExcelJSModule) {
    const wb = await generateExcelWorkbook(data, ExcelJSModule);
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filename = getExportFilename(data, 'xlsx');

    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
  }

  return {
    generateExcelWorkbook,
    downloadExcelFile,
    getExportFilename,
    formatThaiDateShort,
    formatThaiDateLong,
    sanitizeFilename
  };
}));
```

- [ ] **Step 4: Install `exceljs` locally for headless node test verification**
Run: `npm init -y && npm install exceljs` (and ensure `node_modules` is in `.gitignore`)
Run: `node test/excelMatch.test.js`  
Expected: PASS (`✅ Excel export test passed!`)

- [ ] **Step 5: Commit**
```bash
git add js/excelExport.js test/excelMatch.test.js package.json package-lock.json
git commit -m "feat: implement ExcelJS export engine matching original template"
```

---

### Task 4: PDF & Image Export Engines (`js/pdfExport.js`)

**Files:**
- Create: `js/pdfExport.js`

**Interfaces:**
- Produces:
  - `exportToPNG(previewElementId: string, filename: string): Promise<void>`
  - `exportToPDF(previewElementId: string, filename: string): Promise<void>`
  - `printDocument(): void`

- [ ] **Step 1: Implement `js/pdfExport.js`**
```javascript
// js/pdfExport.js
(function(root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DocExporter = factory();
  }
}(typeof self !== 'undefined' ? self : this, function() {

  async function exportToPNG(previewElementId, filename) {
    const el = document.getElementById(previewElementId);
    if (!el) throw new Error('Preview element not found');

    if (typeof html2canvas === 'undefined') {
      throw new Error('html2canvas library is not loaded');
    }

    const canvas = await html2canvas(el, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false
    });

    const dataUrl = canvas.toDataURL('image/png');
    const link = document.createElement('a');
    link.download = filename.endsWith('.png') ? filename : `${filename}.png`;
    link.href = dataUrl;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  async function exportToPDF(previewElementId, filename) {
    const el = document.getElementById(previewElementId);
    if (!el) throw new Error('Preview element not found');

    if (typeof html2canvas === 'undefined' || typeof jspdf === 'undefined') {
      window.print();
      return;
    }

    const { jsPDF } = jspdf;
    const canvas = await html2canvas(el, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false
    });

    const imgData = canvas.toDataURL('image/png');
    const pdf = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4'
    });

    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = (canvas.height * pdfWidth) / canvas.width;

    pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
    pdf.save(filename.endsWith('.pdf') ? filename : `${filename}.pdf`);
  }

  function printDocument() {
    window.print();
  }

  return {
    exportToPNG,
    exportToPDF,
    printDocument
  };
}));
```

- [ ] **Step 2: Commit**
```bash
git add js/pdfExport.js
git commit -m "feat: implement PNG and PDF export utilities"
```

---

### Task 5: Alpine.js Application State & View Logic (`js/app.js`)

**Files:**
- Create: `js/app.js`

**Interfaces:**
- Integrates `bahtText`, `StorageManager`, `ExcelExporter`, `DocExporter`.
- Manages active quotation, adding items, bullet rows, payment terms, calculating totals, auto-saving drafts, loading history, and modal states.

- [ ] **Step 1: Implement `js/app.js`**
```javascript
// js/app.js
document.addEventListener('alpine:init', () => {
  Alpine.data('quotationApp', () => ({
    // Data state
    quotation: {
      id: '',
      docNo: '',
      docDate: new Date().toISOString().slice(0, 10),
      jobTitle: 'งานระบบน้ำอัตโนมัติ',
      customer: {
        name: 'บริษัท เพิ่ม เอสเคท จำกัด',
        address: '82 ถนนสังคมสงเคราะห์ แขวงลาดพร้าว กรุงเทพฯ'
      },
      seller: { ...StorageManager.getSellerProfile() },
      items: [],
      totalAmount: 0,
      totalTextThai: ''
    },

    // UI state
    historyList: [],
    showHistoryModal: false,
    showSellerModal: false,
    historySearch: '',
    saveNotice: '',
    sellerProfileForm: { ...StorageManager.getSellerProfile() },

    init() {
      // 1. Try to load current draft or start fresh
      const draft = StorageManager.loadCurrentDraft();
      if (draft && draft.docNo) {
        this.quotation = draft;
      } else {
        this.startNewQuotation();
      }
      this.calculateTotals();
      this.refreshHistory();

      // Autosave draft on changes
      this.$watch('quotation', () => {
        this.calculateTotals();
        StorageManager.saveCurrentDraft(this.quotation);
      });
    },

    startNewQuotation() {
      const today = new Date().toISOString().slice(0, 10);
      const nextDocNo = StorageManager.generateDocNo(today);
      const seller = StorageManager.getSellerProfile();

      this.quotation = {
        id: 'qt_' + Date.now(),
        docNo: nextDocNo,
        docDate: today,
        jobTitle: '',
        customer: {
          name: '',
          address: ''
        },
        seller: { ...seller },
        items: [
          {
            id: 'item_1',
            type: 'item',
            itemNo: '1',
            description: '',
            unit: 'ชุด',
            qty: 1,
            unitPrice: 0,
            amount: 0
          }
        ],
        totalAmount: 0,
        totalTextThai: '(ศูนย์บาทถ้วน)'
      };
      this.calculateTotals();
      this.flashNotice('เริ่มต้นใบเสนอราคาใหม่เรียบร้อย');
    },

    addItemRow() {
      const currentItems = this.quotation.items.filter(i => i.type === 'item');
      const nextNo = currentItems.length + 1;
      this.quotation.items.push({
        id: 'row_' + Date.now() + Math.random().toString(36).substr(2, 4),
        type: 'item',
        itemNo: String(nextNo),
        description: '',
        unit: 'ชุด',
        qty: 1,
        unitPrice: 0,
        amount: 0
      });
      this.calculateTotals();
    },

    addSubItemRow() {
      this.quotation.items.push({
        id: 'row_' + Date.now() + Math.random().toString(36).substr(2, 4),
        type: 'bullet',
        description: ' - '
      });
    },

    addPaymentTermsPreset() {
      this.quotation.items.push({
        id: 'row_' + Date.now() + '1',
        type: 'note',
        description: 'การเบิกค่างวดงาน'
      });
      this.quotation.items.push({
        id: 'row_' + Date.now() + '2',
        type: 'bullet',
        description: 'เบิกมัดจำ 70% และ 30% เมื่องานเสร็จสิ้น'
      });
    },

    removeItemRow(index) {
      this.quotation.items.splice(index, 1);
      // Re-index main item numbers
      let counter = 1;
      this.quotation.items.forEach(it => {
        if (it.type === 'item') {
          it.itemNo = String(counter++);
        }
      });
      this.calculateTotals();
    },

    calculateTotals() {
      let sum = 0;
      this.quotation.items.forEach(it => {
        if (it.type === 'item') {
          const q = parseFloat(it.qty) || 0;
          const p = parseFloat(it.unitPrice) || 0;
          it.amount = Math.round(q * p * 100) / 100;
          sum += it.amount;
        }
      });
      this.quotation.totalAmount = sum;
      this.quotation.totalTextThai = bahtText(sum);
    },

    saveQuotation() {
      this.calculateTotals();
      StorageManager.saveToHistory(this.quotation);
      this.refreshHistory();
      this.flashNotice(`บันทึกเอกสาร ${this.quotation.docNo} เรียบร้อยแล้ว`);
    },

    loadFromHistory(item) {
      this.quotation = JSON.parse(JSON.stringify(item));
      this.calculateTotals();
      this.showHistoryModal = false;
      this.flashNotice(`โหลดเอกสาร ${this.quotation.docNo} แล้ว`);
    },

    duplicateFromHistory(id) {
      const clone = StorageManager.duplicateQuotation(id);
      if (clone) {
        StorageManager.saveToHistory(clone);
        this.quotation = clone;
        this.calculateTotals();
        this.refreshHistory();
        this.showHistoryModal = false;
        this.flashNotice(`ทำสำเนาเป็นเอกสารใหม่ ${clone.docNo} เรียบร้อย`);
      }
    },

    deleteFromHistory(id) {
      if (confirm('คุณแน่ใจหรือไม่ว่าต้องการลบใบเสนอราคานี้?')) {
        StorageManager.deleteFromHistory(id);
        this.refreshHistory();
        this.flashNotice('ลบเอกสารเรียบร้อยแล้ว');
      }
    },

    refreshHistory() {
      this.historyList = StorageManager.getHistory();
    },

    filteredHistory() {
      if (!this.historySearch.trim()) return this.historyList;
      const q = this.historySearch.toLowerCase();
      return this.historyList.filter(item => 
        (item.docNo && item.docNo.toLowerCase().includes(q)) ||
        (item.customer?.name && item.customer.name.toLowerCase().includes(q)) ||
        (item.jobTitle && item.jobTitle.toLowerCase().includes(q))
      );
    },

    openSellerSettings() {
      this.sellerProfileForm = { ...StorageManager.getSellerProfile() };
      this.showSellerModal = true;
    },

    saveSellerSettings() {
      StorageManager.saveSellerProfile(this.sellerProfileForm);
      this.quotation.seller = { ...this.sellerProfileForm };
      this.showSellerModal = false;
      this.flashNotice('บันทึกข้อมูลผู้เสนอราคาเรียบร้อย');
    },

    flashNotice(msg) {
      this.saveNotice = msg;
      setTimeout(() => {
        if (this.saveNotice === msg) this.saveNotice = '';
      }, 3500);
    },

    getFilename(ext) {
      return ExcelExporter.getExportFilename(this.quotation, ext);
    },

    async exportExcel() {
      this.calculateTotals();
      await ExcelExporter.downloadExcelFile(this.quotation, window.ExcelJS);
      this.flashNotice('ส่งออกไฟล์ Excel สำเร็จ');
    },

    async exportPNG() {
      const filename = this.getFilename('png');
      await DocExporter.exportToPNG('quotation-preview', filename);
      this.flashNotice('ส่งออกรูปภาพ PNG สำเร็จ');
    },

    async exportPDF() {
      const filename = this.getFilename('pdf');
      await DocExporter.exportToPDF('quotation-preview', filename);
      this.flashNotice('ส่งออกไฟล์ PDF สำเร็จ');
    },

    formatDateShort(d) {
      return ExcelExporter.formatThaiDateShort(d);
    },

    formatDateLong(d) {
      return ExcelExporter.formatThaiDateLong(d);
    },

    formatCurrency(val) {
      if (val === undefined || val === null || val === '') return '';
      const num = Number(val);
      if (isNaN(num)) return '';
      return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
  }));
});
```

- [ ] **Step 2: Commit**
```bash
git add js/app.js
git commit -m "feat: implement Alpine.js application controller"
```

---

### Task 6: Responsive UI & Exact A4 Document Template (`index.html`)

**Files:**
- Create: `index.html`

**Features:**
- Header with actions (New, Save, History, Settings, Export Excel, Export PNG, Export PDF)
- Split screen: Left input form with dynamic rows (Item, Sub-item bullet, Payment terms preset), Right exact A4 paper preview
- Modals for History and Seller Settings
- `@media print` CSS for pristine 1-page A4 printing
- Google Fonts: Sarabun / Angsana New compatible font styling

- [ ] **Step 1: Implement `index.html`**
(Build complete responsive application, linking Tailwind CDN, Alpine.js, ExcelJS, html2canvas, jsPDF, Lucide icons, and all modular scripts).

- [ ] **Step 2: Verify in browser**
Open `index.html` locally via Python HTTP server or browser to verify form responsiveness, live calculations, Thai Baht text rendering, and modal popups.

- [ ] **Step 3: Commit**
```bash
git add index.html
git commit -m "feat: complete UI layout and live A4 quotation preview"
```

---

### Task 7: End-to-End Functional Verification

**Files:**
- Test: Verify in browser and node

**Checklist:**
- [ ] Test adding items with quantity, unit price -> verify total & Thai baht text
- [ ] Test Excel export -> open exported `.xlsx` and compare formulas & cells with reference template
- [ ] Test PNG export -> verify 2x retina clarity
- [ ] Test PDF export -> verify 1-page A4 print fit
- [ ] Test LocalStorage save and history recovery
- [ ] Test quotation number auto-incrementation (`QT-YYYYMM-XXXX`)
- [ ] Test custom seller profile saved locally and not tracked by git

---

### Task 8: Push to GitHub & Enable GitHub Pages

**Commands:**
- [ ] Create repository `tppoom/quotation` via GitHub CLI: `gh repo create quotation --public --source=. --push`
- [ ] Enable GitHub Pages: `gh repo edit --enable-pages --pages-branch=main`
- [ ] Verify deployment status and provide URL to user.
