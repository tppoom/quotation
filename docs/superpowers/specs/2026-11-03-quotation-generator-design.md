# Design Specification: Web Quotation Generator (ระบบสร้างใบเสนอราคา)

**Date:** 2026-11-03  
**Target Repository:** `tppoom/quotation` (Path: `/Users/tppoom/Desktop/quotation`)  
**Deployment Target:** GitHub Pages (Static Web Application)

---

## 1. Overview & Objectives

Create a client-side web application for generating, managing, and exporting quotation documents (ใบเสนอราคา) that matches the format, styling, formulas, and visual appearance of the reference Excel file (`ใบเสนอราคา บ.เพิ่ม เอสเคท  จำกัด.xlsx`) with 100% fidelity.

### Key Requirements:
1. **Identical Excel Output**: Exported `.xlsx` file must replicate the template layout, merged cells, fonts, column widths, thin borders, Thai date formatting, and `SUM` formula.
2. **Multi-Format Export**:
   - **Excel (.xlsx)** via ExcelJS
   - **Image (.png)** via html2canvas (retina resolution)
   - **PDF (.pdf)** via jsPDF & printable stylesheet (`@media print` for A4 standard)
3. **Smart Quotation Numbering**:
   - Format: `QT-YYYYMM-XXXX` (e.g. `QT-202611-0001`).
   - Automatically tracks sequence in `localStorage` per month to guarantee non-duplicate running numbers.
4. **Local-First Storage**:
   - All quotations and drafts saved in the user's browser `localStorage`.
   - History list with Search, Edit, Duplicate, and Delete capabilities.
   - Seller Profile saved locally so sensitive business/personal information (name, address, telephone, bank account) is **never** committed or leaked into the public GitHub repository.
5. **Dynamic File Naming**:
   - Standard filename: `ใบเสนอราคา_[ชื่อบริษัทลูกค้า]_[เลขที่เอกสาร].[ext]` (sanitized for safe downloads).
6. **User Interface**:
   - Responsive split-view layout built with Tailwind CSS and Alpine.js.
   - Left pane: Clean, intuitive form for adding/reordering items, sub-items, bullet points, payment terms, and client details.
   - Right pane: Real-time Live Document Preview rendering the exact A4 document view.

---

## 2. System Architecture

```
+---------------------------------------------------------------------------------+
|                                 Web Browser                                     |
|                                                                                 |
|  +-----------------------------------+   +------------------------------------+ |
|  |     Left: Input Form & Controls   |   |     Right: Live A4 Document        | |
|  |  - Document Info (No, Date, Job)  |   |  - Real-time Paper Preview         | |
|  |  - Customer Info (Name, Address)  |   |  - Exact font/table proportions    | |
|  |  - Items, Bullets, Notes          |   |  - Thai Baht text display          | |
|  |  - Payment Terms                  |   |  - Signatures & Totals             | |
|  +-----------------+-----------------+   +------------------+-----------------+ |
|                    |                                        |                   |
|                    v                                        v                   |
|  +----------------------------------------------------------------------------+ |
|  |                         Alpine.js Reactive Store                           | |
|  +-----------------+-------------------+--------------------+-----------------+ |
|                    |                   |                    |                   |
|                    v                   v                    v                   |
|  +--------------------+  +-----------------------+  +-------------------------+ |
|  |    LocalStorage    |  |     Export Engine     |  |   Thai Baht Algorithm   | |
|  |  - Seller Profile  |  |  - ExcelJS (.xlsx)    |  |  - Number -> Thai Text  | |
|  |  - Quotations List |  |  - html2canvas (.png) |  |    (บาทถ้วน / สตางค์)   | |
|  |  - Sequence State  |  |  - jsPDF / Print (PDF)|  |                         | |
|  +--------------------+  +-----------------------+  +-------------------------+ |
+---------------------------------------------------------------------------------+
```

### Zero-Build Deployment:
- Static HTML file (`index.html`) using verified CDN libraries:
  - Alpine.js 3.x (Reactive UI)
  - Tailwind CSS (via Tailwind CDN)
  - ExcelJS (Full Excel formatting support)
  - html2canvas (Retina screenshot export)
  - jsPDF (Vector/Canvas A4 PDF export)
  - Lucide Icons (Clean UI icons)
- Direct deployment to GitHub Pages on branch `main` without requiring complicated build actions.

---

## 3. Data Schema & LocalStorage Model

### 3.1 Quotation Object Schema
```typescript
interface QuotationItem {
  id: string;
  type: 'item' | 'bullet' | 'heading' | 'note';
  itemNo?: string;        // e.g. "1", "2"
  description: string;    // e.g. "ระบบรดน้ำอัตโนมัติ 4 โซน (ด้านหน้า)" or " - หัวป็อบอัพ..."
  unit?: string;          // e.g. "ชุด", "งาน", "จุด"
  qty?: number;           // e.g. 1
  unitPrice?: number;     // e.g. 57552.90
  amount?: number;        // e.g. 57552.90 (auto-calculated qty * unitPrice)
}

interface Quotation {
  id: string;              // uuid / timestamp
  docNo: string;           // e.g. "QT-202611-0001"
  docDate: string;         // YYYY-MM-DD
  jobTitle: string;        // e.g. "งานระบบน้ำอัตโนมัติ"
  
  // Customer details
  customer: {
    name: string;          // e.g. "บริษัท เพิ่ม เอสเคท จำกัด"
    address: string;       // e.g. "82 ถนนสังคมสงเคราะห์ แขวงลาดพร้าว กรุงเทพฯ"
  };
  
  // Seller details (loaded from profile, editable per doc)
  seller: {
    name: string;          // e.g. "นายชัยพจน์ ศีลรัชต์ภัตร์"
    address: string;       // e.g. "89/332 หมู่บ้านฟ้าชมพฤกษ์..."
    tel1: string;          // e.g. "09 0864 3917"
    tel2: string;          // e.g. "09 5742 1247"
    bank: string;          // e.g. "ธ.กรุงเทพ 9330144529"
  };
  
  items: QuotationItem[];
  paymentTerms: string[];  // e.g. ["การเบิกค่างวดงาน", "เบิกมัดจำ 70% และ 30% เมื่องานเสร็จสิ้น"]
  
  // Totals
  subtotal: number;
  discount: number;
  vatRate: number;         // 0 or 7% (default 0 as per template)
  totalAmount: number;
  totalTextThai: string;   // e.g. "(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)"
  
  approver: {
    name: string;
    title: string;
  };
  
  createdAt: string;       // ISO Timestamp
  updatedAt: string;       // ISO Timestamp
}
```

### 3.2 LocalStorage Keys
1. `quotation_seller_profile`: Saved default seller profile so user configures their identity once.
2. `quotation_history`: Array of `Quotation` records.
3. `quotation_sequence`: Object tracking monthly counters `{ "202611": 1 }`.
4. `quotation_current_draft`: Auto-saved active draft to prevent data loss on page refresh.

---

## 4. Excel Template Specification (100% ExcelJS Match)

From analyzing `ใบเสนอราคา บ.เพิ่ม เอสเคท  จำกัด.xlsx`:

| Row(s) | Col A (5.75) | Col B (54.88) | Col C (6.88) | Col D (6.75) | Col E (9.63) | Col F (10.25) | Format / Alignment / Border |
|---|---|---|---|---|---|---|---|
| **1** | <colspan 6> **ใบเสนอราคา** | | | | | | Bold 22pt, Center |
| **2** | <colspan 6> **QUOTATION** | | | | | | Bold 14pt, Center |
| **4** | <colspan 3> เรียน/Attention : [Customer Name] | | | <colspan 2> เลขที่/ No. | | [Doc No] | Left / Right |
| **5** | ที่อยู่  : | [Customer Address] | | <colspan 2> วันที่ Date | | [Thai Date] | Thai date `d mmm yy` |
| **6** | | | | <colspan 3> [Job Title] | | | Bold / Normal |
| **7** | <colspan 3> ชื่อ : [Seller Name] | | | | | | Left |
| **8** | <colspan 3> ที่อยู่ : [Seller Address] | | | <colspan 2> TEL. [Tel 1] | | [Tel 2] | |
| **9** | | | | <colspan 3> [Bank Account] | | | Left |
| **10** | ลำดับที่ | รายการ | หน่วย | จำนวน | ราคา/หน่วย | จำนวนเงิน | Thin outer & inner borders, Bold |
| **11** | ITEM | DESCRIPTION | (merged with 10) | (merged with 10) | (merged with 10) | (merged with 10) | Thin outer & inner borders |
| **12-28** | [No] | [Description / Bullet] | [Unit] | [Qty] | [Unit Price] | [Amount] | Vertical gridlines on cols, top/bottom closed |
| **29** | <colspan 3> ([Total in Thai Baht Text]) | | | <colspan 2> รวมเงิน | | [Total] | Thin borders, formula `=SUM(F12:F28)` |
| **32** | ผู้อนุมัติ : | | | ผู้เสนอ : | <colspan 2> [Seller Name] | Signature line |
| **33** | ตำแหน่ง | | | | <colspan 2> [Thai Full Date] | Signature line |
| **34** | | ( &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ) | | | | | Customer signature placeholder |

---

## 5. Export Engines

1. **Excel Export (`exportExcel()`)**:
   - Uses `ExcelJS.Workbook`.
   - Re-creates exact row heights, column widths, font families (Angsana New / CordiaUPC / Tahoma or modern equivalent Sarabun/Prompt fallback), cell borders, cell merges, and number formats (`#,##0.00`).
   - Inserts Thai Baht text in cell A29.
   - Saves file as: `ใบเสนอราคา_${customerName}_${docNo}.xlsx`.

2. **Image Export (`exportPNG()`)**:
   - Uses `html2canvas(document.getElementById('quotation-preview'), { scale: 2 })`.
   - Converts canvas to PNG blob.
   - Saves file as: `ใบเสนอราคา_${customerName}_${docNo}.png`.

3. **PDF Export (`exportPDF()`)**:
   - Generates vector-grade A4 PDF using jsPDF or direct print dialog with `@media print`.
   - Saves file as: `ใบเสนอราคา_${customerName}_${docNo}.pdf`.

---

## 6. Thai Baht Text Conversion (`bahtText(num)`)

Standard financial Thai Baht converter covering:
- Millions / billions recursions (`ล้าน`)
- Units: `หนึ่ง`, `เอ็ด` (for 21, 31, ...), `ยี่สิบ` (for 20-29)
- Tens: `สิบ`, `ร้อย`, `พัน`, `หมื่น`, `แสน`
- Satang handling: `...บาทถ้วน` if 0 satang, or `...บาท...สตางค์`
- Example: `57552.90` -> `(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)`

---

## 7. Privacy, Security & Repository Setup

1. **No Sensitive Data Leaks**:
   - The repository contains only generic placeholder values (e.g. "ชื่อผู้เสนอราคา / ร้านค้า", "08X-XXX-XXXX", "ธนาคาร...") in default state.
   - The user inputs their actual credentials into the local web application. It remains saved in `window.localStorage` on their machine only.
   - `.gitignore` will ignore any personal `.xlsx` or secret files.
2. **Repository & Hosting**:
   - Directory: `/Users/tppoom/Desktop/quotation`
   - Git remote: `https://github.com/tppoom/quotation.git`
   - GitHub Pages configured on branch `main` root `/`.

---

## 8. Verification & Test Plan

1. **Data Entry & Reactivity**: Verify modifying fields in form instantly updates preview.
2. **Thai Baht Conversion**: Test boundary numbers (`0`, `1`, `21`, `100`, `1001`, `57552.90`, `1000000.50`).
3. **Auto-Numbering**: Test creating successive quotations and verify non-overlapping running IDs (`QT-YYYYMM-0001`, `0002`).
4. **Excel Export**: Compare generated `.xlsx` with `ใบเสนอราคา บ.เพิ่ม เอสเคท  จำกัด.xlsx` cell-by-cell.
5. **Image & PDF Export**: Verify resolution, typography, and page boundaries.
6. **LocalStorage Persistence**: Verify refresh keeps draft; history saves, loads, duplicates, and deletes correctly.
7. **Git & GitHub Pages Deployment**: Verify repo creation, push, and active GitHub Pages URL.
