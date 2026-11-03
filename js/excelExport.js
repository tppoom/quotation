// js/excelExport.js
(function(root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.ExcelExporter = factory();
  }
}(typeof self !== 'undefined' ? self : this, function() {

  function parseDateParts(dateInput) {
    if (!dateInput) return null;
    if (typeof dateInput === 'string') {
      const match = dateInput.trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (match) {
        return {
          year: parseInt(match[1], 10),
          month: parseInt(match[2], 10) - 1,
          day: parseInt(match[3], 10)
        };
      }
    }
    const d = dateInput instanceof Date ? dateInput : new Date(dateInput);
    if (isNaN(d.getTime())) return null;
    return {
      year: d.getFullYear(),
      month: d.getMonth(),
      day: d.getDate()
    };
  }

  function formatThaiDateShort(dateStr) {
    if (!dateStr) return '';
    const parts = parseDateParts(dateStr);
    if (!parts) return dateStr;
    const months = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    const thaiYear = (parts.year + 543).toString().slice(-2);
    return `${parts.day} ${months[parts.month]} ${thaiYear}`;
  }

  function formatThaiDateLong(dateStr) {
    if (!dateStr) return '';
    const parts = parseDateParts(dateStr);
    if (!parts) return dateStr;
    const months = [
      'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
      'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
    ];
    return `${parts.day} ${months[parts.month]} ${parts.year + 543}`;
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
    const ExcelClass = ExcelJSModule || (typeof ExcelJS !== 'undefined' ? ExcelJS : (typeof require !== 'undefined' ? require('exceljs') : null));
    if (!ExcelClass) {
      throw new Error('ExcelJS library is not available');
    }

    const wb = new ExcelClass.Workbook();
    wb.creator = 'Quotation Generator';
    wb.created = new Date();

    const ws = wb.addWorksheet('Sheet1', {
      views: [
        {
          state: 'normal',
          showGridLines: false,
          showRowColHeaders: true,
          zoomScale: 130,
          zoomScaleNormal: 130
        }
      ],
      pageSetup: {
        paperSize: 9,
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1,
        scale: 90,
        horizontalCentered: true,
        printArea: 'A1:F34',
        margins: {
          left: 0.3,
          right: 0.16,
          top: 0.59,
          bottom: 0,
          header: 0.59,
          footer: 0.16
        }
      }
    });

    // Column widths matching original template (A: 5.75, B: 54.88, C: 6.88, D: 6.75, E: 9.63, F: 10.25)
    ws.columns = [
      { key: 'A', width: 5.75 },
      { key: 'B', width: 54.88 },
      { key: 'C', width: 6.88 },
      { key: 'D', width: 6.75 },
      { key: 'E', width: 9.63 },
      { key: 'F', width: 10.25 }
    ];

    // Row heights matching reference template
    ws.getRow(1).height = 30.75;
    ws.getRow(2).height = 23.25;
    ws.getRow(3).height = 4.5;
    ws.getRow(4).height = 25.5;
    ws.getRow(5).height = 21.75;
    ws.getRow(6).height = 26.25;
    ws.getRow(10).height = 24;
    ws.getRow(11).height = 18.75;
    ws.getRow(12).height = 24;
    ws.getRow(17).height = 25.5;
    ws.getRow(19).height = 21.75;
    ws.getRow(21).height = 21.75;
    ws.getRow(22).height = 21.75;
    ws.getRow(23).height = 21.75;
    ws.getRow(24).height = 21.75;
    ws.getRow(29).height = 21.95;
    ws.getRow(30).height = 20.25;
    ws.getRow(31).height = 11.25;
    ws.getRow(33).height = 21.95;
    ws.getRow(34).height = 24;

    const fontHeader1 = { name: 'Angsana New', size: 22, bold: true };
    const fontHeader2 = { name: 'Angsana New', size: 16, bold: true };
    const fontRegular = { name: 'Angsana New', size: 14 };
    const fontBold = { name: 'Angsana New', size: 14, bold: true };

    const thin = { style: 'thin', color: { indexed: 64 } };
    const hair = { style: 'hair', color: { indexed: 64 } };

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
    ws.getCell('A4').alignment = { horizontal: 'left', vertical: 'middle' };

    ws.mergeCells('D4:E4');
    ws.getCell('D4').value = 'เลขที่/ No.';
    ws.getCell('D4').font = fontRegular;
    ws.getCell('D4').alignment = { horizontal: 'left' };
    ws.getCell('F4').value = data.docNo || '';
    ws.getCell('F4').font = fontBold;
    ws.getCell('F4').alignment = { horizontal: 'center', vertical: 'middle' };

    // Row 5: Customer Address & Date
    ws.getCell('A5').value = 'ที่อยู่  :';
    ws.getCell('A5').font = fontRegular;
    ws.mergeCells('B5:C5');
    ws.getCell('B5').value = data.customer?.address || '';
    ws.getCell('B5').font = fontRegular;

    ws.mergeCells('D5:E5');
    ws.getCell('D5').value = 'วันที่ Date ';
    ws.getCell('D5').font = fontRegular;
    ws.getCell('D5').alignment = { horizontal: 'left' };
    ws.getCell('F5').value = formatThaiDateShort(data.docDate);
    ws.getCell('F5').font = fontRegular;
    ws.getCell('F5').alignment = { horizontal: 'center', vertical: 'middle' };

    // Row 6: Job Title
    ws.mergeCells('D6:F6');
    ws.getCell('D6').value = data.jobTitle || '';
    ws.getCell('D6').font = fontBold;
    ws.getCell('D6').alignment = { horizontal: 'center' };

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
    ws.getCell('D8').alignment = { horizontal: 'center' };
    ws.getCell('F8').value = data.seller?.tel2 || '';
    ws.getCell('F8').font = fontRegular;
    ws.getCell('F8').alignment = { horizontal: 'center', vertical: 'middle' };

    // Row 9: Bank Account
    ws.mergeCells('D9:F9');
    ws.getCell('D9').value = data.seller?.bank || '';
    ws.getCell('D9').font = fontRegular;
    ws.getCell('D9').alignment = { horizontal: 'center' };

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
          ws.getCell(`A${r}`).alignment = { horizontal: 'center', vertical: 'middle' };
          ws.getCell(`B${r}`).value = item.description || '';
          ws.getCell(`B${r}`).alignment = { horizontal: 'left', vertical: 'middle' };
          ws.getCell(`C${r}`).value = item.unit || '';
          ws.getCell(`C${r}`).alignment = { horizontal: 'center', vertical: 'middle' };
          ws.getCell(`D${r}`).value = item.qty !== undefined && item.qty !== '' ? Number(item.qty) : '';
          ws.getCell(`D${r}`).alignment = { horizontal: 'center', vertical: 'middle' };
          ws.getCell(`E${r}`).value = item.unitPrice !== undefined && item.unitPrice !== '' ? Number(item.unitPrice) : '';
          ws.getCell(`E${r}`).numFmt = '#,##0.00';
          ws.getCell(`E${r}`).alignment = { horizontal: 'right', vertical: 'middle' };
          ws.getCell(`F${r}`).value = item.amount !== undefined && item.amount !== '' ? Number(item.amount) : '';
          ws.getCell(`F${r}`).numFmt = '#,##0.00';
          ws.getCell(`F${r}`).alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          // Sub-item or note
          ws.getCell(`B${r}`).value = item.description || '';
          ws.getCell(`B${r}`).alignment = { horizontal: 'left', vertical: 'middle' };
        }
      }

      ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => {
        const c = ws.getCell(`${col}${r}`);
        c.font = fontRegular;
        if (!c.alignment) {
          c.alignment = { vertical: 'middle' };
        }
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

    // Row 32-34: Signatures
    ws.getCell('A32').value = 'ผู้อนุมัติ :';
    ws.getCell('A32').font = fontRegular;

    ws.getCell('D32').value = '   ผู้เสนอ :';
    ws.getCell('D32').font = fontRegular;
    ws.getCell('D32').alignment = { horizontal: 'left' };

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
    ws.getCell('B34').alignment = { horizontal: 'center' };

    // =================================================================
    // 100% COMPLETE TABLE & FRAME BORDER RULES (Rows 3 to 34)
    // =================================================================
    // 1. Row 3 bottom border divider
    ['A', 'B', 'C', 'D', 'E', 'F'].forEach(c => {
      ws.getCell(`${c}3`).border = { bottom: thin };
    });
    ws.getCell('A3').border = { bottom: thin, left: thin };
    ws.getCell('F3').border = { bottom: thin, right: thin };

    // 2. Customer Box A4:C6
    ['A', 'B', 'C'].forEach(c => {
      ws.getCell(`${c}4`).border = Object.assign({}, ws.getCell(`${c}4`).border, { top: thin });
      ws.getCell(`${c}6`).border = Object.assign({}, ws.getCell(`${c}6`).border, { bottom: thin });
    });
    [4, 5, 6].forEach(r => {
      ws.getCell(`A${r}`).border = Object.assign({}, ws.getCell(`A${r}`).border, { left: thin });
      ws.getCell(`C${r}`).border = Object.assign({}, ws.getCell(`C${r}`).border, { right: thin });
    });

    // 3. Doc Info Boxes D4:F6
    [['D4', 'E4', 'F4'], ['D5', 'E5', 'F5']].forEach(row => {
      row.forEach(addr => {
        ws.getCell(addr).border = { top: thin, bottom: thin, left: thin, right: thin };
      });
    });
    // D6:F6 is merged Job Title Box
    ['D6', 'E6', 'F6'].forEach(addr => {
      ws.getCell(addr).border = { top: thin, bottom: thin, left: thin, right: thin };
    });

    // 4. Seller Box A7:C9
    ['A', 'B', 'C'].forEach(c => {
      ws.getCell(`${c}7`).border = Object.assign({}, ws.getCell(`${c}7`).border, { top: thin });
      ws.getCell(`${c}9`).border = Object.assign({}, ws.getCell(`${c}9`).border, { bottom: thin });
    });
    [7, 8, 9].forEach(r => {
      ws.getCell(`A${r}`).border = Object.assign({}, ws.getCell(`A${r}`).border, { left: thin });
      ws.getCell(`C${r}`).border = Object.assign({}, ws.getCell(`C${r}`).border, { right: thin });
    });

    // 5. Seller Tel/Bank Boxes
    ws.getCell('D7').border = { top: thin, left: thin };
    ws.getCell('E7').border = { top: thin };
    ws.getCell('F7').border = { top: thin, bottom: thin, left: thin, right: thin };

    ws.getCell('D8').border = { top: thin, left: thin };
    ws.getCell('E8').border = { top: thin };
    ws.getCell('F8').border = { top: thin, bottom: thin, left: thin, right: thin };

    // D9:F9 is merged Bank Account Box
    ['D9', 'E9', 'F9'].forEach(addr => {
      ws.getCell(addr).border = { top: thin, bottom: thin, left: thin, right: thin };
    });

    // 6. Table Header Rows 10 & 11
    ws.getCell('A10').border = { top: thin, left: thin, right: thin };
    ws.getCell('B10').border = { top: thin, left: thin };
    ws.getCell('A11').border = { bottom: thin, left: thin, right: thin };
    ws.getCell('B11').border = { bottom: thin, left: thin };

    // Header merged columns C, D, E, F: Set top, bottom, left, right on merged master
    ['C', 'D', 'E', 'F'].forEach(col => {
      ws.getCell(`${col}10`).border = { top: thin, bottom: thin, left: thin, right: thin };
      ws.getCell(`${col}11`).border = { top: thin, bottom: thin, left: thin, right: thin };
    });

    // 7. Data Rows 12 to 28
    for (let r = startRow; r <= endRow; r++) {
      ws.getCell(`A${r}`).border = (r === 12 || r === 26 || r === 27 || r === 28)
        ? { left: thin, right: thin }
        : { left: thin };
      ws.getCell(`B${r}`).border = { left: thin };
      ['C', 'D', 'E', 'F'].forEach(c => {
        ws.getCell(`${c}${r}`).border = { left: thin, right: thin };
      });
      if (r === endRow) {
        ['A', 'C', 'D', 'E', 'F'].forEach(c => {
          ws.getCell(`${c}${endRow}`).border = { left: thin, right: thin, bottom: thin };
        });
        ws.getCell(`B${endRow}`).border = { left: thin, bottom: thin };
      }
    }

    // 8. Row 29 Total Summary
    ws.getCell('A29').border = { bottom: thin, left: thin, right: thin };
    ws.getCell('B29').border = { bottom: thin, left: thin, right: thin };
    ws.getCell('C29').border = { bottom: thin, left: thin, right: thin };
    ws.getCell('D29').border = { bottom: thin, left: thin, right: thin };
    ws.getCell('E29').border = { bottom: thin, left: thin, right: thin };
    ws.getCell('F29').border = { bottom: thin, left: thin, right: thin };

    // 9. Rows 30 to 34 Signatures Table Frame
    ws.getCell('A30').border = { top: thin, left: thin, right: thin };
    ws.getCell('B30').border = { top: thin, left: thin, right: thin };
    ws.getCell('C30').border = { top: thin, left: thin, right: thin };
    ws.getCell('F30').border = { right: thin };

    ws.getCell('A31').border = { left: thin };
    ws.getCell('C31').border = { right: thin };
    ws.getCell('D31').border = { left: thin };
    ws.getCell('E31').border = { right: thin };
    ws.getCell('F31').border = { right: thin };

    ws.getCell('A32').border = { left: thin };
    ws.getCell('C32').border = { right: thin };
    ws.getCell('D32').border = { left: thin };
    ws.getCell('E32').border = { bottom: hair, right: thin };
    ws.getCell('F32').border = { bottom: hair, right: thin };

    ws.getCell('A33').border = { left: thin };
    ws.getCell('B33').border = { top: hair, bottom: hair };
    ws.getCell('C33').border = { right: thin };
    ws.getCell('D33').border = { left: thin };
    ws.getCell('E33').border = { right: thin };
    ws.getCell('F33').border = { right: thin };

    ws.getCell('A34').border = { bottom: thin, left: thin };
    ws.getCell('B34').border = { bottom: thin };
    ws.getCell('C34').border = { bottom: thin };
    ws.getCell('D34').border = { bottom: thin, left: thin };
    ws.getCell('E34').border = { bottom: thin };
    ws.getCell('F34').border = { bottom: thin, right: thin };

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
