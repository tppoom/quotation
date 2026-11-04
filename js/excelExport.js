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
        printArea: 'A1:F33',
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

    // Row heights matching reference template (Row 3 spacer deleted, all shifted by -1)
    ws.getRow(1).height = 30.75;
    ws.getRow(2).height = 23.25;
    ws.getRow(3).height = 25.5;
    ws.getRow(4).height = 21.75;
    ws.getRow(5).height = 26.25;
    ws.getRow(9).height = 24;
    ws.getRow(10).height = 18.75;
    ws.getRow(11).height = 24;
    ws.getRow(16).height = 25.5;
    ws.getRow(18).height = 21.75;
    ws.getRow(20).height = 21.75;
    ws.getRow(21).height = 21.75;
    ws.getRow(22).height = 21.75;
    ws.getRow(23).height = 21.75;
    ws.getRow(28).height = 21.95;
    ws.getRow(29).height = 20.25;
    ws.getRow(30).height = 11.25;
    ws.getRow(31).height = 24;
    ws.getRow(32).height = 21.95;
    ws.getRow(33).height = 24;

    const fontHeader1 = { name: 'Sarabun', size: 16, bold: true };
    const fontHeader2 = { name: 'Sarabun', size: 12, bold: true };
    const fontRegular = { name: 'Sarabun', size: 10 };
    const fontBold = { name: 'Sarabun', size: 10, bold: true };
    const fontTh = { name: 'Sarabun', size: 9.5, bold: true };

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

    // Row 3: Attention & No (Row 3 spacer deleted)
    ws.mergeCells('A3:C3');
    ws.getCell('A3').value = `เรียน/Attention : ${data.customer?.name || ''}`;
    ws.getCell('A3').font = fontBold;
    ws.getCell('A3').alignment = { horizontal: 'left', vertical: 'middle', shrinkToFit: true };

    ws.mergeCells('D3:E3');
    ws.getCell('D3').value = 'เลขที่/ No.';
    ws.getCell('D3').font = fontRegular;
    ws.getCell('D3').alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getCell('F3').value = data.docNo || '';
    ws.getCell('F3').font = { name: 'Angsana New', size: 10, bold: true };
    ws.getCell('F3').alignment = { horizontal: 'right', vertical: 'middle', shrinkToFit: true };

    // Row 4: Customer Address & Date
    ws.getCell('A4').value = 'ที่อยู่  :';
    ws.getCell('A4').font = fontRegular;
    ws.getCell('A4').alignment = { horizontal: 'left', vertical: 'middle' };
    ws.mergeCells('B4:C4');
    ws.getCell('B4').value = data.customer?.address || '';
    ws.getCell('B4').font = fontRegular;
    ws.getCell('B4').alignment = { horizontal: 'left', vertical: 'middle', shrinkToFit: true };

    ws.mergeCells('D4:E4');
    ws.getCell('D4').value = 'วันที่ Date ';
    ws.getCell('D4').font = fontRegular;
    ws.getCell('D4').alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getCell('F4').value = formatThaiDateShort(data.docDate);
    ws.getCell('F4').font = fontRegular;
    ws.getCell('F4').alignment = { horizontal: 'center', vertical: 'middle' };

    // Row 5: Job Title
    ws.mergeCells('D5:F5');
    ws.getCell('D5').value = data.jobTitle || '';
    ws.getCell('D5').font = { name: 'Sarabun', size: 10.5, bold: true };
    ws.getCell('D5').alignment = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };

    // Row 6: Seller Name
    ws.mergeCells('A6:C6');
    ws.getCell('A6').value = `ชื่อ : ${data.seller?.name || ''}`;
    ws.getCell('A6').font = fontRegular;
    ws.getCell('A6').alignment = { horizontal: 'left', vertical: 'middle', shrinkToFit: true };

    // Row 7: Seller Address & Phone
    ws.mergeCells('A7:C7');
    ws.getCell('A7').value = `ที่อยู่ : ${data.seller?.address || ''}`;
    ws.getCell('A7').font = fontRegular;
    ws.getCell('A7').alignment = { horizontal: 'left', vertical: 'middle', shrinkToFit: true };

    ws.mergeCells('D7:E7');
    ws.getCell('D7').value = data.seller?.tel1 ? `TEL. ${data.seller.tel1}` : '';
    ws.getCell('D7').font = fontRegular;
    ws.getCell('D7').alignment = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };
    ws.getCell('F7').value = data.seller?.tel2 || '';
    ws.getCell('F7').font = fontRegular;
    ws.getCell('F7').alignment = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };

    // Row 8: Bank Account
    ws.mergeCells('D8:F8');
    ws.getCell('D8').value = data.seller?.bank || '';
    ws.getCell('D8').font = fontRegular;
    ws.getCell('D8').alignment = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };

    // Table Header: Row 9 & 10
    ws.getCell('A9').value = 'ลำดับที่';
    ws.getCell('A10').value = 'ITEM';
    ws.getCell('B9').value = 'รายการ';
    ws.getCell('B10').value = 'DESCRIPTION';

    ws.getCell('C9').value = 'หน่วย';
    ws.getCell('D9').value = 'จำนวน';
    ws.getCell('E9').value = 'ราคา/หน่วย';
    ws.getCell('F9').value = 'จำนวนเงิน';

    const thCells = ['A9', 'B9', 'C9', 'D9', 'E9', 'F9', 'A10', 'B10'];
    thCells.forEach(ref => {
      const c = ws.getCell(ref);
      c.font = fontTh;
      c.alignment = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };
    });
    ws.getCell('E9').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true, shrinkToFit: true };

    // Merge C9:C10, D9:D10, E9:E10, F9:F10 with independent top and bottom borders
    ['C', 'D', 'E', 'F'].forEach(col => {
      ws.getCell(`${col}9`).border = { top: thin, left: thin, right: thin };
      ws.getCell(`${col}9`).alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getCell(`${col}10`).border = { bottom: thin, left: thin, right: thin };
      ws.mergeCellsWithoutStyle(`${col}9:${col}10`);
    });

    // Populate Items: Rows 11 to endRow (at least 17 data rows to match A4 single page standard)
    const startRow = 11;
    const items = data.items || [];
    const dataRowCount = Math.max(17, items.length);
    const endRow = startRow + dataRowCount - 1;

    for (let r = startRow; r <= endRow; r++) {
      ws.getRow(r).height = 24;
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

    // Dynamic Total Summary Row
    const totalRow = endRow + 1;
    ws.getRow(totalRow).height = 21.95;
    ws.mergeCells(`A${totalRow}:C${totalRow}`);
    ws.getCell(`A${totalRow}`).value = data.totalTextThai || '';
    ws.getCell(`A${totalRow}`).font = { name: 'Angsana New', size: 14, bold: true };
    ws.getCell(`A${totalRow}`).alignment = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };

    ws.mergeCells(`D${totalRow}:E${totalRow}`);
    ws.getCell(`D${totalRow}`).value = 'รวมเงิน';
    ws.getCell(`D${totalRow}`).font = { name: 'Angsana New', size: 14, bold: true };
    ws.getCell(`D${totalRow}`).alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getCell(`F${totalRow}`).value = { formula: `SUM(F${startRow}:F${endRow})`, result: data.totalAmount || 0 };
    ws.getCell(`F${totalRow}`).font = { name: 'Angsana New', size: 14, bold: true };
    ws.getCell(`F${totalRow}`).numFmt = '#,##0.00';
    ws.getCell(`F${totalRow}`).alignment = { horizontal: 'right', vertical: 'middle' };

    // Dynamic Signature Rows
    const sigSpacerRow = totalRow + 1;
    const sigEmptyRow = totalRow + 2;
    const sigHeaderRow = totalRow + 3;
    const sigPosRow = totalRow + 4;
    const sigParenRow = totalRow + 5;

    ws.getRow(sigSpacerRow).height = 20.25;
    ws.getRow(sigEmptyRow).height = 11.25;
    ws.getRow(sigHeaderRow).height = 24;
    ws.getRow(sigPosRow).height = 21.95;
    ws.getRow(sigParenRow).height = 24;

    ws.getCell(`A${sigHeaderRow}`).value = 'ผู้อนุมัติ :';
    ws.getCell(`A${sigHeaderRow}`).font = { name: 'Angsana New', size: 14, bold: true };
    ws.getCell(`A${sigHeaderRow}`).alignment = { horizontal: 'left', vertical: 'middle' };

    ws.getCell(`D${sigHeaderRow}`).value = '   ผู้เสนอ :';
    ws.getCell(`D${sigHeaderRow}`).font = { name: 'Angsana New', size: 12, bold: true };
    ws.getCell(`D${sigHeaderRow}`).alignment = { horizontal: 'left', vertical: 'middle', shrinkToFit: true };

    ws.mergeCells(`E${sigHeaderRow}:F${sigHeaderRow}`);
    ws.getCell(`E${sigHeaderRow}`).value = data.seller?.name || '';
    const sellerNameLen = (data.seller?.name || '').length;
    ws.getCell(`E${sigHeaderRow}`).font = { name: 'Angsana New', size: sellerNameLen > 25 ? 10 : sellerNameLen > 18 ? 11 : 13, bold: true };
    ws.getCell(`E${sigHeaderRow}`).alignment = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };

    ws.getCell(`A${sigPosRow}`).value = 'ตำแหน่ง ';
    ws.getCell(`A${sigPosRow}`).font = { name: 'Angsana New', size: 14, bold: true };
    ws.getCell(`A${sigPosRow}`).alignment = { horizontal: 'left', vertical: 'middle' };

    ws.mergeCells(`E${sigPosRow}:F${sigPosRow}`);
    ws.getCell(`E${sigPosRow}`).value = formatThaiDateLong(data.docDate);
    ws.getCell(`E${sigPosRow}`).font = { name: 'Angsana New', size: 13 };
    ws.getCell(`E${sigPosRow}`).alignment = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };

    ws.getCell(`B${sigParenRow}`).value = '   (                                                                                )';
    ws.getCell(`B${sigParenRow}`).font = { name: 'CordiaUPC', size: 14 };
    ws.getCell(`B${sigParenRow}`).alignment = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };

    // =================================================================
    // 100% COMPLETE TABLE & FRAME BORDER RULES (Rows 3 to sigParenRow)
    // =================================================================
    // 1. Top border divider under QUOTATION (Row 3 top)
    ['A', 'B', 'C', 'D', 'E', 'F'].forEach(c => {
      ws.getCell(`${c}3`).border = { top: thin };
    });
    ws.getCell('A3').border = { top: thin, left: thin };
    ws.getCell('F3').border = { top: thin, right: thin };

    // 2. Customer Box A3:C5
    ['A', 'B', 'C'].forEach(c => {
      ws.getCell(`${c}3`).border = Object.assign({}, ws.getCell(`${c}3`).border, { top: thin });
      ws.getCell(`${c}5`).border = Object.assign({}, ws.getCell(`${c}5`).border, { bottom: thin });
    });
    [3, 4, 5].forEach(r => {
      ws.getCell(`A${r}`).border = Object.assign({}, ws.getCell(`A${r}`).border, { left: thin });
      ws.getCell(`C${r}`).border = Object.assign({}, ws.getCell(`C${r}`).border, { right: thin });
    });

    // 3. Doc Info Boxes D3:F5
    [['D3', 'E3', 'F3'], ['D4', 'E4', 'F4']].forEach(row => {
      row.forEach(addr => {
        ws.getCell(addr).border = { top: thin, bottom: thin, left: thin, right: thin };
      });
    });
    // D5:F5 is merged Job Title Box
    ['D5', 'E5', 'F5'].forEach(addr => {
      ws.getCell(addr).border = { top: thin, bottom: thin, left: thin, right: thin };
    });

    // 4. Seller Box A6:C8
    ['A', 'B', 'C'].forEach(c => {
      ws.getCell(`${c}6`).border = Object.assign({}, ws.getCell(`${c}6`).border, { top: thin });
      ws.getCell(`${c}8`).border = Object.assign({}, ws.getCell(`${c}8`).border, { bottom: thin });
    });
    [6, 7, 8].forEach(r => {
      ws.getCell(`A${r}`).border = Object.assign({}, ws.getCell(`A${r}`).border, { left: thin });
      ws.getCell(`C${r}`).border = Object.assign({}, ws.getCell(`C${r}`).border, { right: thin });
    });

    // 5. Seller Tel/Bank Boxes (D6:F8)
    ws.getCell('D6').border = { top: thin, left: thin };
    ws.getCell('E6').border = { top: thin };
    ws.getCell('F6').border = { top: thin, bottom: thin, left: thin, right: thin };

    ws.getCell('D7').border = { top: thin, left: thin };
    ws.getCell('E7').border = { top: thin };
    ws.getCell('F7').border = { top: thin, bottom: thin, left: thin, right: thin };

    // D8:F8 is merged Bank Account Box with continuous bottom border
    ws.getCell('D8').border = { top: thin, bottom: thin, left: thin };
    ws.getCell('E8').border = { top: thin, bottom: thin };
    ws.getCell('F8').border = { top: thin, bottom: thin, right: thin };

    // 6. Table Header Rows 9 & 10
    ws.getCell('A9').border = { top: thin, left: thin, right: thin };
    ws.getCell('B9').border = { top: thin, left: thin };
    ws.getCell('A10').border = { bottom: thin, left: thin, right: thin };
    ws.getCell('B10').border = { bottom: thin, left: thin };

    // Header merged columns C, D, E, F: Set top on row 9, bottom on row 10
    ['C', 'D', 'E', 'F'].forEach(col => {
      ws.getCell(`${col}9`).border = { top: thin, left: thin, right: thin };
      ws.getCell(`${col}10`).border = { bottom: thin, left: thin, right: thin };
    });

    // 7. Data Rows 12 to endRow
    for (let r = startRow; r <= endRow; r++) {
      ws.getCell(`A${r}`).border = { left: thin, right: thin };
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

    // 8. Total Summary Row
    ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => {
      ws.getCell(`${col}${totalRow}`).border = { top: thin, bottom: thin, left: thin, right: thin };
    });

    // 9. Signatures Table Frame
    ws.mergeCellsWithoutStyle(`A${sigSpacerRow}:C${sigSpacerRow}`);
    ws.getCell(`A${sigSpacerRow}`).border = { top: thin, left: thin, right: thin };
    ws.getCell(`B${sigSpacerRow}`).border = { top: thin };
    ws.getCell(`C${sigSpacerRow}`).border = { top: thin, right: thin };
    ws.getCell(`D${sigSpacerRow}`).border = { top: thin, left: thin };
    ws.getCell(`E${sigSpacerRow}`).border = { top: thin };
    ws.getCell(`F${sigSpacerRow}`).border = { top: thin, right: thin };

    ws.getCell(`A${sigEmptyRow}`).border = { left: thin };
    ws.getCell(`C${sigEmptyRow}`).border = { right: thin };
    ws.getCell(`D${sigEmptyRow}`).border = { left: thin };
    ws.getCell(`F${sigEmptyRow}`).border = { right: thin };

    ws.getCell(`A${sigHeaderRow}`).border = { left: thin };
    ws.getCell(`B${sigHeaderRow}`).border = { bottom: hair };
    ws.getCell(`C${sigHeaderRow}`).border = { bottom: hair, right: thin };
    ws.getCell(`D${sigHeaderRow}`).border = { left: thin };
    ws.getCell(`E${sigHeaderRow}`).border = { bottom: hair };
    ws.getCell(`F${sigHeaderRow}`).border = { bottom: hair, right: thin };

    ws.getCell(`A${sigPosRow}`).border = { left: thin };
    ws.getCell(`B${sigPosRow}`).border = { top: hair, bottom: hair };
    ws.getCell(`C${sigPosRow}`).border = { right: thin };
    ws.getCell(`D${sigPosRow}`).border = { left: thin };
    ws.getCell(`E${sigPosRow}`).border = {};
    ws.getCell(`F${sigPosRow}`).border = { right: thin };

    ws.getCell(`A${sigParenRow}`).border = { bottom: thin, left: thin };
    ws.getCell(`B${sigParenRow}`).border = { bottom: thin };
    ws.getCell(`C${sigParenRow}`).border = { bottom: thin, right: thin };
    ws.getCell(`D${sigParenRow}`).border = { bottom: thin, left: thin };
    ws.getCell(`E${sigParenRow}`).border = { bottom: thin };
    ws.getCell(`F${sigParenRow}`).border = { bottom: thin, right: thin };

    // Dynamic Page Setup Print Area
    ws.pageSetup.printArea = `A1:F${sigParenRow}`;
    if (items.length > 17) {
      ws.pageSetup.fitToHeight = 0;
    }

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
