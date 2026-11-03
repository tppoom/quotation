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
    const ExcelClass = ExcelJSModule || (typeof ExcelJS !== 'undefined' ? ExcelJS : (typeof require !== 'undefined' ? require('exceljs') : null));
    if (!ExcelClass) {
      throw new Error('ExcelJS library is not available');
    }

    const wb = new ExcelClass.Workbook();
    wb.creator = 'Quotation Generator';
    wb.created = new Date();

    const ws = wb.addWorksheet('Sheet1', {
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1 }
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
    ws.getCell('A4').alignment = { horizontal: 'left', vertical: 'middle' };

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
