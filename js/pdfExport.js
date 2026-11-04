// js/pdfExport.js
(function(root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DocExporter = factory();
  }
}(typeof self !== 'undefined' ? self : this, function() {

  async function exportToPNG(previewElementId, filename) {
    const doc = typeof document !== 'undefined' ? document : null;
    if (!doc) throw new Error('Preview element not found');

    const el = doc.getElementById(previewElementId);
    if (!el) throw new Error('Preview element not found');

    const h2c = typeof html2canvas !== 'undefined' ? html2canvas : (typeof window !== 'undefined' ? window.html2canvas : undefined);
    if (!h2c) {
      throw new Error('html2canvas library is not loaded');
    }

    // Save original styles of element and parents to avoid any clipping during capture
    const parentScaler = el.parentElement;
    const outerContainer = parentScaler ? parentScaler.parentElement : null;

    const origElTransform = el.style ? el.style.transform : '';
    const origElPosition = el.style ? el.style.position : '';
    const origParentWidth = parentScaler && parentScaler.style ? parentScaler.style.width : '';
    const origParentHeight = parentScaler && parentScaler.style ? parentScaler.style.height : '';
    const origParentOverflow = parentScaler && parentScaler.style ? parentScaler.style.overflow : '';
    const origOuterOverflow = outerContainer && outerContainer.style ? outerContainer.style.overflow : '';

    let canvas;
    try {
      // Temporarily expand to natural 794px x 1123px A4 size without clipping
      if (el.style) {
        el.style.transform = 'none';
        el.style.position = 'relative';
      }
      if (parentScaler && parentScaler.style) {
        parentScaler.style.width = '794px';
        parentScaler.style.height = (el.offsetHeight ? el.offsetHeight + 'px' : '1123px');
        parentScaler.style.overflow = 'visible';
      }
      if (outerContainer && outerContainer.style) {
        outerContainer.style.overflow = 'visible';
      }

      // Ensure all web fonts are fully ready before capture
      if (doc.fonts && doc.fonts.ready) {
        await doc.fonts.ready;
      }

      canvas = await h2c(el, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false
      });
    } finally {
      // Restore all original styles
      if (el.style) {
        el.style.transform = origElTransform;
        el.style.position = origElPosition;
      }
      if (parentScaler && parentScaler.style) {
        parentScaler.style.width = origParentWidth;
        parentScaler.style.height = origParentHeight;
        parentScaler.style.overflow = origParentOverflow;
      }
      if (outerContainer && outerContainer.style) {
        outerContainer.style.overflow = origOuterOverflow;
      }
    }

    const dataUrl = canvas.toDataURL('image/png');
    const link = doc.createElement('a');
    const safeFilename = filename || 'quotation';
    link.download = safeFilename.endsWith('.png') ? safeFilename : `${safeFilename}.png`;
    link.href = dataUrl;
    doc.body.appendChild(link);
    link.click();
    doc.body.removeChild(link);
  }

  async function exportToPDF(previewElementId, filename) {
    const doc = typeof document !== 'undefined' ? document : null;
    if (!doc) throw new Error('Preview element not found');

    const el = doc.getElementById(previewElementId);
    if (!el) throw new Error('Preview element not found');

    const h2c = typeof html2canvas !== 'undefined' ? html2canvas : (typeof window !== 'undefined' ? window.html2canvas : undefined);
    const jspdfLib = typeof jspdf !== 'undefined' ? jspdf : (typeof window !== 'undefined' ? window.jspdf : undefined);

    if (!h2c || !jspdfLib) {
      if (typeof window !== 'undefined' && typeof window.print === 'function') {
        window.print();
      }
      return;
    }

    const jsPDFClass = jspdfLib.jsPDF || jspdfLib;

    // Save original styles of element and parents to avoid any clipping during capture
    const parentScaler = el.parentElement;
    const outerContainer = parentScaler ? parentScaler.parentElement : null;

    const origElTransform = el.style ? el.style.transform : '';
    const origElPosition = el.style ? el.style.position : '';
    const origParentWidth = parentScaler && parentScaler.style ? parentScaler.style.width : '';
    const origParentHeight = parentScaler && parentScaler.style ? parentScaler.style.height : '';
    const origParentOverflow = parentScaler && parentScaler.style ? parentScaler.style.overflow : '';
    const origOuterOverflow = outerContainer && outerContainer.style ? outerContainer.style.overflow : '';

    let canvas;
    try {
      // Temporarily expand to natural 794px x 1123px A4 size without clipping
      if (el.style) {
        el.style.transform = 'none';
        el.style.position = 'relative';
      }
      if (parentScaler && parentScaler.style) {
        parentScaler.style.width = '794px';
        parentScaler.style.height = (el.offsetHeight ? el.offsetHeight + 'px' : '1123px');
        parentScaler.style.overflow = 'visible';
      }
      if (outerContainer && outerContainer.style) {
        outerContainer.style.overflow = 'visible';
      }

      // Ensure all web fonts are fully ready before capture
      if (doc.fonts && doc.fonts.ready) {
        await doc.fonts.ready;
      }

      canvas = await h2c(el, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false
      });
    } finally {
      // Restore all original styles
      if (el.style) {
        el.style.transform = origElTransform;
        el.style.position = origElPosition;
      }
      if (parentScaler && parentScaler.style) {
        parentScaler.style.width = origParentWidth;
        parentScaler.style.height = origParentHeight;
        parentScaler.style.overflow = origParentOverflow;
      }
      if (outerContainer && outerContainer.style) {
        outerContainer.style.overflow = origOuterOverflow;
      }
    }

    const imgData = canvas.toDataURL('image/png');
    const pdf = new jsPDFClass({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4'
    });

    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
    const pageHeight = pdf.internal.pageSize.getHeight ? pdf.internal.pageSize.getHeight() : 297;

    // If within 15% of 1 page, fit cleanly onto 1 A4 page (matching Excel fitToPage behavior)
    if (typeof pdf.addPage === 'function' && pdfHeight > pageHeight && pdfHeight <= (pageHeight * 1.15)) {
      const scale = pageHeight / pdfHeight;
      const scaledWidth = pdfWidth * scale;
      const offsetX = (pdfWidth - scaledWidth) / 2;
      pdf.addImage(imgData, 'PNG', offsetX, 0, scaledWidth, pageHeight, undefined, 'FAST');
    } else if (pdfHeight <= pageHeight || typeof pdf.addPage !== 'function') {
      pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight, undefined, 'FAST');
    } else {
      // Multi-page document: Smart row boundary splitting
      const tableRows = el.querySelectorAll ? el.querySelectorAll('tbody tr') : [];
      const elOffsetHeight = el.offsetHeight || 1123;
      const targetPage1Pixels = (pageHeight - 20) * (el.offsetWidth || 794) / pdfWidth;
      let splitOffset = targetPage1Pixels;

      for (let i = 0; i < tableRows.length; i++) {
        const row = tableRows[i];
        const rowBottom = (row.offsetTop || 0) + (row.offsetHeight || 0);
        if (rowBottom <= targetPage1Pixels) {
          splitOffset = rowBottom;
        } else {
          break;
        }
      }

      const canvasScale = canvas.height / elOffsetHeight;
      const splitCanvasY = Math.round(splitOffset * canvasScale);

      // Create Page 1 Canvas
      const page1Canvas = doc.createElement('canvas');
      page1Canvas.width = canvas.width;
      page1Canvas.height = splitCanvasY;
      const ctx1 = page1Canvas.getContext('2d');
      if (ctx1) {
        ctx1.fillStyle = '#ffffff';
        ctx1.fillRect(0, 0, page1Canvas.width, page1Canvas.height);
        ctx1.drawImage(canvas, 0, 0, canvas.width, splitCanvasY, 0, 0, canvas.width, splitCanvasY);
      }

      const page1Img = page1Canvas.toDataURL('image/png');
      const page1HeightMm = (splitCanvasY * pdfWidth) / canvas.width;
      pdf.addImage(page1Img, 'PNG', 0, 0, pdfWidth, page1HeightMm, undefined, 'FAST');

      // Create Page 2 Canvas
      pdf.addPage();
      const page2Canvas = doc.createElement('canvas');
      page2Canvas.width = canvas.width;
      page2Canvas.height = canvas.height - splitCanvasY;
      const ctx2 = page2Canvas.getContext('2d');
      if (ctx2) {
        ctx2.fillStyle = '#ffffff';
        ctx2.fillRect(0, 0, page2Canvas.width, page2Canvas.height);
        ctx2.drawImage(canvas, 0, splitCanvasY, canvas.width, canvas.height - splitCanvasY, 0, 0, canvas.width, page2Canvas.height);
      }

      const page2Img = page2Canvas.toDataURL('image/png');
      const page2HeightMm = ((canvas.height - splitCanvasY) * pdfWidth) / canvas.width;
      pdf.addImage(page2Img, 'PNG', 0, 0, pdfWidth, Math.min(pageHeight, page2HeightMm), undefined, 'FAST');
    }
    const safeFilename = filename || 'quotation';
    pdf.save(safeFilename.endsWith('.pdf') ? safeFilename : `${safeFilename}.pdf`);
  }

  function printDocument() {
    if (typeof window !== 'undefined' && typeof window.print === 'function') {
      window.print();
    }
  }

  return {
    exportToPNG,
    exportToPDF,
    printDocument
  };
}));
