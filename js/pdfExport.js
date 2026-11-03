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

    if (pdfHeight > (pageHeight + 2) && typeof pdf.addPage === 'function') {
      let heightLeft = pdfHeight;
      let position = 0;

      pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight, undefined, 'FAST');
      heightLeft -= pageHeight;

      while (heightLeft > 0) {
        position -= pageHeight;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight, undefined, 'FAST');
        heightLeft -= pageHeight;
      }
    } else {
      pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight, undefined, 'FAST');
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
