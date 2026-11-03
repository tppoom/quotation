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

    const origTransform = (el.style && el.style.transform) || '';
    const origPosition = (el.style && el.style.position) || '';
    if (el.style) {
      el.style.transform = 'none';
      el.style.position = 'relative';
    }

    let canvas;
    try {
      canvas = await h2c(el, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false
      });
    } finally {
      if (el.style) {
        el.style.transform = origTransform;
        el.style.position = origPosition;
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
    const origTransform = (el.style && el.style.transform) || '';
    const origPosition = (el.style && el.style.position) || '';
    if (el.style) {
      el.style.transform = 'none';
      el.style.position = 'relative';
    }

    let canvas;
    try {
      canvas = await h2c(el, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false
      });
    } finally {
      if (el.style) {
        el.style.transform = origTransform;
        el.style.position = origPosition;
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

    pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
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
