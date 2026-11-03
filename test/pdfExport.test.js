const assert = require('assert');
const fs = require('fs');

const DocExporter = require('../js/pdfExport.js');
const { exportToPNG, exportToPDF, printDocument } = DocExporter;

async function runTests() {
  console.log('Testing module exports and function signatures...');
  assert.strictEqual(typeof DocExporter, 'object', 'DocExporter should be an object');
  assert.strictEqual(typeof exportToPNG, 'function', 'exportToPNG should be a function');
  assert.strictEqual(typeof exportToPDF, 'function', 'exportToPDF should be a function');
  assert.strictEqual(typeof printDocument, 'function', 'printDocument should be a function');

  // Helper to setup mock DOM environment
  function setupMockDOM() {
    const appendedChildren = [];
    const removedChildren = [];
    let clickCount = 0;

    const mockElement = {
      id: 'preview-container',
      tagName: 'DIV'
    };

    const mockBody = {
      appendChild: (el) => {
        appendedChildren.push(el);
      },
      removeChild: (el) => {
        removedChildren.push(el);
      }
    };

    const mockDocument = {
      getElementById: (id) => {
        if (id === 'preview-container') return mockElement;
        return null;
      },
      createElement: (tag) => {
        if (tag === 'a') {
          return {
            tagName: 'A',
            download: '',
            href: '',
            click: () => {
              clickCount++;
            }
          };
        }
        return { tagName: tag.toUpperCase() };
      },
      body: mockBody
    };

    return {
      mockElement,
      mockDocument,
      mockBody,
      getAppendedChildren: () => appendedChildren,
      getRemovedChildren: () => removedChildren,
      getClickCount: () => clickCount
    };
  }

  // =========================================================================
  // 1. Error when element not found
  // =========================================================================
  console.log('Testing error when element is not found...');
  {
    const dom = setupMockDOM();
    global.document = dom.mockDocument;

    await assert.rejects(
      async () => {
        await exportToPNG('non-existent-id', 'test');
      },
      {
        name: 'Error',
        message: 'Preview element not found'
      },
      'exportToPNG should reject if preview element is not found'
    );

    await assert.rejects(
      async () => {
        await exportToPDF('non-existent-id', 'test');
      },
      {
        name: 'Error',
        message: 'Preview element not found'
      },
      'exportToPDF should reject if preview element is not found'
    );
  }

  // =========================================================================
  // 2. exportToPNG: Missing html2canvas library
  // =========================================================================
  console.log('Testing exportToPNG when html2canvas is not loaded...');
  {
    const dom = setupMockDOM();
    global.document = dom.mockDocument;
    delete global.html2canvas;
    delete global.window;

    await assert.rejects(
      async () => {
        await exportToPNG('preview-container', 'test');
      },
      {
        name: 'Error',
        message: 'html2canvas library is not loaded'
      },
      'exportToPNG should reject if html2canvas is undefined'
    );
  }

  // =========================================================================
  // 3. exportToPNG: Successful capture and download
  // =========================================================================
  console.log('Testing exportToPNG successful capture and download...');
  {
    const dom = setupMockDOM();
    global.document = dom.mockDocument;

    let capturedElement = null;
    let capturedOptions = null;
    const mockCanvas = {
      toDataURL: (type) => {
        assert.strictEqual(type, 'image/png');
        return 'data:image/png;base64,mockPngBase64Data';
      }
    };

    global.html2canvas = async (el, opts) => {
      capturedElement = el;
      capturedOptions = opts;
      return mockCanvas;
    };

    // Test 3a: filename without .png extension
    await exportToPNG('preview-container', 'my-quotation');
    assert.strictEqual(capturedElement, dom.mockElement);
    assert.deepStrictEqual(capturedOptions, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false
    });

    const appended = dom.getAppendedChildren();
    const removed = dom.getRemovedChildren();
    assert.strictEqual(appended.length, 1);
    assert.strictEqual(removed.length, 1);
    assert.strictEqual(appended[0], removed[0]);
    assert.strictEqual(appended[0].download, 'my-quotation.png');
    assert.strictEqual(appended[0].href, 'data:image/png;base64,mockPngBase64Data');
    assert.strictEqual(dom.getClickCount(), 1);

    // Test 3b: filename already ending with .png
    const dom2 = setupMockDOM();
    global.document = dom2.mockDocument;
    await exportToPNG('preview-container', 'already-has-ext.png');
    const appended2 = dom2.getAppendedChildren();
    assert.strictEqual(appended2[0].download, 'already-has-ext.png', 'Should not append duplicate .png');

    // Test 3c: empty or null filename defaults to quotation.png
    const dom3 = setupMockDOM();
    global.document = dom3.mockDocument;
    await exportToPNG('preview-container', '');
    const appended3 = dom3.getAppendedChildren();
    assert.strictEqual(appended3[0].download, 'quotation.png', 'Empty filename should default to quotation.png');
  }

  // =========================================================================
  // 4. exportToPDF: Fallback handling when html2canvas or jsPDF is missing
  // =========================================================================
  console.log('Testing exportToPDF fallback to window.print()...');
  {
    const dom = setupMockDOM();
    global.document = dom.mockDocument;

    let printCalledCount = 0;
    global.window = {
      print: () => {
        printCalledCount++;
      }
    };

    // Case 4a: html2canvas missing, jspdf present
    delete global.html2canvas;
    global.jspdf = { jsPDF: class {} };
    await exportToPDF('preview-container', 'quote');
    assert.strictEqual(printCalledCount, 1, 'Should call window.print when html2canvas is missing');

    // Case 4b: html2canvas present, jspdf missing
    global.html2canvas = async () => ({});
    delete global.jspdf;
    await exportToPDF('preview-container', 'quote');
    assert.strictEqual(printCalledCount, 2, 'Should call window.print when jspdf is missing');

    // Case 4c: both missing
    delete global.html2canvas;
    delete global.jspdf;
    await exportToPDF('preview-container', 'quote');
    assert.strictEqual(printCalledCount, 3, 'Should call window.print when both libraries are missing');
  }

  // =========================================================================
  // 5. exportToPDF: Successful generation and save
  // =========================================================================
  console.log('Testing exportToPDF successful PDF generation...');
  {
    const dom = setupMockDOM();
    global.document = dom.mockDocument;

    let printCalled = false;
    global.window = {
      print: () => {
        printCalled = true;
      }
    };

    let html2canvasOptions = null;
    const mockCanvas = {
      width: 800,
      height: 1200,
      toDataURL: (type) => {
        assert.strictEqual(type, 'image/png');
        return 'data:image/png;base64,mockPdfCanvasData';
      }
    };

    global.html2canvas = async (el, opts) => {
      html2canvasOptions = opts;
      return mockCanvas;
    };

    let jsPdfOptions = null;
    let addedImageArgs = null;
    let savedFilename = null;

    class MockJsPDF {
      constructor(opts) {
        jsPdfOptions = opts;
        this.internal = {
          pageSize: {
            getWidth: () => 210, // A4 width in mm
            getHeight: () => 297 // A4 height in mm
          }
        };
      }

      addImage(...args) {
        addedImageArgs = args;
      }

      save(fname) {
        savedFilename = fname;
      }
    }

    global.jspdf = { jsPDF: MockJsPDF };

    // Test 5a: filename without .pdf
    await exportToPDF('preview-container', 'invoice-test');
    assert.strictEqual(printCalled, false, 'window.print should NOT be called on success');
    assert.deepStrictEqual(html2canvasOptions, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false
    });
    assert.deepStrictEqual(jsPdfOptions, {
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4'
    });

    // Proportional height check: (canvas.height * pdfWidth) / canvas.width
    // (1200 * 210) / 800 = 315
    const expectedHeight = (1200 * 210) / 800;
    assert.strictEqual(addedImageArgs[0], 'data:image/png;base64,mockPdfCanvasData');
    assert.strictEqual(addedImageArgs[1], 'PNG');
    assert.strictEqual(addedImageArgs[2], 0);
    assert.strictEqual(addedImageArgs[3], 0);
    assert.strictEqual(addedImageArgs[4], 210);
    assert.strictEqual(addedImageArgs[5], expectedHeight);
    assert.strictEqual(savedFilename, 'invoice-test.pdf');

    // Test 5b: filename already ending with .pdf
    await exportToPDF('preview-container', 'document.pdf');
    assert.strictEqual(savedFilename, 'document.pdf', 'Should not append duplicate .pdf');

    // Test 5c: empty filename defaults to quotation.pdf
    await exportToPDF('preview-container', '');
    assert.strictEqual(savedFilename, 'quotation.pdf', 'Empty filename should default to quotation.pdf');

    // Test 5d: jspdf passed directly as constructor rather than { jsPDF }
    global.jspdf = MockJsPDF;
    await exportToPDF('preview-container', 'direct-ctor');
    assert.strictEqual(savedFilename, 'direct-ctor.pdf');
  }

  // =========================================================================
  // 6. printDocument()
  // =========================================================================
  console.log('Testing printDocument()...');
  {
    let printCalled = false;
    global.window = {
      print: () => {
        printCalled = true;
      }
    };
    printDocument();
    assert.strictEqual(printCalled, true, 'printDocument should trigger window.print()');

    // Test when window is undefined - should not throw
    delete global.window;
    assert.doesNotThrow(() => {
      printDocument();
    }, 'printDocument should safely handle undefined window');
  }

  // =========================================================================
  // 7. Browser UMD export verification
  // =========================================================================
  console.log('Testing browser UMD export...');
  {
    const code = fs.readFileSync(require.resolve('../js/pdfExport.js'), 'utf8');
    const fakeWindow = {};
    const fn = new Function('self', 'module', 'exports', code);
    fn(fakeWindow, undefined, undefined);

    assert(fakeWindow.DocExporter, 'window.DocExporter must be defined');
    assert.strictEqual(typeof fakeWindow.DocExporter.exportToPNG, 'function');
    assert.strictEqual(typeof fakeWindow.DocExporter.exportToPDF, 'function');
    assert.strictEqual(typeof fakeWindow.DocExporter.printDocument, 'function');
  }

  // Clean up globals
  delete global.document;
  delete global.window;
  delete global.html2canvas;
  delete global.jspdf;

  console.log('✅ All pdfExport tests passed successfully!');
}

runTests().catch(err => {
  console.error('❌ Test failed:', err);
  process.exit(1);
});
