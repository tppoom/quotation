// js/app.js
(function(root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.quotationApp = factory();
  }
}(typeof self !== 'undefined' ? self : this, function() {

  function getBahtText() {
    if (typeof bahtText === 'function') return bahtText;
    if (typeof window !== 'undefined' && typeof window.bahtText === 'function') return window.bahtText;
    try {
      const mod = require('./bahtText');
      return mod.bahtText || mod;
    } catch (e) {
      return () => '(ศูนย์บาทถ้วน)';
    }
  }

  function getStorageManager() {
    if (typeof StorageManager !== 'undefined') return StorageManager;
    if (typeof window !== 'undefined' && window.StorageManager) return window.StorageManager;
    try {
      return require('./storage');
    } catch (e) {
      return {};
    }
  }

  function getExcelExporter() {
    if (typeof ExcelExporter !== 'undefined') return ExcelExporter;
    if (typeof window !== 'undefined' && window.ExcelExporter) return window.ExcelExporter;
    try {
      return require('./excelExport');
    } catch (e) {
      return {};
    }
  }

  function getDocExporter() {
    if (typeof DocExporter !== 'undefined') return DocExporter;
    if (typeof window !== 'undefined' && window.DocExporter) return window.DocExporter;
    try {
      return require('./pdfExport');
    } catch (e) {
      return {};
    }
  }

  function getLocalDateStr(date) {
    const sm = getStorageManager();
    if (sm && typeof sm.getLocalDateStr === 'function') {
      return sm.getLocalDateStr(date);
    }
    const d = date instanceof Date ? date : (date ? new Date(date) : new Date());
    if (isNaN(d.getTime())) return '';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function parseNum(val) {
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    if (val === null || val === undefined || val === '') return 0;
    const cleaned = String(val).replace(/,/g, '').trim();
    const parsed = parseFloat(cleaned);
    return isNaN(parsed) ? 0 : parsed;
  }

  function quotationApp() {
    const sm = getStorageManager();
    const seller = sm.getSellerProfile ? sm.getSellerProfile() : {};
    const today = getLocalDateStr(new Date());

    return {
      // Data state
      quotation: {
        id: '',
        docNo: '',
        docDate: today,
        jobTitle: '',
        customer: {
          name: '',
          address: ''
        },
        seller: { ...seller },
        items: [],
        totalAmount: 0,
        totalTextThai: '(ศูนย์บาทถ้วน)'
      },

      // UI state
      historyList: [],
      showHistoryModal: false,
      showSellerModal: false,
      historySearch: '',
      saveNotice: '',
      sellerProfileForm: { ...seller },
      _noticeTimeout: null,

      init() {
        const smMgr = getStorageManager();
        // 1. Try to load current draft or start fresh
        const draft = smMgr.loadCurrentDraft ? smMgr.loadCurrentDraft() : null;
        if (draft && draft.docNo) {
          this.quotation = draft;
        } else {
          this.startNewQuotation();
        }
        this.calculateTotals();
        this.refreshHistory();

        // Autosave draft on changes
        if (typeof this.$watch === 'function') {
          this.$watch('quotation', () => {
            this.calculateTotals();
            if (smMgr.saveCurrentDraft) {
              smMgr.saveCurrentDraft(this.quotation);
            }
          });
        }
      },

      startNewQuotation() {
        const currDate = getLocalDateStr(new Date());
        const smMgr = getStorageManager();
        const nextDocNo = smMgr.generateDocNo ? smMgr.generateDocNo(currDate) : `QT-${currDate.replace(/-/g, '').slice(0, 6)}-0001`;
        const sellerProfile = smMgr.getSellerProfile ? smMgr.getSellerProfile() : {};

        this.quotation = {
          id: 'qt_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
          docNo: nextDocNo,
          docDate: currDate,
          jobTitle: '',
          customer: {
            name: '',
            address: ''
          },
          seller: { ...sellerProfile },
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
        if (!Array.isArray(this.quotation.items)) {
          this.quotation.items = [];
        }
        const currentItems = this.quotation.items.filter(i => i.type === 'item');
        const nextNo = currentItems.length + 1;
        this.quotation.items.push({
          id: 'row_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
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
        if (!Array.isArray(this.quotation.items)) {
          this.quotation.items = [];
        }
        this.quotation.items.push({
          id: 'row_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
          type: 'bullet',
          description: ' - '
        });
      },

      addPaymentTermsPreset() {
        if (!Array.isArray(this.quotation.items)) {
          this.quotation.items = [];
        }
        this.quotation.items.push({
          id: 'row_' + Date.now() + '_note',
          type: 'note',
          description: 'การเบิกค่างวดงาน'
        });
        this.quotation.items.push({
          id: 'row_' + Date.now() + '_bullet',
          type: 'bullet',
          description: 'เบิกมัดจำ 70% และ 30% เมื่องานเสร็จสิ้น'
        });
      },

      removeItemRow(index) {
        if (!Array.isArray(this.quotation.items)) return;
        if (index >= 0 && index < this.quotation.items.length) {
          this.quotation.items.splice(index, 1);
          // Re-index main item numbers
          let counter = 1;
          this.quotation.items.forEach(it => {
            if (it && it.type === 'item') {
              it.itemNo = String(counter++);
            }
          });
          this.calculateTotals();
        }
      },

      calculateTotals() {
        let sum = 0;
        if (Array.isArray(this.quotation.items)) {
          this.quotation.items.forEach(it => {
            if (it && it.type === 'item') {
              const q = parseNum(it.qty);
              const p = parseNum(it.unitPrice);
              it.amount = Math.round(q * p * 100) / 100;
              sum += it.amount;
            }
          });
        }
        this.quotation.totalAmount = Math.round(sum * 100) / 100;
        const bText = getBahtText();
        this.quotation.totalTextThai = bText(this.quotation.totalAmount);
      },

      saveQuotation() {
        this.calculateTotals();
        const smMgr = getStorageManager();
        if (smMgr.saveToHistory) {
          smMgr.saveToHistory(this.quotation);
        }
        this.refreshHistory();
        this.flashNotice(`บันทึกเอกสาร ${this.quotation.docNo} เรียบร้อยแล้ว`);
      },

      loadFromHistory(item) {
        if (!item) return;
        this.quotation = JSON.parse(JSON.stringify(item));
        this.calculateTotals();
        this.showHistoryModal = false;
        this.flashNotice(`โหลดเอกสาร ${this.quotation.docNo} แล้ว`);
      },

      duplicateFromHistory(id) {
        const smMgr = getStorageManager();
        if (!smMgr.duplicateQuotation) return;
        const clone = smMgr.duplicateQuotation(id);
        if (clone) {
          if (smMgr.saveToHistory) {
            smMgr.saveToHistory(clone);
          }
          this.quotation = clone;
          this.calculateTotals();
          this.refreshHistory();
          this.showHistoryModal = false;
          this.flashNotice(`ทำสำเนาเป็นเอกสารใหม่ ${clone.docNo} เรียบร้อย`);
        }
      },

      deleteFromHistory(id) {
        const doConfirm = typeof confirm === 'function' ? confirm('คุณแน่ใจหรือไม่ว่าต้องการลบใบเสนอราคานี้?') : true;
        if (doConfirm) {
          const smMgr = getStorageManager();
          if (smMgr.deleteFromHistory) {
            smMgr.deleteFromHistory(id);
          }
          this.refreshHistory();
          this.flashNotice('ลบเอกสารเรียบร้อยแล้ว');
        }
      },

      refreshHistory() {
        const smMgr = getStorageManager();
        this.historyList = (smMgr.getHistory ? smMgr.getHistory() : []) || [];
      },

      filteredHistory() {
        if (!this.historySearch || !this.historySearch.trim()) return this.historyList;
        const q = this.historySearch.toLowerCase().trim();
        return this.historyList.filter(item => {
          if (!item) return false;
          const docNo = (item.docNo || '').toLowerCase();
          const customerName = (item.customer && item.customer.name ? item.customer.name : '').toLowerCase();
          const jobTitle = (item.jobTitle || '').toLowerCase();
          return docNo.includes(q) || customerName.includes(q) || jobTitle.includes(q);
        });
      },

      openSellerSettings() {
        const smMgr = getStorageManager();
        this.sellerProfileForm = smMgr.getSellerProfile ? { ...smMgr.getSellerProfile() } : {};
        this.showSellerModal = true;
      },

      saveSellerSettings() {
        const smMgr = getStorageManager();
        if (smMgr.saveSellerProfile) {
          smMgr.saveSellerProfile(this.sellerProfileForm);
        }
        this.quotation.seller = { ...this.sellerProfileForm };
        this.showSellerModal = false;
        this.flashNotice('บันทึกข้อมูลผู้เสนอราคาเรียบร้อย');
      },

      flashNotice(msg) {
        this.saveNotice = msg;
        if (this._noticeTimeout) {
          clearTimeout(this._noticeTimeout);
        }
        this._noticeTimeout = setTimeout(() => {
          if (this.saveNotice === msg) this.saveNotice = '';
        }, 3500);
      },

      getFilename(ext) {
        const ee = getExcelExporter();
        if (ee && typeof ee.getExportFilename === 'function') {
          return ee.getExportFilename(this.quotation, ext);
        }
        const custName = (this.quotation.customer && this.quotation.customer.name) || 'ลูกค้า';
        const docNo = this.quotation.docNo || 'NO';
        return `ใบเสนอราคา_${custName}_${docNo}.${ext}`;
      },

      async exportExcel() {
        try {
          this.calculateTotals();
          const ee = getExcelExporter();
          const excelJsLib = typeof ExcelJS !== 'undefined' ? ExcelJS : (typeof window !== 'undefined' ? window.ExcelJS : null);
          if (ee && typeof ee.downloadExcelFile === 'function') {
            await ee.downloadExcelFile(this.quotation, excelJsLib);
            this.flashNotice('ส่งออกไฟล์ Excel สำเร็จ');
          }
        } catch (err) {
          console.error('Export Excel failed:', err);
          this.flashNotice('ส่งออก Excel ล้มเหลว: ' + err.message);
        }
      },

      async exportPNG() {
        try {
          const filename = this.getFilename('png');
          const de = getDocExporter();
          if (de && typeof de.exportToPNG === 'function') {
            await de.exportToPNG('quotation-preview', filename);
            this.flashNotice('ส่งออกรูปภาพ PNG สำเร็จ');
          }
        } catch (err) {
          console.error('Export PNG failed:', err);
          this.flashNotice('ส่งออก PNG ล้มเหลว: ' + err.message);
        }
      },

      async exportPDF() {
        try {
          const filename = this.getFilename('pdf');
          const de = getDocExporter();
          if (de && typeof de.exportToPDF === 'function') {
            await de.exportToPDF('quotation-preview', filename);
            this.flashNotice('ส่งออกไฟล์ PDF สำเร็จ');
          }
        } catch (err) {
          console.error('Export PDF failed:', err);
          this.flashNotice('ส่งออก PDF ล้มเหลว: ' + err.message);
        }
      },

      printDocument() {
        const de = getDocExporter();
        if (de && typeof de.printDocument === 'function') {
          de.printDocument();
        } else if (typeof window !== 'undefined' && typeof window.print === 'function') {
          window.print();
        }
      },

      formatDateShort(d) {
        const ee = getExcelExporter();
        return (ee && typeof ee.formatThaiDateShort === 'function') ? ee.formatThaiDateShort(d) : (d || '');
      },

      formatDateLong(d) {
        const ee = getExcelExporter();
        return (ee && typeof ee.formatThaiDateLong === 'function') ? ee.formatThaiDateLong(d) : (d || '');
      },

      formatCurrency(val) {
        if (val === undefined || val === null || val === '') return '';
        const num = typeof val === 'number' ? val : Number(String(val).replace(/,/g, '').trim());
        if (isNaN(num)) return '';
        return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      }
    };
  }

  // Auto-register with Alpine if Alpine is present
  function registerAlpine() {
    const alpineInstance = (typeof Alpine !== 'undefined') ? Alpine : (typeof window !== 'undefined' ? window.Alpine : null);
    if (alpineInstance && typeof alpineInstance.data === 'function') {
      alpineInstance.data('quotationApp', quotationApp);
    }
  }

  if (typeof document !== 'undefined') {
    document.addEventListener('alpine:init', () => {
      registerAlpine();
    });
    registerAlpine();
  }

  quotationApp.quotationApp = quotationApp;
  return quotationApp;
}));
