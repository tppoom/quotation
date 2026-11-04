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
      showMobileExportSheet: false,
      showMobileMenu: false,
      activeMobileTab: 'form', // 'form' | 'preview'
      previewZoom: 'fit', // 'fit' | '100' | '75'
      computedScale: 1,
      showSellerFields: false,
      isAutoSaved: true,
      pastCustomers: [],
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
          const savedSeller = smMgr.getSellerProfile ? smMgr.getSellerProfile() : {};
          if (!this.quotation.seller || !this.quotation.seller.name) {
            this.quotation.seller = { ...savedSeller };
          }
        } else {
          this.startNewQuotation();
        }
        this.calculateTotals();
        this.refreshHistory();
        this.loadPastCustomers();

        // Autosave draft on changes
        if (typeof this.$watch === 'function') {
          this.$watch('quotation', () => {
            this.calculateTotals();
            if (smMgr.saveCurrentDraft) {
              smMgr.saveCurrentDraft(this.quotation);
            }
            this.isAutoSaved = true;
          });
        }

        if (typeof window !== 'undefined') {
          window.addEventListener('resize', () => {
            this.computeScale();
          });
          window.addEventListener('orientationchange', () => {
            setTimeout(() => this.computeScale(), 150);
          });
          setTimeout(() => {
            this.computeScale();
          }, 100);
        }
      },

      autoSaveSeller() {
        const smMgr = getStorageManager();
        if (smMgr && smMgr.saveSellerProfile && this.quotation.seller) {
          smMgr.saveSellerProfile(this.quotation.seller);
        }
        this.sellerProfileForm = { ...(this.quotation.seller || {}) };
        this.isAutoSaved = true;
      },

      loadPastCustomers() {
        const smMgr = getStorageManager();
        const history = (smMgr.getHistory ? smMgr.getHistory() : []) || [];
        const map = new Map();
        history.forEach(item => {
          if (item && item.customer && item.customer.name && item.customer.name.trim()) {
            const name = item.customer.name.trim();
            if (!map.has(name)) {
              map.set(name, item.customer.address || '');
            }
          }
        });
        this.pastCustomers = Array.from(map.entries()).map(([name, address]) => ({ name, address }));
      },

      onCustomerSelect(name) {
        if (!name || typeof name !== 'string') return;
        const match = this.pastCustomers.find(c => c.name.toLowerCase() === name.trim().toLowerCase());
        if (match && match.address && (!this.quotation.customer.address || !this.quotation.customer.address.trim())) {
          this.quotation.customer.address = match.address;
          this.flashNotice(`ดึงที่อยู่ลูกค้า ${match.name} อัตโนมัติ`);
        }
      },

      computeScale() {
        if (typeof document === 'undefined') return;
        const container = document.getElementById('preview-outer-container');
        let containerWidth = container ? container.clientWidth : 0;
        if (!containerWidth || containerWidth < 50) {
          if (typeof window !== 'undefined') {
            containerWidth = Math.max(260, window.innerWidth - 24);
          } else {
            containerWidth = 360;
          }
        }
        // Inner available width subtracting container padding (p-2 = 16px, plus 8px safety)
        const availableWidth = Math.max(100, containerWidth - 24);
        if (this.previewZoom === 'fit') {
          if (availableWidth < 800) {
            // Math.floor ensures scaled width (794 * computedScale) never exceeds available width even by 0.1px
            this.computedScale = Math.max(0.2, Math.floor((availableWidth / 794) * 100) / 100);
          } else {
            this.computedScale = 1.0;
          }
        } else if (this.previewZoom === '75') {
          this.computedScale = 0.75;
        } else {
          this.computedScale = 1.0;
        }
      },

      setPreviewZoom(zoom) {
        this.previewZoom = zoom;
        this.computeScale();
      },

      setMobileTab(tab) {
        this.activeMobileTab = tab;
        if (tab === 'preview') {
          this.computeScale();
          setTimeout(() => this.computeScale(), 40);
          setTimeout(() => this.computeScale(), 150);
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

      loadDemoTemplate() {
        this.quotation = {
          id: 'qt_demo_' + Date.now(),
          docNo: 'QT-202204-0001',
          docDate: '2022-04-24',
          jobTitle: 'งานระบบน้ำอัตโนมัติ',
          customer: {
            name: 'บริษัท เพิ่ม เอสเคท จำกัด',
            address: '82 ถนนสังคมสงเคราะห์ แขวงลาดพร้าว กรุงเทพฯ'
          },
          seller: {
            name: 'นายสมชาย ใจดี (ผู้เสนอราคาตัวอย่าง)',
            address: '89/332 ถนนประชาอุทิศ แขวงดอนเมือง เขตดอนเมือง กรุงเทพฯ 10210',
            tel1: '081 234 5678',
            tel2: '089 876 5432',
            bank: 'ธ.กรุงเทพ 123-4-56789-0'
          },
          items: [
            {
              id: 'demo_item_1',
              type: 'item',
              itemNo: '1',
              description: 'ระบบรดน้ำอัตโนมัติ 4 โซน (ด้านหน้า)',
              unit: 'ชุด',
              qty: 1,
              unitPrice: 57552.9,
              amount: 57552.9
            },
            { id: 'demo_b_1', type: 'bullet', description: ' - หัวป็อบอัพ 4 หุนจำนวน 5 หัว' },
            { id: 'demo_b_2', type: 'bullet', description: ' - ท่อ PE 20 มิลลิเมตร' },
            { id: 'demo_b_3', type: 'bullet', description: ' - ท่อไมโคร' },
            { id: 'demo_b_4', type: 'bullet', description: ' - ตู้ไฟตั้งเวลาเปิด/ปิดโซลินอยวาล์ว 2 ชุด' },
            { id: 'demo_b_5', type: 'bullet', description: ' - หัวมินิสปริงเกอร์ + หัวมินิสปริงเกอร์พร้อมเสาปัก' },
            { id: 'demo_b_6', type: 'bullet', description: ' - อุปกรณ์ข้อต่อ' },
            { id: 'demo_b_7', type: 'bullet', description: 'ระบบรดน้ำอัตโนมัติ 1 โซน (ด้านหลัง)' },
            { id: 'demo_b_8', type: 'bullet', description: ' - ตู้ไฟตั้งเวลาเปิด/ปิดโซลินอยวาล์ว 2 ชุด' },
            { id: 'demo_b_9', type: 'bullet', description: ' - ท่อ + อุปกรณ์' },
            { id: 'demo_n_1', type: 'note', description: 'การเบิกค่างวดงาน' },
            { id: 'demo_b_10', type: 'bullet', description: 'เบิกมัดจำ 70% และ 30% เมื่องานเสร็จสิ้น' }
          ],
          totalAmount: 57552.9,
          totalTextThai: '(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)'
        };
        this.calculateTotals();
        setTimeout(() => this.computeScale(), 100);
        this.flashNotice('⚡ โหลดตัวอย่างใบเสนอราคาตามไฟล์ต้นแบบเรียบร้อยแล้ว');
      },

      loadLongDemoTemplate() {
        this.quotation = {
          id: 'qt_demo_long_' + Date.now(),
          docNo: 'QT-202204-LONG',
          docDate: '2022-04-24',
          jobTitle: 'งานติดตั้งระบบรดน้ำอัตโนมัติ สวนเกษตรเฟส 1-2 (โครงการใหญ่)',
          customer: {
            name: 'บริษัท เพิ่ม เอสเคท จำกัด',
            address: '82 ถนนสังคมสงเคราะห์ แขวงลาดพร้าว เขตลาดพร้าว กรุงเทพฯ 10230'
          },
          seller: {
            name: 'นายสมชาย ใจดี (ผู้เสนอราคาตัวอย่าง)',
            address: '89/332 ถนนประชาอุทิศ แขวงดอนเมือง เขตดอนเมือง กรุงเทพฯ 10210',
            tel1: '081 234 5678',
            tel2: '089 876 5432',
            bank: 'ธ.กรุงเทพ 123-4-56789-0'
          },
          items: [
            { id: 'l_1', type: 'item', itemNo: '1', description: 'ระบบรดน้ำอัตโนมัติ โซน 1 สวนหน้าบ้าน', unit: 'ชุด', qty: 1, unitPrice: 28500, amount: 28500 },
            { id: 'l_2', type: 'bullet', description: ' - หัวป็อบอัพ Hunter Pro-Spray 4 หุน 8 หัว' },
            { id: 'l_3', type: 'bullet', description: ' - ท่อ HDPE PN6 ขนาด 25 มม. ความยาว 60 เมตร' },
            { id: 'l_4', type: 'bullet', description: ' - ข้อต่อสวมอัด HDPE Compression Fit 12 ตัว' },
            { id: 'l_5', type: 'bullet', description: ' - โซลินอยด์วาล์ว 1 นิ้ว 24VAC พร้อมกล่องวาล์ว' },
            { id: 'l_6', type: 'item', itemNo: '2', description: 'ระบบมินิสปริงเกอร์ โซน 2 สวนหย่อมด้านข้าง', unit: 'ชุด', qty: 1, unitPrice: 19200, amount: 19200 },
            { id: 'l_7', type: 'bullet', description: ' - มินิสปริงเกอร์พร้อมก้านปักและสายไมโคร 24 จุด' },
            { id: 'l_8', type: 'bullet', description: ' - ท่อ PE 20 มม. ความยาว 45 เมตร' },
            { id: 'l_9', type: 'bullet', description: ' - วาล์วปรับแรงดันน้ำอัตโนมัติ 1 ชุด' },
            { id: 'l_10', type: 'item', itemNo: '3', description: 'ระบบน้ำหยด โซน 3 กระถางต้นไม้และแนวรั้ว', unit: 'ชุด', qty: 1, unitPrice: 14500, amount: 14500 },
            { id: 'l_11', type: 'bullet', description: ' - หัวน้ำหยดปรับอัตราการไหลชดเชยแรงดัน 40 จุด' },
            { id: 'l_12', type: 'bullet', description: ' - ท่อไมโคร PE 4/7 มม. ความยาว 80 เมตร' },
            { id: 'l_13', type: 'item', itemNo: '4', description: 'ระบบควบคุมอัจฉริยะ WiFi & Sensor', unit: 'ชุด', qty: 1, unitPrice: 18500, amount: 18500 },
            { id: 'l_14', type: 'bullet', description: ' - ตู้คอนโทรล 4 สถานี สั่งงานผ่านสมาร์ตโฟน' },
            { id: 'l_15', type: 'bullet', description: ' - เซนเซอร์วัดความชื้นในดินและฝนตกอัตโนมัติ' },
            { id: 'l_16', type: 'bullet', description: ' - ระบบป้องกันไฟกระชากและเบรกเกอร์กันดูด' },
            { id: 'l_17', type: 'item', itemNo: '5', description: 'งานปั๊มน้ำและระบบกรองน้ำเกษตร', unit: 'ชุด', qty: 1, unitPrice: 16800, amount: 16800 },
            { id: 'l_18', type: 'bullet', description: ' - ปั๊มหอยโข่งแสตนเลส 1.5 แรงม้า' },
            { id: 'l_19', type: 'bullet', description: ' - กรองตะกอนแบบดิสก์ 1.5 นิ้ว 120 เมช' },
            { id: 'l_20', type: 'bullet', description: ' - เช็ควาล์วและเกจวัดแรงดันน้ำ 0-10 บาร์' },
            { id: 'l_21', type: 'item', itemNo: '6', description: 'ค่าบริการติดตั้ง ขุดร่องดิน และทดสอบระบบ', unit: 'งาน', qty: 1, unitPrice: 15000, amount: 15000 },
            { id: 'l_22', type: 'bullet', description: ' - งานเดินท่อใต้ดินและเก็บความเรียบร้อยหน้างาน' },
            { id: 'l_23', type: 'bullet', description: ' - การทดสอบแรงดันน้ำและปรับรัศมีสปริงเกอร์' },
            { id: 'l_24', type: 'bullet', description: ' - คู่มือการใช้งานและฝึกอบรมการตั้งโปรแกรม' },
            { id: 'l_25', type: 'note', description: 'เงื่อนไขการรับประกันและการชำระเงิน' },
            { id: 'l_26', type: 'bullet', description: ' - รับประกันอุปกรณ์และงานติดตั้ง 1 ปีเต็ม' },
            { id: 'l_27', type: 'bullet', description: ' - มัดจำ 50% เมื่อสั่งซื้อ, 50% เมื่องานแล้วเสร็จส่งมอบ' }
          ],
          totalAmount: 112500,
          totalTextThai: '(หนึ่งแสนหนึ่งหมื่นสองพันห้าร้อยบาทถ้วน)'
        };
        this.calculateTotals();
        setTimeout(() => this.computeScale(), 100);
        this.flashNotice('⚡ โหลดตัวอย่างใบเสนอราคา 27 รายการ (ทดสอบรายการยาว) เรียบร้อยแล้ว');
      },

      addItemRow() {
        if (!Array.isArray(this.quotation.items)) {
          this.quotation.items = [];
        }
        const currentItems = this.quotation.items.filter(i => i.type === 'item');
        const nextNo = currentItems.length + 1;

        // Auto-inherit unit from the latest item row that has a unit defined
        let defaultUnit = 'ชุด';
        for (let i = this.quotation.items.length - 1; i >= 0; i--) {
          const it = this.quotation.items[i];
          if (it && it.type === 'item' && it.unit && String(it.unit).trim()) {
            defaultUnit = String(it.unit).trim();
            break;
          }
        }

        this.quotation.items.push({
          id: 'row_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
          type: 'item',
          itemNo: String(nextNo),
          description: '',
          unit: defaultUnit,
          qty: 1,
          unitPrice: 0,
          amount: 0
        });
        this.calculateTotals();
      },

      duplicateItemRow(index) {
        if (!Array.isArray(this.quotation.items)) return;
        const target = this.quotation.items[index];
        if (!target) return;
        const clone = JSON.parse(JSON.stringify(target));
        clone.id = 'row_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
        this.quotation.items.splice(index + 1, 0, clone);
        // Re-index main item numbers
        let counter = 1;
        this.quotation.items.forEach(it => {
          if (it && it.type === 'item') {
            it.itemNo = String(counter++);
          }
        });
        this.calculateTotals();
        this.flashNotice('📋 คัดลอกรายการเรียบร้อยแล้ว');
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
        if (smMgr.saveSellerProfile && this.quotation.seller) {
          smMgr.saveSellerProfile(this.quotation.seller);
        }
        this.refreshHistory();
        this.loadPastCustomers();
        this.flashNotice(`บันทึกเอกสาร ${this.quotation.docNo} ลงประวัติเรียบร้อยแล้ว`);
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
        const prevTab = this.activeMobileTab;
        try {
          if (this.activeMobileTab !== 'preview') {
            this.activeMobileTab = 'preview';
            await new Promise(r => setTimeout(r, 60));
          }
          const filename = this.getFilename('png');
          const de = getDocExporter();
          if (de && typeof de.exportToPNG === 'function') {
            await de.exportToPNG('quotation-preview', filename);
            this.flashNotice('ส่งออกรูปภาพ PNG สำเร็จ');
          }
        } catch (err) {
          console.error('Export PNG failed:', err);
          this.flashNotice('ส่งออก PNG ล้มเหลว: ' + err.message);
        } finally {
          this.activeMobileTab = prevTab;
        }
      },

      async exportPDF() {
        const prevTab = this.activeMobileTab;
        try {
          if (this.activeMobileTab !== 'preview') {
            this.activeMobileTab = 'preview';
            await new Promise(r => setTimeout(r, 60));
          }
          const filename = this.getFilename('pdf');
          const de = getDocExporter();
          if (de && typeof de.exportToPDF === 'function') {
            await de.exportToPDF('quotation-preview', filename);
            this.flashNotice('ส่งออกไฟล์ PDF สำเร็จ');
          }
        } catch (err) {
          console.error('Export PDF failed:', err);
          this.flashNotice('ส่งออก PDF ล้มเหลว: ' + err.message);
        } finally {
          this.activeMobileTab = prevTab;
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
