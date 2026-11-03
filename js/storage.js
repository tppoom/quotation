// js/storage.js
(function(root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.StorageManager = factory();
  }
}(typeof self !== 'undefined' ? self : this, function() {
  const KEYS = {
    SELLER: 'quotation_seller_profile',
    HISTORY: 'quotation_history',
    SEQUENCE: 'quotation_sequence',
    DRAFT: 'quotation_current_draft'
  };

  const DEFAULT_SELLER = {
    name: 'ชื่อผู้เสนอราคา / ร้านค้า',
    address: 'ที่อยู่ผู้เสนอราคา...',
    tel1: '08X-XXX-XXXX',
    tel2: '08X-XXX-XXXX',
    bank: 'ธนาคาร... เลขที่บัญชี...'
  };

  function getStorage() {
    if (typeof localStorage !== 'undefined') return localStorage;
    if (typeof window !== 'undefined' && window.localStorage) return window.localStorage;
    if (typeof global !== 'undefined' && global.localStorage) return global.localStorage;
    return null;
  }

  function getLocalDateStr(date) {
    const d = date instanceof Date ? date : (date ? new Date(date) : new Date());
    if (isNaN(d.getTime())) return '';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function getSellerProfile() {
    try {
      const storage = getStorage();
      const data = storage ? storage.getItem(KEYS.SELLER) : null;
      if (!data) return { ...DEFAULT_SELLER };
      const parsed = JSON.parse(data);
      if (parsed && typeof parsed === 'object') {
        return Object.assign({}, DEFAULT_SELLER, parsed);
      }
      return { ...DEFAULT_SELLER };
    } catch (e) {
      return { ...DEFAULT_SELLER };
    }
  }

  function saveSellerProfile(profile) {
    try {
      const storage = getStorage();
      if (storage) {
        storage.setItem(KEYS.SELLER, JSON.stringify(profile));
      }
    } catch (e) {
      console.error('Failed to save seller profile:', e);
    }
  }

  function getHistory() {
    try {
      const storage = getStorage();
      const data = storage ? storage.getItem(KEYS.HISTORY) : null;
      const parsed = data ? JSON.parse(data) : [];
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) {
      return [];
    }
  }

  function getSequenceMap() {
    try {
      const storage = getStorage();
      const data = storage ? storage.getItem(KEYS.SEQUENCE) : null;
      const parsed = data ? JSON.parse(data) : {};
      return (parsed && typeof parsed === 'object') ? parsed : {};
    } catch (e) {
      return {};
    }
  }

  function parseYearMonth(dateStr) {
    if (typeof dateStr === 'string') {
      const match = dateStr.match(/^(\d{4})[-/.]?(\d{2})/);
      if (match) {
        return `${match[1]}${match[2]}`;
      }
    }
    const d = dateStr ? new Date(dateStr) : new Date();
    if (isNaN(d.getTime())) {
      const fallback = new Date();
      return `${fallback.getFullYear()}${String(fallback.getMonth() + 1).padStart(2, '0')}`;
    }
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    return `${yyyy}${mm}`;
  }

  function generateDocNo(dateStr) {
    const ymKey = parseYearMonth(dateStr);
    const prefix = `QT-${ymKey}-`;

    // Inspect existing history to make sure we don't collide
    const history = getHistory();
    let maxSeq = 0;
    history.forEach(item => {
      if (item && item.docNo && typeof item.docNo === 'string' && item.docNo.startsWith(prefix)) {
        const seq = parseInt(item.docNo.slice(prefix.length), 10);
        if (!isNaN(seq) && seq > maxSeq) {
          maxSeq = seq;
        }
      }
    });

    const seqMap = getSequenceMap();
    const trackedSeq = typeof seqMap[ymKey] === 'number' ? seqMap[ymKey] : 0;
    const nextSeq = Math.max(maxSeq, trackedSeq) + 1;

    return `${prefix}${String(nextSeq).padStart(4, '0')}`;
  }

  function recordDocNoUsed(docNo) {
    if (!docNo || typeof docNo !== 'string' || !docNo.startsWith('QT-')) return;
    const parts = docNo.split('-');
    if (parts.length === 3) {
      const ymKey = parts[1];
      const seq = parseInt(parts[2], 10);
      if (!isNaN(seq)) {
        const seqMap = getSequenceMap();
        const current = typeof seqMap[ymKey] === 'number' ? seqMap[ymKey] : 0;
        if (seq > current) {
          seqMap[ymKey] = seq;
          try {
            const storage = getStorage();
            if (storage) {
              storage.setItem(KEYS.SEQUENCE, JSON.stringify(seqMap));
            }
          } catch (e) {}
        }
      }
    }
  }

  function saveToHistory(quotation) {
    if (!quotation || typeof quotation !== 'object') return;

    if (!quotation.id) {
      quotation.id = 'qt_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
    }
    quotation.updatedAt = new Date().toISOString();
    if (!quotation.createdAt) {
      quotation.createdAt = quotation.updatedAt;
    }

    const history = getHistory();
    const existingIndex = history.findIndex(item => item && item.id === quotation.id);
    if (existingIndex >= 0) {
      history[existingIndex] = quotation;
    } else {
      history.unshift(quotation);
    }

    try {
      const storage = getStorage();
      if (storage) {
        storage.setItem(KEYS.HISTORY, JSON.stringify(history));
      }
    } catch (e) {
      console.error('Failed to save to history:', e);
    }

    recordDocNoUsed(quotation.docNo);
  }

  function deleteFromHistory(id) {
    const history = getHistory().filter(item => item && item.id !== id);
    try {
      const storage = getStorage();
      if (storage) {
        storage.setItem(KEYS.HISTORY, JSON.stringify(history));
      }
    } catch (e) {
      console.error('Failed to delete from history:', e);
    }
  }

  function duplicateQuotation(id) {
    const history = getHistory();
    const source = history.find(item => item && item.id === id);
    if (!source) return null;

    const clone = JSON.parse(JSON.stringify(source));
    clone.id = 'qt_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
    const today = getLocalDateStr(new Date());
    clone.docNo = generateDocNo(today);
    clone.docDate = today;
    clone.createdAt = new Date().toISOString();
    clone.updatedAt = clone.createdAt;
    return clone;
  }

  function loadCurrentDraft() {
    try {
      const storage = getStorage();
      const data = storage ? storage.getItem(KEYS.DRAFT) : null;
      return data ? JSON.parse(data) : null;
    } catch (e) {
      return null;
    }
  }

  function saveCurrentDraft(draft) {
    try {
      const storage = getStorage();
      if (storage) {
        storage.setItem(KEYS.DRAFT, JSON.stringify(draft));
      }
    } catch (e) {
      console.error('Failed to save draft:', e);
    }
  }

  function clearCurrentDraft() {
    try {
      const storage = getStorage();
      if (storage) {
        storage.removeItem(KEYS.DRAFT);
      }
    } catch (e) {}
  }

  return {
    getSellerProfile,
    loadSellerProfile: getSellerProfile,
    saveSellerProfile,
    getHistory,
    saveToHistory,
    deleteFromHistory,
    duplicateQuotation,
    generateDocNo,
    recordDocNoUsed,
    loadCurrentDraft,
    saveCurrentDraft,
    clearCurrentDraft,
    getLocalDateStr,
    DEFAULT_SELLER,
    KEYS
  };
}));
