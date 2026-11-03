const assert = require('assert');
const fs = require('fs');

// Mock localStorage for node environment
global.localStorage = (function() {
  let store = {};
  return {
    getItem: (key) => (Object.prototype.hasOwnProperty.call(store, key) ? store[key] : null),
    setItem: (key, val) => { store[key] = String(val); },
    removeItem: (key) => { delete store[key]; },
    clear: () => { store = {}; }
  };
})();

const StorageManager = require('../js/storage.js');
const {
  generateDocNo,
  recordDocNoUsed,
  getSellerProfile,
  loadSellerProfile,
  saveSellerProfile,
  getHistory,
  saveToHistory,
  deleteFromHistory,
  duplicateQuotation,
  loadCurrentDraft,
  saveCurrentDraft,
  clearCurrentDraft,
  getLocalDateStr,
  DEFAULT_SELLER,
  KEYS
} = StorageManager;

function testStorage() {
  localStorage.clear();

  // =========================================================================
  // 1. Doc Number Generation Sequence and Collision Protection
  // =========================================================================
  console.log('Testing doc number generation...');
  const doc1 = generateDocNo('2026-11-03');
  assert.strictEqual(doc1, 'QT-202611-0001', 'First doc in 2026-11 should be 0001');

  // Saving doc1 to history should advance sequence
  saveToHistory({ id: 'qt_1', docNo: doc1, customer: { name: 'บ.เอ' } });

  const doc2 = generateDocNo('2026-11-03');
  assert.strictEqual(doc2, 'QT-202611-0002', 'Next doc in 2026-11 should be 0002');

  // Different month starts sequence at 0001
  const docDec = generateDocNo('2026-12-01');
  assert.strictEqual(docDec, 'QT-202612-0001', 'First doc in 2026-12 should be 0001');

  // No arguments generates for current year-month
  const now = new Date();
  const currentYm = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}`;
  const docCurrent = generateDocNo();
  assert(docCurrent.startsWith(`QT-${currentYm}-`), 'DocNo without dateStr uses current YYYYMM');

  // Collision prevention when quotation is deleted from history
  saveToHistory({ id: 'qt_2', docNo: doc2, customer: { name: 'บ.บี' } });
  deleteFromHistory('qt_2'); // history length drops, but sequence tracker remains
  const doc3 = generateDocNo('2026-11-15');
  assert.strictEqual(doc3, 'QT-202611-0003', 'Sequence must not repeat even if history item is deleted');

  // Collision prevention if sequence tracker is wiped but history has existing items
  localStorage.removeItem(KEYS.SEQUENCE);
  // History still has doc1 (0001)
  const docRecovered = generateDocNo('2026-11-20');
  assert.strictEqual(docRecovered, 'QT-202611-0002', 'Sequence recovers next number from history if sequence tracker wiped');

  // Direct sequence recording
  recordDocNoUsed('QT-202611-0050');
  assert.strictEqual(generateDocNo('2026-11-25'), 'QT-202611-0051', 'Direct recordDocNoUsed advances sequence');

  // =========================================================================
  // 2. Seller Profile & Partial Key Fallback
  // =========================================================================
  console.log('Testing seller profile...');
  localStorage.clear();

  // Default generic profile returned when empty
  const defaultProfile = getSellerProfile();
  assert.strictEqual(defaultProfile.name, 'ชื่อผู้เสนอราคา / ร้านค้า');
  assert.strictEqual(defaultProfile.bank, 'ธนาคาร... เลขที่บัญชี...');
  assert.deepStrictEqual(loadSellerProfile(), defaultProfile, 'loadSellerProfile is alias to getSellerProfile');

  // Mutating returned default profile does not mutate constant
  defaultProfile.name = 'Changed Name';
  assert.strictEqual(getSellerProfile().name, 'ชื่อผู้เสนอราคา / ร้านค้า');

  // Custom profile saving and loading
  saveSellerProfile({ name: 'ทดสอบ', tel1: '0812345678' });
  const profile = getSellerProfile();
  assert.strictEqual(profile.name, 'ทดสอบ');
  assert.strictEqual(profile.tel1, '0812345678');
  // Partial profile fallback: unspecified keys fallback to DEFAULT_SELLER
  assert.strictEqual(profile.bank, DEFAULT_SELLER.bank, 'Missing fields in profile fallback to DEFAULT_SELLER');
  assert.strictEqual(profile.address, DEFAULT_SELLER.address);

  // Malformed JSON falls back safely to default profile
  localStorage.setItem(KEYS.SELLER, '{malformed json}');
  assert.strictEqual(getSellerProfile().name, 'ชื่อผู้เสนอราคา / ร้านค้า');

  // =========================================================================
  // 3. History CRUD Operations
  // =========================================================================
  console.log('Testing history operations...');
  localStorage.clear();
  assert.deepStrictEqual(getHistory(), [], 'Initially history is empty');

  // Save new quotation without id, createdAt, or updatedAt
  const quote1 = { docNo: 'QT-202611-0001', customer: { name: 'ลูกค้า ก' } };
  saveToHistory(quote1);

  let history = getHistory();
  assert.strictEqual(history.length, 1);
  assert(history[0].id && history[0].id.startsWith('qt_'), 'Auto-assigns unique id');
  assert(history[0].createdAt, 'Auto-assigns createdAt');
  assert(history[0].updatedAt, 'Auto-assigns updatedAt');
  assert.strictEqual(history[0].customer.name, 'ลูกค้า ก');

  const savedId1 = history[0].id;

  // Save another quotation - should be prepended (most recent first)
  const quote2 = { id: 'qt_custom_2', docNo: 'QT-202611-0002', customer: { name: 'ลูกค้า ข' } };
  saveToHistory(quote2);

  history = getHistory();
  assert.strictEqual(history.length, 2);
  assert.strictEqual(history[0].id, 'qt_custom_2', 'Newest item prepended to top of history');

  // Update existing quotation
  const updatedQuote2 = { ...quote2, customer: { name: 'ลูกค้า ข (แก้ไข)' } };
  saveToHistory(updatedQuote2);

  history = getHistory();
  assert.strictEqual(history.length, 2, 'Updating existing id should not duplicate row');
  assert.strictEqual(history[0].customer.name, 'ลูกค้า ข (แก้ไข)');

  // Delete from history
  deleteFromHistory(savedId1);
  history = getHistory();
  assert.strictEqual(history.length, 1);
  assert.strictEqual(history[0].id, 'qt_custom_2');

  // Malformed history JSON recovers gracefully to empty array
  localStorage.setItem(KEYS.HISTORY, 'not json');
  assert.deepStrictEqual(getHistory(), []);

  // =========================================================================
  // 4. Duplicate Quotation & Local Date Handling
  // =========================================================================
  console.log('Testing duplicate quotation...');
  localStorage.clear();
  const original = {
    id: 'orig_1',
    docNo: 'QT-202610-0099',
    docDate: '2026-10-15',
    customer: { name: 'บริษัท ออริจินัล' },
    items: [{ no: 1, desc: 'ต้นไม้ 1', qty: 2, price: 1000 }]
  };
  saveToHistory(original);

  const nonExistent = duplicateQuotation('non_existent_id');
  assert.strictEqual(nonExistent, null, 'Duplicate non-existent returns null');

  const duplicated = duplicateQuotation('orig_1');
  assert(duplicated !== null, 'Duplicated object created');
  assert.notStrictEqual(duplicated.id, 'orig_1', 'Duplicated has new id');
  assert.notStrictEqual(duplicated.docNo, 'QT-202610-0099', 'Duplicated gets fresh docNo');
  assert(duplicated.docNo.startsWith('QT-'), 'DocNo follows QT- pattern');
  assert.strictEqual(duplicated.customer.name, 'บริษัท ออริจินัล', 'Clones deep properties');
  assert.strictEqual(duplicated.items[0].desc, 'ต้นไม้ 1', 'Clones item array');
  assert.strictEqual(duplicated.items.length, 1);

  // Verify local date string format YYYY-MM-DD
  const expectedToday = getLocalDateStr(new Date());
  assert.strictEqual(duplicated.docDate, expectedToday, 'Uses local date string rather than UTC slice');

  // Modifying duplicated does not mutate original in history
  duplicated.items[0].desc = 'ต้นไม้แก้ไข';
  const checkOrig = getHistory().find(q => q.id === 'orig_1');
  assert.strictEqual(checkOrig.items[0].desc, 'ต้นไม้ 1');

  // =========================================================================
  // 5. Draft Management (Auto-save & restore)
  // =========================================================================
  console.log('Testing draft auto-save and restore...');
  localStorage.clear();

  assert.strictEqual(loadCurrentDraft(), null, 'Initially draft is null');

  const draftData = {
    customer: { name: 'ร่างใบเสนอราคา' },
    items: [{ no: 1, desc: 'ทดสอบร่าง', qty: 5, price: 200 }]
  };
  saveCurrentDraft(draftData);

  const loadedDraft = loadCurrentDraft();
  assert.deepStrictEqual(loadedDraft, draftData, 'Loaded draft matches saved draft');

  if (typeof clearCurrentDraft === 'function') {
    clearCurrentDraft();
    assert.strictEqual(loadCurrentDraft(), null, 'clearCurrentDraft removes draft');
  }

  // Malformed JSON returns null safely
  localStorage.setItem(KEYS.DRAFT, 'corrupted draft');
  assert.strictEqual(loadCurrentDraft(), null);

  // =========================================================================
  // 6. UMD / Browser Compatibility
  // =========================================================================
  console.log('Testing UMD browser export...');
  const code = fs.readFileSync(require.resolve('../js/storage.js'), 'utf8');
  const fakeWindow = { localStorage: global.localStorage };
  const fn = new Function('self', 'module', 'exports', code);
  fn(fakeWindow, undefined, undefined);

  assert.strictEqual(typeof fakeWindow.StorageManager, 'object');
  assert.strictEqual(typeof fakeWindow.StorageManager.generateDocNo, 'function');
  assert.strictEqual(typeof fakeWindow.StorageManager.saveToHistory, 'function');
  assert.strictEqual(typeof fakeWindow.StorageManager.loadSellerProfile, 'function');
  assert.strictEqual(typeof fakeWindow.StorageManager.getLocalDateStr, 'function');

  console.log('✅ All storage tests passed successfully!');
}

testStorage();
