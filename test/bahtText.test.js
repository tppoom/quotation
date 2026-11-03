const assert = require('assert');
const { bahtText } = require('../js/bahtText.js');

function runTests() {
  // Original specification cases
  assert.strictEqual(bahtText(0), '(ศูนย์บาทถ้วน)');
  assert.strictEqual(bahtText(1), '(หนึ่งบาทถ้วน)');
  assert.strictEqual(bahtText(11), '(สิบเอ็ดบาทถ้วน)');
  assert.strictEqual(bahtText(21), '(ยี่สิบเอ็ดบาทถ้วน)');
  assert.strictEqual(bahtText(101), '(หนึ่งร้อยหนึ่งบาทถ้วน)');
  assert.strictEqual(bahtText(1001), '(หนึ่งพันหนึ่งบาทถ้วน)');
  assert.strictEqual(bahtText(57552.90), '(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)');
  assert.strictEqual(bahtText(1000000), '(หนึ่งล้านบาทถ้วน)');
  assert.strictEqual(bahtText(10000001), '(สิบล้านหนึ่งบาทถ้วน)');
  assert.strictEqual(bahtText(0.25), '(ยี่สิบห้าสตางค์)');

  // Additional boundary and edge cases
  assert.strictEqual(bahtText(111), '(หนึ่งร้อยสิบเอ็ดบาทถ้วน)');
  assert.strictEqual(bahtText(121), '(หนึ่งร้อยยี่สิบเอ็ดบาทถ้วน)');
  assert.strictEqual(bahtText(0.01), '(หนึ่งสตางค์)');
  assert.strictEqual(bahtText(0.11), '(สิบเอ็ดสตางค์)');
  assert.strictEqual(bahtText(0.21), '(ยี่สิบเอ็ดสตางค์)');
  assert.strictEqual(bahtText(-50), '(ลบห้าสิบบาทถ้วน)');
  assert.strictEqual(bahtText('57552.90'), '(ห้าหมื่นเจ็ดพันห้าร้อยห้าสิบสองบาทเก้าสิบสตางค์)');
  assert.strictEqual(bahtText(null), '(ศูนย์บาทถ้วน)');
  assert.strictEqual(bahtText(undefined), '(ศูนย์บาทถ้วน)');
  assert.strictEqual(bahtText(''), '(ศูนย์บาทถ้วน)');
  assert.strictEqual(bahtText('invalid'), '(ศูนย์บาทถ้วน)');

  // Browser global export check
  const fs = require('fs');
  const code = fs.readFileSync(require.resolve('../js/bahtText.js'), 'utf8');
  const fakeWindow = {};
  const fn = new Function('self', 'module', 'exports', code);
  fn(fakeWindow, undefined, undefined);
  assert.strictEqual(typeof fakeWindow.bahtText, 'function');
  assert.strictEqual(fakeWindow.bahtText(100), '(หนึ่งร้อยบาทถ้วน)');

  console.log('✅ All bahtText tests passed!');
}

runTests();
