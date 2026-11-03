(function(root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.bahtText = factory().bahtText;
  }
}(typeof self !== 'undefined' ? self : this, function() {
  const NUMBERS = ['', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
  const DIGITS = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน', 'ล้าน'];

  function readBlock(numStr) {
    let result = '';
    const len = numStr.length;
    const hasTens = len > 1 && numStr.charAt(len - 2) !== '0';
    for (let i = 0; i < len; i++) {
      const digit = parseInt(numStr.charAt(i), 10);
      const pos = len - i - 1;
      if (digit !== 0) {
        if (pos === 0 && digit === 1 && hasTens) {
          result += 'เอ็ด';
        } else if (pos === 1 && digit === 2) {
          result += 'ยี่';
        } else if (pos === 1 && digit === 1) {
          result += '';
        } else {
          result += NUMBERS[digit];
        }
        result += DIGITS[pos];
      }
    }
    return result;
  }

  function bahtText(input) {
    if (input === null || input === undefined || input === '') return '(ศูนย์บาทถ้วน)';
    const cleanInput = typeof input === 'string' ? input.replace(/,/g, '').trim() : input;
    const num = parseFloat(cleanInput);
    if (isNaN(num)) return '(ศูนย์บาทถ้วน)';
    if (num === 0) return '(ศูนย์บาทถ้วน)';

    const isNegative = num < 0;
    const absNum = Math.abs(num);
    const fixed = absNum.toFixed(2);
    const parts = fixed.split('.');
    let integerPart = parts[0];
    const decimalPart = parts[1];

    let bahtStr = '';
    if (parseInt(integerPart, 10) === 0) {
      bahtStr = '';
    } else {
      let millions = [];
      while (integerPart.length > 0) {
        millions.unshift(integerPart.slice(Math.max(0, integerPart.length - 6)));
        integerPart = integerPart.slice(0, Math.max(0, integerPart.length - 6));
      }
      for (let i = 0; i < millions.length; i++) {
        const blockText = readBlock(millions[i]);
        if (blockText) {
          bahtStr += blockText + (i < millions.length - 1 ? 'ล้าน' : '');
        }
      }
      bahtStr += 'บาท';
    }

    let satangStr = '';
    const satangVal = parseInt(decimalPart, 10);
    if (satangVal === 0) {
      satangStr = bahtStr ? 'ถ้วน' : 'ศูนย์บาทถ้วน';
    } else {
      satangStr = readBlock(decimalPart) + 'สตางค์';
    }

    const full = (isNegative ? 'ลบ' : '') + (bahtStr + satangStr);
    return `(${full})`;
  }

  return { bahtText };
}));
