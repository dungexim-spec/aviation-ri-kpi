// Kiểm thử engine bằng dữ liệu giả lập (không dùng dữ liệu thật). Chạy: npm install && npm test
const assert = require('assert');
const XLSX = require('xlsx');
const E = require('../js/engine.js');

function wb(sheets) { const b = XLSX.utils.book_new(); Object.entries(sheets).forEach(([n, rows]) => XLSX.utils.book_append_sheet(b, XLSX.utils.aoa_to_sheet(rows), n)); return b; }
const P = ['MRO', 'Broker', 'PVI ref number', 'Original Insured', 'Region', 'Domicile', 'Risk offer date', 'Status', 'Risk bound date', 'Inception', 'Expiry', 'Leader', 'Type', 'Channel'];
const R = ['MRO', 'PVI share', 'Hull', 'CSL', 'A/C', 'AFV'];
const M = ['MRO', 'Deposit GWP', 'PVI deposite GWP', 'PVI deposite NWP', 'Total deduction', 'Facility fee', 'Other Deduction', 'Fund', 'PVI actual earning', 'Policy status'];
const C = ['MRO', 'Claim number', 'Date claim reported', 'Occurrence date', 'Coverage', 'Type of loss', 'Gross Incurred', 'PVI share', 'PVI Paid Claim', 'PVI Paid Fee', 'PVI reserved claim', 'PVI reserved Fee', 'Status', 'Original Insured'];
const I = ['MRO', 'Broker', 'Original Insured', 'Inst 1 Amount', 'Inst 1 Due Date', 'Inst 1 Paid', 'Inst 1 Actual NWP', 'Inst 2 Amount', 'Inst 2 Due Date', 'Inst 2 Paid', 'Inst 2 Actual NWP'];

const book1 = wb({
  Policy: [P,
    ['M1', 'AJG', 'PVI1', 'Alpha Air', 'Asia', 'Vietnam', '2025-01-01', 'Bound', '2025-01-05', '2025-01-10', '2026-01-09', 'TMK', 'RI', 'A1'],
    ['M2', 'Marsh', 'PVI2', 'Beta Cargo', 'Africa', 'Libya', '2025-03-01', 'Bound', '2025-04-03', '2025-04-01', '2025-03-31', 'AXA XL', 'RI', 'MAAPS']],
  Risk: [R, ['M1', 0.5, 100000000, 1000000000, 20, 800000000], ['M2', 1.5, 20000000, 500000000, 3, 50000000]],
  Premium: [M, ['M1', 2000000, 10000, 7500, 25, 0, 0, 0, 7500, 'Bound'], ['M2', 400000, 6000, 4800, 20, 5, 100, 0, 4460, 'Bound']],
  Claim: [C, ['M1', 'C1', '2025-06-02', '2025-06-01', 'Hull', 'Engine', 1000000, 0.5, 2000, 0, 3000, 0, 'Open', 'Alpha Air'],
    ['M2', 'C2', '2026-07-01', '2026-06-30', 'Hull', 'Total loss', 2000000, 1.5, 0, 0, 30000, 0, 'Open', 'Beta Cargo']],
  Installment: [I, ['M1', 'AJG', 'Alpha Air', 3750, '2025-02-10', 'Paid', 3750, 3750, '2025-08-10', '', 0]],
});

const asOf = E.parseDate('2025-07-10');
const b = E.readWorkbook(XLSX, book1, E.uyFromFilename('UY_1_test.xlsx'));
assert.strictEqual(b.uy, 'UY1');
assert.deepStrictEqual(b.warnings, []);
const m = E.build([b], { asOf });
const [p1, p2] = m.policies;

// chuẩn hoá
assert.strictEqual(p1.region, 'Asia Pacific RIM');
assert.strictEqual(p1.leader, 'Tokio Marine Kiln');
// phí đã hưởng: 10/01/2025 → 10/07/2025 = 182 ngày / 365
assert.strictEqual(p1.dur, 365);
assert.strictEqual(p1.onRisk, 182);
assert.ok(Math.abs(p1.ep - 7500 * 182 / 365) < 1e-6);
// hết hạn lỗi → hiệu lực + 12 tháng
assert.strictEqual(E.fmtDate(p2.expC), '31/03/2026');
// bind sau hiệu lực + vùng nhạy cảm + line vượt trần
assert.ok(p2.flags.includes('Bind sau hiệu lực'));
assert.ok(p2.flags.includes('Vùng nhạy cảm'));
assert.ok(p2.flags.some((f) => f.startsWith('Line ')));
// bồi thường
assert.strictEqual(p1.incurred, 5000);
const c2 = m.claims.find((c) => c.claimNo === 'C2');
assert.strictEqual(c2.flag, 'Ngoài kỳ HĐ');
// tuổi nợ: kỳ 2 hết hạn 10/08/2025 > asOf → chưa quá hạn
assert.strictEqual(m.insts.filter((i) => i.days > 0).length, 0);
const later = E.build([b], { asOf: E.parseDate('2025-09-10') });
assert.strictEqual(later.insts.find((i) => i.n === 2).days, 31);
assert.strictEqual(later.insts.find((i) => i.n === 2).bucket, '31-90');
// tổng hợp
const s = E.summarize(m);
assert.strictEqual(s.n, 2);
assert.strictEqual(s.gwp, 16000);
assert.strictEqual(s.incurred, 35000);
// kiểm tra dữ liệu
assert.ok(m.issues.some((i) => i.field === 'Expiry' && i.sev === 'Cao'));
assert.ok(m.issues.some((i) => i.field === 'Lịch thanh toán'));
// ngày
assert.strictEqual(E.fmtDate(E.parseDate('05/10/2026')), '05/10/2026');
assert.strictEqual(E.parseDate('2026-21-08'), null);
assert.strictEqual(E.fmtDate(E.parseDate(46300)), E.fmtDate(E.parseDate('2026-10-05')));
console.log('OK – tất cả kiểm thử engine đều đạt');
