/*
 * Engine tính chỉ số danh mục nhận tái bảo hiểm hàng không.
 * Đọc workbook theo định dạng UY_x_update (sheet Policy, Risk, Premium, Claim, Installment),
 * chuẩn hoá dữ liệu, tính chỉ số và phát hiện lỗi dữ liệu.
 * Chạy được cả trên trình duyệt (window.KpiEngine) và Node (module.exports) để kiểm thử.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.KpiEngine = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------- Thiết lập mặc định ----------
  const DEFAULT_SETTINGS = {
    asOf: null,                 // số ngày (UTC) – null = hôm nay
    lineCapSmall: 0.005,        // trần line cho hãng nhỏ / vùng rủi ro cao
    lineCapGeneral: 0.01,       // trần line chung
    smallFleet: 10,             // < số tàu này coi là hãng nhỏ
    leaderCap: 0.30,            // tỷ trọng GWP tối đa cho một leader trong một UY
    minPremium: 2500,           // GWP tối thiểu mỗi line
    overdueAlertDays: 90,
    upcomingDays: 30,           // nhắc kỳ phí sắp đến hạn trong số ngày này
    highRiskCountries: ['Syria', 'Ukraine', 'Libya', 'Iraq', 'Myanmar', 'Mali', 'Afghanistan', 'Iran', 'Yemen', 'Sudan', 'Russia', 'Belarus', 'Venezuela', 'North Korea', 'Cuba'],
    regionMap: { 'Asia': 'Asia Pacific RIM', 'LATAM': 'Americas', 'South America': 'Americas' },
    leaderMap: { 'TMK': 'Tokio Marine Kiln' },
  };

  // ---------- Ngày tháng (số ngày kể từ 1970-01-01 UTC) ----------
  const DAY = 86400000;
  const PLACEHOLDER = Date.UTC(2025, 0, 1) / DAY;
  function ymd(y, m, d) {
    if (m < 1 || m > 12 || d < 1 || d > 31) return null;
    const t = Date.UTC(y, m - 1, d);
    const dt = new Date(t);
    if (dt.getUTCMonth() !== m - 1) return null;
    return t / DAY;
  }
  function parseDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : Math.round(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()) / DAY);
    if (typeof v === 'number') return v > 20000 && v < 80000 ? Math.round(v) - 25569 : null; // serial Excel
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$/);
    if (m) return ymd(+m[1], +m[2], +m[3]);
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return ymd(+m[3], +m[2], +m[1]);
    return null;
  }
  function addYearMinusDay(d) {
    const dt = new Date(d * DAY);
    return Date.UTC(dt.getUTCFullYear() + 1, dt.getUTCMonth(), dt.getUTCDate()) / DAY - 1;
  }
  function fmtDate(d) {
    if (d === null || d === undefined) return '';
    const dt = new Date(d * DAY);
    return String(dt.getUTCDate()).padStart(2, '0') + '/' + String(dt.getUTCMonth() + 1).padStart(2, '0') + '/' + dt.getUTCFullYear();
  }
  function todayDay() { const n = new Date(); return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) / DAY; }

  // ---------- Tiện ích ----------
  const norm = (s) => String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim();
  const clean = (s) => (s === null || s === undefined) ? '' : String(s).replace(/\s+/g, ' ').trim();
  const num = (v) => { if (v === null || v === undefined || v === '') return 0; const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, '')); return isFinite(n) ? n : 0; };
  const sum = (arr, f) => arr.reduce((a, x) => a + (f ? f(x) : x), 0);

  function findSheet(wb, name) {
    const k = wb.SheetNames.find((n) => norm(n) === norm(name));
    return k ? wb.Sheets[k] : null;
  }
  function sheetObjects(XLSX, wb, name) {
    const ws = findSheet(wb, name);
    if (!ws) return null;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    if (!rows.length) return [];
    const hdr = rows[0].map(norm);
    return rows.slice(1).map((r) => { const o = {}; hdr.forEach((h, i) => { if (h && !(h in o)) o[h] = r[i]; }); return o; });
  }
  const g = (o, ...names) => { for (const n of names) { const v = o[norm(n)]; if (v !== undefined) return v; } return null; };

  function uyFromFilename(fn) {
    const m = String(fn).match(/UY[\s_-]*(\d{1,4})/i);
    return m ? 'UY' + m[1] : String(fn).replace(/\.[^.]+$/, '').slice(0, 20);
  }
  // Ngày chốt số liệu (kỳ báo cáo) từ tên file: ..._update_05102026 → 05/10/2026 (ddmmyyyy), hoặc yyyymmdd / yyyy-mm-dd
  function snapFromFilename(fn) {
    const s = String(fn);
    let m = s.match(/update[_\s-]*(\d{2})(\d{2})(\d{4})/i);
    if (m) { const d = ymd(+m[3], +m[2], +m[1]); if (d !== null) return d; }
    m = s.match(/(20\d{2})[-_]?(\d{2})[-_]?(\d{2})/);
    if (m) { const d = ymd(+m[1], +m[2], +m[3]); if (d !== null) return d; }
    m = s.match(/(\d{2})(\d{2})(20\d{2})/);
    if (m) { const d = ymd(+m[3], +m[2], +m[1]); if (d !== null) return d; }
    return null;
  }
  function quarterLabel(d) {
    if (d === null || d === undefined) return '(không rõ)';
    const dt = new Date(d * DAY);
    return dt.getUTCFullYear() + '-Q' + (Math.floor(dt.getUTCMonth() / 3) + 1);
  }

  // ---------- Đọc một workbook ----------
  function readWorkbook(XLSX, wb, uy) {
    const warnings = [];
    const pol = sheetObjects(XLSX, wb, 'Policy');
    const risk = sheetObjects(XLSX, wb, 'Risk') || [];
    const prem = sheetObjects(XLSX, wb, 'Premium') || [];
    const clm = sheetObjects(XLSX, wb, 'Claim') || [];
    const inst = sheetObjects(XLSX, wb, 'Installment') || [];
    if (!pol) throw new Error('Không tìm thấy sheet "Policy" – file không đúng định dạng.');
    const hasMro = (r) => clean(g(r, 'MRO')) !== '';
    const P = pol.filter(hasMro), R = risk.filter(hasMro), M = prem.filter(hasMro);
    if (R.length !== P.length) warnings.push(`Sheet Risk có ${R.length} dòng, Policy có ${P.length} dòng.`);
    if (M.length !== P.length) warnings.push(`Sheet Premium có ${M.length} dòng, Policy có ${P.length} dòng.`);
    let misalign = 0;
    P.forEach((p, i) => { if ((R[i] && clean(g(R[i], 'MRO')) !== clean(g(p, 'MRO'))) || (M[i] && clean(g(M[i], 'MRO')) !== clean(g(p, 'MRO')))) misalign++; });
    if (misalign) warnings.push(`${misalign} dòng có MRO khác nhau giữa Policy/Risk/Premium – các sheet phải cùng thứ tự dòng.`);
    return { uy, policies: P.map((p, i) => ({ p, r: R[i] || {}, m: M[i] || {} })), claims: clm.filter(hasMro), installments: inst.filter(hasMro), warnings };
  }

  // ---------- Xây dựng mô hình ----------
  function build(books, userSettings) {
    const S = Object.assign({}, DEFAULT_SETTINGS, userSettings || {});
    const asOf = S.asOf == null ? todayDay() : S.asOf;
    const high = new Set(S.highRiskCountries.map(norm));
    const issues = [];
    const iss = (uy, sheet, rec, field, msg, sev, fix) => issues.push({ uy, sheet, rec: clean(rec).slice(0, 90), field, msg, sev, fix });
    const policies = [], claims = [], insts = [];
    const mroKey = {};

    // --- Hợp đồng
    for (const b of books) {
      const uy = b.uy; const used = {};
      const list = b.policies.map(({ p, r, m }) => {
        const ref = clean(g(p, 'PVI ref number', 'PVI ref'));
        let key = uy + '|' + (ref || clean(g(p, 'MRO')));
        const o = {
          uy, ref, mro: clean(g(p, 'MRO')), broker: clean(g(p, 'Broker')), insured: clean(g(p, 'Original Insured')),
          cedant: clean(g(p, 'Cedant Name')), region: clean(g(p, 'Region')), domicile: clean(g(p, 'Domicile')),
          channel: clean(g(p, 'Channel')) || clean(g(m, 'Facility')), leader: clean(g(p, 'Leader')), type: clean(g(p, 'Type')),
          riskType: clean(g(p, 'Risk type')), category: clean(g(p, 'Category')), status: clean(g(p, 'Status')) || clean(g(m, 'Policy status')),
          offerRaw: g(p, 'Risk offer date'), boundRaw: g(p, 'Risk bound date'), incRaw: g(p, 'Inception'), expRaw: g(p, 'Expiry'),
          share: num(g(r, 'PVI share')) / 100, hull: num(g(r, 'Hull')), csl: num(g(r, 'CSL')), ac: num(g(r, 'A/C')), afv: num(g(r, 'AFV')),
          gwp100: num(g(m, 'Deposit GWP')), gwp: num(g(m, 'PVI deposite GWP', 'PVI deposit GWP')), nwp: num(g(m, 'PVI deposite NWP', 'PVI deposit NWP')),
          ded: num(g(m, 'Total deduction')) / 100, ff: num(g(m, 'Facility fee')) / 100, od: num(g(m, 'Other Deduction')), fund: num(g(m, 'Fund')),
          earningSrc: g(m, 'PVI actual earning'), diffLead: num(g(m, 'Diffirental from lead', 'Differential from lead')), premStatus: clean(g(m, 'Premium status')),
        };
        o.region = S.regionMap[o.region] || o.region; o.leader = S.leaderMap[o.leader] || o.leader;
        o.offer = parseDate(o.offerRaw); o.bound = parseDate(o.boundRaw); o.inc = parseDate(o.incRaw); o.exp = parseDate(o.expRaw);
        o.earning = o.earningSrc === null || o.earningSrc === '' ? o.nwp * (1 - o.ff) - o.od - o.fund : num(o.earningSrc);
        if (used[key]) { used[key]++; key += '#' + used[key]; iss(uy, 'Policy', o.insured, 'PVI ref number', `Số tham chiếu PVI ${ref} bị trùng`, 'Cao', 'Cấp số tham chiếu riêng cho từng hợp đồng'); }
        else used[key] = 1;
        o.key = key;
        return o;
      });
      // khoá trùng: đánh dấu cả dòng đầu tiên
      Object.entries(used).filter(([, n]) => n > 1).forEach(([k]) => { const f = list.find((x) => x.key === k); if (f) iss(uy, 'Policy', f.insured, 'PVI ref number', `Số tham chiếu PVI ${f.ref} bị trùng`, 'Cao', 'Cấp số tham chiếu riêng cho từng hợp đồng'); });
      // MRO -> hợp đồng ưu tiên (Bound, không phải sửa đổi bổ sung)
      const byMro = {};
      list.forEach((x) => { (byMro[x.mro] = byMro[x.mro] || []).push(x); });
      mroKey[uy] = {};
      Object.entries(byMro).forEach(([mro, arr]) => {
        const pref = arr.find((x) => x.status === 'Bound' && !/Endt|\(AP/i.test(x.insured)) || arr.find((x) => x.status === 'Bound') || arr[0];
        mroKey[uy][mro] = pref.key;
        if (arr.length > 1) iss(uy, 'Policy', arr.map((x) => x.insured.slice(0, 30)).join(' / '), 'MRO', `MRO ${mro} xuất hiện ${arr.length} dòng (${arr.map((x) => x.status).join(', ')})`, 'TB', 'Dùng PVI ref làm khoá; sửa đổi bổ sung ghi ở bảng riêng');
      });
      policies.push(...list);
    }

    // --- Tính toán & kiểm tra từng hợp đồng
    for (const o of policies) {
      o.expC = o.exp !== null && o.inc !== null ? (o.exp <= o.inc ? addYearMinusDay(o.inc) : o.exp) : null;
      o.dur = o.expC !== null ? o.expC - o.inc + 1 : 0;
      o.onRisk = o.expC !== null ? Math.max(0, Math.min(asOf, o.expC) - o.inc + 1) : 0;
      o.earnPct = o.dur > 0 ? Math.min(1, o.onRisk / o.dur) : 0;
      o.ep = o.earning * o.earnPct;
      o.cap = o.share * (o.hull + o.csl);
      o.acq = o.gwp > 0 ? 1 - o.earning / o.gwp : null;
      o.acqUsd = o.gwp - o.earning;
      o.incQ = quarterLabel(o.inc);
      o.bound_ = o.status === 'Bound';
      o.highRisk = high.has(norm(o.domicile));
      o.smallOp = o.ac > 0 && o.ac < S.smallFleet;
      o.flags = [];
      if (!o.bound_) continue;
      const nm = o.insured, uy = o.uy;
      if (o.inc === null || o.exp === null) { iss(uy, 'Policy', nm, 'Inception/Expiry', `Ngày hiệu lực/hết hạn không đọc được ("${o.incRaw}" / "${o.expRaw}")`, 'Cao', 'Nhập lại dạng ngày'); continue; }
      const d = o.exp - o.inc;
      if (d <= 0) iss(uy, 'Policy', nm, 'Expiry', `Ngày hết hạn ${fmtDate(o.exp)} không sau ngày hiệu lực ${fmtDate(o.inc)}`, 'Cao', 'Sửa ngày hết hạn (app tạm dùng hiệu lực + 12 tháng)');
      else if (d > 400) iss(uy, 'Policy', nm, 'Expiry', `Thời hạn ${d + 1} ngày (>13 tháng)`, 'Thấp', 'Xác nhận thời hạn kéo dài là đúng');
      else if (d < 300 && !/Endt|\(AP/i.test(nm)) iss(uy, 'Policy', nm, 'Expiry', `Thời hạn chỉ ${d + 1} ngày`, 'Thấp', 'Xác nhận kỳ ngắn hạn là đúng');
      [['offer', 'Risk offer date'], ['bound', 'Risk bound date']].forEach(([k, lab]) => { if (o[k] === PLACEHOLDER) iss(uy, 'Policy', nm, lab, 'Ngày 01/01/2025 có vẻ là giá trị điền tạm', 'TB', 'Cập nhật ngày thực tế'); });
      if (o.bound !== null && o.bound !== PLACEHOLDER && o.bound > o.inc) {
        o.flags.push('Bind sau hiệu lực');
        iss(uy, 'Policy', nm, 'Risk bound date', `Bind ${fmtDate(o.bound)}, sau ngày hiệu lực ${fmtDate(o.inc)} (${o.bound - o.inc} ngày)`, 'Cao', 'Lấy xác nhận không tổn thất (NKLL/held covered) hoặc sửa ngày');
      }
      if (o.offer !== null && o.bound !== null && o.offer !== PLACEHOLDER && o.bound !== PLACEHOLDER && o.offer > o.bound) iss(uy, 'Policy', nm, 'Risk offer date', `Ngày chào (${fmtDate(o.offer)}) sau ngày bind (${fmtDate(o.bound)})`, 'TB', 'Kiểm tra lại thứ tự ngày');
      if (o.share <= 0) iss(uy, 'Risk', nm, 'PVI share', 'Hợp đồng Bound nhưng tỷ lệ PVI = 0', 'Cao', 'Nhập tỷ lệ tham gia thực tế');
      if (o.ac > 1000) iss(uy, 'Risk', nm, 'A/C, AFV', `Số tàu = ${Math.round(o.ac).toLocaleString('en')}, AFV = ${Math.round(o.afv).toLocaleString('en')}: có thể bị đảo cột`, 'TB', 'Đổi chỗ hai giá trị');
      else if (o.ac > 0 && o.afv / o.ac > 500e6) iss(uy, 'Risk', nm, 'AFV', `AFV ${(o.afv / 1e6).toFixed(0)} triệu USD cho ${o.ac} tàu: bất thường`, 'TB', 'Kiểm tra lại AFV');
      if (o.highRisk) { o.flags.push('Vùng nhạy cảm'); iss(uy, 'Policy', nm, 'Domicile', `Rủi ro tại ${o.domicile}: cần hồ sơ sàng lọc cấm vận và kiểm tra phạm vi lãnh thổ treaty`, 'TB', 'Lưu bằng chứng screening; đối chiếu loại trừ treaty'); }
      if (o.gwp100 && Math.abs(o.gwp100 * o.share * (1 - o.ded) - o.nwp) > 1) iss(uy, 'Premium', nm, 'PVI NWP', `NWP ${o.nwp.toFixed(2)} ≠ GWP × (1 − khấu trừ) = ${(o.gwp100 * o.share * (1 - o.ded)).toFixed(2)}`, 'TB', 'Kiểm tra % khấu trừ hoặc NWP');
      if (o.earningSrc !== null && o.earningSrc !== '' && Math.abs(o.nwp * (1 - o.ff) - o.od - o.fund - o.earning) > 1) iss(uy, 'Premium', nm, 'PVI actual earning', 'Net earning không khớp NWP × (1 − facility fee) − khấu trừ khác − quỹ', 'TB', 'Kiểm tra công thức');
      // cờ danh mục
      const cap = (o.smallOp || o.highRisk) ? S.lineCapSmall : S.lineCapGeneral;
      if (o.share > cap + 1e-9 && (o.hull > 0 || o.csl > 0)) o.flags.push(`Line ${(o.share * 100).toFixed(2)}% > trần ${(cap * 100).toFixed(2)}%`);
      if (o.gwp > 0 && o.gwp < S.minPremium) o.flags.push('Phí dưới mức tối thiểu');
    }
    const policyByKey = Object.fromEntries(policies.map((o) => [o.key, o]));

    // --- Bồi thường
    for (const b of books) {
      const uy = b.uy;
      const list = b.claims.map((c) => {
        const mro = clean(g(c, 'MRO'));
        const o = {
          uy, mro, key: mroKey[uy][mro] || null, claimNo: clean(g(c, 'Claim number')), broker: clean(g(c, 'Broker')), insured: clean(g(c, 'Original Insured')),
          coverage: clean(g(c, 'Coverage')), lossType: clean(g(c, 'Type of loss')), occRaw: g(c, 'Occurrence date'), repRaw: g(c, 'Date claim reported'),
          gross100: num(g(c, 'Gross Incurred')), share: num(g(c, 'PVI share')) / 100,
          paid: num(g(c, 'PVI Paid Claim')), paidFee: num(g(c, 'PVI Paid Fee')), res: num(g(c, 'PVI reserved claim')), resFee: num(g(c, 'PVI reserved Fee')),
          status: clean(g(c, 'Status')), paidStatus: clean(g(c, 'Paid Status')), details: clean(g(c, 'Loss details')),
        };
        o.occ = parseDate(o.occRaw); o.rep = parseDate(o.repRaw);
        o.incurred = o.paid + o.paidFee + o.res + o.resFee; o.paidT = o.paid + o.paidFee; o.os = o.res + o.resFee;
        o.flag = '';
        const p = o.key ? policyByKey[o.key] : null;
        o.policy = p;
        if (!p) iss(uy, 'Claim', o.insured, 'MRO', `Không tìm thấy hợp đồng có MRO ${mro}`, 'Cao', 'Kiểm tra MRO');
        if (o.occ === null) { o.flag = 'Ngày không hợp lệ'; iss(uy, 'Claim', o.insured, 'Occurrence date', `Ngày tổn thất "${clean(o.occRaw)}" không phải dạng ngày`, 'TB', 'Nhập ngày chuẩn'); }
        if (o.rep === null && clean(o.repRaw)) iss(uy, 'Claim', o.insured, 'Date claim reported', `Ngày thông báo "${clean(o.repRaw)}" không hợp lệ`, 'TB', 'Nhập ngày chuẩn');
        if (o.occ !== null && o.rep !== null && o.rep < o.occ) iss(uy, 'Claim', o.insured, 'Date claim reported', `Ngày thông báo ${fmtDate(o.rep)} trước ngày tổn thất ${fmtDate(o.occ)}`, 'TB', 'Kiểm tra lại ngày');
        if (p && o.occ !== null && p.inc !== null && p.expC !== null && (o.occ < p.inc || o.occ > p.expC)) {
          o.flag = 'Ngoài kỳ HĐ';
          const alt = policies.find((q) => q.bound_ && q.key !== p.key && q.inc !== null && q.expC !== null && q.inc <= o.occ && q.expC >= o.occ && norm(q.insured).slice(0, 12) === norm(o.insured).slice(0, 12));
          const sug = alt ? `→ thuộc hợp đồng ${alt.uy} (${alt.ref}, ${fmtDate(alt.inc)}–${fmtDate(alt.expC)})` : '→ không có hợp đồng nào đã tải lên hiệu lực vào ngày này';
          o.suggest = alt ? alt.uy : '';
          iss(uy, 'Claim', o.insured, 'Occurrence date', `Tổn thất ${fmtDate(o.occ)} ngoài kỳ hợp đồng ${fmtDate(p.inc)}–${fmtDate(p.expC)}, PVI incurred ${Math.round(o.incurred).toLocaleString('en')} ${sug}`, o.incurred > 0 ? 'Cao' : 'TB', 'Chuyển sang đúng hợp đồng/UY hoặc sửa ngày');
        }
        if ((o.coverage === 'Liability' && /^(Engine|Spare|Total loss)$/.test(o.lossType)) || (o.coverage === 'Hull' && /^(Human|Passenger)$/.test(o.lossType)))
          iss(uy, 'Claim', o.insured, 'Coverage / Type of loss', `Không khớp: ${o.coverage} + ${o.lossType}`, 'Thấp', 'Chuẩn hoá danh mục loại tổn thất');
        if (o.gross100 === 0 && o.incurred > 0) iss(uy, 'Claim', o.insured, 'Gross Incurred', `Gross 100% = 0 nhưng PVI incurred = ${o.incurred.toFixed(2)}`, 'TB', 'Bổ sung số 100%');
        if (o.share === 0 && o.gross100 > 0) iss(uy, 'Claim', o.insured, 'PVI share', `Tổn thất 100% ${Math.round(o.gross100).toLocaleString('en')} nhưng tỷ lệ PVI = 0`, 'Cao', 'Xác nhận PVI có tham gia không');
        return o;
      });
      const zeroOpen = list.filter((c) => c.incurred === 0 && /open/i.test(c.status)).length;
      if (zeroOpen) iss(uy, 'Claim', `${zeroOpen} hồ sơ`, 'Status', `${zeroOpen}/${list.length} hồ sơ đang mở nhưng incurred = 0`, 'TB', 'Đặt dự phòng ban đầu hoặc đóng hồ sơ');
      const paidZero = list.filter((c) => /paid/i.test(c.paidStatus) && c.paidT === 0).length;
      if (paidZero) iss(uy, 'Claim', `${paidZero} hồ sơ`, 'Paid Status', `${paidZero} hồ sơ ghi "Paid" nhưng số PVI đã trả = 0`, 'Thấp', 'Thống nhất định nghĩa trạng thái');
      const noNo = list.filter((c) => !c.claimNo).length;
      if (noNo) iss(uy, 'Claim', `${noNo} hồ sơ`, 'Claim number', `${noNo} hồ sơ chưa có số bồi thường`, 'TB', 'Cấp số hồ sơ');
      claims.push(...list);
    }

    // --- Kỳ thanh toán (bảng ngang → dọc)
    for (const b of books) {
      const uy = b.uy; const list = [];
      b.installments.forEach((x) => {
        const mro = clean(g(x, 'MRO'));
        for (let n = 1; n <= 24; n++) {
          const amtRaw = g(x, `Inst ${n} Amount`);
          if (amtRaw === null || amtRaw === undefined) { if (n > 12) break; continue; }
          const amount = num(amtRaw); if (!amount) continue;
          const due = parseDate(g(x, `Inst ${n} Due Date`));
          const paid = clean(g(x, `Inst ${n} Paid`));
          const o = { uy, mro, key: mroKey[uy][mro] || null, broker: clean(g(x, 'Broker')), insured: clean(g(x, 'Original Insured')), n, amount, due, paid, actual: num(g(x, `Inst ${n} Actual NWP`)) };
          o.isPaid = /^paid$/i.test(paid);
          o.days = !o.isPaid && due !== null && due < asOf ? asOf - due : 0;
          o.bucket = o.days === 0 ? '' : o.days <= 30 ? '0-30' : o.days <= 90 ? '31-90' : o.days <= 180 ? '91-180' : o.days <= 365 ? '181-365' : 'Trên 365';
          // Thực thu: số thực thu ghi nhận; kỳ ghi "Paid" mà chưa nhập số thực thu thì lấy số tiền kỳ
          o.collected = o.isPaid ? (o.actual > 0 ? o.actual : amount) : 0;
          o.isDue = due !== null && due <= asOf;
          o.daysToDue = !o.isPaid && due !== null && due >= asOf ? due - asOf : null;
          list.push(o);
        }
      });
      const pa0 = list.filter((i) => i.isPaid && !i.actual).length;
      if (pa0) iss(uy, 'Installment', `${pa0} kỳ`, 'Actual NWP', `${pa0} kỳ ghi "Paid" nhưng Actual NWP = 0`, 'TB', 'Ghi số thực thu');
      const bound = policies.filter((p) => p.uy === uy && p.bound_);
      const withSched = new Set(list.map((i) => i.key));
      const miss = bound.filter((p) => !withSched.has(p.key)).length;
      if (miss) iss(uy, 'Installment', `${miss} hợp đồng`, 'Lịch thanh toán', `${miss}/${bound.length} hợp đồng Bound chưa có lịch thanh toán; tổng kỳ phí ${Math.round(sum(list, (i) => i.amount)).toLocaleString('en')} so với NWP ${Math.round(sum(bound, (p) => p.nwp)).toLocaleString('en')}`, 'Cao', 'Lập lịch thanh toán cho mọi hợp đồng');
      insts.push(...list);
    }

    // --- Gắn bồi thường vào hợp đồng
    policies.forEach((p) => { p.incurred = 0; p.paidT = 0; p.os = 0; p.nClaims = 0; });
    claims.forEach((c) => { if (c.policy) { c.policy.incurred += c.incurred; c.policy.paidT += c.paidT; c.policy.os += c.os; c.policy.nClaims++; } });
    policies.forEach((p) => { p.lr = p.ep > 0 ? p.incurred / p.ep : null; p.lossMultiple = p.earning > 0 ? p.incurred / p.earning : null; });

    // --- Tập trung leader
    const uys = [...new Set(books.map((b) => b.uy))].sort();
    uys.forEach((uy) => {
      const bp = policies.filter((p) => p.uy === uy && p.bound_);
      const tot = sum(bp, (p) => p.gwp);
      const by = groupSum(bp, (p) => p.leader || '(trống)', (p) => p.gwp);
      Object.entries(by).forEach(([l, v]) => { if (tot > 0 && v / tot > S.leaderCap) iss(uy, 'Danh mục', l, 'Leader', `Leader ${l} chiếm ${(v / tot * 100).toFixed(1)}% GWP (trần ${(S.leaderCap * 100).toFixed(0)}%)`, 'TB', 'Đa dạng hoá leader'); });
    });

    const sevOrder = { 'Cao': 0, 'TB': 1, 'Thấp': 2 };
    issues.sort((a, b) => sevOrder[a.sev] - sevOrder[b.sev] || a.uy.localeCompare(b.uy) || a.sheet.localeCompare(b.sheet));
    return { settings: S, asOf, uys, policies, claims, insts, issues };
  }

  function groupSum(arr, kf, vf) { const o = {}; arr.forEach((x) => { const k = kf(x); o[k] = (o[k] || 0) + vf(x); }); return o; }

  // ---------- Tổng hợp ----------
  function summarize(model, filter) {
    const f = filter || (() => true);
    const bp = model.policies.filter((p) => p.bound_ && f(p));
    const keys = new Set(bp.map((p) => p.key));
    const cl = model.claims.filter((c) => c.policy && keys.has(c.key));
    const ins = model.insts.filter((i) => keys.has(i.key));
    const s = {
      n: bp.length, gwp: sum(bp, (p) => p.gwp), nwp: sum(bp, (p) => p.nwp), earning: sum(bp, (p) => p.earning), ep: sum(bp, (p) => p.ep),
      incurred: sum(cl, (c) => c.incurred), paid: sum(cl, (c) => c.paidT), os: sum(cl, (c) => c.os), nClaims: cl.length,
      nZero: cl.filter((c) => c.incurred === 0).length, cap: sum(bp, (p) => p.cap),
      overdue: sum(ins.filter((i) => i.days > 0), (i) => i.amount), overdue90: sum(ins.filter((i) => i.days > model.settings.overdueAlertDays), (i) => i.amount),
      boundLate: bp.filter((p) => p.flags.includes('Bind sau hiệu lực')).length,
      cancelled: model.policies.filter((p) => f(p) && /cancel/i.test(p.status)).length,
    };
    // Thu phí
    const up = model.settings.upcomingDays || 30;
    const sched = new Set(ins.map((i) => i.key));
    s.schedAmt = sum(ins, (i) => i.amount);
    s.dueAmt = sum(ins.filter((i) => i.isDue), (i) => i.amount);
    s.collected = sum(ins, (i) => i.collected);
    s.collectedPctGwp = s.gwp > 0 ? s.collected / s.gwp : null;
    s.collectedPctNwp = s.nwp > 0 ? s.collected / s.nwp : null;
    s.collectedPctDue = s.dueAmt > 0 ? sum(ins.filter((i) => i.isDue), (i) => i.collected) / s.dueAmt : null;
    s.upcoming = sum(ins.filter((i) => i.daysToDue !== null && i.daysToDue <= up), (i) => i.amount);
    s.nUpcoming = ins.filter((i) => i.daysToDue !== null && i.daysToDue <= up).length;
    s.nOverdue = ins.filter((i) => i.days > 0).length;
    s.overduePctNwp = s.nwp > 0 ? s.overdue / s.nwp : null;
    s.noSchedN = bp.filter((p) => !sched.has(p.key)).length;
    s.noSchedNwp = sum(bp.filter((p) => !sched.has(p.key)), (p) => p.nwp);
    s.acq = s.gwp > 0 ? 1 - s.earning / s.gwp : null;
    s.earnedPct = s.earning > 0 ? s.ep / s.earning : null;
    s.lrEP = s.ep > 0 ? s.incurred / s.ep : null;
    s.lrFull = s.earning > 0 ? s.incurred / s.earning : null;
    // Chi phí khai thác = GWP − net earning (brokerage, SMB, facility fee, khấu trừ khác, quỹ)
    s.acqUsd = s.gwp - s.earning;
    s.brokSmbUsd = s.gwp - s.nwp;          // phần trừ từ GWP xuống NWP (brokerage + SMB …)
    s.otherCostUsd = s.nwp - s.earning;    // facility fee, khấu trừ khác, quỹ
    s.acqPct = s.gwp > 0 ? s.acqUsd / s.gwp : null;
    s.paidPct = s.gwp > 0 ? s.paid / s.gwp : null;
    s.osPct = s.gwp > 0 ? s.os / s.gwp : null;
    s.claimsPct = s.gwp > 0 ? s.incurred / s.gwp : null;
    s.combinedUsd = s.acqUsd + s.incurred;
    s.combinedPct = s.gwp > 0 ? s.combinedUsd / s.gwp : null;
    s.resultUsd = s.gwp - s.combinedUsd;   // kết quả trước chi phí quản lý, IBNR và tái bảo hiểm
    s.combined = s.combinedPct;
    const top2 = bp.map((p) => p.incurred).sort((a, b) => b - a).slice(0, 2);
    s.lrExTop2 = s.ep > 0 ? (s.incurred - sum(top2)) / s.ep : null;
    return s;
  }

  // So sánh theo một chiều (broker, channel, region…). baseFilter giới hạn phạm vi (ví dụ một UY).
  function segment(model, dim, baseFilter) {
    const bf = baseFilter || (() => true);
    const bp = model.policies.filter((p) => p.bound_ && bf(p));
    const cats = [...new Set(bp.map((p) => p[dim] || '(trống)'))];
    const all = summarize(model, bf);
    const rows = cats.map((cat) => {
      const r = { cat, byUy: {} };
      model.uys.forEach((uy) => { r.byUy[uy] = summarize(model, (p) => bf(p) && p.uy === uy && (p[dim] || '(trống)') === cat); });
      r.total = summarize(model, (p) => bf(p) && (p[dim] || '(trống)') === cat);
      r.shareGwp = all.gwp > 0 ? r.total.gwp / all.gwp : null;
      return r;
    });
    rows.sort((a, b) => b.total.gwp - a.total.gwp);
    return { rows, all };
  }

  // ---------- Tái tục ----------
  const STOP = new Set(['pt', 'the', 'co', 'ltd', 'llc', 'jsc', 'inc', 'limited', 'company', 'and', 'or', 'of', 'a', 'b', 'c', 'd', 'i', 'ii', 'iii', 'pvt', 'plc', 'sa', 'corp', 'corporation', 'group', 'ojsc', 'cjsc', 'fze', 'fzc', 'tbk', 'persero']);
  function insuredKey(name) {
    const w = norm(name).replace(/\([^)]{0,6}\)/g, ' ').replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter((x) => x && !STOP.has(x));
    return w.slice(0, 2).join(' ');
  }
  // Hợp đồng hết hạn trong [ngày báo cáo − lookback, ngày báo cáo + days]; tự nhận diện hợp đồng tái tục đã có trong dữ liệu.
  function renewals(model, days, lookback) {
    const base = model.asOf; const lb = lookback == null ? 30 : lookback;
    const bound = model.policies.filter((p) => p.bound_ && p.expC !== null);
    const list = bound.filter((p) => p.expC >= base - lb && p.expC <= base + days && !/Endt|\(AP /i.test(p.insured));
    return list.map((p) => {
      const k = insuredKey(p.insured);
      const ren = bound.find((q) => q !== p && q.inc !== null && insuredKey(q.insured) === k && q.inc >= p.expC - 45 && q.inc <= p.expC + 45);
      const daysLeft = p.expC - base;
      const lr = p.ep > 0 ? p.incurred / p.ep : null;
      const comb = p.gwp > 0 ? (p.acqUsd + p.incurred) / p.gwp : null;
      const notes = [];
      if (comb !== null && comb > 1) notes.push('Kết hợp > 100%: xem lại phí/điều kiện');
      else if (comb !== null && comb > 0.8) notes.push('Kết hợp > 80%');
      if (p.flags.some((f) => f.startsWith('Line '))) notes.push('Line vượt trần');
      if (p.highRisk) notes.push('Vùng nhạy cảm: sàng lọc lại cấm vận');
      const od = model.insts.filter((i) => i.key === p.key && i.days > 0).reduce((a, i) => a + i.amount, 0);
      if (od > 0) notes.push(`Còn ${Math.round(od).toLocaleString('en')} USD phí quá hạn`);
      return {
        p, daysLeft, renewed: ren || null, lr, comb,
        status: ren ? 'Đã có HĐ tái tục' : daysLeft < 0 ? 'Đã hết hạn, chưa thấy tái tục' : daysLeft <= 30 ? 'Trong 30 ngày' : daysLeft <= 60 ? 'Trong 31–60 ngày' : 'Trong 61–' + days + ' ngày',
        notes,
      };
    }).sort((a, b) => a.daysLeft - b.daysLeft);
  }

  return { DEFAULT_SETTINGS, readWorkbook, build, summarize, segment, renewals, insuredKey, parseDate, fmtDate, todayDay, uyFromFilename, snapFromFilename, quarterLabel, DAY };
});
