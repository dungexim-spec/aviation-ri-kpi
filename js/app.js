/* Giao diện: tải file → KpiEngine.build → hiển thị các tab báo cáo. Không gửi dữ liệu ra ngoài. */
(function () {
  'use strict';
  const E = window.KpiEngine;
  const $ = (s, r) => (r || document).querySelector(s);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fN = (v) => (v == null || !isFinite(v)) ? '–' : Math.round(v).toLocaleString('vi-VN');
  const fN2 = (v) => (v == null || !isFinite(v)) ? '–' : v.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fP = (v, d = 1) => (v == null || !isFinite(v)) ? '–' : (v * 100).toFixed(d).replace('.', ',') + '%';
  const fD = (d) => E.fmtDate(d);
  const SERIES = ['--s1', '--s2', '--s3', '--s4', '--s5', '--s6', '--s7', '--s8'];

  const state = { books: [], files: [], model: null, tab: 'overview', settings: loadSettings(), segDim: 'broker', claimF: { uy: '', st: '', flag: '' }, issueF: { sev: '', uy: '', q: '' } };

  // ---------- Thiết lập (lưu trong trình duyệt, chỉ là tham số – không lưu dữ liệu) ----------
  function loadSettings() {
    const d = Object.assign({}, E.DEFAULT_SETTINGS, { asOf: E.todayDay() });
    try { const s = JSON.parse(localStorage.getItem('kpi-settings') || 'null'); if (s) Object.assign(d, s, { asOf: d.asOf }); } catch (e) { /* bỏ qua */ }
    return d;
  }
  function saveSettings() {
    try { const s = Object.assign({}, state.settings); delete s.asOf; localStorage.setItem('kpi-settings', JSON.stringify(s)); } catch (e) { /* bỏ qua */ }
  }
  function uyColor(uy) { const m = state.model; const i = m ? m.uys.indexOf(uy) : 0; return `var(${SERIES[Math.max(0, i) % SERIES.length]})`; }
  function toast(msg) { const t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toast._t); toast._t = setTimeout(() => { t.hidden = true; }, 3200); }

  // ---------- Tải file ----------
  async function addFiles(fileList) {
    const warns = [];
    for (const f of fileList) {
      try {
        const buf = await f.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        let uy = E.uyFromFilename(f.name);
        const book = E.readWorkbook(XLSX, wb, uy);
        book.file = f.name;
        state.books.push(book);
        book.warnings.forEach((w) => warns.push(`${f.name}: ${w}`));
      } catch (err) { warns.push(`${f.name}: ${err.message || err}`); }
    }
    $('#loadWarn').innerHTML = warns.map((w) => `<div class="warnline">⚠ ${esc(w)}</div>`).join('');
    recompute();
  }
  function renderFiles() {
    $('#files').innerHTML = state.books.map((b, i) => `<span class="chip"><span class="sw" style="background:${uyColor(b.uy)}"></span><input aria-label="Năm nghiệp vụ của ${esc(b.file)}" data-i="${i}" value="${esc(b.uy)}"><span class="fn" title="${esc(b.file)}">${esc(b.file)}</span><span class="muted">${b.policies.length} HĐ</span><button class="x" data-rm="${i}" aria-label="Bỏ file ${esc(b.file)}">×</button></span>`).join('');
  }
  function recompute() {
    state.model = state.books.length ? E.build(state.books, state.settings) : null;
    $('#export').disabled = !state.model;
    renderFiles(); renderTabs(); render();
  }

  // ---------- Tabs ----------
  const TABS = [
    ['overview', 'Tổng quan'], ['segments', 'Phân khúc'], ['claims', 'Bồi thường'], ['collections', 'Thu phí'],
    ['risk', 'Cỡ line & tập trung'], ['issues', 'Kiểm tra dữ liệu'], ['settings', 'Thiết lập'], ['guide', 'Hướng dẫn'],
  ];
  function renderTabs() {
    const m = state.model; const nCrit = m ? m.issues.filter((i) => i.sev === 'Cao').length : 0;
    $('#tabs').innerHTML = TABS.map(([k, l]) => `<button role="tab" aria-selected="${state.tab === k}" data-tab="${k}">${l}${k === 'issues' && nCrit ? `<span class="count">${nCrit}</span>` : ''}</button>`).join('');
  }

  // ---------- Bảng & biểu đồ ----------
  function table(cols, rows, opts = {}) {
    const id = 't' + Math.random().toString(36).slice(2, 8);
    const sort = opts.sort; let data = rows.slice();
    if (sort) { const c = cols[sort.i]; data.sort((a, b) => { const x = c.v(a), y = c.v(b); return (x > y ? 1 : x < y ? -1 : 0) * sort.dir; }); }
    const head = cols.map((c, i) => `<th class="${c.num ? 'r' : ''} sortable" data-sort="${i}">${esc(c.l)}</th>`).join('');
    const body = data.map((r) => `<tr class="${opts.rowClass ? opts.rowClass(r) : ''}">${cols.map((c) => `<td class="${c.num ? 'r' : ''} ${c.wrap ? 'tw' : ''}">${c.h ? c.h(r) : esc(c.f ? c.f(c.v(r)) : c.v(r))}</td>`).join('')}</tr>`).join('');
    const foot = opts.total ? `<tr class="total">${cols.map((c, i) => `<td class="${c.num ? 'r' : ''}">${i === 0 ? 'Tổng' : (c.t ? esc(c.t(rows)) : '')}</td>`).join('')}</tr>` : '';
    TABLES[id] = { cols, rows, opts };
    return `<div class="tbl ${opts.tall ? 'tall' : ''}" id="${id}"><table><thead><tr>${head}</tr></thead><tbody>${body || `<tr><td colspan="${cols.length}" class="muted">Không có dòng nào.</td></tr>`}${foot}</tbody></table></div>`;
  }
  const TABLES = {};
  function bars(rows, uys, max, fmt) {
    if (!rows.length) return '<p class="muted">Chưa có dữ liệu.</p>';
    return `<div class="bars">${rows.map((r) => `<div class="cat" title="${esc(r.cat)}">${esc(r.cat)}</div><div class="track">${uys.map((uy) => { const v = r.vals[uy] || 0; const w = max > 0 ? Math.max(0.3, Math.min(100, v / max * 100)) * 0.8 : 0; return `<div class="bline"><div class="bar" style="width:${w}%;background:${uyColor(uy)}" title="${esc(r.cat)} · ${esc(uy)}: ${esc(fmt(v))}"></div><span class="val">${esc(fmt(v))}</span></div>`; }).join('')}</div>`).join('')}</div>`;
  }
  const legend = (uys) => `<div class="legend">${uys.map((u) => `<span><span class="sw" style="background:${uyColor(u)}"></span>${esc(u)}</span>`).join('')}</div>`;
  const sevPill = (s) => `<span class="pill ${s === 'Cao' ? 'crit' : s === 'TB' ? 'warn' : 'good'}">● ${s === 'TB' ? 'Trung bình' : esc(s)}</span>`;

  // ---------- Views ----------
  function render() {
    const v = $('#view'); const m = state.model;
    if (!m && !['settings', 'guide'].includes(state.tab)) { v.innerHTML = emptyView(); return; }
    v.innerHTML = ({ overview, segments, claims, collections, risk, issues, settings, guide })[state.tab]();
  }

  function emptyView() {
    return `<div class="panel empty"><h2>Bắt đầu</h2>
      <ol><li>Bấm <b>Chọn file Excel</b> và chọn một hoặc nhiều file theo dõi, ví dụ <code>UY_1_update_05102026.xlsx</code>, <code>UY_2_update_05102026.xlsx</code>.</li>
      <li>App tự nhận năm nghiệp vụ từ tên file (<code>UY_1</code> → UY1). Có thể sửa nhãn ngay trên thẻ file.</li>
      <li>Chọn <b>Ngày báo cáo</b>: mọi chỉ số phí đã hưởng, loss ratio, tuổi nợ tính đến ngày này.</li>
      <li>Xem các tab báo cáo, hoặc bấm <b>Xuất báo cáo Excel</b>.</li></ol>
      <p class="muted small">File cần các sheet <code>Policy</code>, <code>Risk</code>, <code>Premium</code>, <code>Claim</code>, <code>Installment</code> với tiêu đề cột như file theo dõi hiện tại. Bấm <b>Tải file mẫu</b> để lấy khung trống.</p></div>`;
  }

  function overview() {
    const m = state.model; const uys = m.uys;
    const S = {}; uys.forEach((u) => { S[u] = E.summarize(m, (p) => p.uy === u); }); const T = E.summarize(m);
    const kpi = (lab, f, fmt) => `<div class="panel kpi"><div class="lab">${lab}</div><div class="vals">${uys.map((u) => `<div class="v"><span class="sw" style="background:${uyColor(u)}"></span><span class="u">${esc(u)}</span>${fmt(f(S[u]))}</div>`).join('')}</div></div>`;
    const rowsDef = [
      ['Số hợp đồng Bound', (s) => s.n, (v) => fN(v)], ['Hợp đồng huỷ', (s) => s.cancelled, fN], ['PVI GWP', (s) => s.gwp, fN], ['PVI NWP', (s) => s.nwp, fN],
      ['PVI net earning', (s) => s.earning, fN], ['Chi phí khai thác % GWP', (s) => s.acq, fP], ['Phí đã hưởng', (s) => s.ep, fN], ['% đã hưởng', (s) => s.earnedPct, fP],
      ['PVI incurred', (s) => s.incurred, fN], ['  đã trả', (s) => s.paid, fN], ['  dự phòng', (s) => s.os, fN], ['Loss ratio / phí đã hưởng', (s) => s.lrEP, fP],
      ['Loss ratio / net earning cả năm', (s) => s.lrFull, fP], ['LR bỏ 2 vụ lớn nhất', (s) => s.lrExTop2, fP], ['(Incurred + chi phí KT) / GWP', (s) => s.combined, fP],
      ['Số hồ sơ bồi thường', (s) => s.nClaims, fN], ['  incurred = 0', (s) => s.nZero, fN], ['Tổng capacity PVI', (s) => s.cap, fN],
      ['Kỳ phí quá hạn', (s) => s.overdue, fN], [`  quá hạn > ${m.settings.overdueAlertDays} ngày`, (s) => s.overdue90, fN], ['Bind sau hiệu lực', (s) => s.boundLate, fN],
    ];
    const mt = `<div class="tbl"><table><thead><tr><th>Chỉ tiêu</th>${uys.map((u) => `<th class="r">${esc(u)}</th>`).join('')}${uys.length > 1 ? '<th class="r">Tổng</th>' : ''}</tr></thead><tbody>${rowsDef.map(([l, f, fmt]) => `<tr><td>${esc(l)}</td>${uys.map((u) => `<td class="r">${fmt(f(S[u]))}</td>`).join('')}${uys.length > 1 ? `<td class="r">${fmt(f(T))}</td>` : ''}</tr>`).join('')}</tbody></table></div>`;
    // điểm cần chú ý
    const ins = [];
    uys.forEach((u) => {
      const s = S[u];
      if (s.lrEP != null && s.lrEP > 1) ins.push(['Cao', `${u}: loss ratio trên phí đã hưởng ${fP(s.lrEP)} (đã hưởng ${fP(s.earnedPct, 0)} phí).`]);
      else if (s.lrEP != null && s.lrEP > 0.65) ins.push(['TB', `${u}: loss ratio trên phí đã hưởng ${fP(s.lrEP)}.`]);
      const bp = m.policies.filter((p) => p.uy === u && p.bound_).sort((a, b) => b.incurred - a.incurred);
      const top2 = bp.slice(0, 2).filter((p) => p.incurred > 0);
      if (s.incurred > 0 && top2.length && (top2[0].incurred + (top2[1] ? top2[1].incurred : 0)) / s.incurred > 0.6)
        ins.push(['TB', `${u}: ${top2.map((p) => `${p.insured.slice(0, 32)} (${fN(p.incurred)})`).join(' và ')} chiếm ${fP((top2[0].incurred + (top2[1] ? top2[1].incurred : 0)) / s.incurred, 0)} incurred. Bỏ 2 vụ này LR còn ${fP(s.lrExTop2)}.`]);
      if (s.overdue90 > 0) ins.push(['Cao', `${u}: ${fN(s.overdue90)} USD kỳ phí quá hạn trên ${m.settings.overdueAlertDays} ngày chưa ghi nhận thu.`]);
      if (s.boundLate) ins.push(['Cao', `${u}: ${s.boundLate} hợp đồng bind sau ngày hiệu lực.`]);
      const over = bp.filter((p) => p.flags.some((f) => f.startsWith('Line '))).length;
      if (over) ins.push(['TB', `${u}: ${over} hợp đồng có line vượt trần (xem tab Cỡ line & tập trung).`]);
    });
    m.issues.filter((i) => i.sheet === 'Danh mục').forEach((i) => ins.push(['TB', `${i.uy}: ${i.msg}.`]));
    const nCrit = m.issues.filter((i) => i.sev === 'Cao').length;
    if (nCrit) ins.push(['Cao', `${nCrit} lỗi dữ liệu mức cao (tab Kiểm tra dữ liệu).`]);
    const ord = { 'Cao': 0, 'TB': 1, 'Thấp': 2 }; ins.sort((a, b) => ord[a[0]] - ord[b[0]]);
    const lrRows = [{ cat: 'LR / phí đã hưởng', vals: Object.fromEntries(uys.map((u) => [u, S[u].lrEP || 0])) }, { cat: 'LR / net earning cả năm', vals: Object.fromEntries(uys.map((u) => [u, S[u].lrFull || 0])) }, { cat: 'Chi phí khai thác', vals: Object.fromEntries(uys.map((u) => [u, S[u].acq || 0])) }];
    const lrMax = Math.max(1, ...lrRows.flatMap((r) => Object.values(r.vals)));
    const top = m.policies.filter((p) => p.bound_ && p.incurred > 0).sort((a, b) => b.incurred - a.incurred).slice(0, 10);
    return `<div class="grid kpis">
        ${kpi('PVI GWP', (s) => s.gwp, fN)}${kpi('Phí đã hưởng', (s) => s.ep, fN)}${kpi('PVI incurred', (s) => s.incurred, fN)}${kpi('Loss ratio / phí đã hưởng', (s) => s.lrEP, (v) => fP(v))}
      </div>
      <div class="grid two section">
        <div class="panel"><h2>Điểm cần chú ý</h2><ul class="insights" style="margin-top:10px">${ins.length ? ins.map(([s, t]) => `<li>${sevPill(s)}<span>${esc(t)}</span></li>`).join('') : '<li><span class="pill good">● Ổn</span><span>Không có cảnh báo theo ngưỡng hiện tại.</span></li>'}</ul></div>
        <div class="panel"><h2>Tỷ lệ chính theo UY</h2>${legend(uys)}${bars(lrRows, uys, lrMax, (v) => fP(v))}<div class="refline">Tính đến ${fD(m.asOf)}. Chưa gồm IBNR.</div></div>
      </div>
      <div class="section"><h2>Kết quả theo năm nghiệp vụ</h2>${mt}</div>
      <div class="section"><h2>10 hợp đồng tổn thất lớn nhất</h2>${table([
        { l: 'UY', v: (p) => p.uy }, { l: 'Original insured', v: (p) => p.insured, wrap: true }, { l: 'Broker', v: (p) => p.broker }, { l: 'Leader', v: (p) => p.leader },
        { l: 'Line', v: (p) => p.share, f: (v) => fP(v, 2), num: true }, { l: 'Net earning', v: (p) => p.earning, f: fN, num: true }, { l: 'Incurred', v: (p) => p.incurred, f: fN, num: true },
        { l: 'Incurred / phí', v: (p) => p.lossMultiple, f: (v) => v == null ? '–' : v.toFixed(1).replace('.', ',') + '×', num: true }], top)}</div>`;
  }

  const DIMS = [['broker', 'Broker'], ['channel', 'Kênh'], ['region', 'Vùng'], ['leader', 'Leader'], ['domicile', 'Quốc gia'], ['cedant', 'Cedant'], ['type', 'Loại (RI/XOL…)'], ['riskType', 'Risk type']];
  function segments() {
    const m = state.model; const uys = m.uys; const dim = state.segDim; const rows = E.segment(m, dim);
    const top = rows.slice(0, 12).map((r) => ({ cat: r.cat, vals: Object.fromEntries(uys.map((u) => [u, r.byUy[u].gwp])) }));
    const max = Math.max(0, ...top.flatMap((r) => Object.values(r.vals)));
    const cols = [{ l: DIMS.find((d) => d[0] === dim)[1], v: (r) => r.cat, wrap: true }];
    uys.forEach((u) => {
      cols.push({ l: `GWP ${u}`, v: (r) => r.byUy[u].gwp, f: fN, num: true, t: (rs) => fN(rs.reduce((a, r) => a + r.byUy[u].gwp, 0)) });
      cols.push({ l: `HĐ ${u}`, v: (r) => r.byUy[u].n, f: fN, num: true, t: (rs) => fN(rs.reduce((a, r) => a + r.byUy[u].n, 0)) });
      cols.push({ l: `Chi phí KT ${u}`, v: (r) => r.byUy[u].acq, f: fP, num: true });
      cols.push({ l: `Incurred ${u}`, v: (r) => r.byUy[u].incurred, f: fN, num: true, t: (rs) => fN(rs.reduce((a, r) => a + r.byUy[u].incurred, 0)) });
      cols.push({ l: `LR đã hưởng ${u}`, v: (r) => r.byUy[u].lrEP, f: fP, num: true });
    });
    cols.push({ l: 'Tỷ trọng GWP', v: (r) => r.total.gwp / (E.summarize(m).gwp || 1), f: fP, num: true });
    return `<div class="toolbar"><label class="fld" for="segDim">Phân theo<select id="segDim">${DIMS.map(([k, l]) => `<option value="${k}" ${k === dim ? 'selected' : ''}>${l}</option>`).join('')}</select></label></div>
      <div class="panel"><h2>PVI GWP theo ${esc(DIMS.find((d) => d[0] === dim)[1].toLowerCase())}${rows.length > 12 ? ' (12 nhóm lớn nhất)' : ''}</h2>${legend(uys)}${bars(top, uys, max, fN)}</div>
      <div class="section">${table(cols, rows, { total: true })}</div>`;
  }

  function claims() {
    const m = state.model; const F = state.claimF;
    let rows = m.claims.filter((c) => (!F.uy || c.uy === F.uy) && (!F.st || c.status === F.st) && (!F.flag || (F.flag === 'any' ? c.flag : c.flag === F.flag)));
    const sts = [...new Set(m.claims.map((c) => c.status).filter(Boolean))];
    const byCov = {}; rows.forEach((c) => { const k = (c.coverage || '(trống)') + ' · ' + (c.lossType || '(trống)'); byCov[k] = byCov[k] || { cat: k, n: 0, inc: 0 }; byCov[k].n++; byCov[k].inc += c.incurred; });
    return `<div class="toolbar">
        <label class="fld" for="cUy">UY<select id="cUy"><option value="">Tất cả</option>${m.uys.map((u) => `<option ${F.uy === u ? 'selected' : ''}>${esc(u)}</option>`).join('')}</select></label>
        <label class="fld" for="cSt">Trạng thái<select id="cSt"><option value="">Tất cả</option>${sts.map((s) => `<option ${F.st === s ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></label>
        <label class="fld" for="cFl">Kiểm tra kỳ<select id="cFl"><option value="">Tất cả</option><option value="any" ${F.flag === 'any' ? 'selected' : ''}>Có cảnh báo</option><option ${F.flag === 'Ngoài kỳ HĐ' ? 'selected' : ''}>Ngoài kỳ HĐ</option><option ${F.flag === 'Ngày không hợp lệ' ? 'selected' : ''}>Ngày không hợp lệ</option></select></label>
        <span class="muted small">${rows.length} hồ sơ · incurred ${fN(rows.reduce((a, c) => a + c.incurred, 0))} USD</span></div>
      ${table([
        { l: 'UY', v: (c) => c.uy }, { l: 'Số hồ sơ', v: (c) => c.claimNo }, { l: 'Original insured', v: (c) => c.insured, wrap: true }, { l: 'Coverage', v: (c) => c.coverage }, { l: 'Loại', v: (c) => c.lossType },
        { l: 'Ngày tổn thất', v: (c) => c.occ ?? -1, h: (c) => c.occ != null ? fD(c.occ) : esc(c.occRaw) }, { l: 'Gross 100%', v: (c) => c.gross100, f: fN, num: true }, { l: 'Line', v: (c) => c.share, f: (v) => fP(v, 2), num: true },
        { l: 'Đã trả', v: (c) => c.paidT, f: fN2, num: true }, { l: 'Dự phòng', v: (c) => c.os, f: fN2, num: true }, { l: 'Incurred', v: (c) => c.incurred, f: fN2, num: true }, { l: 'Status', v: (c) => c.status },
        { l: 'Kiểm tra', v: (c) => c.flag, h: (c) => c.flag ? `<span class="pill crit">${esc(c.flag)}</span>${c.suggest ? ` <span class="muted small">→ ${esc(c.suggest)}</span>` : ''}` : '' }],
        rows, { tall: true, sort: { i: 10, dir: -1 } })}
      <div class="section"><h2>Theo nghiệp vụ và loại tổn thất</h2>${table([{ l: 'Nhóm', v: (r) => r.cat }, { l: 'Số hồ sơ', v: (r) => r.n, f: fN, num: true }, { l: 'Incurred', v: (r) => r.inc, f: fN, num: true }], Object.values(byCov), { sort: { i: 2, dir: -1 } })}</div>`;
  }

  function collections() {
    const m = state.model; const uys = m.uys; const B = ['0-30', '31-90', '91-180', '181-365', 'Trên 365'];
    const ag = B.map((b) => ({ cat: b, vals: Object.fromEntries(uys.map((u) => [u, m.insts.filter((i) => i.uy === u && i.bucket === b).reduce((a, i) => a + i.amount, 0)])) }));
    const max = Math.max(0, ...ag.flatMap((r) => Object.values(r.vals)));
    const cov = uys.map((u) => { const bp = m.policies.filter((p) => p.uy === u && p.bound_); const ks = new Set(m.insts.filter((i) => i.uy === u).map((i) => i.key)); const ins = m.insts.filter((i) => i.uy === u); return { u, n: bp.length, sched: bp.filter((p) => ks.has(p.key)).length, nwp: bp.reduce((a, p) => a + p.nwp, 0), amt: ins.reduce((a, i) => a + i.amount, 0), paid: ins.filter((i) => i.isPaid).reduce((a, i) => a + i.amount, 0), od: ins.filter((i) => i.days > 0).reduce((a, i) => a + i.amount, 0) }; });
    const od = m.insts.filter((i) => i.days > 0);
    return `<div class="grid two">
        <div class="panel"><h2>Tuổi nợ kỳ phí chưa ghi nhận thu</h2>${legend(uys)}${bars(ag, uys, max, fN)}<div class="refline">Kỳ đã đến hạn trước ${fD(m.asOf)} mà chưa ghi "Paid".</div></div>
        <div class="panel"><h2>Độ phủ lịch thanh toán</h2>${table([{ l: 'UY', v: (r) => r.u }, { l: 'HĐ có lịch', v: (r) => r.sched, h: (r) => `${r.sched}/${r.n}`, num: true }, { l: 'NWP', v: (r) => r.nwp, f: fN, num: true }, { l: 'Tổng kỳ phí', v: (r) => r.amt, f: fN, num: true }, { l: 'Đã thu', v: (r) => r.paid, f: fN, num: true }, { l: 'Quá hạn', v: (r) => r.od, f: fN, num: true }], cov)}
          <p class="muted small" style="margin-top:8px">Tổng kỳ phí thấp hơn NWP nghĩa là còn hợp đồng chưa lập lịch thanh toán.</p></div>
      </div>
      <div class="section"><h2>Kỳ phí quá hạn (${od.length})</h2>${table([{ l: 'UY', v: (i) => i.uy }, { l: 'Original insured', v: (i) => i.insured, wrap: true }, { l: 'Broker', v: (i) => i.broker }, { l: 'Kỳ', v: (i) => i.n, num: true }, { l: 'Hạn', v: (i) => i.due, f: fD }, { l: 'Số tiền', v: (i) => i.amount, f: fN2, num: true }, { l: 'Ngày quá hạn', v: (i) => i.days, f: fN, num: true }, { l: 'Nhóm', v: (i) => i.bucket }], od, { tall: true, sort: { i: 6, dir: -1 }, total: false })}</div>`;
  }

  function risk() {
    const m = state.model; const S = m.settings; const bp = m.policies.filter((p) => p.bound_);
    const over = bp.filter((p) => p.flags.some((f) => f.startsWith('Line ')));
    const hr = bp.filter((p) => p.highRisk);
    const small = bp.filter((p) => p.flags.includes('Phí dưới mức tối thiểu'));
    const late = bp.filter((p) => p.flags.includes('Bind sau hiệu lực'));
    const lead = []; m.uys.forEach((u) => { const q = bp.filter((p) => p.uy === u); const tot = q.reduce((a, p) => a + p.gwp, 0); const inc = q.reduce((a, p) => a + p.incurred, 0); const by = {}; q.forEach((p) => { const k = p.leader || '(trống)'; by[k] = by[k] || { u, l: k, gwp: 0, inc: 0, n: 0 }; by[k].gwp += p.gwp; by[k].inc += p.incurred; by[k].n++; }); Object.values(by).forEach((r) => { r.sg = tot ? r.gwp / tot : 0; r.si = inc ? r.inc / inc : 0; lead.push(r); }); });
    const pc = [{ l: 'UY', v: (p) => p.uy }, { l: 'Original insured', v: (p) => p.insured, wrap: true }, { l: 'Quốc gia', v: (p) => p.domicile }, { l: 'Leader', v: (p) => p.leader }, { l: 'Số tàu', v: (p) => p.ac, f: fN, num: true }, { l: 'Line', v: (p) => p.share, f: (v) => fP(v, 2), num: true }, { l: 'Hull PVI', v: (p) => p.share * p.hull, f: fN, num: true }, { l: 'Net earning', v: (p) => p.earning, f: fN, num: true }, { l: 'Incurred', v: (p) => p.incurred, f: fN, num: true }];
    return `<p class="muted small">Ngưỡng hiện tại: trần line ${fP(S.lineCapGeneral, 2)} (chung), ${fP(S.lineCapSmall, 2)} (hãng dưới ${S.smallFleet} tàu hoặc quốc gia rủi ro cao) · leader tối đa ${fP(S.leaderCap, 0)} GWP · phí tối thiểu ${fN(S.minPremium)} USD. Đổi ở tab Thiết lập.</p>
      <div class="section"><h2>Line vượt trần (${over.length})</h2>${table(pc, over, { sort: { i: 6, dir: -1 }, tall: true })}</div>
      <div class="section"><h2>Tập trung leader</h2>${table([{ l: 'UY', v: (r) => r.u }, { l: 'Leader', v: (r) => r.l }, { l: 'Số HĐ', v: (r) => r.n, num: true }, { l: 'GWP', v: (r) => r.gwp, f: fN, num: true }, { l: '% GWP', v: (r) => r.sg, h: (r) => `${fP(r.sg)}${r.sg > S.leaderCap ? ' <span class="pill warn">vượt trần</span>' : ''}`, num: true }, { l: 'Incurred', v: (r) => r.inc, f: fN, num: true }, { l: '% incurred', v: (r) => r.si, f: fP, num: true }], lead, { sort: { i: 3, dir: -1 }, tall: true })}</div>
      <div class="section"><h2>Quốc gia nhạy cảm cấm vận (${hr.length})</h2>${table(pc.concat([{ l: 'Cedant', v: (p) => p.cedant, wrap: true }]), hr, { tall: true })}</div>
      <div class="section"><h2>Bind sau ngày hiệu lực (${late.length})</h2>${table([{ l: 'UY', v: (p) => p.uy }, { l: 'Original insured', v: (p) => p.insured, wrap: true }, { l: 'Hiệu lực', v: (p) => p.inc, f: fD }, { l: 'Bind', v: (p) => p.bound, f: fD }, { l: 'Số ngày trễ', v: (p) => p.bound - p.inc, num: true }, { l: 'Incurred', v: (p) => p.incurred, f: fN, num: true }], late, { sort: { i: 4, dir: -1 } })}</div>
      <div class="section"><h2>Phí dưới mức tối thiểu (${small.length})</h2>${table([{ l: 'UY', v: (p) => p.uy }, { l: 'Original insured', v: (p) => p.insured, wrap: true }, { l: 'Broker', v: (p) => p.broker }, { l: 'GWP', v: (p) => p.gwp, f: fN, num: true }, { l: 'Net earning', v: (p) => p.earning, f: fN, num: true }], small, { sort: { i: 3, dir: 1 }, total: true, tall: true })}</div>`;
  }

  function issues() {
    const m = state.model; const F = state.issueF; const q = F.q.toLowerCase();
    const rows = m.issues.filter((i) => (!F.sev || i.sev === F.sev) && (!F.uy || i.uy === F.uy) && (!q || (i.rec + ' ' + i.msg + ' ' + i.field).toLowerCase().includes(q)));
    const cnt = (s) => m.issues.filter((i) => i.sev === s).length;
    return `<div class="toolbar">
        <label class="fld" for="iSev">Mức<select id="iSev"><option value="">Tất cả (${m.issues.length})</option>${['Cao', 'TB', 'Thấp'].map((s) => `<option value="${s}" ${F.sev === s ? 'selected' : ''}>${s === 'TB' ? 'Trung bình' : s} (${cnt(s)})</option>`).join('')}</select></label>
        <label class="fld" for="iUy">UY<select id="iUy"><option value="">Tất cả</option>${m.uys.map((u) => `<option ${F.uy === u ? 'selected' : ''}>${esc(u)}</option>`).join('')}</select></label>
        <label class="fld" for="iQ">Tìm<input type="search" id="iQ" value="${esc(F.q)}" placeholder="Tên hãng, trường…"></label></div>
      ${table([{ l: 'Mức', v: (i) => ({ 'Cao': 0, 'TB': 1, 'Thấp': 2 })[i.sev], h: (i) => sevPill(i.sev) }, { l: 'UY', v: (i) => i.uy }, { l: 'Sheet', v: (i) => i.sheet }, { l: 'Đối tượng', v: (i) => i.rec, wrap: true }, { l: 'Trường', v: (i) => i.field }, { l: 'Vấn đề', v: (i) => i.msg, wrap: true }, { l: 'Hướng xử lý', v: (i) => i.fix, wrap: true }], rows, { tall: true })}`;
  }

  function settings() {
    const S = state.settings;
    const n = (id, l, v, step, hint) => `<label class="fld" for="${id}">${l}<input type="number" id="${id}" value="${v}" step="${step}"><span class="small muted">${hint || ''}</span></label>`;
    return `<div class="panel"><h2>Ngưỡng cảnh báo</h2><p class="muted small">Lưu trong trình duyệt này. Chỉ lưu các ngưỡng, không lưu dữ liệu.</p>
      <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr));margin-top:10px">
        ${n('sGen', 'Trần line chung (%)', +(S.lineCapGeneral * 100).toFixed(3), '0.05')}
        ${n('sSmall', 'Trần line hãng nhỏ / vùng rủi ro (%)', +(S.lineCapSmall * 100).toFixed(3), '0.05')}
        ${n('sFleet', 'Hãng nhỏ: dưới bao nhiêu tàu', S.smallFleet, '1')}
        ${n('sLead', 'Tỷ trọng tối đa một leader (%)', +(S.leaderCap * 100).toFixed(1), '1')}
        ${n('sMin', 'GWP tối thiểu mỗi line (USD)', S.minPremium, '100')}
        ${n('sOd', 'Cảnh báo quá hạn sau (ngày)', S.overdueAlertDays, '1')}
      </div>
      <label class="fld" for="sHigh" style="margin-top:12px">Quốc gia rủi ro cao / nhạy cảm cấm vận (phân cách bằng dấu phẩy)<textarea id="sHigh">${esc(S.highRiskCountries.join(', '))}</textarea></label>
      <div class="toolbar" style="margin-top:12px"><button class="primary" id="sSave">Áp dụng</button><button id="sReset">Khôi phục mặc định</button></div></div>`;
  }

  function guide() {
    return `<div class="panel"><h2>Định dạng file đầu vào</h2>
      <p>Mỗi file là một năm nghiệp vụ, gồm các sheet và cột dưới đây (tên cột không phân biệt hoa thường, thứ tự cột tuỳ ý). Các sheet <code>Policy</code>, <code>Risk</code>, <code>Premium</code> phải cùng thứ tự dòng.</p>
      <div class="steps">
        <div><h3>Policy</h3><p class="small">MRO, Broker, PVI ref number, Original Insured, Cedant Name, Region, Domicile, Risk offer date, Status, Risk bound date, Inception, Expiry, Leader, Risk type, Type, Category, Channel</p></div>
        <div><h3>Risk</h3><p class="small">MRO, PVI share (% – 0,5 nghĩa là 0,5%), Hull, CSL, A/C, AFV</p></div>
        <div><h3>Premium</h3><p class="small">MRO, Deposit GWP, PVI deposite GWP, PVI deposite NWP, Total deduction, Facility fee, Other Deduction, Fund, PVI actual earning, Premium status</p></div>
        <div><h3>Claim</h3><p class="small">MRO, Claim number, Date claim reported, Occurrence date, Coverage, Type of loss, Gross Incurred, PVI share, PVI Paid Claim, PVI Paid Fee, PVI reserved claim, PVI reserved Fee, Status, Paid Status</p></div>
        <div><h3>Installment</h3><p class="small">MRO, Broker, Original Insured, Inst n Amount, Inst n Due Date, Inst n Paid ("Paid" khi đã thu), Inst n Actual NWP (n = 1…12)</p></div>
      </div>
      <h2 style="margin-top:18px">Cách tính</h2>
      <ul class="small" style="max-width:80ch;display:grid;gap:4px">
        <li>Phí đã hưởng = net earning × số ngày đã qua / số ngày của hợp đồng (tối đa 100%).</li>
        <li>Incurred = đã trả + phí đã trả + dự phòng + phí dự phòng. Bồi thường gắn với hợp đồng qua MRO trong cùng UY; MRO trùng (sửa đổi bổ sung) gắn vào dòng gốc đang Bound.</li>
        <li>Chi phí khai thác = 1 − net earning / GWP (brokerage, SMB, facility fee, khấu trừ khác, quỹ).</li>
        <li>Ngày nhập dạng chữ phải theo <code>yyyy-mm-dd</code> hoặc <code>dd/mm/yyyy</code>; ô dạng ngày của Excel đọc trực tiếp.</li>
        <li>Region chuẩn hoá: "Asia" → "Asia Pacific RIM"; "LATAM", "South America" → "Americas". Leader: "TMK" → "Tokio Marine Kiln".</li>
      </ul>
      <h2 style="margin-top:18px">Bảo mật</h2>
      <p class="small">File được đọc bằng JavaScript ngay trong trình duyệt. App không có máy chủ, không gửi dữ liệu đi đâu; đóng tab là dữ liệu mất. Không đưa file dữ liệu thật vào repository.</p></div>`;
  }

  // ---------- Xuất Excel ----------
  function exportWorkbook() {
    const m = state.model; if (!m) return;
    const wb = XLSX.utils.book_new();
    const add = (name, rows) => { const ws = XLSX.utils.json_to_sheet(rows); ws['!cols'] = Object.keys(rows[0] || { a: 1 }).map((k) => ({ wch: Math.min(48, Math.max(10, k.length + 2)) })); XLSX.utils.book_append_sheet(wb, ws, name); };
    const d = (x) => x == null ? '' : fD(x);
    const r2 = (x) => x == null || !isFinite(x) ? '' : Math.round(x * 100) / 100;
    const r4 = (x) => x == null || !isFinite(x) ? '' : Math.round(x * 10000) / 10000;
    const sumRows = []; const keys = [['Số HĐ Bound', 'n'], ['HĐ huỷ', 'cancelled'], ['PVI GWP', 'gwp'], ['PVI NWP', 'nwp'], ['PVI net earning', 'earning'], ['Chi phí KT % GWP', 'acq'], ['Phí đã hưởng', 'ep'], ['% đã hưởng', 'earnedPct'], ['PVI incurred', 'incurred'], ['Đã trả', 'paid'], ['Dự phòng', 'os'], ['LR / phí đã hưởng', 'lrEP'], ['LR / net earning cả năm', 'lrFull'], ['LR bỏ 2 vụ lớn nhất', 'lrExTop2'], ['(Incurred + chi phí KT) / GWP', 'combined'], ['Số hồ sơ', 'nClaims'], ['Capacity PVI', 'cap'], ['Kỳ phí quá hạn', 'overdue'], ['Quá hạn > ngưỡng', 'overdue90'], ['Bind sau hiệu lực', 'boundLate']];
    const S = Object.fromEntries(m.uys.map((u) => [u, E.summarize(m, (p) => p.uy === u)])); const T = E.summarize(m);
    keys.forEach(([l, k]) => { const r = { 'Chỉ tiêu': l }; m.uys.forEach((u) => { r[u] = r4(S[u][k]); }); r['Tổng'] = r4(T[k]); sumRows.push(r); });
    sumRows.push({}); sumRows.push({ 'Chỉ tiêu': 'Ngày báo cáo', [m.uys[0]]: fD(m.asOf) });
    add('Tong quan', sumRows);
    add('Hop dong', m.policies.map((p) => ({ UY: p.uy, 'PVI ref': p.ref, MRO: p.mro, Broker: p.broker, 'Original insured': p.insured, Cedant: p.cedant, Region: p.region, Domicile: p.domicile, Channel: p.channel, Leader: p.leader, Type: p.type, Status: p.status, 'Ngày bind': d(p.bound), 'Hiệu lực': d(p.inc), 'Hết hạn': d(p.exp), 'Hết hạn tính toán': d(p.expC), 'Earned %': r4(p.earnPct), 'PVI share': r4(p.share), Hull: p.hull, CSL: p.csl, 'Capacity PVI': r2(p.cap), 'PVI GWP': r2(p.gwp), 'PVI NWP': r2(p.nwp), 'Net earning': r2(p.earning), 'Phí đã hưởng': r2(p.ep), Incurred: r2(p.incurred), 'Đã trả': r2(p.paidT), 'Dự phòng': r2(p.os), 'LR đã hưởng': r4(p.lr), 'Chi phí KT %': r4(p.acq), 'Cờ': p.flags.join('; ') })));
    add('Boi thuong', m.claims.map((c) => ({ UY: c.uy, 'Số hồ sơ': c.claimNo, MRO: c.mro, 'Original insured': c.insured, Coverage: c.coverage, 'Loại tổn thất': c.lossType, 'Ngày tổn thất': c.occ != null ? fD(c.occ) : clean(c.occRaw), 'Ngày thông báo': d(c.rep), 'Gross 100%': c.gross100, 'PVI share': r4(c.share), 'Đã trả': r2(c.paidT), 'Dự phòng': r2(c.os), Incurred: r2(c.incurred), Status: c.status, 'Kiểm tra': c.flag, 'Gợi ý UY': c.suggest || '' })));
    add('Ky thanh toan', m.insts.map((i) => ({ UY: i.uy, MRO: i.mro, Broker: i.broker, 'Original insured': i.insured, 'Kỳ': i.n, 'Số tiền': r2(i.amount), 'Hạn': d(i.due), 'Trạng thái': i.paid, 'Thực thu': r2(i.actual), 'Ngày quá hạn': i.days, 'Nhóm tuổi nợ': i.bucket })));
    [['broker', 'PK Broker'], ['channel', 'PK Kenh'], ['region', 'PK Vung'], ['leader', 'PK Leader']].forEach(([dim, name]) => {
      add(name, E.segment(m, dim).map((r) => { const o = { 'Nhóm': r.cat }; m.uys.forEach((u) => { o[`GWP ${u}`] = r2(r.byUy[u].gwp); o[`Incurred ${u}`] = r2(r.byUy[u].incurred); o[`LR đã hưởng ${u}`] = r4(r.byUy[u].lrEP); o[`Chi phí KT ${u}`] = r4(r.byUy[u].acq); }); return o; }));
    });
    add('Kiem tra du lieu', m.issues.map((i, k) => ({ '#': k + 1, 'Mức': i.sev, UY: i.uy, Sheet: i.sheet, 'Đối tượng': i.rec, 'Trường': i.field, 'Vấn đề': i.msg, 'Hướng xử lý': i.fix, 'Đã xử lý?': '' })));
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const dt = new Date(m.asOf * E.DAY); const stamp = `${dt.getUTCFullYear()}${String(dt.getUTCMonth() + 1).padStart(2, '0')}${String(dt.getUTCDate()).padStart(2, '0')}`;
    saveFile(`Bao_cao_chi_so_TBH_hang_khong_${stamp}.xlsx`, new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }
  function clean(s) { return s == null ? '' : String(s).trim(); }
  async function saveFile(name, blob) {
    // Trong claude.ai: dùng capability downloads; trên GitHub Pages: link tải thông thường.
    if (window.claude && typeof window.claude.use === 'function') {
      try { const dl = await window.claude.use('downloads'); if (dl) { await dl.save({ filename: name, data: blob }); toast('Đã lưu ' + name); return; } } catch (e) { if (e && e.code === 'declined') return; }
    }
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  function exportTemplate() {
    const wb = XLSX.utils.book_new();
    const ex = 'VÍ DỤ – xoá dòng này';
    const sheets = {
      Policy: [['MRO', 'Broker', 'PIAS Ref', 'PVI ref number', 'Original Insured', 'Cedant Name', 'Region', 'Domicile', 'Risk offer date', 'Status', 'Risk bound date', 'Inception', 'Expiry', 'Flying/Non-flying', 'Leader', 'Risk type', 'Type', 'Category', 'Channel'], ['B0000EXAMPLE01', 'AJG', '', 'PVI202600001', ex, 'Example Insurance Co', 'Asia Pacific RIM', 'Vietnam', '2026-09-01', 'Bound', '2026-09-20', '2026-10-01', '2027-09-30', 'Flying', 'AXA XL', 'HSL', 'RI', '', 'A1']],
      Risk: [['MRO', 'Broker', 'Original Insured', 'PVI share', 'Capacity deployment', 'Hull', 'Spare limit', 'CSL', 'PA', 'AVN52E', 'A/C', 'AFV', 'Pax', 'Dept', 'Spare value'], ['B0000EXAMPLE01', 'AJG', ex, 0.5, 5500000, 100000000, 0, 1000000000, 0, 0, 12, 400000000, 0, 0, 0]],
      Premium: [['MRO', 'Broker', 'Original Insured', 'Region', 'Domicile', 'PVI share', 'Deposit GWP', 'Deposit NWP', 'PVI deposite GWP', 'PVI deposite NWP', 'Brokerage', 'SMB', 'Total deduction', 'Facility fee', 'Other Deduction', 'Fund', 'PVI actual earning', 'Diffirental from lead', 'Facility', 'Month', 'Year', 'Policy status', 'Premium status'], ['B0000EXAMPLE01', 'AJG', ex, 'Asia Pacific RIM', 'Vietnam', 0.5, 2000000, 1500000, 10000, 7500, 21.5, 3.5, 25, 0, 0, 0, 7500, 0, 'A1', 10, 2026, 'Bound', 'Estimated']],
      Claim: [['MRO', 'Broker', 'Original Insured', 'Claim number', 'Date claim reported', 'Occurrence date', 'Location', 'Loss details', 'Coverage', 'Type of loss', 'Loss adjuster/ Lawyer', 'Paid Claim (100%)', 'Paid Fee (100%)', 'Reserved claim', 'Resered Fee', 'Gross Incurred', 'PVI share', 'PVI Paid Claim', 'PVI Paid Fee', 'PVI reserved claim', 'PVI reserved Fee', "PVI's Gross Incured", 'Status', 'Lastest Movement Date', 'Refund amout', 'Paid Status'], ['B0000EXAMPLE01', 'AJG', ex, 'PVIAVC2026000001', '2026-11-05', '2026-11-01', '', 'Mô tả tổn thất', 'Hull', 'Engine', '', 0, 0, 500000, 0, 500000, 0.5, 0, 0, 2500, 0, 2500, 'Open', '2026-11-05', 0, '']],
    };
    const ih = ['MRO', 'Broker', 'Original Insured', 'Closing Ref', 'Closing Date']; for (let n = 1; n <= 12; n++) ih.push(`Inst ${n} Date`, `Inst ${n} Share %`, `Inst ${n} Amount`, `Inst ${n} Due Date`, `Inst ${n} Paid`, `Inst ${n} Actual NWP`); ih.push('Last Movement Date');
    sheets.Installment = [ih, ['B0000EXAMPLE01', 'AJG', ex, '', '', '2026-10-01', 0, 3750, '2026-11-01', 'Paid', 3750, '2027-04-01', 0, 3750, '2027-05-01', '', 0]];
    Object.entries(sheets).forEach(([n, rows]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), n));
    saveFile('Mau_UY_x_theo_doi.xlsx', new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }

  // ---------- Sự kiện ----------
  function setAsOfInput() { const d = new Date(state.settings.asOf * E.DAY); $('#asof').value = d.toISOString().slice(0, 10); }
  $('#pick').addEventListener('click', () => $('#file').click());
  $('#file').addEventListener('change', (e) => { addFiles([...e.target.files]); e.target.value = ''; });
  const drop = $('#drop');
  ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => { const f = [...(e.dataTransfer.files || [])].filter((x) => /\.xls[xm]?$/i.test(x.name)); if (f.length) addFiles(f); });
  $('#asof').addEventListener('change', (e) => { const d = E.parseDate(e.target.value); if (d != null) { state.settings.asOf = d; recompute(); } });
  $('#export').addEventListener('click', exportWorkbook);
  $('#template').addEventListener('click', exportTemplate);
  $('#files').addEventListener('click', (e) => { const i = e.target.getAttribute('data-rm'); if (i != null) { state.books.splice(+i, 1); recompute(); } });
  $('#files').addEventListener('change', (e) => { const i = e.target.getAttribute('data-i'); if (i != null) { state.books[+i].uy = e.target.value.trim() || state.books[+i].uy; recompute(); } });
  $('#tabs').addEventListener('click', (e) => { const b = e.target.closest('[data-tab]'); if (b) { state.tab = b.dataset.tab; try { localStorage.setItem('kpi-tab', state.tab); } catch (x) { /* */ } renderTabs(); render(); } });
  const view = $('#view');
  view.addEventListener('click', (e) => {
    const th = e.target.closest('th[data-sort]');
    if (th) { const box = th.closest('.tbl'); const t = TABLES[box.id]; const i = +th.dataset.sort; const cur = t.opts.sort; t.opts.sort = { i, dir: cur && cur.i === i ? -cur.dir : (t.cols[i].num ? -1 : 1) }; box.outerHTML = table(t.cols, t.rows, t.opts); return; }
    if (e.target.id === 'sSave') {
      const S = state.settings; const v = (id) => parseFloat($('#' + id).value);
      S.lineCapGeneral = v('sGen') / 100; S.lineCapSmall = v('sSmall') / 100; S.smallFleet = v('sFleet'); S.leaderCap = v('sLead') / 100; S.minPremium = v('sMin'); S.overdueAlertDays = v('sOd');
      S.highRiskCountries = $('#sHigh').value.split(',').map((s) => s.trim()).filter(Boolean);
      saveSettings(); recompute(); toast('Đã áp dụng ngưỡng mới');
    }
    if (e.target.id === 'sReset') { const a = state.settings.asOf; state.settings = Object.assign({}, E.DEFAULT_SETTINGS, { asOf: a }); saveSettings(); recompute(); toast('Đã khôi phục mặc định'); }
  });
  view.addEventListener('change', (e) => {
    const id = e.target.id;
    if (id === 'segDim') { state.segDim = e.target.value; render(); }
    if (id === 'cUy') { state.claimF.uy = e.target.value; render(); }
    if (id === 'cSt') { state.claimF.st = e.target.value; render(); }
    if (id === 'cFl') { state.claimF.flag = e.target.value; render(); }
    if (id === 'iSev') { state.issueF.sev = e.target.value; render(); }
    if (id === 'iUy') { state.issueF.uy = e.target.value; render(); }
  });
  view.addEventListener('input', (e) => { if (e.target.id === 'iQ') { state.issueF.q = e.target.value; const pos = e.target.selectionStart; render(); const q = $('#iQ'); q.focus(); q.setSelectionRange(pos, pos); } });

  try { const t = localStorage.getItem('kpi-tab'); if (t && TABS.some((x) => x[0] === t)) state.tab = t; } catch (e) { /* */ }
  setAsOfInput(); renderTabs(); render();
})();
