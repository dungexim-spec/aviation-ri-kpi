/* Giao diện: tải file → KpiEngine → các báo cáo. Dữ liệu chỉ xử lý trong trình duyệt. */
(function () {
  'use strict';
  const E = window.KpiEngine;
  const $ = (s, r) => (r || document).querySelector(s);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fin = (v) => v != null && isFinite(v);
  const fN = (v) => fin(v) ? Math.round(v).toLocaleString('vi-VN') : '–';
  const fN2 = (v) => fin(v) ? v.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '–';
  const fP = (v, d = 1) => fin(v) ? (v * 100).toFixed(d).replace('.', ',') + '%' : '–';
  const fPP = (v) => fin(v) ? (v >= 0 ? '+' : '') + (v * 100).toFixed(1).replace('.', ',') + ' đ%' : '–';
  const fDelta = (v) => fin(v) ? (v >= 0 ? '+' : '−') + Math.abs(Math.round(v)).toLocaleString('vi-VN') : '–';
  const fD = (d) => E.fmtDate(d);
  const SERIES = ['--s1', '--s2', '--s3', '--s4', '--s5', '--s6', '--s7', '--s8'];
  const ORD = { 'Cao': 0, 'TB': 1, 'Thấp': 2 };

  const state = {
    books: [], main: null, snaps: [], tab: 'overview', settings: loadSettings(), asOfManual: false,
    scope: { broker: '', channel: '', region: '', coll: '' }, collDim: 'broker', renDim: 'broker', channelDim: 'channel', regionDim: 'region', renewDays: 90,
    claimF: { uy: '', st: '', flag: '' }, issueF: { sev: '', uy: '', q: '' },
  };

  // ---------- Thiết lập (chỉ lưu ngưỡng, không lưu dữ liệu) ----------
  function loadSettings() {
    const d = Object.assign({}, E.DEFAULT_SETTINGS, { asOf: E.todayDay(), renewDays: 90 });
    try { const s = JSON.parse(localStorage.getItem('kpi-settings') || 'null'); if (s) Object.assign(d, s, { asOf: d.asOf }); } catch (e) { /* bỏ qua */ }
    return d;
  }
  function saveSettings() { try { const s = Object.assign({}, state.settings); delete s.asOf; localStorage.setItem('kpi-settings', JSON.stringify(s)); } catch (e) { /* bỏ qua */ } }
  function uyColor(uy) { const m = state.main; const i = m ? m.uys.indexOf(uy) : 0; return `var(${SERIES[Math.max(0, i) % SERIES.length]})`; }
  function toast(msg) { const t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toast._t); toast._t = setTimeout(() => { t.hidden = true; }, 3200); }

  // ---------- Tải file ----------
  async function addFiles(fileList) {
    const warns = [];
    for (const f of fileList) {
      try {
        const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
        const book = E.readWorkbook(XLSX, wb, E.uyFromFilename(f.name));
        book.file = f.name;
        book.snap = E.snapFromFilename(f.name);
        if (book.snap == null && f.lastModified) { const d = new Date(f.lastModified); book.snap = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / E.DAY; }
        state.books.push(book);
        book.warnings.forEach((w) => warns.push(`${f.name}: ${w}`));
      } catch (err) { warns.push(`${f.name}: ${err.message || err}`); }
    }
    $('#loadWarn').innerHTML = warns.map((w) => `<div class="warnline">⚠ ${esc(w)}</div>`).join('');
    if (!state.asOfManual) { const mx = Math.max(...state.books.map((b) => b.snap || 0)); if (mx > 0) { state.settings.asOf = mx; setAsOfInput(); } }
    recompute();
  }
  function renderFiles() {
    $('#files').innerHTML = state.books.map((b, i) => {
      const iso = b.snap != null ? new Date(b.snap * E.DAY).toISOString().slice(0, 10) : '';
      const latest = state.main && state.main.bookSet.has(b);
      return `<span class="chip ${latest ? '' : 'old'}" title="${latest ? 'Đang dùng cho báo cáo chính' : 'Bản cũ hơn – chỉ dùng ở tab Kỳ báo cáo'}"><span class="sw" style="background:${uyColor(b.uy)}"></span>
        <input aria-label="Năm nghiệp vụ" data-uy="${i}" value="${esc(b.uy)}"><input type="date" aria-label="Ngày chốt số liệu" data-snap="${i}" value="${iso}">
        <span class="fn" title="${esc(b.file)}">${esc(b.file)}</span><span class="muted">${b.policies.length} HĐ</span><button class="x" data-rm="${i}" aria-label="Bỏ file">×</button></span>`;
    }).join('');
  }
  function recompute() {
    const S = state.settings;
    if (!state.books.length) { state.main = null; state.snaps = []; }
    else {
      // báo cáo chính: bản mới nhất của mỗi UY
      const latest = {};
      state.books.forEach((b) => { const cur = latest[b.uy]; if (!cur || (b.snap || 0) >= (cur.snap || 0)) latest[b.uy] = b; });
      const set = Object.values(latest);
      state.main = E.build(set, S); state.main.bookSet = new Set(set);
      // từng bản (cho so sánh kỳ báo cáo)
      state.snaps = state.books.map((b) => { const m = E.build([b], Object.assign({}, S, { asOf: b.snap != null ? b.snap : S.asOf })); return { b, s: E.summarize(m) }; })
        .sort((a, b) => a.b.uy.localeCompare(b.b.uy) || (a.b.snap || 0) - (b.b.snap || 0));
    }
    $('#export').disabled = !state.main;
    renderFiles(); renderTabs(); render();
  }

  // ---------- Tabs ----------
  const TABS = [['overview', 'Tổng quan'], ['broker', 'Môi giới'], ['channel', 'Kênh khai thác'], ['region', 'Khu vực'], ['period', 'Kỳ báo cáo'], ['renew', 'Tái tục'],
    ['claims', 'Bồi thường'], ['collections', 'Thu phí'], ['risk', 'Cỡ line & tập trung'], ['issues', 'Kiểm tra dữ liệu'], ['settings', 'Thiết lập'], ['guide', 'Hướng dẫn']];
  function renderTabs() {
    const m = state.main;
    const nCrit = m ? m.issues.filter((i) => i.sev === 'Cao').length : 0;
    const nRen = m ? E.renewals(m, 30, 30).filter((r) => !r.renewed).length : 0;
    const nOd = m ? m.insts.filter((i) => i.days > 0).length : 0;
    $('#tabs').innerHTML = TABS.map(([k, l]) => `<button role="tab" aria-selected="${state.tab === k}" data-tab="${k}">${l}${k === 'issues' && nCrit ? `<span class="count">${nCrit}</span>` : ''}${k === 'renew' && nRen ? `<span class="count warn">${nRen}</span>` : ''}${k === 'collections' && nOd ? `<span class="count">${nOd}</span>` : ''}</button>`).join('');
  }

  // ---------- Bảng ----------
  const TABLES = {};
  function table(cols, rows, opts = {}) {
    const id = 't' + Math.random().toString(36).slice(2, 8);
    let data = rows.slice(); const sort = opts.sort;
    if (sort) { const c = cols[sort.i]; data.sort((a, b) => { let x = c.v(a), y = c.v(b); if (x == null) x = -Infinity; if (y == null) y = -Infinity; return (x > y ? 1 : x < y ? -1 : 0) * sort.dir; }); }
    const grp = opts.groups ? `<tr>${opts.groups.map(([l, n]) => `<th colspan="${n}" class="grp">${esc(l)}</th>`).join('')}</tr>` : '';
    const head = cols.map((c, i) => `<th class="${c.num ? 'r' : ''} sortable" data-sort="${i}" ${c.tip ? `title="${esc(c.tip)}"` : ''}>${esc(c.l)}</th>`).join('');
    const body = data.map((r) => `<tr class="${opts.rowClass ? opts.rowClass(r) : ''}">${cols.map((c) => `<td class="${c.num ? 'r' : ''} ${c.wrap ? 'tw' : ''} ${c.cls ? c.cls(r) : ''}">${c.h ? c.h(r) : esc(c.f ? c.f(c.v(r)) : c.v(r))}</td>`).join('')}</tr>`).join('');
    const foot = opts.foot ? `<tr class="total">${cols.map((c, i) => `<td class="${c.num ? 'r' : ''}">${i === 0 ? esc(opts.footLabel || 'Tổng') : (c.t ? esc(c.t()) : '')}</td>`).join('')}</tr>` : '';
    TABLES[id] = { cols, rows, opts };
    return `<div class="tbl ${opts.tall ? 'tall' : ''}" id="${id}"><table><thead>${grp}<tr>${head}</tr></thead><tbody>${body || `<tr><td colspan="${cols.length}" class="muted">Không có dòng nào.</td></tr>`}${foot}</tbody></table></div>`;
  }
  // Bộ cột chuẩn: GWP, NWP, chi phí khai thác, bồi thường, kết hợp
  function metricCols(get, foot, withShare) {
    const col = (l, k, f, tip) => ({ l, v: (r) => get(r)[k], f, num: true, t: foot ? () => f(foot[k]) : null, tip });
    const cols = [
      col('Số HĐ', 'n', fN), col('GWP', 'gwp', fN), col('NWP', 'nwp', fN),
      col('Chi phí KT', 'acqUsd', fN, 'GWP − net earning: brokerage, SMB, facility fee, khấu trừ khác, quỹ'), col('CP KT % GWP', 'acqPct', fP),
      col('BT đã trả', 'paid', fN), col('BT dự phòng', 'os', fN), col('Bồi thường', 'incurred', fN), col('BT % GWP', 'claimsPct', fP),
      { l: 'Kết hợp', v: (r) => get(r).combinedUsd, f: fN, num: true, t: foot ? () => fN(foot.combinedUsd) : null, tip: 'Chi phí khai thác + bồi thường' },
      { l: 'Kết hợp % GWP', v: (r) => get(r).combinedPct, h: (r) => ratioCell(get(r).combinedPct), num: true, t: foot ? () => fP(foot.combinedPct) : null },
    ];
    if (withShare) cols.push({ l: 'Tỷ trọng GWP', v: (r) => r.shareGwp, f: fP, num: true, t: foot ? () => '100%' : null });
    return cols;
  }
  const ratioCell = (v) => !fin(v) ? '–' : `<span class="${v > 1 ? 'neg' : v > 0.8 ? 'warnc' : ''}">${fP(v)}</span>`;
  const GROUPS = (lead, extra) => [[lead, 1], ['Phí', 3], ['Chi phí khai thác', 2], ['Bồi thường', 4], ['Kết hợp', 2]].concat(extra ? [['', extra]] : []);

  // Biểu đồ tỷ lệ kết hợp: chi phí KT % + bồi thường % (thanh chồng), vạch 100%
  function comboChart(rows, label) {
    if (!rows.length) return '<p class="muted">Chưa có dữ liệu.</p>';
    const max = Math.max(1.2, ...rows.map((r) => (r.s.acqPct || 0) + (r.s.claimsPct || 0))) * 1.02;
    const pct = (v) => (Math.max(0, v) / max * 100).toFixed(2) + '%';
    return `<div class="legend"><span><span class="sw" style="background:var(--c-acq)"></span>Chi phí khai thác % GWP</span><span><span class="sw" style="background:var(--c-clm)"></span>Bồi thường % GWP</span><span><span class="sw line"></span>100%</span></div>
      <div class="cbars">${rows.map((r) => `<div class="cat" title="${esc(r.cat)}">${esc(r.cat)}</div>
        <div class="ctrack"><div class="ref" style="left:${pct(1)}"></div>
          <div class="cstack"><div class="a" style="width:${pct(r.s.acqPct || 0)}" title="${esc(r.cat)} · chi phí KT ${fP(r.s.acqPct)} (${fN(r.s.acqUsd)} USD)"></div><div class="c" style="width:${pct(r.s.claimsPct || 0)}" title="${esc(r.cat)} · bồi thường ${fP(r.s.claimsPct)} (${fN(r.s.incurred)} USD)"></div>
          <span class="cval">${ratioCell(r.s.combinedPct)} <span class="muted">· GWP ${fN(r.s.gwp)}</span></span></div></div>`).join('')}</div>
      <div class="refline">${esc(label || '')}</div>`;
  }
  const legend = (uys) => `<div class="legend">${uys.map((u) => `<span><span class="sw" style="background:${uyColor(u)}"></span>${esc(u)}</span>`).join('')}</div>`;
  const sevPill = (s) => `<span class="pill ${s === 'Cao' ? 'crit' : s === 'TB' ? 'warn' : 'good'}">● ${s === 'TB' ? 'Trung bình' : esc(s)}</span>`;
  const uySelect = (id, val, allLabel) => `<label class="fld" for="${id}">Năm nghiệp vụ<select id="${id}"><option value="">${allLabel || 'Tất cả UY'}</option>${state.main.uys.map((u) => `<option ${val === u ? 'selected' : ''}>${esc(u)}</option>`).join('')}</select></label>`;

  // ---------- Views ----------
  function render() {
    const v = $('#view'); const m = state.main;
    if (!m && !['settings', 'guide'].includes(state.tab)) { v.innerHTML = emptyView(); return; }
    const fn = { overview, broker: () => segView('broker', 'Môi giới', 'broker'), channel: channelView, region: regionView, period, renew, claims, collections, risk, issues, settings, guide }[state.tab];
    v.innerHTML = fn();
  }
  function emptyView() {
    return `<div class="panel empty"><h2>Bắt đầu</h2>
      <ol><li>Bấm <b>Chọn file Excel</b> và chọn các file theo dõi, ví dụ <code>UY_1_update_05102026.xlsx</code>, <code>UY_2_update_05102026.xlsx</code>.</li>
      <li>App nhận năm nghiệp vụ (<code>UY_1</code> → UY1) và ngày chốt số liệu (<code>update_05102026</code> → 05/10/2026) từ tên file; có thể sửa trên thẻ file.</li>
      <li>Muốn so sánh giữa các kỳ báo cáo, tải thêm các bản cập nhật cũ hơn của cùng UY. Báo cáo chính luôn dùng bản mới nhất của mỗi UY.</li>
      <li>Xem các tab, hoặc bấm <b>Xuất báo cáo Excel</b>.</li></ol>
      <p class="muted small">Cần các sheet <code>Policy</code>, <code>Risk</code>, <code>Premium</code>, <code>Claim</code>, <code>Installment</code> như file theo dõi hiện tại. Bấm <b>Tải file mẫu</b> để lấy khung trống.</p></div>`;
  }

  function overview() {
    const m = state.main; const uys = m.uys;
    const S = {}; uys.forEach((u) => { S[u] = E.summarize(m, (p) => p.uy === u); }); const T = E.summarize(m);
    const cols = uys.concat(uys.length > 1 ? ['Tổng'] : []); const get = (c) => c === 'Tổng' ? T : S[c];
    const kpi = (lab, a, b) => `<div class="panel kpi"><div class="lab">${lab}</div><div class="vals">${uys.map((u) => `<div class="v"><span class="sw" style="background:${uyColor(u)}"></span><span class="u">${esc(u)}</span><span>${a(S[u])}</span>${b ? `<span class="pct">${b(S[u])}</span>` : ''}</div>`).join('')}</div></div>`;
    const sec = (t) => `<tr class="sec"><td colspan="${cols.length + 1}">${t}</td></tr>`;
    const row = (l, f, cls) => `<tr class="${cls || ''}"><td>${l}</td>${cols.map((c) => `<td class="r">${f(get(c))}</td>`).join('')}</tr>`;
    const mt = `<div class="tbl"><table><thead><tr><th>Chỉ tiêu (USD)</th>${cols.map((c) => `<th class="r">${esc(c)}</th>`).join('')}</tr></thead><tbody>
      ${sec('Phí')}${row('Số hợp đồng Bound', (s) => fN(s.n))}${row('GWP', (s) => fN(s.gwp), 'b')}${row('NWP', (s) => fN(s.nwp))}${row('NWP % GWP', (s) => fP(s.gwp ? s.nwp / s.gwp : null))}
      ${sec('Chi phí khai thác')}${row('Brokerage, SMB (GWP − NWP)', (s) => fN(s.brokSmbUsd))}${row('Facility fee, khấu trừ khác, quỹ', (s) => fN(s.otherCostUsd))}${row('Chi phí khai thác', (s) => fN(s.acqUsd), 'b')}${row('Chi phí khai thác % GWP', (s) => fP(s.acqPct), 'b')}
      ${sec('Bồi thường')}${row('Đã trả', (s) => fN(s.paid))}${row('Đã trả % GWP', (s) => fP(s.paidPct))}${row('Dự phòng', (s) => fN(s.os))}${row('Dự phòng % GWP', (s) => fP(s.osPct))}${row('Bồi thường (đã trả + dự phòng)', (s) => fN(s.incurred), 'b')}${row('Bồi thường % GWP', (s) => fP(s.claimsPct), 'b')}
      ${sec('Kết hợp chi phí khai thác và bồi thường')}${row('Kết hợp', (s) => fN(s.combinedUsd), 'b')}${row('Tỷ lệ kết hợp % GWP', (s) => ratioCell(s.combinedPct), 'b')}${row('Còn lại sau chi phí KT và bồi thường', (s) => fN(s.resultUsd))}
      ${sec('Thu phí')}${row('Phí theo lịch thanh toán', (s) => fN(s.schedAmt))}${row('Thực thu', (s) => fN(s.collected), 'b')}${row('Thực thu % GWP', (s) => fP(s.collectedPctGwp), 'b')}${row('Thực thu % NWP', (s) => fP(s.collectedPctNwp))}${row('Phí quá hạn chưa thu', (s) => fN(s.overdue))}${row(`Phí đến hạn trong ${state.settings.upcomingDays || 30} ngày tới`, (s) => fN(s.upcoming))}${row('NWP chưa có lịch thanh toán', (s) => fN(s.noSchedNwp))}
      ${sec('Tái tục')}${(() => { const st = E.renewalStats(m); const by = Object.fromEntries(st.map((x) => [x.to, x])); const cell = (c, f) => { if (c === 'Tổng') { if (!st.length) return '–'; const t = st.reduce((a, x) => ({ r: a.r + x.nRenewed, d: a.d + x.nDue, n: a.n + x.nNew, nx: a.nx + x.nNext, gr: a.gr + x.gwpRenewedPrev, gd: a.gd + x.gwpDue }), { r: 0, d: 0, n: 0, nx: 0, gr: 0, gd: 0 }); return f({ nRenewed: t.r, nDue: t.d, ratioN: t.d ? t.r / t.d : null, ratioG: t.gd ? t.gr / t.gd : null, nNew: t.n, nNext: t.nx }); } return by[c] ? f(by[c]) : '–'; };
        const r2 = (l, f, cls) => `<tr class="${cls || ''}"><td>${l}</td>${cols.map((c) => `<td class="r">${cell(c, f)}</td>`).join('')}</tr>`;
        return r2('Đơn tái tục / đơn năm trước đã đến hạn', (x) => `${x.nRenewed}/${x.nDue}`) + r2('Tỷ lệ tái tục (số đơn)', (x) => fP(x.ratioN), 'b') + r2('Tỷ lệ tái tục (GWP)', (x) => fP(x.ratioG)) + r2('Đơn mới / tổng số đơn', (x) => `${x.nNew}/${x.nNext}`); })()}
      ${sec('Tham khảo')}${row('Phí đã hưởng (net earning pro-rata)', (s) => fN(s.ep))}${row('Loss ratio / phí đã hưởng', (s) => fP(s.lrEP))}${row('Số hồ sơ bồi thường', (s) => fN(s.nClaims))}
      </tbody></table></div>`;
    const rn = E.renewals(m, state.settings.renewDays || 90, 30);
    const n30 = rn.filter((r) => !r.renewed && r.daysLeft >= 0 && r.daysLeft <= 30), exp = rn.filter((r) => !r.renewed && r.daysLeft < 0), nAll = rn.filter((r) => !r.renewed && r.daysLeft >= 0);
    const banner = (n30.length || exp.length || nAll.length) ? `<div class="panel banner"><div><b>Nhắc tái tục</b> · ${n30.length} hợp đồng hết hạn trong 30 ngày tới (GWP ${fN(n30.reduce((a, r) => a + r.p.gwp, 0))} USD) · ${nAll.length} trong ${state.settings.renewDays || 90} ngày · ${exp.length} đã hết hạn chưa thấy tái tục.
      ${n30.length ? `<div class="small muted" style="margin-top:4px">${n30.slice(0, 6).map((r) => `${esc(r.p.insured.slice(0, 32))} (${fD(r.p.expC)})`).join(' · ')}${n30.length > 6 ? ' …' : ''}</div>` : ''}</div><button data-goto="renew">Xem danh sách</button></div>` : '';
    // điểm chú ý
    const ins = [];
    uys.forEach((u) => {
      const s = S[u];
      if (fin(s.combinedPct) && s.combinedPct > 1) ins.push(['Cao', `${u}: tỷ lệ kết hợp ${fP(s.combinedPct)} GWP (chi phí KT ${fP(s.acqPct)} + bồi thường ${fP(s.claimsPct)}).`]);
      else if (fin(s.combinedPct) && s.combinedPct > 0.8) ins.push(['TB', `${u}: tỷ lệ kết hợp ${fP(s.combinedPct)} GWP.`]);
      const bp = m.policies.filter((p) => p.uy === u && p.bound_).sort((a, b) => b.incurred - a.incurred).filter((p) => p.incurred > 0).slice(0, 2);
      const t2 = bp.reduce((a, p) => a + p.incurred, 0);
      if (s.incurred > 0 && t2 / s.incurred > 0.6) ins.push(['TB', `${u}: ${bp.map((p) => `${p.insured.slice(0, 30)} (${fN(p.incurred)})`).join(' và ')} chiếm ${fP(t2 / s.incurred, 0)} bồi thường.`]);
      if (s.overdue90 > 0) ins.push(['Cao', `${u}: ${fN(s.overdue90)} USD kỳ phí quá hạn trên ${m.settings.overdueAlertDays} ngày.`]);
    });
    const seg = E.segment(m, 'broker').rows.filter((r) => r.total.gwp > 0.05 * T.gwp && fin(r.total.combinedPct) && r.total.combinedPct > 0.8);
    seg.forEach((r) => ins.push(['TB', `Broker ${r.cat}: tỷ lệ kết hợp ${fP(r.total.combinedPct)} trên GWP ${fN(r.total.gwp)} USD.`]));
    const nCrit = m.issues.filter((i) => i.sev === 'Cao').length;
    if (nCrit) ins.push(['Cao', `${nCrit} lỗi dữ liệu mức cao (tab Kiểm tra dữ liệu).`]);
    ins.sort((a, b) => ORD[a[0]] - ORD[b[0]]);
    const pb = (T.overdue > 0 || T.upcoming > 0) ? `<div class="panel banner danger"><div><b>Nhắc thu phí</b> · ${fN(T.upcoming)} USD (${T.nUpcoming} kỳ) đến hạn trong ${state.settings.upcomingDays || 30} ngày tới · ${fN(T.overdue)} USD (${T.nOverdue} kỳ) đã quá hạn, trong đó ${fN(T.overdue90)} USD quá ${m.settings.overdueAlertDays} ngày.</div><button data-goto="collections">Xem chi tiết</button></div>` : '';
    return `${banner}${pb}
      <div class="grid kpis">
        ${kpi('GWP · NWP', (s) => fN(s.gwp), (s) => fN(s.nwp))}
        ${kpi('Chi phí khai thác', (s) => fN(s.acqUsd), (s) => fP(s.acqPct))}
        ${kpi('Bồi thường (đã trả + DP)', (s) => fN(s.incurred), (s) => fP(s.claimsPct))}
        ${kpi('Kết hợp CP KT + bồi thường', (s) => fN(s.combinedUsd), (s) => ratioCell(s.combinedPct))}
        ${kpi('Thực thu phí', (s) => fN(s.collected), (s) => fP(s.collectedPctGwp) + ' GWP')}
      </div>
      <div class="grid two section">
        <div class="panel"><h2>Tỷ lệ kết hợp theo năm nghiệp vụ</h2>${comboChart(cols.map((c) => ({ cat: c, s: get(c) })), `% trên GWP. Số liệu đến ${fD(m.asOf)}; bồi thường chưa gồm IBNR.`)}</div>
        <div class="panel"><h2>Điểm cần chú ý</h2><ul class="insights" style="margin-top:10px">${ins.length ? ins.map(([s, t]) => `<li>${sevPill(s)}<span>${esc(t)}</span></li>`).join('') : '<li><span class="pill good">● Ổn</span><span>Không có cảnh báo theo ngưỡng hiện tại.</span></li>'}</ul></div>
      </div>
      <div class="section"><h2>Báo cáo theo năm nghiệp vụ</h2>${mt}</div>`;
  }

  // Báo cáo so sánh theo một chiều
  function segView(dim, title, scopeKey, extraCtl) {
    const m = state.main; const sc = state.scope[scopeKey] || '';
    const bf = sc ? (p) => p.uy === sc : null;
    const { rows, all } = E.segment(m, dim, bf);
    const chartRows = rows.slice(0, 14).map((r) => ({ cat: r.cat, s: r.total }));
    const cols = [{ l: title, v: (r) => r.cat, wrap: true }].concat(metricCols((r) => r.total, all, true));
    // so sánh giữa các UY
    const uys = sc ? [sc] : m.uys;
    const pcols = [{ l: title, v: (r) => r.cat, wrap: true }];
    uys.forEach((u) => {
      pcols.push({ l: `GWP ${u}`, v: (r) => r.byUy[u].gwp, f: fN, num: true });
      pcols.push({ l: `CP KT % ${u}`, v: (r) => r.byUy[u].acqPct, f: fP, num: true });
      pcols.push({ l: `BT % ${u}`, v: (r) => r.byUy[u].claimsPct, f: fP, num: true });
      pcols.push({ l: `Kết hợp % ${u}`, v: (r) => r.byUy[u].combinedPct, h: (r) => ratioCell(r.byUy[u].combinedPct), num: true });
    });
    if (uys.length === 2) pcols.push({ l: `Δ GWP ${uys[1]}/${uys[0]}`, v: (r) => r.byUy[uys[1]].gwp - r.byUy[uys[0]].gwp, f: fDelta, num: true });
    return `<div class="toolbar">${uySelect('scope-' + scopeKey, sc)}${extraCtl || ''}<span class="muted small">${rows.length} nhóm · GWP ${fN(all.gwp)} USD · kết hợp ${fP(all.combinedPct)}</span></div>
      <div class="panel"><h2>Tỷ lệ kết hợp theo ${esc(title.toLowerCase())}${rows.length > 14 ? ' (14 nhóm GWP lớn nhất)' : ''}</h2>${comboChart(chartRows, '% trên GWP; vạch dọc là 100%.')}</div>
      <div class="section"><h2>Bảng so sánh ${sc ? '– ' + esc(sc) : '– tất cả UY'}</h2>${table(cols, rows, { foot: true, groups: GROUPS(title, 1), sort: { i: 2, dir: -1 } })}</div>
      ${uys.length > 1 ? `<div class="section"><h2>So sánh giữa các năm nghiệp vụ</h2>${legend(uys)}${table(pcols, rows, { sort: { i: 1, dir: -1 } })}</div>` : ''}`;
  }
  function channelView() {
    const ctl = `<label class="fld" for="channelDim">Phân theo<select id="channelDim"><option value="channel" ${state.channelDim === 'channel' ? 'selected' : ''}>Kênh (A1, MAAPS, Open market)</option><option value="source" ${state.channelDim === 'source' ? 'selected' : ''}>Nguồn (Major PVI / Non-major Alif)</option></select></label>`;
    return segView(state.channelDim, state.channelDim === 'source' ? 'Nguồn' : 'Kênh khai thác', 'channel', ctl);
  }
  function regionView() {
    const ctl = `<label class="fld" for="regionDim">Phân theo<select id="regionDim"><option value="region" ${state.regionDim === 'region' ? 'selected' : ''}>Khu vực</option><option value="domicile" ${state.regionDim === 'domicile' ? 'selected' : ''}>Quốc gia</option></select></label>`;
    return segView(state.regionDim, state.regionDim === 'region' ? 'Khu vực' : 'Quốc gia', 'region', ctl);
  }

  function period() {
    const m = state.main; const uys = m.uys;
    const S = uys.map((u) => ({ cat: u, s: E.summarize(m, (p) => p.uy === u) }));
    const keys = [['Số HĐ', 'n', fN], ['GWP', 'gwp', fN], ['NWP', 'nwp', fN], ['Chi phí KT', 'acqUsd', fN], ['CP KT % GWP', 'acqPct', fP], ['BT đã trả', 'paid', fN], ['BT dự phòng', 'os', fN], ['Bồi thường', 'incurred', fN], ['BT % GWP', 'claimsPct', fP], ['Kết hợp', 'combinedUsd', fN], ['Kết hợp % GWP', 'combinedPct', fP]];
    const isPct = (k) => /Pct$/.test(k);
    // A. giữa các UY
    let ta = `<div class="tbl"><table><thead><tr><th>Chỉ tiêu</th>${S.map((x) => `<th class="r">${esc(x.cat)}</th>`).join('')}${S.length > 1 ? S.slice(1).map((x, i) => `<th class="r">${esc(x.cat)} so với ${esc(S[i].cat)}</th>`).join('') : ''}</tr></thead><tbody>`;
    keys.forEach(([l, k, f]) => {
      ta += `<tr><td>${l}</td>${S.map((x) => `<td class="r">${k === 'combinedPct' ? ratioCell(x.s[k]) : f(x.s[k])}</td>`).join('')}`;
      ta += S.slice(1).map((x, i) => { const a = S[i].s[k], b = x.s[k]; if (isPct(k)) return `<td class="r">${fPP(b - a)}</td>`; return `<td class="r">${fDelta(b - a)}${a ? ` <span class="muted">(${fP((b - a) / a, 0)})</span>` : ''}</td>`; }).join('');
      ta += '</tr>';
    });
    ta += '</tbody></table></div>';
    // B. các lần cập nhật
    const byUy = {}; state.snaps.forEach((x) => { (byUy[x.b.uy] = byUy[x.b.uy] || []).push(x); });
    const multi = Object.values(byUy).some((a) => a.length > 1);
    const srows = []; Object.values(byUy).forEach((arr) => arr.forEach((x, i) => srows.push({ uy: x.b.uy, snap: x.b.snap, file: x.b.file, s: x.s, prev: i ? arr[i - 1].s : null })));
    const scols = [{ l: 'UY', v: (r) => r.uy }, { l: 'Ngày chốt', v: (r) => r.snap, f: fD }]
      .concat(metricCols((r) => r.s, null, false).slice(1))
      .concat([{ l: 'Δ GWP', v: (r) => r.prev ? r.s.gwp - r.prev.gwp : null, f: fDelta, num: true }, { l: 'Δ Bồi thường', v: (r) => r.prev ? r.s.incurred - r.prev.incurred : null, f: fDelta, num: true }, { l: 'Δ Kết hợp', v: (r) => r.prev ? r.s.combinedPct - r.prev.combinedPct : null, f: fPP, num: true }]);
    // C. theo quý hiệu lực
    const q = E.segment(m, 'incQ'); q.rows.sort((a, b) => a.cat.localeCompare(b.cat));
    const qcols = [{ l: 'Quý hiệu lực', v: (r) => r.cat }].concat(metricCols((r) => r.total, q.all, true));
    return `<div class="panel"><h2>So sánh giữa các năm nghiệp vụ</h2>${comboChart(S, 'Cùng ngày báo cáo ' + fD(m.asOf) + '. Năm sau còn ít thời gian phát triển bồi thường hơn năm trước.')}</div>
      <div class="section">${ta}</div>
      <div class="section"><h2>So sánh giữa các lần cập nhật số liệu</h2>
        ${multi ? '' : '<p class="muted small">Mỗi UY hiện chỉ có một bản. Tải thêm bản cập nhật của kỳ trước (ví dụ <code>UY_2_update_05092026.xlsx</code>) để xem số liệu thay đổi giữa các kỳ báo cáo.</p>'}
        ${table(scols, srows, { groups: [['Kỳ', 2], ['Phí', 2], ['Chi phí khai thác', 2], ['Bồi thường', 4], ['Kết hợp', 2], ['Thay đổi so với kỳ trước', 3]] })}</div>
      <div class="section"><h2>Theo quý hiệu lực của hợp đồng</h2>${table(qcols, q.rows, { foot: true, groups: GROUPS('Quý', 1) })}</div>`;
  }

  // Tỷ lệ tái tục: đơn UY trước đã đến hạn có tái tục sang UY sau
  function renewalRatio() {
    const m = state.main; const stats = E.renewalStats(m);
    if (!stats.length) return `<div class="panel"><h2>Tỷ lệ tái tục</h2><p class="muted small">Cần tải ít nhất hai năm nghiệp vụ liên tiếp (ví dụ UY1 và UY2) để tính số đơn tái tục trên tổng số đơn đến hạn.</p></div>`;
    const dim = state.renDim; const dimLab = { broker: 'Môi giới (năm trước)', channel: 'Kênh', region: 'Khu vực', domicile: 'Quốc gia', leader: 'Leader' }[dim];
    return stats.map((st) => {
      const groups = {};
      st.rows.forEach((r) => { const k = r.p[dim] || '(trống)'; const g = groups[k] = groups[k] || { cat: k, nDue: 0, nRen: 0, gDue: 0, gRenPrev: 0, gRenNew: 0, chg: 0 }; g.nDue++; g.gDue += r.p.gwp; if (r.q) { g.nRen++; g.gRenPrev += r.p.gwp; g.gRenNew += r.q.gwp; if (r.q.broker !== r.p.broker) g.chg++; } });
      const rows = Object.values(groups).map((g) => Object.assign(g, { rN: g.nDue ? g.nRen / g.nDue : null, rG: g.gDue ? g.gRenPrev / g.gDue : null, pc: g.gRenPrev ? g.gRenNew / g.gRenPrev - 1 : null }));
      const lost = st.rows.filter((r) => !r.q).map((r) => r.p);
      const chgB = st.rows.filter((r) => r.q && r.q.broker !== r.p.broker);
      const tile = (l, v, sub) => `<div class="panel tile"><div class="lab">${l}</div><div class="big num">${v}</div><div class="small muted">${sub}</div></div>`;
      const pctCell = (v) => !fin(v) ? '–' : `<span class="${v < 0.5 ? 'neg' : v < 0.7 ? 'warnc' : ''}">${fP(v)}</span>`;
      return `<div class="panel"><h2>Tỷ lệ tái tục ${esc(st.from)} → ${esc(st.to)}</h2>
        <p class="muted small">Đơn ${esc(st.from)} đã hết hạn đến ${fD(m.asOf)} (không tính đơn huỷ và sửa đổi bổ sung) có hợp đồng cùng insured trong ${esc(st.to)}, hiệu lực trong ±90 ngày quanh ngày hết hạn.${st.nNotDue ? ` Còn ${st.nNotDue} đơn chưa đến hạn, chưa tính.` : ''}</p>
        <div class="grid tiles" style="margin-top:10px">
          ${tile('Đơn tái tục / đơn đến hạn', `${st.nRenewed}/${st.nDue}`, `${fP(st.ratioN)} số đơn · ${fP(st.ratioG)} GWP`)}
          ${tile('Không tái tục', fN(st.nLost), `GWP năm trước ${fN(st.gwpLost)} USD`)}
          ${tile('Phí đơn tái tục', fPP(st.premChange).replace(' đ%', '%'), `${fN(st.gwpRenewedPrev)} → ${fN(st.gwpRenewedNew)} USD`)}
          ${tile(`Đơn mới trong ${esc(st.to)}`, `${st.nNew}/${st.nNext}`, `${fP(st.newShareN)} số đơn · ${fP(st.newShareG)} GWP`)}
          ${tile('Tái tục qua broker khác', fN(chgB.length), chgB.slice(0, 3).map((r) => `${r.p.insured.slice(0, 16)}: ${r.p.broker}→${r.q.broker}`).join(' · ') || 'không có')}
        </div>
        <div class="toolbar" style="margin-top:14px"><h3 style="margin:0">Theo</h3><label class="fld" for="renDim"><select id="renDim">${[['broker', 'Môi giới'], ['channel', 'Kênh'], ['region', 'Khu vực'], ['domicile', 'Quốc gia'], ['leader', 'Leader']].map(([k, l]) => `<option value="${k}" ${k === dim ? 'selected' : ''}>${l}</option>`).join('')}</select></label></div>
        ${table([{ l: dimLab, v: (r) => r.cat, wrap: true }, { l: 'Đơn đến hạn', v: (r) => r.nDue, f: fN, num: true, t: () => fN(st.nDue) }, { l: 'Đã tái tục', v: (r) => r.nRen, f: fN, num: true, t: () => fN(st.nRenewed) }, { l: 'Tỷ lệ số đơn', v: (r) => r.rN, h: (r) => pctCell(r.rN), num: true, t: () => fP(st.ratioN) },
          { l: 'GWP đến hạn', v: (r) => r.gDue, f: fN, num: true, t: () => fN(st.gwpDue) }, { l: 'GWP đã tái tục (năm trước)', v: (r) => r.gRenPrev, f: fN, num: true, t: () => fN(st.gwpRenewedPrev) }, { l: 'Tỷ lệ GWP', v: (r) => r.rG, h: (r) => pctCell(r.rG), num: true, t: () => fP(st.ratioG) },
          { l: `GWP tái tục (${st.to})`, v: (r) => r.gRenNew, f: fN, num: true, t: () => fN(st.gwpRenewedNew) }, { l: 'Thay đổi phí', v: (r) => r.pc, f: (v) => fPP(v).replace(' đ%', '%'), num: true, t: () => fPP(st.premChange).replace(' đ%', '%') }, { l: 'Đổi broker', v: (r) => r.chg, f: fN, num: true }], rows, { foot: true, sort: { i: 1, dir: -1 } })}
        <h3 style="margin-top:14px">Đơn không tái tục (${lost.length})</h3>
        ${table([{ l: 'Original insured', v: (p) => p.insured, wrap: true }, { l: 'Broker', v: (p) => p.broker }, { l: 'Kênh', v: (p) => p.channel }, { l: 'Quốc gia', v: (p) => p.domicile }, { l: 'Hết hạn', v: (p) => p.expC, f: fD }, { l: 'GWP', v: (p) => p.gwp, f: fN, num: true }, { l: 'Bồi thường', v: (p) => p.incurred, f: fN, num: true }, { l: 'Kết hợp % GWP', v: (p) => p.gwp ? (p.acqUsd + p.incurred) / p.gwp : null, h: (p) => ratioCell(p.gwp ? (p.acqUsd + p.incurred) / p.gwp : null), num: true }], lost, { tall: true, sort: { i: 5, dir: -1 } })}
        <p class="muted small" style="margin-top:6px">Kết hợp thấp nghĩa là mất đơn có lãi, cần tìm hiểu lý do; kết hợp cao thì việc không tái tục có thể là chủ động.</p></div>`;
    }).join('');
  }

  function renew() {
    const m = state.main; const days = state.settings.renewDays || 90;
    const rn = E.renewals(m, days, 30);
    const pend = rn.filter((r) => !r.renewed);
    const sumG = (arr) => arr.reduce((a, r) => a + r.p.gwp, 0);
    const b = (lo, hi) => pend.filter((r) => r.daysLeft >= lo && r.daysLeft <= hi);
    const tiles = [['Đã hết hạn, chưa thấy tái tục', pend.filter((r) => r.daysLeft < 0)], ['Trong 30 ngày', b(0, 30)], ['31–60 ngày', b(31, 60)], [`61–${days} ngày`, b(61, days)], ['Đã có HĐ tái tục', rn.filter((r) => r.renewed)]];
    const pill = (r) => r.renewed ? '<span class="pill good">Đã tái tục</span>' : r.daysLeft < 0 ? '<span class="pill crit">Đã hết hạn</span>' : r.daysLeft <= 30 ? '<span class="pill crit">≤ 30 ngày</span>' : r.daysLeft <= 60 ? '<span class="pill warn">≤ 60 ngày</span>' : '<span class="pill info">≤ ' + days + ' ngày</span>';
    const cols = [
      { l: 'Trạng thái', v: (r) => r.renewed ? 9999 : r.daysLeft, h: pill }, { l: 'Hết hạn', v: (r) => r.p.expC, f: fD }, { l: 'Còn (ngày)', v: (r) => r.daysLeft, f: fN, num: true },
      { l: 'UY', v: (r) => r.p.uy }, { l: 'Original insured', v: (r) => r.p.insured, wrap: true }, { l: 'Broker', v: (r) => r.p.broker }, { l: 'Kênh', v: (r) => r.p.channel }, { l: 'Leader', v: (r) => r.p.leader },
      { l: 'Line', v: (r) => r.p.share, f: (v) => fP(v, 2), num: true }, { l: 'GWP', v: (r) => r.p.gwp, f: fN, num: true }, { l: 'Bồi thường', v: (r) => r.p.incurred, f: fN, num: true },
      { l: 'Kết hợp % GWP', v: (r) => r.comb, h: (r) => ratioCell(r.comb), num: true },
      { l: 'HĐ tái tục', v: (r) => r.renewed ? r.renewed.uy + ' ' + r.renewed.ref : '', h: (r) => r.renewed ? `${esc(r.renewed.uy)} · ${esc(r.renewed.ref)} · từ ${fD(r.renewed.inc)}` : '' },
      { l: 'Lưu ý khi tái tục', v: (r) => r.notes.join('; '), wrap: true },
    ];
    return `${renewalRatio()}
      <h2 style="margin-top:24px">Hợp đồng sắp hết hạn</h2>
      <div class="toolbar"><label class="fld" for="renewDays">Hết hạn trong<select id="renewDays">${[30, 60, 90, 120, 180].map((d) => `<option value="${d}" ${d === days ? 'selected' : ''}>${d} ngày tới</option>`).join('')}</select></label>
        <span class="muted small">Tính từ ngày báo cáo ${fD(m.asOf)}; gồm cả hợp đồng đã hết hạn trong 30 ngày qua. App tự tìm hợp đồng tái tục cùng tên insured có hiệu lực trong ±90 ngày quanh ngày hết hạn, trong các file đã tải.</span></div>
      <div class="grid tiles">${tiles.map(([l, arr]) => `<div class="panel tile"><div class="lab">${l}</div><div class="big num">${arr.length}</div><div class="small muted">GWP ${fN(sumG(arr))} USD</div></div>`).join('')}</div>
      <div class="section">${table(cols, rn, { tall: true, sort: { i: 0, dir: 1 } })}</div>`;
  }

  function claims() {
    const m = state.main; const F = state.claimF;
    const rows = m.claims.filter((c) => (!F.uy || c.uy === F.uy) && (!F.st || c.status === F.st) && (!F.flag || (F.flag === 'any' ? c.flag : c.flag === F.flag)));
    const sts = [...new Set(m.claims.map((c) => c.status).filter(Boolean))];
    const by = {}; rows.forEach((c) => { const k = (c.coverage || '(trống)') + ' · ' + (c.lossType || '(trống)'); by[k] = by[k] || { cat: k, n: 0, paid: 0, os: 0, inc: 0 }; by[k].n++; by[k].paid += c.paidT; by[k].os += c.os; by[k].inc += c.incurred; });
    return `<div class="toolbar">
        <label class="fld" for="cUy">UY<select id="cUy"><option value="">Tất cả</option>${m.uys.map((u) => `<option ${F.uy === u ? 'selected' : ''}>${esc(u)}</option>`).join('')}</select></label>
        <label class="fld" for="cSt">Trạng thái<select id="cSt"><option value="">Tất cả</option>${sts.map((s) => `<option ${F.st === s ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></label>
        <label class="fld" for="cFl">Kiểm tra kỳ<select id="cFl"><option value="">Tất cả</option><option value="any" ${F.flag === 'any' ? 'selected' : ''}>Có cảnh báo</option><option ${F.flag === 'Ngoài kỳ HĐ' ? 'selected' : ''}>Ngoài kỳ HĐ</option><option ${F.flag === 'Ngày không hợp lệ' ? 'selected' : ''}>Ngày không hợp lệ</option></select></label>
        <span class="muted small">${rows.length} hồ sơ · đã trả ${fN(rows.reduce((a, c) => a + c.paidT, 0))} · dự phòng ${fN(rows.reduce((a, c) => a + c.os, 0))} · tổng ${fN(rows.reduce((a, c) => a + c.incurred, 0))} USD</span></div>
      ${table([
        { l: 'UY', v: (c) => c.uy }, { l: 'Số hồ sơ', v: (c) => c.claimNo }, { l: 'Original insured', v: (c) => c.insured, wrap: true }, { l: 'Broker', v: (c) => c.broker }, { l: 'Coverage', v: (c) => c.coverage }, { l: 'Loại', v: (c) => c.lossType },
        { l: 'Ngày tổn thất', v: (c) => c.occ ?? -1, h: (c) => c.occ != null ? fD(c.occ) : esc(c.occRaw) }, { l: 'Gross 100%', v: (c) => c.gross100, f: fN, num: true }, { l: 'Line', v: (c) => c.share, f: (v) => fP(v, 2), num: true },
        { l: 'Đã trả', v: (c) => c.paidT, f: fN2, num: true }, { l: 'Dự phòng', v: (c) => c.os, f: fN2, num: true }, { l: 'Bồi thường', v: (c) => c.incurred, f: fN2, num: true }, { l: 'Status', v: (c) => c.status },
        { l: 'Kiểm tra', v: (c) => c.flag, h: (c) => c.flag ? `<span class="pill crit">${esc(c.flag)}</span>${c.suggest ? ` <span class="muted small">→ ${esc(c.suggest)}</span>` : ''}` : '' }],
        rows, { tall: true, sort: { i: 11, dir: -1 } })}
      <div class="section"><h2>Theo nghiệp vụ và loại tổn thất</h2>${table([{ l: 'Nhóm', v: (r) => r.cat }, { l: 'Số hồ sơ', v: (r) => r.n, f: fN, num: true }, { l: 'Đã trả', v: (r) => r.paid, f: fN, num: true }, { l: 'Dự phòng', v: (r) => r.os, f: fN, num: true }, { l: 'Bồi thường', v: (r) => r.inc, f: fN, num: true }], Object.values(by), { sort: { i: 4, dir: -1 } })}</div>`;
  }

  function collections() {
    const m = state.main; const S = m.settings; const up = S.upcomingDays || 30; const sc = state.scope.coll || ''; const dim = state.collDim || 'broker';
    const inScope = (x) => !sc || x.uy === sc;
    const T = E.summarize(m, (p) => inScope(p));
    const uys = sc ? [sc] : m.uys;
    const B = ['0-30', '31-90', '91-180', '181-365', 'Trên 365'];
    const ag = B.map((b) => { const o = { cat: b }; uys.forEach((u) => { o[u] = m.insts.filter((i) => i.uy === u && i.bucket === b).reduce((a, i) => a + i.amount, 0); }); o.t = uys.reduce((a, u) => a + o[u], 0); return o; });
    const keyOk = new Set(m.policies.filter((p) => p.bound_ && inScope(p)).map((p) => p.key));
    const od = m.insts.filter((i) => i.days > 0 && keyOk.has(i.key));
    const upc = m.insts.filter((i) => i.daysToDue !== null && i.daysToDue <= up && keyOk.has(i.key));
    const pol = Object.fromEntries(m.policies.map((p) => [p.key, p]));
    const cc = (get, foot) => {
      const c = (l, k, f, tip) => ({ l, v: (r) => get(r)[k], f, num: true, t: foot ? () => f(foot[k]) : null, tip });
      return [c('GWP', 'gwp', fN), c('NWP', 'nwp', fN), c('Phí theo lịch', 'schedAmt', fN, 'Tổng các kỳ phí trong sheet Installment'), c('Đã đến hạn', 'dueAmt', fN, 'Các kỳ có hạn thanh toán đến ngày báo cáo'),
        c('Thực thu', 'collected', fN), c('Thu % GWP', 'collectedPctGwp', fP), c('Thu % NWP', 'collectedPctNwp', fP), c('Thu % đến hạn', 'collectedPctDue', fP),
        { l: 'Quá hạn', v: (r) => get(r).overdue, h: (r) => get(r).overdue > 0 ? `<span class="neg">${fN(get(r).overdue)}</span>` : fN(get(r).overdue), num: true, t: foot ? () => fN(foot.overdue) : null },
        c(`Đến hạn ≤ ${up} ngày`, 'upcoming', fN), c('HĐ chưa có lịch', 'noSchedN', fN), c('NWP chưa có lịch', 'noSchedNwp', fN)];
    };
    const byUy = uys.map((u) => ({ cat: u, s: E.summarize(m, (p) => p.uy === u) }));
    const seg = E.segment(m, dim, sc ? (p) => p.uy === sc : null);
    const dimLab = { broker: 'Môi giới', channel: 'Kênh', region: 'Khu vực', domicile: 'Quốc gia' }[dim];
    const tile = (l, v, sub, cls) => `<div class="panel tile ${cls || ''}"><div class="lab">${l}</div><div class="big num">${v}</div><div class="small muted">${sub}</div></div>`;
    return `<div class="toolbar">${uySelect('scope-coll', sc)}
        <label class="fld" for="upDays">Nhắc phí sắp đến hạn trong<select id="upDays">${[7, 15, 30, 60, 90].map((d) => `<option value="${d}" ${d === up ? 'selected' : ''}>${d} ngày tới</option>`).join('')}</select></label>
        <span class="muted small">Tính đến ngày báo cáo ${fD(m.asOf)}. Thực thu = Actual NWP của kỳ ghi "Paid" (nếu chưa nhập thì lấy số tiền kỳ).</span></div>
      <div class="grid tiles">
        ${tile('Thực thu', fN(T.collected), `${fP(T.collectedPctGwp)} GWP · ${fP(T.collectedPctNwp)} NWP`)}
        ${tile('Đã đến hạn', fN(T.dueAmt), `đã thu ${fP(T.collectedPctDue)} số đến hạn`)}
        ${tile('Quá hạn chưa thu', fN(T.overdue), `${T.nOverdue} kỳ · ${fP(T.overduePctNwp)} NWP`, T.overdue > 0 ? 'alert' : '')}
        ${tile(`Sắp đến hạn (≤ ${up} ngày)`, fN(T.upcoming), `${T.nUpcoming} kỳ`, T.upcoming > 0 ? 'soon' : '')}
        ${tile('Chưa có lịch thanh toán', fN(T.noSchedNwp), `${T.noSchedN} HĐ · NWP chưa theo dõi được`)}
      </div>
      <div class="section"><h2>Thu phí theo năm nghiệp vụ</h2>${table([{ l: 'UY', v: (r) => r.cat }].concat(cc((r) => r.s, uys.length > 1 ? T : null)), byUy, { foot: uys.length > 1, groups: [['', 1], ['Doanh thu', 2], ['Kỳ phí', 2], ['Thực thu', 4], ['Cần thu', 2], ['Chưa theo dõi', 2]] })}</div>
      <div class="section"><div class="toolbar"><h2 style="margin:0">Thu phí theo</h2><label class="fld" for="collDim"><select id="collDim">${[['broker', 'Môi giới'], ['channel', 'Kênh'], ['region', 'Khu vực'], ['domicile', 'Quốc gia']].map(([k, l]) => `<option value="${k}" ${k === dim ? 'selected' : ''}>${l}</option>`).join('')}</select></label></div>
        ${table([{ l: dimLab, v: (r) => r.cat, wrap: true }].concat(cc((r) => r.total, seg.all)), seg.rows, { foot: true, sort: { i: 9, dir: -1 } })}</div>
      <div class="grid section">
        <div class="panel"><h2>Sắp đến kỳ thu (${upc.length})</h2>${table([{ l: 'Hạn', v: (i) => i.due, f: fD }, { l: 'Còn (ngày)', v: (i) => i.daysToDue, f: fN, num: true }, { l: 'UY', v: (i) => i.uy }, { l: 'Original insured', v: (i) => i.insured, wrap: true }, { l: 'Broker', v: (i) => i.broker }, { l: 'Kỳ', v: (i) => i.n, num: true }, { l: 'Số tiền', v: (i) => i.amount, f: fN2, num: true }], upc, { tall: true, sort: { i: 0, dir: 1 } })}</div>
        <div class="panel"><h2>Tuổi nợ phí quá hạn</h2>${table([{ l: 'Số ngày quá hạn', v: (r) => r.cat }].concat(uys.map((u) => ({ l: u, v: (r) => r[u], f: fN, num: true }))).concat(uys.length > 1 ? [{ l: 'Tổng', v: (r) => r.t, f: fN, num: true }] : []), ag)}</div>
      </div>
      <div class="section"><h2>Phí đã quá hạn (${od.length} kỳ · ${fN(od.reduce((a, i) => a + i.amount, 0))} USD)</h2>${table([{ l: 'UY', v: (i) => i.uy }, { l: 'Original insured', v: (i) => i.insured, wrap: true }, { l: 'Broker', v: (i) => i.broker }, { l: 'Kênh', v: (i) => (pol[i.key] || {}).channel || '' }, { l: 'Kỳ', v: (i) => i.n, num: true }, { l: 'Hạn', v: (i) => i.due, f: fD }, { l: 'Số tiền', v: (i) => i.amount, f: fN2, num: true }, { l: 'Ngày quá hạn', v: (i) => i.days, h: (i) => `<span class="${i.days > S.overdueAlertDays ? 'neg' : 'warnc'}">${fN(i.days)}</span>`, num: true }, { l: 'Nhóm', v: (i) => i.bucket }], od, { tall: true, sort: { i: 7, dir: -1 } })}</div>`;
  }

  function risk() {
    const m = state.main; const S = m.settings; const bp = m.policies.filter((p) => p.bound_);
    const over = bp.filter((p) => p.flags.some((f) => f.startsWith('Line '))); const hr = bp.filter((p) => p.highRisk); const late = bp.filter((p) => p.flags.includes('Bind sau hiệu lực'));
    const lead = []; m.uys.forEach((u) => { const q = bp.filter((p) => p.uy === u); const tot = q.reduce((a, p) => a + p.gwp, 0); const inc = q.reduce((a, p) => a + p.incurred, 0); const by = {}; q.forEach((p) => { const k = p.leader || '(trống)'; by[k] = by[k] || { u, l: k, gwp: 0, inc: 0, n: 0 }; by[k].gwp += p.gwp; by[k].inc += p.incurred; by[k].n++; }); Object.values(by).forEach((r) => { r.sg = tot ? r.gwp / tot : 0; r.si = inc ? r.inc / inc : 0; lead.push(r); }); });
    const pc = [{ l: 'UY', v: (p) => p.uy }, { l: 'Original insured', v: (p) => p.insured, wrap: true }, { l: 'Quốc gia', v: (p) => p.domicile }, { l: 'Leader', v: (p) => p.leader }, { l: 'Số tàu', v: (p) => p.ac, f: fN, num: true }, { l: 'Line', v: (p) => p.share, f: (v) => fP(v, 2), num: true }, { l: 'Hull PVI', v: (p) => p.share * p.hull, f: fN, num: true }, { l: 'GWP', v: (p) => p.gwp, f: fN, num: true }, { l: 'Bồi thường', v: (p) => p.incurred, f: fN, num: true }];
    return `<p class="muted small">Trần line ${fP(S.lineCapGeneral, 2)} (chung), ${fP(S.lineCapSmall, 2)} (hãng dưới ${S.smallFleet} tàu hoặc quốc gia rủi ro cao) · leader tối đa ${fP(S.leaderCap, 0)} GWP. Đổi ở tab Thiết lập.</p>
      <div class="section"><h2>Line vượt trần (${over.length})</h2>${table(pc, over, { tall: true, sort: { i: 6, dir: -1 } })}</div>
      <div class="section"><h2>Tập trung leader</h2>${table([{ l: 'UY', v: (r) => r.u }, { l: 'Leader', v: (r) => r.l }, { l: 'Số HĐ', v: (r) => r.n, num: true }, { l: 'GWP', v: (r) => r.gwp, f: fN, num: true }, { l: '% GWP', v: (r) => r.sg, h: (r) => `${fP(r.sg)}${r.sg > S.leaderCap ? ' <span class="pill warn">vượt trần</span>' : ''}`, num: true }, { l: 'Bồi thường', v: (r) => r.inc, f: fN, num: true }, { l: '% bồi thường', v: (r) => r.si, f: fP, num: true }], lead, { tall: true, sort: { i: 3, dir: -1 } })}</div>
      <div class="section"><h2>Quốc gia nhạy cảm cấm vận (${hr.length})</h2>${table(pc.concat([{ l: 'Cedant', v: (p) => p.cedant, wrap: true }]), hr, { tall: true })}</div>
      <div class="section"><h2>Bind sau ngày hiệu lực (${late.length})</h2>${table([{ l: 'UY', v: (p) => p.uy }, { l: 'Original insured', v: (p) => p.insured, wrap: true }, { l: 'Hiệu lực', v: (p) => p.inc, f: fD }, { l: 'Bind', v: (p) => p.bound, f: fD }, { l: 'Số ngày trễ', v: (p) => p.bound - p.inc, num: true }, { l: 'Bồi thường', v: (p) => p.incurred, f: fN, num: true }], late, { sort: { i: 4, dir: -1 } })}</div>`;
  }

  function issues() {
    const m = state.main; const F = state.issueF; const q = F.q.toLowerCase();
    const rows = m.issues.filter((i) => (!F.sev || i.sev === F.sev) && (!F.uy || i.uy === F.uy) && (!q || (i.rec + ' ' + i.msg + ' ' + i.field).toLowerCase().includes(q)));
    const cnt = (s) => m.issues.filter((i) => i.sev === s).length;
    return `<div class="toolbar">
        <label class="fld" for="iSev">Mức<select id="iSev"><option value="">Tất cả (${m.issues.length})</option>${['Cao', 'TB', 'Thấp'].map((s) => `<option value="${s}" ${F.sev === s ? 'selected' : ''}>${s === 'TB' ? 'Trung bình' : s} (${cnt(s)})</option>`).join('')}</select></label>
        <label class="fld" for="iUy">UY<select id="iUy"><option value="">Tất cả</option>${m.uys.map((u) => `<option ${F.uy === u ? 'selected' : ''}>${esc(u)}</option>`).join('')}</select></label>
        <label class="fld" for="iQ">Tìm<input type="search" id="iQ" value="${esc(F.q)}" placeholder="Tên hãng, trường…"></label></div>
      ${table([{ l: 'Mức', v: (i) => ORD[i.sev], h: (i) => sevPill(i.sev) }, { l: 'UY', v: (i) => i.uy }, { l: 'Sheet', v: (i) => i.sheet }, { l: 'Đối tượng', v: (i) => i.rec, wrap: true }, { l: 'Trường', v: (i) => i.field }, { l: 'Vấn đề', v: (i) => i.msg, wrap: true }, { l: 'Hướng xử lý', v: (i) => i.fix, wrap: true }], rows, { tall: true })}`;
  }

  function settings() {
    const S = state.settings;
    const n = (id, l, v, step) => `<label class="fld" for="${id}">${l}<input type="number" id="${id}" value="${v}" step="${step}"></label>`;
    return `<div class="panel"><h2>Ngưỡng cảnh báo</h2><p class="muted small">Lưu trong trình duyệt này; chỉ lưu ngưỡng, không lưu dữ liệu.</p>
      <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr));margin-top:10px">
        ${n('sRenew', 'Nhắc tái tục trước (ngày)', S.renewDays || 90, '1')}
        ${n('sUp', 'Nhắc phí sắp đến hạn trước (ngày)', S.upcomingDays || 30, '1')}
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
    return `<div class="panel"><h2>Các chỉ số</h2>
      <ul class="small" style="max-width:85ch;display:grid;gap:4px">
        <li><b>GWP</b>: PVI deposite GWP. <b>NWP</b>: PVI deposite NWP (sau brokerage, SMB…).</li>
        <li><b>Chi phí khai thác</b> = GWP − net earning = (GWP − NWP) + facility fee + khấu trừ khác + quỹ; tỷ lệ tính trên GWP.</li>
        <li><b>Bồi thường</b> = đã trả (gồm phí) + dự phòng (gồm phí), phần của PVI; tỷ lệ tính trên GWP. Chưa gồm IBNR.</li>
        <li><b>Tỷ lệ kết hợp</b> = (chi phí khai thác + bồi thường) / GWP. Trên 100% là lỗ trước chi phí quản lý và tái bảo hiểm.</li>
        <li><b>Kỳ báo cáo</b>: mỗi file mang một ngày chốt số liệu. Báo cáo chính dùng bản mới nhất của mỗi UY; tab Kỳ báo cáo so sánh giữa các UY, giữa các bản cập nhật và theo quý hiệu lực.</li>
        <li><b>Tái tục</b>: hợp đồng hết hạn trong khoảng đã chọn kể từ ngày báo cáo; app tự tìm hợp đồng tái tục cùng tên insured có hiệu lực quanh ngày hết hạn.</li>
      </ul>
      <h2 style="margin-top:18px">Định dạng file đầu vào</h2>
      <p class="small">Mỗi file một UY, các sheet <code>Policy</code>, <code>Risk</code>, <code>Premium</code>, <code>Claim</code>, <code>Installment</code>; tên cột như file theo dõi hiện tại (không phân biệt hoa thường). <code>Policy</code>, <code>Risk</code>, <code>Premium</code> phải cùng thứ tự dòng. Ngày dạng <code>yyyy-mm-dd</code>, <code>dd/mm/yyyy</code> hoặc ô ngày Excel. Tên file nên có dạng <code>UY_2_update_ddmmyyyy.xlsx</code>.</p>
      <h2 style="margin-top:18px">Bảo mật</h2>
      <p class="small">File được đọc và tính ngay trong trình duyệt; không gửi dữ liệu đi đâu. Không đưa file dữ liệu thật lên repository.</p></div>`;
  }

  // ---------- Xuất Excel ----------
  function exportWorkbook() {
    const m = state.main; if (!m) return;
    const wb = XLSX.utils.book_new();
    const add = (name, rows) => { if (!rows.length) rows = [{ ' ': 'Không có dữ liệu' }]; const ws = XLSX.utils.json_to_sheet(rows); ws['!cols'] = Object.keys(rows[0]).map((k) => ({ wch: Math.min(46, Math.max(10, k.length + 2)) })); XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31)); };
    const d = (x) => x == null ? '' : fD(x); const r2 = (x) => fin(x) ? Math.round(x * 100) / 100 : ''; const r4 = (x) => fin(x) ? Math.round(x * 10000) / 10000 : '';
    const met = (s) => ({ 'Số HĐ': s.n, GWP: r2(s.gwp), NWP: r2(s.nwp), 'Chi phí KT': r2(s.acqUsd), 'CP KT % GWP': r4(s.acqPct), 'BT đã trả': r2(s.paid), 'BT dự phòng': r2(s.os), 'Bồi thường': r2(s.incurred), 'BT % GWP': r4(s.claimsPct), 'Kết hợp': r2(s.combinedUsd), 'Kết hợp % GWP': r4(s.combinedPct) });
    const uyRows = m.uys.map((u) => Object.assign({ UY: u }, met(E.summarize(m, (p) => p.uy === u)))); uyRows.push(Object.assign({ UY: 'Tổng' }, met(E.summarize(m))));
    uyRows.push({}); uyRows.push({ UY: 'Ngày báo cáo', 'Số HĐ': fD(m.asOf) });
    add('Tong quan', uyRows);
    const segRows = (dim, lab) => { const { rows, all } = E.segment(m, dim); const out = rows.map((r) => { const o = Object.assign({ [lab]: r.cat }, met(r.total), { 'Tỷ trọng GWP': r4(r.shareGwp) }); m.uys.forEach((u) => { o[`GWP ${u}`] = r2(r.byUy[u].gwp); o[`Kết hợp % ${u}`] = r4(r.byUy[u].combinedPct); }); return o; }); out.push(Object.assign({ [lab]: 'Tổng' }, met(all))); return out; };
    add('Moi gioi', segRows('broker', 'Broker')); add('Kenh', segRows('channel', 'Kênh')); add('Khu vuc', segRows('region', 'Khu vực')); add('Quoc gia', segRows('domicile', 'Quốc gia')); add('Quy hieu luc', segRows('incQ', 'Quý'));
    add('Ky bao cao', state.snaps.map((x) => Object.assign({ UY: x.b.uy, 'Ngày chốt': d(x.b.snap), File: x.b.file }, met(x.s))));
    add('Tai tuc', E.renewals(m, state.settings.renewDays || 90, 30).map((r) => ({ 'Trạng thái': r.status, 'Hết hạn': d(r.p.expC), 'Còn (ngày)': r.daysLeft, UY: r.p.uy, 'PVI ref': r.p.ref, 'Original insured': r.p.insured, Broker: r.p.broker, Kênh: r.p.channel, Leader: r.p.leader, Line: r4(r.p.share), GWP: r2(r.p.gwp), 'Bồi thường': r2(r.p.incurred), 'Kết hợp % GWP': r4(r.comb), 'HĐ tái tục': r.renewed ? `${r.renewed.uy} ${r.renewed.ref}` : '', 'Lưu ý': r.notes.join('; ') })));
    const cm = (s) => ({ GWP: r2(s.gwp), NWP: r2(s.nwp), 'Phí theo lịch': r2(s.schedAmt), 'Đã đến hạn': r2(s.dueAmt), 'Thực thu': r2(s.collected), 'Thu % GWP': r4(s.collectedPctGwp), 'Thu % NWP': r4(s.collectedPctNwp), 'Thu % đến hạn': r4(s.collectedPctDue), 'Quá hạn': r2(s.overdue), 'Sắp đến hạn': r2(s.upcoming), 'HĐ chưa có lịch': s.noSchedN, 'NWP chưa có lịch': r2(s.noSchedNwp) });
    const coll = m.uys.map((u) => Object.assign({ 'Nhóm': u }, cm(E.summarize(m, (p) => p.uy === u)))); coll.push(Object.assign({ 'Nhóm': 'Tổng' }, cm(E.summarize(m)))); coll.push({});
    E.segment(m, 'broker').rows.forEach((r) => coll.push(Object.assign({ 'Nhóm': 'Broker: ' + r.cat }, cm(r.total))));
    add('Thu phi', coll);
    const rs = E.renewalStats(m); const rr = [];
    rs.forEach((st) => { rr.push({ 'Kỳ': `${st.from} → ${st.to}`, 'Đơn đến hạn': st.nDue, 'Đã tái tục': st.nRenewed, 'Tỷ lệ số đơn': r4(st.ratioN), 'GWP đến hạn': r2(st.gwpDue), 'GWP đã tái tục (năm trước)': r2(st.gwpRenewedPrev), 'Tỷ lệ GWP': r4(st.ratioG), 'GWP tái tục (năm sau)': r2(st.gwpRenewedNew), 'Thay đổi phí': r4(st.premChange), 'Đơn mới': st.nNew, 'Tổng đơn năm sau': st.nNext, 'GWP đơn mới': r2(st.gwpNew) }); });
    add('Ty le tai tuc', rr);
    const rl = []; rs.forEach((st) => st.rows.forEach((r) => rl.push({ 'Kỳ': `${st.from} → ${st.to}`, 'Trạng thái': r.q ? 'Đã tái tục' : 'Không tái tục', 'Original insured': r.p.insured, Broker: r.p.broker, 'Hết hạn': d(r.p.expC), 'GWP năm trước': r2(r.p.gwp), 'HĐ tái tục': r.q ? r.q.ref : '', 'Broker mới': r.q ? r.q.broker : '', 'GWP tái tục': r.q ? r2(r.q.gwp) : '' })));
    add('Chi tiet tai tuc', rl);
    add('Sap den han', m.insts.filter((i) => i.daysToDue !== null && i.daysToDue <= (state.settings.upcomingDays || 30)).sort((a, b) => a.due - b.due).map((i) => ({ 'Hạn': d(i.due), 'Còn (ngày)': i.daysToDue, UY: i.uy, 'Original insured': i.insured, Broker: i.broker, 'Kỳ': i.n, 'Số tiền': r2(i.amount) })));
    add('Qua han', m.insts.filter((i) => i.days > 0).sort((a, b) => b.days - a.days).map((i) => ({ UY: i.uy, 'Original insured': i.insured, Broker: i.broker, 'Kỳ': i.n, 'Hạn': d(i.due), 'Số tiền': r2(i.amount), 'Ngày quá hạn': i.days, 'Nhóm': i.bucket })));
    add('Hop dong', m.policies.map((p) => ({ UY: p.uy, 'PVI ref': p.ref, MRO: p.mro, Broker: p.broker, 'Original insured': p.insured, Region: p.region, Domicile: p.domicile, Channel: p.channel, Leader: p.leader, 'Nguồn': p.source, 'Prior PVI ref': p.priorRef, 'Endorsement of': p.endtOf, Status: p.status, 'Hiệu lực': d(p.inc), 'Hết hạn': d(p.exp), 'PVI share': r4(p.share), GWP: r2(p.gwp), NWP: r2(p.nwp), 'Chi phí KT': r2(p.acqUsd), 'CP KT %': r4(p.acq), 'Đã trả': r2(p.paidT), 'Dự phòng': r2(p.os), 'Bồi thường': r2(p.incurred), 'Kết hợp % GWP': p.gwp ? r4((p.acqUsd + p.incurred) / p.gwp) : '', 'Cờ': p.flags.join('; ') })));
    add('Boi thuong', m.claims.map((c) => ({ UY: c.uy, 'Số hồ sơ': c.claimNo, MRO: c.mro, 'Original insured': c.insured, Coverage: c.coverage, 'Loại': c.lossType, 'Ngày tổn thất': c.occ != null ? fD(c.occ) : String(c.occRaw ?? ''), 'Đã trả': r2(c.paidT), 'Dự phòng': r2(c.os), 'Bồi thường': r2(c.incurred), Status: c.status, 'Kiểm tra': c.flag })));
    add('Ky thanh toan', m.insts.map((i) => ({ UY: i.uy, MRO: i.mro, Broker: i.broker, 'Original insured': i.insured, 'Kỳ': i.n, 'Số tiền': r2(i.amount), 'Hạn': d(i.due), 'Trạng thái': i.paid, 'Thực thu': r2(i.collected), 'Ngày quá hạn': i.days, 'Nhóm': i.bucket })));
    add('Kiem tra du lieu', m.issues.map((i, k) => ({ '#': k + 1, 'Mức': i.sev, UY: i.uy, Sheet: i.sheet, 'Đối tượng': i.rec, 'Trường': i.field, 'Vấn đề': i.msg, 'Hướng xử lý': i.fix, 'Đã xử lý?': '' })));
    const dt = new Date(m.asOf * E.DAY); const stamp = `${dt.getUTCFullYear()}${String(dt.getUTCMonth() + 1).padStart(2, '0')}${String(dt.getUTCDate()).padStart(2, '0')}`;
    saveFile(`Bao_cao_TBH_hang_khong_${stamp}.xlsx`, new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }
  async function saveFile(name, blob) {
    if (window.claude && typeof window.claude.use === 'function') {
      try { const dl = await window.claude.use('downloads'); if (dl) { await dl.save({ filename: name, data: blob }); toast('Đã lưu ' + name); return; } } catch (e) { if (e && e.code === 'declined') return; }
    }
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  async function exportTemplate() {
    try {
      const r = await fetch('templates/Mau_nhap_lieu_UY_x_update_ddmmyyyy.xlsx');
      if (r.ok) { saveFile('Mau_nhap_lieu_UY_x_update_ddmmyyyy.xlsx', await r.blob()); return; }
    } catch (e) { /* chạy ngoài GitHub Pages: dùng mẫu rút gọn */ }
    exportTemplateBasic();
  }
  function exportTemplateBasic() {
    const wb = XLSX.utils.book_new(); const ex = 'VÍ DỤ – xoá dòng này';
    const sheets = {
      Policy: [['MRO', 'Broker', 'PIAS Ref', 'PVI ref number', 'Original Insured', 'Cedant Name', 'Region', 'Domicile', 'Risk offer date', 'Status', 'Risk bound date', 'Inception', 'Expiry', 'Flying/Non-flying', 'Leader', 'Risk type', 'Type', 'Category', 'Channel'], ['B0000EXAMPLE01', 'AJG', '', 'PVI202600001', ex, 'Example Insurance Co', 'Asia Pacific RIM', 'Vietnam', '2026-09-01', 'Bound', '2026-09-20', '2026-10-01', '2027-09-30', 'Flying', 'AXA XL', 'HSL', 'RI', '', 'A1']],
      Risk: [['MRO', 'Broker', 'Original Insured', 'PVI share', 'Capacity deployment', 'Hull', 'Spare limit', 'CSL', 'PA', 'AVN52E', 'A/C', 'AFV', 'Pax', 'Dept', 'Spare value'], ['B0000EXAMPLE01', 'AJG', ex, 0.5, 5500000, 100000000, 0, 1000000000, 0, 0, 12, 400000000, 0, 0, 0]],
      Premium: [['MRO', 'Broker', 'Original Insured', 'Region', 'Domicile', 'PVI share', 'Deposit GWP', 'Deposit NWP', 'PVI deposite GWP', 'PVI deposite NWP', 'Brokerage', 'SMB', 'Total deduction', 'Facility fee', 'Other Deduction', 'Fund', 'PVI actual earning', 'Diffirental from lead', 'Facility', 'Month', 'Year', 'Policy status', 'Premium status'], ['B0000EXAMPLE01', 'AJG', ex, 'Asia Pacific RIM', 'Vietnam', 0.5, 2000000, 1500000, 10000, 7500, 21.5, 3.5, 25, 0, 0, 0, 7500, 0, 'A1', 10, 2026, 'Bound', 'Estimated']],
      Claim: [['MRO', 'Broker', 'Original Insured', 'Claim number', 'Date claim reported', 'Occurrence date', 'Location', 'Loss details', 'Coverage', 'Type of loss', 'Loss adjuster/ Lawyer', 'Paid Claim (100%)', 'Paid Fee (100%)', 'Reserved claim', 'Resered Fee', 'Gross Incurred', 'PVI share', 'PVI Paid Claim', 'PVI Paid Fee', 'PVI reserved claim', 'PVI reserved Fee', "PVI's Gross Incured", 'Status', 'Lastest Movement Date', 'Refund amout', 'Paid Status'], ['B0000EXAMPLE01', 'AJG', ex, 'PVIAVC2026000001', '2026-11-05', '2026-11-01', '', 'Mô tả tổn thất', 'Hull', 'Engine', '', 0, 0, 500000, 0, 500000, 0.5, 0, 0, 2500, 0, 2500, 'Open', '2026-11-05', 0, '']],
    };
    const ih = ['MRO', 'Broker', 'Original Insured', 'Closing Ref', 'Closing Date']; for (let n = 1; n <= 12; n++) ih.push(`Inst ${n} Date`, `Inst ${n} Share %`, `Inst ${n} Amount`, `Inst ${n} Due Date`, `Inst ${n} Paid`, `Inst ${n} Actual NWP`); ih.push('Last Movement Date');
    sheets.Installment = [ih, ['B0000EXAMPLE01', 'AJG', ex, '', '', '2026-10-01', 0, 3750, '2026-11-01', 'Paid', 3750, '2027-04-01', 0, 3750, '2027-05-01', '', 0]];
    Object.entries(sheets).forEach(([n, rows]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), n));
    saveFile('UY_x_update_ddmmyyyy_mau.xlsx', new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }

  // ---------- Sự kiện ----------
  function setAsOfInput() { $('#asof').value = new Date(state.settings.asOf * E.DAY).toISOString().slice(0, 10); }
  $('#pick').addEventListener('click', () => $('#file').click());
  $('#file').addEventListener('change', (e) => { addFiles([...e.target.files]); e.target.value = ''; });
  const drop = $('#drop');
  ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => { const f = [...(e.dataTransfer.files || [])].filter((x) => /\.xls[xm]?$/i.test(x.name)); if (f.length) addFiles(f); });
  $('#asof').addEventListener('change', (e) => { const d = E.parseDate(e.target.value); if (d != null) { state.settings.asOf = d; state.asOfManual = true; recompute(); } });
  $('#export').addEventListener('click', exportWorkbook);
  $('#template').addEventListener('click', exportTemplate);
  $('#files').addEventListener('click', (e) => { const i = e.target.getAttribute('data-rm'); if (i != null) { state.books.splice(+i, 1); recompute(); } });
  $('#files').addEventListener('change', (e) => {
    const i = e.target.getAttribute('data-uy'); if (i != null) { state.books[+i].uy = e.target.value.trim() || state.books[+i].uy; recompute(); }
    const j = e.target.getAttribute('data-snap'); if (j != null) { const d = E.parseDate(e.target.value); state.books[+j].snap = d; recompute(); }
  });
  $('#tabs').addEventListener('click', (e) => { const b = e.target.closest('[data-tab]'); if (b) go(b.dataset.tab); });
  function go(t) { state.tab = t; try { localStorage.setItem('kpi-tab', t); } catch (x) { /* */ } renderTabs(); render(); }
  const view = $('#view');
  view.addEventListener('click', (e) => {
    const th = e.target.closest('th[data-sort]');
    if (th) { const box = th.closest('.tbl'); const t = TABLES[box.id]; const i = +th.dataset.sort; const cur = t.opts.sort; t.opts.sort = { i, dir: cur && cur.i === i ? -cur.dir : (t.cols[i].num ? -1 : 1) }; box.outerHTML = table(t.cols, t.rows, t.opts); return; }
    const g = e.target.closest('[data-goto]'); if (g) { go(g.dataset.goto); return; }
    if (e.target.id === 'sSave') {
      const S = state.settings; const v = (id) => parseFloat($('#' + id).value);
      S.renewDays = v('sRenew') || 90; S.upcomingDays = v('sUp') || 30; S.lineCapGeneral = v('sGen') / 100; S.lineCapSmall = v('sSmall') / 100; S.smallFleet = v('sFleet'); S.leaderCap = v('sLead') / 100; S.minPremium = v('sMin'); S.overdueAlertDays = v('sOd');
      S.highRiskCountries = $('#sHigh').value.split(',').map((s) => s.trim()).filter(Boolean);
      saveSettings(); recompute(); toast('Đã áp dụng ngưỡng mới');
    }
    if (e.target.id === 'sReset') { const a = state.settings.asOf; state.settings = Object.assign({}, E.DEFAULT_SETTINGS, { asOf: a, renewDays: 90 }); saveSettings(); recompute(); toast('Đã khôi phục mặc định'); }
  });
  view.addEventListener('change', (e) => {
    const id = e.target.id; const v = e.target.value;
    if (id.startsWith('scope-')) { state.scope[id.slice(6)] = v; render(); }
    if (id === 'regionDim') { state.regionDim = v; render(); }
    if (id === 'channelDim') { state.channelDim = v; render(); }
    if (id === 'collDim') { state.collDim = v; render(); }
    if (id === 'renDim') { state.renDim = v; render(); }
    if (id === 'upDays') { state.settings.upcomingDays = +v; saveSettings(); recompute(); }
    if (id === 'renewDays') { state.settings.renewDays = +v; saveSettings(); renderTabs(); render(); }
    if (id === 'cUy') { state.claimF.uy = v; render(); }
    if (id === 'cSt') { state.claimF.st = v; render(); }
    if (id === 'cFl') { state.claimF.flag = v; render(); }
    if (id === 'iSev') { state.issueF.sev = v; render(); }
    if (id === 'iUy') { state.issueF.uy = v; render(); }
  });
  view.addEventListener('input', (e) => { if (e.target.id === 'iQ') { state.issueF.q = e.target.value; const pos = e.target.selectionStart; render(); const q = $('#iQ'); q.focus(); q.setSelectionRange(pos, pos); } });

  try { const t = localStorage.getItem('kpi-tab'); if (t && TABS.some((x) => x[0] === t)) state.tab = t; } catch (e) { /* */ }
  setAsOfInput(); renderTabs(); render();
})();
