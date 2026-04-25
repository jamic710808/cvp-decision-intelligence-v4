/* ╔══════════════════════════════════════════════════════════╗
   ║  CVP V4.0 — 工具函數 & UI 輔助                           ║
   ╚══════════════════════════════════════════════════════════╝ */

/* ──────────────────── 數值格式化 ──────────────────── */
/** 整數（千分位）*/
const fmt0 = v => Math.round(v).toLocaleString('zh-TW');
/** 一位小數 */
const fmt1 = v => (+v).toLocaleString('zh-TW', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
/** 兩位小數（萬元用）*/
const fmt2 = v => (v != null && !isNaN(v)) ? Number(v).toFixed(2) : '—';
/** 百分比一位小數 */
const pct1 = v => (v * 100).toFixed(1) + '%';
/** 百分比零位小數 */
const pct0 = v => Math.round(v * 100) + '%';
/** 現金符號顯示（有正負號）*/
const fmtDelta = v => (v >= 0 ? '+' : '') + fmt2(v);
const fmtDeltaPct = v => (v >= 0 ? '+' : '') + v.toFixed(1) + '%';
/** 短日期 */
const fmtDate = d => `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
/** 完整日期時間 */
const fmtDateTime = d => `${fmtDate(d)} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
/** 安全轉數字 */
const toNum = v => parseFloat(v) || 0;
/** 夾值 */
const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

/* ──────────────────── 顏色判斷 ──────────────────── */
/** 正負決定顏色（純文字 class）*/
const posNegClass = (v, threshold = 0) => v >= threshold ? 'success' : 'danger';
/** 邊際貢獻率色彩評等 */
const cmRatingClass = (cmRate) => {
    const p = AppState.params;
    if (cmRate >= p.cmGoodThreshold) return 'badge-green';
    if (cmRate >= p.industryAvgCM) return 'badge-yellow';
    return 'badge-red';
};
const cmRatingText = (cmRate) => {
    const p = AppState.params;
    if (cmRate >= p.cmGoodThreshold) return '優良';
    if (cmRate >= p.industryAvgCM) return '達標';
    return '待改善';
};

/* ──────────────────── Toast 通知 ──────────────────── */
let _toastTimer = null;
function showToast(msg, type = 'info') {
    const t = document.getElementById('toast');
    if (!t) return;
    clearTimeout(_toastTimer);
    t.textContent = msg;
    t.className = `show ${type}`;
    _toastTimer = setTimeout(() => { t.className = ''; }, 3000);
}

/* ──────────────────── 操作日誌 ──────────────────── */
function logChange(msg) {
    const log = AppState.changeLog;
    log.unshift(`${fmtDateTime(new Date())}　${msg}`);
    if (log.length > 100) log.pop();
    const el = document.getElementById('change-log');
    if (el) el.innerHTML = log.slice(0, 30).map(l => `<div>${l}</div>`).join('');
}

/* ──────────────────── 頁首資訊更新 ──────────────────── */
function updateHeader() {
    const p = AppState.params;
    const compEl = document.getElementById('hdr-company');
    const perEl = document.getElementById('hdr-period');
    const updEl = document.getElementById('last-update');
    if (compEl) compEl.innerHTML = `公司：<span style="color:var(--sky-lt);font-weight:700">${p.company}</span>`;
    if (perEl) perEl.innerHTML = `期間：<span style="color:var(--violet-lt)">${p.period}</span>`;
    if (updEl) updEl.textContent = AppState.lastUpdate ? fmtDateTime(new Date(AppState.lastUpdate)) : fmtDate(new Date());
}

/* ──────────────────── CountUp 動畫 ──────────────────── */
/**
 * 讓目標元素以動畫方式顯示最終數值（簡易計數器）
 * @param {HTMLElement} el - 目標元素
 * @param {string} finalText - 最終文字（如 "1,234.5"）
 */
function animateKPI(el, finalText) {
    if (!el) return;
    el.classList.remove('animating');
    void el.offsetWidth; // reflow
    el.textContent = finalText;
    el.classList.add('animating');
}

/* ──────────────────── Sparkline（迷你折線圖）──────────────────── */
/**
 * 在容器內畫一條 SVG 迷你折線
 * @param {HTMLElement} container - 容器元素
 * @param {number[]} data         - 數據陣列
 * @param {string} color          - 線條顏色
 */
function drawSparkline(container, data, color = '#0ea5e9') {
    if (!container || data.length < 2) return;
    const W = container.clientWidth || 120, H = 30;
    const min = Math.min(...data), max = Math.max(...data);
    const range = max - min || 1;
    const pts = data.map((v, i) => {
        const x = (i / (data.length - 1)) * W;
        const y = H - ((v - min) / range) * (H - 4) - 2;
        return `${x.toFixed(1)},${y.toFixed(1)}`;
    }).join(' ');
    const last = data[data.length - 1];
    const isUp = last >= data[0];
    const finalColor = isUp ? '#10b981' : '#ef4444';
    container.innerHTML = `
    <svg width="${W}" height="${H}" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
      <polyline points="${pts}" fill="none" stroke="${finalColor}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" opacity="0.85"/>
      <circle cx="${pts.split(' ').at(-1).split(',')[0]}" cy="${pts.split(' ').at(-1).split(',')[1]}" r="2.5" fill="${finalColor}"/>
    </svg>`;
}

/* ──────────────────── 頁籤切換 ──────────────────── */
function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
    const tab = document.getElementById(tabId);
    if (tab) tab.classList.add('active');
    const navBtn = document.querySelector(`[data-tab="${tabId}"]`);
    if (navBtn) navBtn.classList.add('active');
    renderTab(tabId);
}

/* ──────────────────── 分類面板手風琴 ──────────────────── */
function toggleNavGroup(groupEl) {
    groupEl.classList.toggle('open');
}
function initNav() {
    document.querySelectorAll('.nav-group-btn').forEach(btn => {
        btn.addEventListener('click', () => toggleNavGroup(btn.closest('.nav-group')));
    });
    // 預設展開第一個分組（基礎分析）
    const first = document.querySelector('.nav-group');
    if (first) first.classList.add('open');
}

/* ──────────────────── 主題切換 ──────────────────── */
function toggleTheme() {
    AppState.theme = AppState.theme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', AppState.theme);
    const btn = document.getElementById('theme-btn');
    if (btn) btn.textContent = AppState.theme === 'dark' ? '☀ 亮色' : '🌙 暗色';
    saveState();
}

/* ──────────────────── 行動裝置開關側邊欄 ──────────────────── */
function toggleSidebar() {
    const nav = document.getElementById('nav-sidebar');
    if (nav) nav.classList.toggle('open');
}

/* ──────────────────── 差異徽章 ──────────────────── */
function varianceBadge(val, pct) {
    const cls = val >= 0 ? 'badge-green' : 'badge-red';
    const symbol = val >= 0 ? '▲' : '▼';
    return `<span class="badge ${cls}">${symbol} ${Math.abs(pct).toFixed(1)}%</span>`;
}

/* ──────────────────── 告警條目 ──────────────────── */
function alertItem(cls, icon, level, msg) {
    return `
    <div style="display:flex;align-items:center;gap:10px;padding:10px 14px;background:rgba(0,0,0,0.18);border-radius:8px;margin-bottom:8px">
      <span class="badge ${cls}">${icon} ${level}</span>
      <span style="font-size:13px;color:var(--text-primary)">${msg}</span>
    </div>`;
}

/* ──────────────────── PDF 列印 ──────────────────── */
function printReport(tabId) {
    window.print();
}

/* ──────────────────── 圖表截圖下載 ──────────────────── */
function downloadChartPng(chartId, filename) {
    const ch = _charts[chartId];
    if (!ch) { showToast('⚠ 圖表未就緒', 'warn'); return; }
    const a = document.createElement('a');
    a.href = ch.toBase64Image();
    a.download = filename || `chart_${chartId}.png`;
    a.click();
}

/* ──────────────────── 全局渲染協調器 ──────────────────── */
function fullRender() {
    updateHeader();
    const activeTab = document.querySelector('.tab-content.active');
    if (activeTab) renderTab(activeTab.id);
}
