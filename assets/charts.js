/* ╔══════════════════════════════════════════════════════════╗
   ║  CVP V4.0 — 圖表工廠 & 通用圖表函數                      ║
   ╚══════════════════════════════════════════════════════════╝ */

/* ─── 色彩系統 ─── */
const PALETTE = [
    '#0ea5e9', '#8b5cf6', '#10b981', '#f59e0b',
    '#ef4444', '#06b6d4', '#ec4899', '#f97316',
    '#a3e635', '#818cf8', '#34d399', '#fbbf24',
];

/* ─── 圖表實例管理 ─── */
const _charts = {};

/**
 * 統一圖表工廠：自動摧毀同 ID 舊實例，注入響應式設定
 * @param {string} id  - canvas 元素 ID
 * @param {object} cfg - Chart.js 設定物件
 * @returns {Chart|null}
 */
function mkChart(id, cfg) {
    if (_charts[id]) {
        _charts[id].destroy();
        delete _charts[id];
    }
    const ctx = document.getElementById(id);
    if (!ctx) return null;

    /* 注入全局響應式設定 */
    cfg.options = cfg.options || {};
    cfg.options.responsive = true;
    cfg.options.maintainAspectRatio = false;
    cfg.options.animation = cfg.options.animation ?? { duration: 400 };

    /* 注入全局字型顏色 */
    cfg.options.plugins = cfg.options.plugins || {};
    cfg.options.plugins.legend = cfg.options.plugins.legend || {};
    cfg.options.plugins.legend.labels = {
        color: 'var(--text-secondary)',
        font: { size: 11 },
        ...cfg.options.plugins.legend.labels,
    };

    _charts[id] = new Chart(ctx, cfg);
    return _charts[id];
}

/* ──────────────────── 全局預設值 ──────────────────── */
const SCALE_DEFAULTS = {
    x: { ticks: { color: 'var(--tick-color)', font: { size: 11 } }, grid: { color: 'var(--grid-color)' } },
    y: { ticks: { color: 'var(--tick-color)', font: { size: 11 } }, grid: { color: 'var(--grid-color)' } },
};

/**
 * 生成標準 XY 軸設定（可自訂標題）
 */
function xyScales(xTitle = '', yTitle = '') {
    const s = JSON.parse(JSON.stringify(SCALE_DEFAULTS));
    if (xTitle) s.x.title = { display: true, text: xTitle, color: 'var(--text-muted)', font: { size: 11 } };
    if (yTitle) s.y.title = { display: true, text: yTitle, color: 'var(--text-muted)', font: { size: 11 } };
    return s;
}

/* ──────────────────── 常用圖表快捷函數 ──────────────────── */

/**
 * 快速建立長條圖
 */
function barChart(id, labels, datasets, opts = {}) {
    datasets = datasets.map((d, i) => ({
        borderRadius: 5,
        backgroundColor: PALETTE[i % PALETTE.length] + (d.alpha ?? 'B3'),
        ...d,
    }));
    return mkChart(id, {
        type: 'bar',
        data: { labels, datasets },
        options: {
            plugins: { legend: { display: datasets.length > 1 ? true : false } },
            scales: xyScales(opts.xTitle, opts.yTitle),
            ...opts,
        },
    });
}

/**
 * 快速建立折線圖
 */
function lineChart(id, labels, datasets, opts = {}) {
    datasets = datasets.map((d, i) => ({
        tension: 0.35,
        fill: false,
        borderWidth: 2,
        pointRadius: 3,
        pointHoverRadius: 6,
        borderColor: PALETTE[i % PALETTE.length],
        pointBackgroundColor: PALETTE[i % PALETTE.length],
        ...d,
    }));
    return mkChart(id, {
        type: 'line',
        data: { labels, datasets },
        options: {
            scales: xyScales(opts.xTitle, opts.yTitle),
            ...opts,
        },
    });
}

/**
 * 快速建立甜甜圈 / 圓餅圖
 */
function doughnutChart(id, labels, data, opts = {}) {
    return mkChart(id, {
        type: opts.pie ? 'pie' : 'doughnut',
        data: {
            labels,
            datasets: [{
                data,
                backgroundColor: PALETTE.map(c => c + 'CC'),
                borderColor: 'rgba(0,0,0,0.25)',
                borderWidth: 2,
                hoverOffset: 6,
            }],
        },
        options: {
            plugins: { legend: { position: 'bottom', labels: { color: 'var(--text-secondary)', padding: 12, font: { size: 11 } } } },
            ...opts,
        },
    });
}

/**
 * 快速建立雷達圖
 */
function radarChart(id, labels, datasets, opts = {}) {
    datasets = datasets.map((d, i) => ({
        borderColor: PALETTE[i % PALETTE.length],
        backgroundColor: PALETTE[i % PALETTE.length] + '33',
        pointBackgroundColor: PALETTE[i % PALETTE.length],
        borderWidth: 2,
        pointRadius: 3,
        ...d,
    }));
    return mkChart(id, {
        type: 'radar',
        data: { labels, datasets },
        options: {
            scales: {
                r: {
                    angleLines: { color: 'rgba(255,255,255,0.10)' },
                    grid: { color: 'rgba(255,255,255,0.10)' },
                    pointLabels: { color: 'var(--text-secondary)', font: { size: 11 } },
                    ticks: { backdropColor: 'transparent', color: 'var(--text-muted)', stepSize: 20, font: { size: 10 } },
                    suggestedMin: 0,
                    suggestedMax: 120,
                },
            },
            plugins: { legend: { position: 'bottom', labels: { color: 'var(--text-secondary)', padding: 14, font: { size: 11 } } } },
            ...opts,
        },
    });
}

/**
 * 瀑布圖（利用 bar chart 堆疊模擬）
 */
function waterfallChart(id, labels, values, opts = {}) {
    const posColors = values.map(v => v >= 0 ? 'rgba(16,185,129,0.80)' : 'rgba(239,68,68,0.80)');
    return mkChart(id, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: opts.label || '金額（萬）',
                data: values,
                backgroundColor: posColors,
                borderRadius: 6,
                borderColor: posColors.map(c => c.replace('0.80', '1')),
                borderWidth: 1,
            }],
        },
        options: {
            plugins: { legend: { display: false } },
            scales: xyScales(opts.xTitle, opts.yTitle || '萬元'),
        },
    });
}

/**
 * 散點氣泡圖（BCG 矩陣用）
 */
function bubbleChart(id, datasets, opts = {}) {
    datasets = datasets.map((d, i) => ({
        backgroundColor: PALETTE[i % PALETTE.length] + '99',
        borderColor: PALETTE[i % PALETTE.length],
        borderWidth: 2,
        ...d,
    }));
    return mkChart(id, {
        type: 'bubble',
        data: { datasets },
        options: {
            scales: {
                x: { ...SCALE_DEFAULTS.x, title: { display: true, text: opts.xTitle || 'X', color: 'var(--text-muted)' }, suggestedMin: 0 },
                y: { ...SCALE_DEFAULTS.y, title: { display: true, text: opts.yTitle || 'Y', color: 'var(--text-muted)' }, suggestedMin: 0 },
            },
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: ctx => {
                            const d = ctx.raw;
                            return `${ctx.dataset.label}: (${fmt2(d.x)}, ${fmt2(d.y)}) 銷量:${fmt0(d.r * 10)}`;
                        },
                    },
                },
            },
            ...opts,
        },
    });
}

/**
 * 直方圖（用 bar chart 模擬，for 蒙地卡羅）
 */
function histogramChart(id, data, bins = 40, opts = {}) {
    if (!data || data.length === 0) return null;
    const sorted = [...data].sort((a, b) => a - b);
    const minV = sorted[0], maxV = sorted[sorted.length - 1];
    const binW = (maxV - minV) / bins || 1;
    const freq = new Array(bins).fill(0);
    const mids = [];
    sorted.forEach(v => {
        const idx = Math.min(Math.floor((v - minV) / binW), bins - 1);
        freq[idx]++;
    });
    for (let i = 0; i < bins; i++) mids.push(+(minV + (i + 0.5) * binW).toFixed(1));

    const colors = mids.map(m => m < 0 ? 'rgba(239,68,68,0.75)' : 'rgba(139,92,246,0.75)');

    return mkChart(id, {
        type: 'bar',
        data: {
            labels: mids.map(m => fmt2(m)),
            datasets: [{
                label: '模擬頻次',
                data: freq,
                backgroundColor: colors,
                borderWidth: 0,
                borderRadius: 2,
                barPercentage: 1.0,
                categoryPercentage: 1.0,
            }],
        },
        options: {
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        title: items => `利潤區間 ~${items[0].label} 萬`,
                        label: item => `出現次數：${item.raw}（${(item.raw / data.length * 100).toFixed(1)}%）`,
                    },
                },
                annotation: opts.annotation, // 可傳入垂直線標記
            },
            scales: {
                x: { ...SCALE_DEFAULTS.x, ticks: { maxTicksLimit: 12, color: 'var(--tick-color)', font: { size: 10 } } },
                y: { ...SCALE_DEFAULTS.y },
            },
        },
    });
}
