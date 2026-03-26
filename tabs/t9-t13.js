/* ╔════════════════════════════════════════════════════════╗
   ║  T9  — 產品組合優化器                                  ║
   ║  T10 — 資金效率（杜邦分析）                            ║
   ║  T11 — 月度趨勢追蹤                                   ║
   ║  T12 — 風險儀表板（VaR + 10000 次模擬）                ║
   ║  T13 — 資料管理中心                                   ║
   ╚════════════════════════════════════════════════════════╝ */

/* ── T9: 產品組合優化器 ── */
let _optResult = null;
function runOptimizerUI() {
    _optResult = runOptimizer();
    const { optProds, origCM, optCM, origProfit, optProfit } = _optResult;
    const prods = AppState.products;
    const gain = optCM - origCM;

    document.getElementById('opt-kpis').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">現行邊際貢獻(萬)</div><div class="kpi-val">${fmt2(origCM)}</div></div>
    <div class="kpi-card"><div class="kpi-label">最佳邊際貢獻(萬)</div><div class="kpi-val success">${fmt2(optCM)}</div></div>
    <div class="kpi-card"><div class="kpi-label">CM 增益(萬)</div><div class="kpi-val success">+${fmt2(gain)}</div></div>
    <div class="kpi-card"><div class="kpi-label">CM 增益幅度</div><div class="kpi-val success">+${origCM > 0 ? ((gain / origCM) * 100).toFixed(1) : 0}%</div></div>`;

    barChart('ch-opt',
        prods.map(p => p.name.split(' ')[0]),
        [
            { label: '現行銷量', data: prods.map(p => p.vol) },
            { label: '最佳銷量', data: optProds.map(p => p.vol) },
        ]
    );

    document.getElementById('opt-table-body').innerHTML = prods.map((p, i) => {
        const oVol = optProds[i].vol;
        const diff = oVol - p.vol;
        const cmU = p.price - p.vc;
        const cmGain = cmU * diff;
        return `<tr>
      <td style="text-align:left">${p.name}</td>
      <td>${fmt0(p.vol)}</td><td>${fmt0(oVol)}</td>
      <td><span class="badge ${diff > 0 ? 'badge-green' : 'badge-yellow'}">${diff >= 0 ? '+' : ''}${fmt0(diff)}</span></td>
      <td>${fmt2(cmU)}</td>
      <td class="${posNegClass(cmGain)}">${cmGain >= 0 ? '+' : ''}${fmt2(cmGain)}</td>
    </tr>`;
    }).join('');
    showToast('✅ 優化完成！', 'ok');
}
function renderT9() { if (!_optResult) runOptimizerUI(); }

/* ── T10: 杜邦分析 ── */
function renderT10() {
    const cap = runCapital();
    const c = AppState.capital;
    document.getElementById('cap-kpis').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">ROE</div>
      <div class="kpi-val ${posNegClass(cap.roe - c.indAvgROE)}">${(cap.roe * 100).toFixed(1)}%</div>
      <div class="kpi-sub">行業均：${pct0(c.indAvgROE)}</div></div>
    <div class="kpi-card"><div class="kpi-label">ROA</div>
      <div class="kpi-val ${posNegClass(cap.roa - c.indAvgROA)}">${(cap.roa * 100).toFixed(1)}%</div>
      <div class="kpi-sub">行業均：${pct0(c.indAvgROA)}</div></div>
    <div class="kpi-card"><div class="kpi-label">總資產周轉率</div>
      <div class="kpi-val ${posNegClass(cap.assetTurnover - c.indAvgTurnover)}">${cap.assetTurnover.toFixed(2)}次</div>
      <div class="kpi-sub">行業均：${c.indAvgTurnover.toFixed(2)} 次</div></div>
    <div class="kpi-card"><div class="kpi-label">財務槓桿係數</div>
      <div class="kpi-val">${cap.leverage.toFixed(2)}×</div>
      <div class="kpi-sub">總資產÷股東權益</div></div>`;

    document.getElementById('dupont-tree').innerHTML = `
    <div style="display:flex;align-items:center;justify-content:center;gap:8px;flex-wrap:wrap;padding:12px 0">
      <div class="dupont-box" style="border-color:rgba(139,92,246,0.5)">
        <div style="font-size:11px;color:var(--violet-lt);margin-bottom:4px">ROE</div>
        <div style="font-size:22px;font-weight:900;color:#c4b5fd">${(cap.roe * 100).toFixed(1)}%</div>
        <div style="font-size:10px;color:var(--text-muted)">股東權益報酬率</div>
      </div>
      <div class="dupont-op">=</div>
      <div class="dupont-box">
        <div style="font-size:11px;color:var(--sky-lt);margin-bottom:4px">利潤率</div>
        <div style="font-size:18px;font-weight:800;color:#7dd3fc">${(cap.profitMargin * 100).toFixed(1)}%</div>
        <div style="font-size:10px;color:var(--text-muted)">利潤÷收入</div>
      </div>
      <div class="dupont-op">×</div>
      <div class="dupont-box">
        <div style="font-size:11px;color:var(--emerald);margin-bottom:4px">資產周轉</div>
        <div style="font-size:18px;font-weight:800;color:#6ee7b7">${cap.assetTurnover.toFixed(2)}×</div>
        <div style="font-size:10px;color:var(--text-muted)">收入÷總資產</div>
      </div>
      <div class="dupont-op">×</div>
      <div class="dupont-box">
        <div style="font-size:11px;color:var(--amber);margin-bottom:4px">財務槓桿</div>
        <div style="font-size:18px;font-weight:800;color:#fde68a">${cap.leverage.toFixed(2)}×</div>
        <div style="font-size:10px;color:var(--text-muted)">總資產÷股東權益</div>
      </div>
    </div>`;

    radarChart('ch-cap-radar',
        ['ROE', 'ROA', '資產周轉', '利潤率', '財務槓桿(相對)'],
        [
            { label: '本公司', data: [cap.roe / 0.2, cap.roa / 0.12, cap.assetTurnover / 1.5, cap.profitMargin / 0.15, cap.leverage / 3].map(v => +(v * 100).toFixed(0)) },
            { label: '行業基準', data: [100, 100, 100, 100, 100], borderDash: [5, 3] },
        ]
    );

    /* ROE 月度趨勢 */
    lineChart('ch-roe-decomp',
        AppState.monthlyTrend.map(m => m.month),
        [{
            label: '月度年化 ROE(%)', data:
                AppState.monthlyTrend.map(m => +((m.profit / cap.equity) * 12 * 100).toFixed(2)),
            fill: true, backgroundColor: 'rgba(139,92,246,0.12)'
        }]
    );
}

/* ── T11: 月度趨勢 ── */
function renderT11() {
    const td = AppState.monthlyTrend;
    if (!td.length) return;
    const totalRev = td.reduce((s, m) => s + m.rev, 0);
    const totalCM = td.reduce((s, m) => s + m.cm, 0);
    const totalProfit = td.reduce((s, m) => s + m.profit, 0);
    const avgCMRate = totalRev > 0 ? totalCM / totalRev : 0;

    document.getElementById('trend-kpis').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">年度總收入(萬)</div><div class="kpi-val">${fmt2(totalRev)}</div></div>
    <div class="kpi-card"><div class="kpi-label">年度邊際貢獻(萬)</div><div class="kpi-val success">${fmt2(totalCM)}</div></div>
    <div class="kpi-card"><div class="kpi-label">年度平均 CM 率</div><div class="kpi-val">${(avgCMRate * 100).toFixed(1)}%</div></div>
    <div class="kpi-card"><div class="kpi-label">年度營業利潤(萬)</div><div class="kpi-val ${posNegClass(totalProfit)}">${fmt2(totalProfit)}</div></div>`;

    lineChart('ch-trend-main', td.map(m => m.month), [
        { label: '收入(萬)', data: td.map(m => +m.rev.toFixed(0)), fill: true, backgroundColor: 'rgba(14,165,233,0.08)' },
        { label: '邊際貢獻(萬)', data: td.map(m => +m.cm.toFixed(0)), fill: false },
        { label: '利潤(萬)', data: td.map(m => +m.profit.toFixed(0)), fill: false },
    ]);

    barChart('ch-trend-cm', td.map(m => m.month),
        [{
            label: 'CM率(%)', data: td.map(m => +(m.cmRate * 100).toFixed(1)),
            backgroundColor: td.map(m =>
                m.cmRate >= AppState.params.cmGoodThreshold ? 'rgba(16,185,129,0.75)' :
                    m.cmRate >= AppState.params.industryAvgCM ? 'rgba(14,165,233,0.75)' :
                        'rgba(239,68,68,0.70)')
        }]
    );

    document.getElementById('trend-table-body').innerHTML = td.map((m, i) => {
        const prev = td[i - 1];
        const mom = prev && prev.rev > 0 ? ((m.rev - prev.rev) / prev.rev * 100).toFixed(1) : '—';
        return `<tr>
      <td>${m.month}</td><td>${fmt2(m.rev)}</td><td>${fmt2(m.cm)}</td>
      <td>${(m.cmRate * 100).toFixed(1)}%</td><td class="${posNegClass(m.profit)}">${fmt2(m.profit)}</td>
      <td>${mom === '—' ? '—' : `<span class="badge ${parseFloat(mom) >= 0 ? 'badge-green' : 'badge-red'}">${parseFloat(mom) >= 0 ? '+' : ''}${mom}%</span>`}</td>
    </tr>`;
    }).join('');
}

/* ── T12: 風險儀表板 ── */
let _mcResult = null;
function runMonteCarloUI() {
    const n = 10000;
    const revStd = +document.getElementById('mc-rev-std').value;
    const costStd = +document.getElementById('mc-cost-std').value;
    _mcResult = runMonteCarlo(n, revStd, costStd);
    const { mean, p5, p95, var95, var99, cvar95, lossProb } = _mcResult;

    document.getElementById('mc-kpis').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">期望利潤(萬)</div><div class="kpi-val ${posNegClass(mean)}">${fmt2(mean)}</div></div>
    <div class="kpi-card"><div class="kpi-label">P5 ~ P95 區間(萬)</div>
      <div style="font-size:15px;font-weight:700;color:var(--amber)">${fmt2(p5)} ~ ${fmt2(p95)}</div></div>
    <div class="kpi-card"><div class="kpi-label">虧損機率</div><div class="kpi-val ${lossProb > 0.3 ? 'danger' : 'success'}">${(lossProb * 100).toFixed(1)}%</div></div>
    <div class="kpi-card"><div class="kpi-label">95% VaR (最大可能損失)</div><div class="kpi-val danger">${fmt2(Math.max(var95, 0))}</div><div class="kpi-sub">萬元</div></div>
    <div class="kpi-card"><div class="kpi-label">99% VaR</div><div class="kpi-val danger">${fmt2(Math.max(var99, 0))}</div><div class="kpi-sub">萬元</div></div>
    <div class="kpi-card"><div class="kpi-label">CVaR（條件 VaR）</div><div class="kpi-val danger">${fmt2(Math.max(cvar95, 0))}</div><div class="kpi-sub">95% 尾部平均損失</div></div>`;

    histogramChart('ch-mc-hist', Array.from(_mcResult.profits), 40);
}

function renderT12() {
    const risks = AppState.risks;
    /* 5×5 風險矩陣 */
    const levels = [[0, 0.2], [0.2, 0.4], [0.4, 0.6], [0.6, 0.8], [0.8, 1.01]];
    const cells = [];
    for (let r = 4; r >= 0; r--) {
        for (let c = 0; c < 5; c++) {
            const matches = risks.filter(k =>
                k.prob >= levels[c][0] && k.prob < levels[c][1] &&
                k.impact >= levels[r][0] && k.impact < levels[r][1]);
            const score = (r + c) / 8;
            const bg = score > 0.7 ? 'rgba(239,68,68,0.65)' : score > 0.4 ? 'rgba(245,158,11,0.55)' : 'rgba(16,185,129,0.45)';
            cells.push(`<div class="risk-cell" style="background:${bg}" title="${matches.map(k => k.name).join('\n')}">
        ${matches.length ? `<div style="font-size:18px;font-weight:900">${matches.length}</div><div style="font-size:9px">${matches[0].name.slice(0, 6)}</div>` : ''}
      </div>`);
        }
    }
    document.getElementById('risk-matrix-container').innerHTML =
        `<div style="font-size:11px;color:var(--text-muted);margin-bottom:6px;text-align:center">↑ 財務衝擊度 × 發生機率 →</div>
     <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:4px">${cells.join('')}</div>
     <div style="display:flex;gap:12px;margin-top:8px;font-size:11px;flex-wrap:wrap">
       <span><span style="display:inline-block;width:12px;height:12px;background:rgba(239,68,68,0.65);border-radius:2px;margin-right:4px"></span>高風險</span>
       <span><span style="display:inline-block;width:12px;height:12px;background:rgba(245,158,11,0.55);border-radius:2px;margin-right:4px"></span>中風險</span>
       <span><span style="display:inline-block;width:12px;height:12px;background:rgba(16,185,129,0.45);border-radius:2px;margin-right:4px"></span>低風險</span>
     </div>`;

    document.getElementById('risk-table-body').innerHTML = risks.map(r => {
        const score = r.prob * r.impact;
        const cls = score > 0.3 ? 'badge-red' : score > 0.12 ? 'badge-yellow' : 'badge-green';
        const lvl = score > 0.3 ? '高' : score > 0.12 ? '中' : '低';
        return `<tr>
      <td style="text-align:left;font-size:11px">${r.name}</td>
      <td>${(r.prob * 100).toFixed(0)}%</td><td>${(r.impact * 100).toFixed(0)}%</td>
      <td><span class="badge ${cls}">${lvl}</span></td>
      <td style="font-size:11px">${r.cat}</td>
      <td style="font-size:11px;color:var(--text-muted)">${r.mitigation || '—'}</td>
    </tr>`;
    }).join('');

    if (!_mcResult) runMonteCarloUI();
}

/* ── T13: 資料管理 ── */
function renderT13() {
    const base = runEngine();
    document.getElementById('data-status').innerHTML = `
    <div>產品種類：<b style="color:var(--sky-lt)">${AppState.products.length}</b> 項</div>
    <div>固定成本：<b style="color:var(--amber)">${fmt2(AppState.params.fc)}</b> 萬元</div>
    <div>總收入：<b style="color:var(--sky-lt)">${fmt2(base.totalRev)}</b> 萬元</div>
    <div>情境記錄：<b style="color:var(--violet-lt)">${AppState.scenarios.length + AppState.userScenarios.length}</b> 個</div>
    <div>月度資料：<b style="color:var(--emerald)">${AppState.monthlyTrend.length}</b> 筆</div>
    <div>預算資料：<b style="color:var(--amber)">${AppState.budgetData.length}</b> 項產品</div>
    <div>風險事件：<b style="color:var(--red)">${AppState.risks.length}</b> 項</div>
    <div>多年度資料：<b style="color:var(--violet-lt)">${AppState.yoyData.length}</b> 年度</div>`;
    logChange('檢視資料管理頁');
}
