/* ╔════════════════════════════════════════════════════════╗
   ║  T14 — 預算差異分析（Budget Variance Analysis）         ║
   ║  T15 — 現金流量預測                                    ║
   ║  T16 — 龍捲風敏感性圖                                  ║
   ║  T17 — 邊際貢獻矩陣（BCG 四象限）                      ║
   ║  T18 — 多年度比較（YoY）                               ║
   ║  T19 — AI 決策摘要報告                                 ║
   ╚════════════════════════════════════════════════════════╝ */

/* ══════════════════════════════════════
   T14 — 預算差異分析
   ══════════════════════════════════════ */
function renderT14() {
    const va = runBudgetVariance();

    /* 匯總 KPI */
    document.getElementById('bv-kpis').innerHTML = `
    <div class="kpi-card">
      <div class="kpi-label">實際收入(萬)</div>
      <div class="kpi-val">${fmt2(va.actRev)}</div>
      <div class="kpi-sub">預算：${fmt2(va.budRev)} &nbsp; ${varianceBadge(va.revVar, (va.revVar / Math.abs(va.budRev || 1) * 100))}</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">實際邊際貢獻(萬)</div>
      <div class="kpi-val ${posNegClass(va.cmVar)}">${fmt2(va.actCM)}</div>
      <div class="kpi-sub">預算：${fmt2(va.budCM)} &nbsp; ${varianceBadge(va.cmVar, (va.cmVar / Math.abs(va.budCM || 1) * 100))}</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">實際利潤(萬)</div>
      <div class="kpi-val ${posNegClass(va.actProfit)}">${fmt2(va.actProfit)}</div>
      <div class="kpi-sub">預算：${fmt2(va.budProfit)} &nbsp; ${varianceBadge(va.profitVar, (va.profitVar / Math.abs(va.budProfit || 1) * 100))}</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">CM 率差異</div>
      <div class="kpi-val ${posNegClass(va.actCMRate - va.budCMRate)}">${(va.actCMRate * 100).toFixed(1)}%</div>
      <div class="kpi-sub">預算：${(va.budCMRate * 100).toFixed(1)}%</div>
    </div>`;

    /* 差異瀑布圖 */
    waterfallChart('ch-bv-waterfall',
        ['預算收入', '收入差異', '預算CM', 'CM差異', '固定成本差異', '實際利潤'],
        [va.budRev, va.revVar, va.budCM, va.cmVar, 0, va.actProfit],
        { label: '萬元' }
    );

    /* 產品差異明細表 */
    document.getElementById('bv-table').innerHTML =
        `<thead><tr>
      <th style="text-align:left">產品</th>
      <th>實際收入(萬)</th><th>預算收入(萬)</th><th>收入差異</th>
      <th>實際CM(萬)</th><th>預算CM(萬)</th><th>CM差異</th>
    </tr></thead><tbody>` +
        va.prodVariance.map(p => `<tr>
      <td style="text-align:left">${p.name.split(' ')[0]}</td>
      <td>${fmt2(p.actRev)}</td><td>${fmt2(p.budRev)}</td>
      <td>${varianceBadge(p.revVar, p.revVarPct)}</td>
      <td>${fmt2(p.actCM)}</td><td>${fmt2(p.budCM)}</td>
      <td>${varianceBadge(p.cmVar, p.cmVarPct)}</td>
    </tr>`).join('') + '</tbody>';
}

/* ══════════════════════════════════════
   T15 — 現金流預測
   ══════════════════════════════════════ */
function renderT15() {
    const cf = runCashflowForecast();
    const cp = AppState.cashflowParams;
    const below = cf.filter(m => m.isBelowSafety);

    /* KPI */
    const totalNet = cf.reduce((s, m) => s + m.netCF, 0);
    const minCum = Math.min(...cf.map(m => m.cumulative));
    document.getElementById('cf-kpis').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">全年淨現金流(萬)</div>
      <div class="kpi-val ${posNegClass(totalNet)}">${fmtDelta(totalNet)}</div></div>
    <div class="kpi-card"><div class="kpi-label">最低現金水位(萬)</div>
      <div class="kpi-val ${posNegClass(minCum - cp.safetyBuffer)}">${fmt2(minCum)}</div>
      <div class="kpi-sub">安全水位：${fmt2(cp.safetyBuffer)} 萬</div></div>
    <div class="kpi-card"><div class="kpi-label">低於安全水位月份</div>
      <div class="kpi-val ${below.length > 0 ? 'danger' : 'success'}">${below.length}</div>
      <div class="kpi-sub">個月</div></div>
    <div class="kpi-card"><div class="kpi-label">應收帳款天數</div>
      <div class="kpi-val">${cp.arDays}</div><div class="kpi-sub">天</div></div>`;

    /* 現金流折線 */
    lineChart('ch-cf-forecast', cf.map(m => m.month), [
        { label: '營運現金流(萬)', data: cf.map(m => m.operatingCF), borderColor: '#0ea5e9' },
        { label: '投資現金流(萬)', data: cf.map(m => m.investingCF), borderColor: '#f59e0b', borderDash: [5, 3] },
        { label: '淨現金流(萬)', data: cf.map(m => m.netCF), borderColor: '#10b981' },
        { label: '累計現金水位(萬)', data: cf.map(m => m.cumulative), borderColor: '#8b5cf6', fill: true, backgroundColor: 'rgba(139,92,246,0.08)' },
    ]);

    /* 警示 */
    const alertEl = document.getElementById('cf-alerts');
    if (alertEl) alertEl.innerHTML = below.length > 0
        ? below.map(m => alertItem('badge-red', '💧', '危急', `${m.month} 累計現金 <b>${fmt2(m.cumulative)}</b> 萬低於安全水位 ${fmt2(cp.safetyBuffer)} 萬，需緊急融資。`)).join('')
        : alertItem('badge-green', '✅', '正常', '全年現金水位維持在安全線以上。');
}

/* ══════════════════════════════════════
   T16 — 龍捲風敏感性圖
   ══════════════════════════════════════ */
function renderT16() {
    const delta = +(document.getElementById('tornado-delta')?.value || 10);
    const sens = runSensitivity(delta);
    const maxAbs = Math.max(...sens.map(s => s.maxAbs)) || 1;

    document.getElementById('tornado-chart').innerHTML = sens.map(s => `
    <div class="tornado-row">
      <div class="tornado-label">${s.label}</div>
      <div class="tornado-bars">
        <div class="tornado-bar-pos" style="width:${Math.abs(s.impactPos) / maxAbs * 48}%"></div>
        <div class="tornado-bar-neg" style="width:${Math.abs(s.impactNeg) / maxAbs * 48}%"></div>
        <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);width:2px;height:100%;background:rgba(255,255,255,0.3)"></div>
      </div>
      <div class="tornado-val text-pos">+${s.impactPos}%</div>
      <div class="tornado-val text-neg">${s.impactNeg}%</div>
    </div>`).join('');

    /* 也以 barChart 呈現 */
    barChart('ch-tornado',
        sens.map(s => s.label),
        [
            { label: `+${delta}% 影響`, data: sens.map(s => s.impactPos) },
            { label: `-${delta}% 影響`, data: sens.map(s => s.impactNeg) },
        ]
    );

    document.getElementById('tornado-insight').innerHTML = `
    <div class="ai-box">📊 <b style="color:var(--sky-lt)">敏感性排行：</b>
    ${sens.map((s, i) => `<b>${i + 1}. ${s.label}</b>（最大影響 ±${s.maxAbs.toFixed(1)}%）`).join('、')}。
    管理優先順序：聚焦最排名靠前的變數，可獲得最大的利潤改善效果。</div>`;
}

/* ══════════════════════════════════════
   T17 — CM 矩陣（BCG 四象限氣泡圖）
   ══════════════════════════════════════ */
function renderT17() {
    const base = runEngine();
    const prods = base.prodCalc;
    const avgCMR = base.cmRate;
    const totalRev = base.totalRev;

    /* 四象限分割線 */
    const xMid = avgCMR * 100;            // 平均 CM 率（%）
    const yMid = 100 / prods.length;      // 平均收入佔比（%）

    /* 氣泡數據 */
    const datasets = prods.map((p, i) => ({
        label: p.name.split(' ')[0],
        data: [{
            x: +(p.cmRate * 100).toFixed(1),              // X = CM率
            y: +(p.rev / totalRev * 100).toFixed(1),       // Y = 收入佔比
            r: Math.max(5, Math.sqrt(p.vol) * 0.5),        // r = 銷量（氣泡大小）
        }],
    }));
    bubbleChart('ch-bcg', datasets, {
        xTitle: '邊際貢獻率（%）',
        yTitle: '收入佔比（%）',
    });

    /* 四象限文字指引 */
    document.getElementById('bcg-legend').innerHTML = `
    <div class="kpi-grid-4" style="font-size:12px;margin-top:10px">
      <div class="matrix-quadrant" style="background:rgba(16,185,129,0.1);border-color:rgba(16,185,129,0.3)">
        ⭐ <b>明星產品</b><br>CM率高 + 收入佔比大</div>
      <div class="matrix-quadrant" style="background:rgba(245,158,11,0.1);border-color:rgba(245,158,11,0.3)">
        🐄 <b>金牛產品</b><br>CM率低 + 收入佔比大</div>
      <div class="matrix-quadrant" style="background:rgba(14,165,233,0.1);border-color:rgba(14,165,233,0.3)">
        ❓ <b>問號產品</b><br>CM率高 + 收入佔比小</div>
      <div class="matrix-quadrant" style="background:rgba(239,68,68,0.1);border-color:rgba(239,68,68,0.3)">
        🐕 <b>瘦狗產品</b><br>CM率低 + 收入佔比小</div>
    </div>`;

    /* 各產品分類說明 */
    document.getElementById('bcg-table').innerHTML =
        `<thead><tr><th>產品</th><th>CM率</th><th>收入佔比</th><th>分類</th><th>建議</th></tr></thead><tbody>` +
        prods.map(p => {
            const cmrPct = p.cmRate * 100;
            const revPct = p.rev / totalRev * 100;
            const isStar = cmrPct >= xMid && revPct >= yMid;
            const isCow = cmrPct < xMid && revPct >= yMid;
            const isQ = cmrPct >= xMid && revPct < yMid;
            const isDog = !isStar && !isCow && !isQ;
            const [icon, type, tip] =
                isStar ? ['⭐', '明星', '持續投資，擴大優勢'] :
                    isCow ? ['🐄', '金牛', '維持量能，改善成本結構'] :
                        isQ ? ['❓', '問號', '評估擴大銷量，提升市場佔比'] :
                            ['🐕', '瘦狗', '考慮調整定價或退出'];
            return `<tr>
        <td style="text-align:left">${p.name}</td>
        <td>${cmrPct.toFixed(1)}%</td>
        <td>${revPct.toFixed(1)}%</td>
        <td>${icon} ${type}</td>
        <td style="font-size:11px;color:var(--text-muted)">${tip}</td>
      </tr>`;
        }).join('') + '</tbody>';
}

/* ══════════════════════════════════════
   T18 — 多年度比較（YoY + CAGR）
   ══════════════════════════════════════ */
function renderT18() {
    const yoy = AppState.yoyData;
    const noDataEl = document.getElementById('yoy-nodata');
    const chartsEl = document.getElementById('yoy-charts');
    if (!yoy.length) {
        if (noDataEl) noDataEl.style.display = '';
        if (chartsEl) chartsEl.style.display = 'none';
        return;
    }
    if (noDataEl) noDataEl.style.display = 'none';
    if (chartsEl) chartsEl.style.display = '';

    const years = yoy.map(d => d.year);

    /* 年度總計 */
    const annuals = yoy.map(d => ({
        year: d.year,
        rev: d.months.reduce((s, m) => s + m.rev, 0),
        cm: d.months.reduce((s, m) => s + m.cm, 0),
        profit: d.months.reduce((s, m) => s + m.profit, 0),
    }));

    /* CAGR */
    const cagr = (first, last, n) => n > 1 && first > 0 ? (Math.pow(last / first, 1 / (n - 1)) - 1) * 100 : 0;
    const n = annuals.length;
    const cagrRev = cagr(annuals[0].rev, annuals[n - 1].rev, n);
    const cagrProfit = cagr(Math.abs(annuals[0].profit), Math.abs(annuals[n - 1].profit), n);

    document.getElementById('yoy-kpis').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">分析年度</div><div class="kpi-val">${n}</div><div class="kpi-sub">年（${years[0]} - ${years[n - 1]}）</div></div>
    <div class="kpi-card"><div class="kpi-label">收入 CAGR</div><div class="kpi-val ${posNegClass(cagrRev)}">${fmtDeltaPct(cagrRev)}</div><div class="kpi-sub">複合年成長率</div></div>
    <div class="kpi-card"><div class="kpi-label">利潤 CAGR</div><div class="kpi-val ${posNegClass(cagrProfit)}">${fmtDeltaPct(cagrProfit)}</div></div>
    <div class="kpi-card"><div class="kpi-label">最新年度利潤</div><div class="kpi-val ${posNegClass(annuals[n - 1].profit)}">${fmt2(annuals[n - 1].profit)}</div><div class="kpi-sub">萬元</div></div>`;

    lineChart('ch-yoy', years, [
        { label: '年度總收入(萬)', data: annuals.map(d => +d.rev.toFixed(0)) },
        { label: '年度邊際貢獻(萬)', data: annuals.map(d => +d.cm.toFixed(0)) },
        { label: '年度利潤(萬)', data: annuals.map(d => +d.profit.toFixed(0)) },
    ]);
}

/* ══════════════════════════════════════
   T19 — AI 決策摘要報告（可列印）
   ══════════════════════════════════════ */
function renderT19() {
    const r = runEngine();
    const cap = runCapital();
    const inv = runInvestment();
    const va = runBudgetVariance();
    const p = AppState.params;
    const mc = runMonteCarlo(5000, 15, 10); // 快速版 5000 次

    /* 評等算法 */
    const profitScore = r.profit >= p.targetProfit ? 100 : Math.max(0, r.profit / p.targetProfit * 100);
    const cmScore = r.cmRate >= p.cmGoodThreshold ? 100 : r.cmRate >= p.industryAvgCM ? 65 : 30;
    const safetyScore = r.safetyMarginRate >= p.safetyWarn ? 100 : r.safetyMarginRate / p.safetyWarn * 100;
    const riskScore = mc.lossProb < 0.05 ? 100 : mc.lossProb < 0.15 ? 70 : 40;
    const overall = ((profitScore + cmScore + safetyScore + riskScore) / 4).toFixed(0);
    const grade = overall >= 80 ? '⭐ 優秀' : overall >= 60 ? '📊 良好' : overall >= 40 ? '⚠ 一般' : '🔴 需改善';

    /* 三大機會 */
    const opps = [];
    if (r.cmRate < p.cmGoodThreshold) opps.push('提升產品定價或降低原物料成本，改善邊際貢獻率');
    if (r.safetyMarginRate < 0.25) opps.push('擴大銷量以提升安全邊際，降低損益平衡風險');
    if (inv.npv > 0 && inv.irr > p.targetProfit / 100) opps.push(`投資計畫 NPV 為正（${fmt2(inv.npv)} 萬），IRR = ${(inv.irr * 100).toFixed(1)}%，可行性高`);
    if (va.cmVar > 0) opps.push(`邊際貢獻超出預算 ${fmt2(va.cmVar)} 萬，可維持當前策略並加速擴張`);
    while (opps.length < 3) opps.push('維持現有成本效率，持續追蹤月度趨勢');

    /* 三大風險 */
    const top3Risks = [...AppState.risks]
        .sort((a, b) => b.prob * b.impact - a.prob * a.impact)
        .slice(0, 3);

    document.getElementById('ai-report').innerHTML = `
    <div style="max-width:900px;margin:0 auto">
      <!-- 報告標題 -->
      <div style="text-align:center;margin-bottom:28px;padding:20px;background:linear-gradient(135deg,rgba(14,165,233,0.15),rgba(139,92,246,0.12));border-radius:16px">
        <div style="font-size:28px;font-weight:900;background:linear-gradient(90deg,var(--sky-lt),var(--violet-lt));-webkit-background-clip:text;-webkit-text-fill-color:transparent">
          CVP 一頁式決策摘要
        </div>
        <div style="font-size:14px;color:var(--text-muted);margin-top:6px">${p.company}・${p.period}・生成時間：${fmtDateTime(new Date())}</div>
        <div style="font-size:32px;margin-top:10px">${grade}</div>
        <div style="font-size:14px;color:var(--text-muted)">綜合評分：${overall} / 100</div>
      </div>

      <!-- 核心 KPI -->
      <div class="kpi-grid-4 mb-6">
        <div class="kpi-card"><div class="kpi-label">總收入</div><div class="kpi-val">${fmt2(r.totalRev)}</div><div class="kpi-sub">萬元</div></div>
        <div class="kpi-card"><div class="kpi-label">邊際貢獻率</div><div class="kpi-val ${posNegClass(r.cmRate - p.industryAvgCM)}">${(r.cmRate * 100).toFixed(1)}%</div><div class="kpi-sub">行業均 ${pct0(p.industryAvgCM)}</div></div>
        <div class="kpi-card"><div class="kpi-label">營業利潤</div><div class="kpi-val ${posNegClass(r.profit - p.targetProfit)}">${fmt2(r.profit)}</div><div class="kpi-sub">目標 ${fmt2(p.targetProfit)} 萬</div></div>
        <div class="kpi-card"><div class="kpi-label">安全邊際率</div><div class="kpi-val ${posNegClass(r.safetyMarginRate - p.safetyWarn)}">${(r.safetyMarginRate * 100).toFixed(1)}%</div><div class="kpi-sub">警示線 ${pct0(p.safetyWarn)}</div></div>
      </div>

      <!-- 三大機會 -->
      <div class="glass mb-4" style="padding:20px">
        <div class="sec-hd">🚀 三大成長機會</div>
        ${opps.slice(0, 3).map((o, i) => `
          <div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:10px;padding:10px;background:rgba(16,185,129,0.06);border-radius:8px;border-left:3px solid var(--emerald)">
            <span style="color:var(--emerald);font-weight:900;font-size:16px">${i + 1}</span>
            <span style="font-size:13px;color:var(--text-primary)">${o}</span>
          </div>`).join('')}
      </div>

      <!-- 三大風險 -->
      <div class="glass mb-4" style="padding:20px">
        <div class="sec-hd">⚠ 三大潛在風險</div>
        ${top3Risks.map((r, i) => `
          <div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:10px;padding:10px;background:rgba(239,68,68,0.06);border-radius:8px;border-left:3px solid var(--red)">
            <span style="color:var(--red);font-weight:900;font-size:16px">${i + 1}</span>
            <div style="flex:1">
              <div style="font-size:13px;font-weight:600;color:var(--text-primary)">${r.name}</div>
              <div style="font-size:11px;color:var(--text-muted)">機率 ${(r.prob * 100).toFixed(0)}% × 衝擊 ${(r.impact * 100).toFixed(0)}%・應對：${r.mitigation || '待制定'}</div>
            </div>
          </div>`).join('')}
      </div>

      <!-- 行動方案 -->
      <div class="glass mb-4" style="padding:20px">
        <div class="sec-hd">📋 建議行動方案（30天內）</div>
        <div class="ai-box">
          <ol style="padding-left:18px;line-height:2">
            ${r.profit < p.targetProfit ? `<li>啟動<b>目標補缺計畫</b>：利潤缺口 ${fmt2(p.targetProfit - r.profit)} 萬，透過銷量提升 ${fmt2((p.targetProfit - r.profit) / r.totalCM * r.totalRev / r.totalRev * 100)}% 或降低成本可彌補。</li>` : ''}
            ${r.safetyMarginRate < p.safetyWarn ? `<li>緊急審視<b>固定成本結構</b>：安全邊際率偏低（${(r.safetyMarginRate * 100).toFixed(1)}%），建議將部分固定成本轉為變動成本。</li>` : ''}
            <li>每月追蹤月度趨勢 KPI，當 CM 率月變動 > ±3% 時立即分析原因。</li>
            <li>更新風險矩陣中的<b>高風險事件</b>應對計畫，確保緊急預案就位。</li>
          </ol>
        </div>
      </div>

      <!-- 列印按鈕 -->
      <div style="text-align:center;margin-top:20px">
        <button class="btn-primary" onclick="window.print()">🖨 列印為 PDF</button>
      </div>
    </div>`;
}
