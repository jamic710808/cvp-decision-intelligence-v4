/* ╔════════════════════════════════════════════════════╗
   ║  T5 — 多情境平行比較（含機率加權期望利潤）          ║
   ║  T6 — 投資 ROI                                     ║
   ║  T7 — AI 目標搜尋                                  ║
   ║  T8 — 經營槓桿度分析                               ║
   ╚════════════════════════════════════════════════════╝ */

/* ── T5: 多情境 ── */
function renderT5() {
    const analysis = runScenarioAnalysis();
    const { scenarios, expectedProfit } = analysis;
    const base = scenarios.find(s => s.name === '基準情境')?.result || runEngine();

    /* 情境表格 */
    const tbody = document.getElementById('scenario-tbody');
    if (!tbody) return;
    tbody.innerHTML = scenarios.map((s, i) => {
        const r = s.result;
        const chg = ((r.profit - base.profit) / Math.abs(base.profit || 1) * 100);
        const isBase = i === 0 && s.name === '基準情境';
        const isUserDefined = i >= AppState.scenarios.length;
        return `<tr>
      <td style="text-align:left;font-weight:${isBase ? 700 : 400}">${s.name}</td>
      <td>${(s.prob * 100).toFixed(0)}%</td>
      <td>${(s.vol >= 0 ? '+' : '')}${(s.vol || 0).toFixed(0)}%</td>
      <td>${(s.price >= 0 ? '+' : '')}${(s.price || 0).toFixed(0)}%</td>
      <td>${fmt2(r.totalRev)}</td>
      <td>${fmt2(r.totalCM)}</td>
      <td>${pct1(r.cmRate)}</td>
      <td class="${posNegClass(r.profit)}">${fmt2(r.profit)}</td>
      <td>${isBase ? '<span class="badge badge-blue">基準</span>' :
                `<span class="badge ${chg > 5 ? 'badge-green' : chg < -5 ? 'badge-red' : 'badge-yellow'}">${fmtDeltaPct(chg)}</span>`}</td>
      <td>${isUserDefined
                ? `<button onclick="removeUserScenario(${i - AppState.scenarios.length})" style="color:var(--red);cursor:pointer;border:none;background:none;font-size:11px">✕ 刪除</button>`
                : '—'}</td>
    </tr>`;
    }).join('');

    /* 加權期望利潤顯示 */
    const elpEl = document.getElementById('expected-profit');
    if (elpEl) elpEl.innerHTML = `
    <div class="kpi-label">📊 加權期望利潤</div>
    <div class="kpi-val ${posNegClass(expectedProfit)}">${fmt2(expectedProfit)}</div>
    <div class="kpi-sub">萬元（依各情境發生機率加權）</div>`;

    /* 情境利潤對比直方圖 */
    const profits = scenarios.map(s => s.result.profit);
    const labels = scenarios.map(s => s.name);
    barChart('ch-scenario', labels,
        [{
            label: '利潤（萬元）', data: profits,
            backgroundColor: profits.map(v => v >= 0 ? 'rgba(16,185,129,0.75)' : 'rgba(239,68,68,0.75)')
        }],
        { plugins: { legend: { display: false } } }
    );

    /* 目標利潤水平線（Chart.js annotation 暫用 dataset 近似）*/

    /* 雷達圖：前 4 情境 */
    const topFour = scenarios.slice(0, 4);
    const maxRev = Math.max(...topFour.map(s => s.result.totalRev)) || 1;
    const maxCM = Math.max(...topFour.map(s => s.result.totalCM)) || 1;
    const maxProf = Math.max(...topFour.map(s => Math.abs(s.result.profit))) || 1;
    radarChart('ch-scenario-radar',
        ['總收入(相對)', '邊際貢獻(相對)', '淨利(相對)', 'CM率(%)', '安全邊際(%)'],
        topFour.map(s => ({
            label: s.name,
            data: [
                s.result.totalRev / maxRev * 100,
                s.result.totalCM / maxCM * 100,
                s.result.profit / maxProf * 100,
                s.result.cmRate * 100,
                s.result.safetyMarginRate * 100,
            ],
        }))
    );
}
function removeUserScenario(idx) {
    AppState.userScenarios.splice(idx, 1);
    saveState();
    renderT5();
}

/* ── T6: 投資 ROI ── */
function renderT6() {
    const inv = runInvestment();
    document.getElementById('inv-kpis').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">NPV（萬元）</div>
      <div class="kpi-val ${posNegClass(inv.npv)}">${fmt2(inv.npv)}</div></div>
    <div class="kpi-card"><div class="kpi-label">IRR</div>
      <div class="kpi-val">${(inv.irr * 100).toFixed(1)}%</div>
      <div class="kpi-sub">WACC：${(AppState.investment.wacc * 100).toFixed(1)}%</div></div>
    <div class="kpi-card"><div class="kpi-label">回收期（PBP）</div>
      <div class="kpi-val">${isFinite(inv.pbp) ? inv.pbp.toFixed(1) : '∞'}</div>
      <div class="kpi-sub">年</div></div>
    <div class="kpi-card"><div class="kpi-label">投資報酬率</div>
      <div class="kpi-val ${posNegClass(inv.roi)}">${(inv.roi * 100).toFixed(1)}%</div></div>`;
    /* 現金流瀑布 */
    const cfLabels = ['第0年（投資）', ...inv.cfArr.slice(1).map((_, i) => `第${i + 1}年`)];
    waterfallChart('ch-cf', cfLabels, inv.cfArr);
    /* 累計現金流（回收期線） */
    lineChart('ch-payback', cfLabels, [
        { label: '累計現金流（萬）', data: inv.cumulativeCF, borderColor: '#8b5cf6', fill: true, backgroundColor: 'rgba(139,92,246,0.10)' },
    ]);
    /* NPV 敏感度 */
    lineChart('ch-npv-sens',
        inv.waccRange.map(w => (w * 100).toFixed(0) + '%'),
        [{
            label: 'NPV（萬）', data: inv.npvSens, borderColor: '#a78bfa', fill: true, backgroundColor: 'rgba(139,92,246,0.12)',
            pointBackgroundColor: inv.npvSens.map(v => v >= 0 ? 'var(--emerald)' : 'var(--red)'), pointRadius: 5
        }],
        { xTitle: '折現率（WACC）', yTitle: 'NPV（萬元）' }
    );
}

/* ── T7: AI 目標搜尋 ── */
function runGoalSeekUI() {
    const targetType = document.getElementById('gs-target-type').value;
    const targetVal = parseFloat(document.getElementById('gs-target-val').value);
    const varType = document.getElementById('gs-var').value;
    if (isNaN(targetVal)) { showToast('⚠ 請輸入有效目標值', 'warn'); return; }
    const res = runGoalSeek(targetType, targetVal, varType);
    const { bestAdj, finalR, varType: vt, targetType: tt, targetVal: tv } = res;
    const varName = { vol: '銷量', price: '單價', vc: '變動成本', fc: '固定成本' }[vt];
    const targetName = { profit: '目標利潤', margin: '邊際貢獻率', safety: '安全邊際率' }[tt];
    document.getElementById('gs-result').innerHTML = `
    <div class="sec-hd">🤖 AI 搜尋結果</div>
    <div class="kpi-grid-4 mb-4">
      <div class="kpi-card"><div class="kpi-label">${varName} 需調整</div>
        <div class="kpi-val">${bestAdj >= 0 ? '+' : ''}${bestAdj.toFixed(1)}%</div></div>
      <div class="kpi-card"><div class="kpi-label">目標：${targetName}</div>
        <div class="kpi-val success">${tt === 'profit' ? fmt2(tv) + ' 萬' : tv.toFixed(1) + '%'}</div></div>
      <div class="kpi-card"><div class="kpi-label">達成後利潤(萬)</div>
        <div class="kpi-val ${posNegClass(finalR.profit)}">${fmt2(finalR.profit)}</div></div>
      <div class="kpi-card"><div class="kpi-label">安全邊際率</div>
        <div class="kpi-val">${(finalR.safetyMarginRate * 100).toFixed(1)}%</div></div>
    </div>
    <div class="ai-box">💡 <b style="color:var(--violet-lt)">AI 建議：</b>
      若將 <b style="color:var(--sky-lt)">${varName}</b> 調整
      <b style="color:var(--amber)">${bestAdj >= 0 ? '+' : ''}${bestAdj.toFixed(1)}%</b>，
      可使 ${targetName} 達到目標值。建議搭配多情境分析進一步評估可行性。</div>`;
}

/* ── T8: DOL ── */
function renderDOL() {
    const r = runEngine();
    _setKPIVal('dol-val', isFinite(r.dol) ? r.dol.toFixed(2) : '∞');
    _setKPIVal('dol-cm', fmt2(r.totalCM));
    _setKPIVal('dol-profit', fmt2(r.profit));
    const impact = document.getElementById('dol-impact');
    if (impact) impact.textContent = isFinite(r.dol) ? r.dol.toFixed(2) : '∞';

    const range = +document.getElementById('dol-range').value;
    const STEPS = 15;
    const labels = [], revChg = [], profChg = [];
    for (let i = -STEPS; i <= STEPS; i++) {
        const pct = (i / STEPS) * range;
        labels.push((pct >= 0 ? '+' : '') + pct.toFixed(0) + '%');
        const sim = runEngine({ volAdj: pct });
        revChg.push(+((sim.totalRev - r.totalRev) / r.totalRev * 100).toFixed(1));
        profChg.push(r.profit !== 0 ? +((sim.profit - r.profit) / Math.abs(r.profit) * 100).toFixed(1) : 0);
    }
    lineChart('ch-dol', labels, [
        { label: '收入變動%', data: revChg, borderColor: '#0ea5e9' },
        { label: '利潤變動(DOL放大)%', data: profChg, borderColor: '#f59e0b', borderWidth: 2.5 },
    ]);

    document.getElementById('dol-table-body').innerHTML = r.prodCalc.map(p => {
        const share = r.totalCM > 0 ? p.cm / r.totalCM : 0;
        const lev = r.profit !== 0 ? p.cm / r.profit : 0;
        return `<tr>
      <td style="text-align:left">${p.name}</td>
      <td>${fmt2(p.cm)}</td><td>${pct1(p.cmRate)}</td>
      <td>${(share * 100).toFixed(1)}%</td>
      <td><span class="badge ${lev > 3 ? 'badge-red' : lev > 1.5 ? 'badge-yellow' : 'badge-green'}">${lev.toFixed(2)}×</span></td>
    </tr>`;
    }).join('');
}
function renderT8() { renderDOL(); }
