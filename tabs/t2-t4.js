/* ╔════════════════════════════════════════╗
   ║  T2 — 管理損益表 + 瀑布圖              ║
   ╚════════════════════════════════════════╝ */

function renderT2() {
    const r = runEngine();
    const prods = r.prodCalc;

    /* 損益表 */
    let hdr = `<thead><tr><th style="text-align:left">項目</th>
    ${prods.map(p => `<th>${p.name.split(' ')[0]}</th>`).join('')}
    <th>合計</th></tr></thead>`;
    const rows = [
        { label: '收入（萬元）', vals: prods.map(p => fmt2(p.rev)), total: fmt2(r.totalRev), cls: 'text-sky' },
        { label: '變動成本（萬元）', vals: prods.map(p => fmt2(p.vcTotal)), total: fmt2(r.totalVC), cls: '' },
        { label: '邊際貢獻（萬元）', vals: prods.map(p => fmt2(p.cm)), total: fmt2(r.totalCM), cls: 'text-pos text-bold' },
        { label: '邊際貢獻率', vals: prods.map(p => pct1(p.cmRate)), total: pct1(r.cmRate), cls: 'warn' },
        { label: '固定成本（萬元）', vals: prods.map(() => '—'), total: fmt2(r.fc), cls: '' },
        { label: '營業利潤（萬元）', vals: prods.map(() => '—'), total: fmt2(r.profit), cls: r.profit >= 0 ? 'text-pos' : 'text-neg' },
    ];
    document.getElementById('pl-table').innerHTML = hdr + '<tbody>' +
        rows.map(row => `<tr>
      <td style="text-align:left;font-weight:600;color:var(--text-secondary)">${row.label}</td>
      ${row.vals.map(v => `<td>${v}</td>`).join('')}
      <td class="${row.cls}" style="font-weight:700">${row.total}</td>
    </tr>`).join('') + '</tbody>';

    /* 瀑布圖 */
    waterfallChart('ch-waterfall',
        ['總收入', '−變動成本', '＝邊際貢獻', '−固定成本', '＝營業利潤'],
        [r.totalRev, -r.totalVC, r.totalCM, -r.fc, r.profit],
        { label: '金額（萬元）' }
    );

    /* 成本結構圓餅 */
    doughnutChart('ch-cost-struct',
        ['變動成本', '固定成本', '營業利潤'],
        [+r.totalVC.toFixed(0), +r.fc.toFixed(0), Math.max(+r.profit.toFixed(0), 0)],
        {}
    );
}

/* ╔════════════════════════════════════════╗
   ║  T3 — CVP 靜態分析                     ║
   ╚════════════════════════════════════════╝ */
function renderT3() {
    const r = runEngine();

    /* KPI */
    _setKPIVal('bep-rev', fmt2(r.bepRev));
    _setKPIVal('safety-m', (r.safetyMarginRate * 100).toFixed(1));
    const smEl = document.getElementById('safety-m');
    if (smEl) smEl.className = 'kpi-val ' + posNegClass(r.safetyMarginRate - AppState.params.safetyWarn);

    /* CVP 損益平衡折線圖（0% ~ 180% 銷量）*/
    const STEPS = 24;
    const labels = [], revLine = [], totalCostLine = [], fcLine = [];
    for (let i = 0; i <= STEPS; i++) {
        const f = (i / STEPS) * 1.8;
        labels.push((f * 100).toFixed(0) + '%');
        revLine.push(+(r.totalRev * f).toFixed(0));
        fcLine.push(+r.fc.toFixed(0));
        totalCostLine.push(+(r.totalVC * f + r.fc).toFixed(0));
    }
    lineChart('ch-cvp', labels, [
        { label: '總收入', data: revLine, borderColor: '#0ea5e9', backgroundColor: 'rgba(14,165,233,0.10)', fill: true },
        { label: '總成本', data: totalCostLine, borderColor: '#ef4444', backgroundColor: 'rgba(239,68,68,0.08)', fill: true },
        { label: '固定成本', data: fcLine, borderColor: '#f59e0b', borderDash: [6, 4], fill: false },
    ], { plugins: { legend: { display: true } }, annotation: {} });

    /* 產品貢獻明細表 */
    document.getElementById('product-table').innerHTML =
        `<thead><tr>
      <th style="text-align:left">產品</th><th>售價(萬)</th><th>變動成本(萬)</th>
      <th>CM/件(萬)</th><th>CM率</th><th>銷量(件)</th><th>CM合計(萬)</th><th>狀態</th>
    </tr></thead><tbody>` +
        r.prodCalc.map(p => `<tr>
      <td style="text-align:left">${p.name}</td>
      <td>${fmt2(p.price)}</td><td>${fmt2(p.vc)}</td>
      <td>${fmt2(p.cmPerUnit)}</td><td>${pct1(p.cmRate)}</td>
      <td>${fmt0(p.vol)}</td><td>${fmt2(p.cm)}</td>
      <td><span class="badge ${cmRatingClass(p.cmRate)}">${cmRatingText(p.cmRate)}</span></td>
    </tr>`).join('') + '</tbody>';
}

function _setKPIVal(id, text) {
    const el = document.getElementById(id);
    if (el) animateKPI(el, text);
}

/* ╔════════════════════════════════════════╗
   ║  T4 — 動態情境模擬器                   ║
   ╚════════════════════════════════════════╝ */

function simUpdate() {
    const v = +document.getElementById('vol-slider').value;
    const p = +document.getElementById('price-slider').value;
    const vc = +document.getElementById('vc-slider').value;
    const fc = +document.getElementById('fc-slider').value;

    document.getElementById('vol-lbl').textContent = (v >= 0 ? '+' : '') + v + '%';
    document.getElementById('price-lbl').textContent = (p >= 0 ? '+' : '') + p + '%';
    document.getElementById('vc-lbl').textContent = (vc >= 0 ? '+' : '') + vc + '%';
    document.getElementById('fc-lbl').textContent = (fc >= 0 ? '+' : '') + fc + '%';

    const base = runEngine();
    const sim = runEngine({ volAdj: v, priceAdj: p, vcAdj: vc, fcAdj: fc });
    const chg = (sim.profit - base.profit) / Math.abs(base.profit || 1) * 100;

    /* 敏感係數 */
    const prof = Math.abs(sim.profit) > 0.001 ? sim.profit : 0.001;
    const sens = {
        price: sim.totalRev / prof,
        vol: sim.totalCM / prof,
        vc: -sim.totalVC / prof,
        fc: -sim.fc / prof,
    };

    document.getElementById('sensitivity-factors').innerHTML = [
        ['售價', sens.price, 'text-sky'],
        ['銷量', sens.vol, 'text-pos'],
        ['單位變動成本', sens.vc, 'warn'],
        ['固定成本', sens.fc, 'text-neg'],
    ].map(([l, v, cls]) => `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
      <span style="color:var(--text-secondary);font-size:12px">${l}</span>
      <span class="${cls}" style="font-family:monospace;font-size:13px;font-weight:700">${v.toFixed(2)}</span>
    </div>
    <div style="margin-bottom:10px">
      <div class="impact-label" style="font-size:10px;color:var(--text-muted);text-align:right">
        此調整預計影響利潤：<b>${fmtDelta(Math.abs(v) * 0.01 * Math.abs(base.profit))}</b> 萬
      </div>
    </div>`
    ).join('');

    document.getElementById('sim-kpis').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">模擬收入(萬)</div><div class="kpi-val">${fmt2(sim.totalRev)}</div></div>
    <div class="kpi-card"><div class="kpi-label">模擬 CM 率</div><div class="kpi-val">${(sim.cmRate * 100).toFixed(1)}%</div></div>
    <div class="kpi-card"><div class="kpi-label">模擬利潤(萬)</div><div class="kpi-val ${posNegClass(sim.profit)}">${fmt2(sim.profit)}</div></div>
    <div class="kpi-card"><div class="kpi-label">利潤變動</div><div class="kpi-val ${posNegClass(chg)}">${fmtDeltaPct(chg)}</div></div>`;

    barChart('ch-sim',
        ['總收入', '變動成本', '邊際貢獻', '固定成本', '利潤'],
        [
            { label: '基準', data: [base.totalRev, base.totalVC, base.totalCM, base.fc, base.profit] },
            { label: '模擬', data: [sim.totalRev, sim.totalVC, sim.totalCM, sim.fc, sim.profit] },
        ]
    );
}
function simReset() {
    ['vol', 'price', 'vc', 'fc'].forEach(k => {
        const s = document.getElementById(k + '-slider');
        const l = document.getElementById(k + '-lbl');
        if (s) s.value = 0;
        if (l) l.textContent = '0%';
    });
    simUpdate();
}
function addScenario() {
    const v = +document.getElementById('vol-slider').value;
    const p = +document.getElementById('price-slider').value;
    const vc = +document.getElementById('vc-slider').value;
    const fc = +document.getElementById('fc-slider').value;
    const n = AppState.userScenarios.length + 1;
    AppState.userScenarios.push({ name: `自訂情境 ${n}`, vol: v, price: p, vc, fc, prob: 0 });
    saveState();
    showToast(`✅ 自訂情境 ${n} 已儲存`, 'ok');
    logChange(`儲存情境：銷量${v >= 0 ? '+' : ''}${v}% 單價${p >= 0 ? '+' : ''}${p}%`);
}
function renderT4() { simUpdate(); }
