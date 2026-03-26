/* ╔══════════════════════════════════════════════════════════╗
   ║  T1 — 執行長總覽（含 Sparklines & 分層警示）              ║
   ╚══════════════════════════════════════════════════════════╝ */

function renderT1() {
    const r = runEngine();
    const p = AppState.params;
    const td = AppState.monthlyTrend;

    /* ── KPI 卡片 ── */
    const profitDelta = r.profit - p.targetProfit;
    const kpiGrid = document.getElementById('kpi-grid');
    if (!kpiGrid) return;
    kpiGrid.innerHTML = `
    <div class="kpi-card" id="kpi-rev">
      <div class="kpi-label">💰 總收入</div>
      <div class="kpi-val" id="kv-rev">${fmt2(r.totalRev)}</div>
      <div class="kpi-sub">萬元</div>
      <div class="sparkline-wrap" id="spark-rev"></div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">📊 邊際貢獻率</div>
      <div class="kpi-val ${posNegClass(r.cmRate - p.industryAvgCM)}" id="kv-cm">${(r.cmRate * 100).toFixed(1)}%</div>
      <div class="kpi-sub">行業均：${pct0(p.industryAvgCM)}</div>
      <div class="sparkline-wrap" id="spark-cm"></div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">💵 營業利潤</div>
      <div class="kpi-val ${posNegClass(r.profit - p.targetProfit)}" id="kv-profit">${fmt2(r.profit)}</div>
      <div class="kpi-sub">目標：${fmt2(p.targetProfit)}　差異：<span class="${posNegClass(profitDelta)}">${fmtDelta(profitDelta)}</span></div>
      <div class="sparkline-wrap" id="spark-profit"></div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">🛡 安全邊際率</div>
      <div class="kpi-val ${posNegClass(r.safetyMarginRate - p.safetyWarn)}" id="kv-safety">${(r.safetyMarginRate * 100).toFixed(1)}%</div>
      <div class="kpi-sub">警示門檻：${pct0(p.safetyWarn)}</div>
      <div class="sparkline-wrap" id="spark-safety"></div>
    </div>`;

    /* ── Sparklines（從月度趨勢取資料）── */
    if (td.length > 0) {
        requestAnimationFrame(() => {
            drawSparkline(document.getElementById('spark-rev'), td.map(m => m.rev));
            drawSparkline(document.getElementById('spark-cm'), td.map(m => m.cmRate * 100));
            drawSparkline(document.getElementById('spark-profit'), td.map(m => m.profit));
            drawSparkline(document.getElementById('spark-safety'), td.map(m => {
                const b = runEngine(); // 每月安全邊際暫用全年數值替代（月度傾向誤差，僅做趨勢參考）
                return m.cm > 0 ? (m.cm - b.bepRev / 12) / m.rev * 100 : 0;
            }));
        });
    }

    /* ── 收入 vs 成本結構分組長條圖 ── */
    barChart('ch-overview',
        AppState.products.map(p => p.name.split(' ')[0]),
        [
            { label: '收入(萬)', data: AppState.products.map(p => +(p.price * p.vol).toFixed(0)) },
            { label: '邊際貢獻(萬)', data: AppState.products.map(p => +((p.price - p.vc) * p.vol).toFixed(0)) },
        ],
        { plugins: { legend: { display: true } } }
    );

    /* ── 產品貢獻甜甜圈 ── */
    doughnutChart('ch-donut',
        AppState.products.map(p => p.name.split(' ')[0]),
        AppState.products.map(p => +(p.price * p.vol).toFixed(0))
    );

    /* ── 警示中心（分層：危急/注意/正常）── */
    const alerts = [];

    // 危急
    if (r.safetyMarginRate < p.safetyWarn)
        alerts.push({
            cls: 'badge-red', icon: '🚨', level: '危急',
            msg: `安全邊際率 <b>${(r.safetyMarginRate * 100).toFixed(1)}%</b> 低於警戒值 <b>${(p.safetyWarn * 100).toFixed(0)}%</b>，建議立即提升銷量或降低成本。`
        });
    if (r.profit < 0)
        alerts.push({
            cls: 'badge-red', icon: '🔴', level: '危急',
            msg: `公司處於<b>虧損狀態</b>（利潤 ${fmt2(r.profit)} 萬），需緊急審視成本結構或訂價。`
        });

    // 注意
    if (r.cmRate < p.industryAvgCM)
        alerts.push({
            cls: 'badge-yellow', icon: '⚠', level: '注意',
            msg: `邊際貢獻率 <b>${(r.cmRate * 100).toFixed(1)}%</b> 低於行業均值 <b>${(p.industryAvgCM * 100).toFixed(0)}%</b>，考慮提升售價或降低變動成本。`
        });
    if (r.profit >= 0 && r.profit < p.targetProfit)
        alerts.push({
            cls: 'badge-yellow', icon: '💸', level: '注意',
            msg: `利潤 <b>${fmt2(r.profit)} 萬</b> 未達目標 ${fmt2(p.targetProfit)} 萬，缺口 <b>${fmt2(p.targetProfit - r.profit)} 萬</b>。`
        });
    if (r.dol > 5 && isFinite(r.dol))
        alerts.push({
            cls: 'badge-yellow', icon: '⚡', level: '注意',
            msg: `經營槓桿度 <b>${r.dol.toFixed(1)} 倍</b>偏高，銷量波動對利潤衝擊巨大，建議降低固定成本佔比。`
        });

    // 產能瓶頸
    AppState.products.forEach(p => {
        if (p.capMax < 9999 && p.vol / p.capMax > 0.85)
            alerts.push({
                cls: 'badge-blue', icon: '🏭', level: '注意',
                msg: `${p.name.split(' ')[0]} 產能利用率 <b>${(p.vol / p.capMax * 100).toFixed(0)}%</b>，接近瓶頸，考慮擴產或優化排程。`
            });
    });

    if (!alerts.length)
        alerts.push({ cls: 'badge-green', icon: '✅', level: '正常', msg: '各項指標健康，無立即風險。' });

    document.getElementById('alert-center').innerHTML =
        alerts.map(a => alertItem(a.cls, a.icon, a.level, a.msg)).join('');
}
