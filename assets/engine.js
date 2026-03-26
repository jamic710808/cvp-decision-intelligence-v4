/* ╔══════════════════════════════════════════════════════════╗
   ║  CVP V4.0 — 核心計算引擎                                 ║
   ║  包含：CVP / 投資 / 資金效率 / 蒙地卡羅（常態分佈）        ║
   ╚══════════════════════════════════════════════════════════╝ */

/* ══════════════════════════════════════
   1. 主 CVP 引擎
   ══════════════════════════════════════ */
/**
 * 核心 CVP 計算
 * @param {object} opts - { volAdj, priceAdj, vcAdj, fcAdj }（百分比，如 10 = +10%）
 * @returns {object} 完整計算結果
 */
function runEngine(opts = {}) {
    const { volAdj = 0, priceAdj = 0, vcAdj = 0, fcAdj = 0 } = opts;
    const vf = 1 + volAdj / 100;
    const pf = 1 + priceAdj / 100;
    const cf = 1 + vcAdj / 100;
    const ff = 1 + fcAdj / 100;

    const prods = AppState.products;
    let totalRev = 0, totalVC = 0, totalCM = 0;

    const prodCalc = prods.map(p => {
        const price = p.price * pf;
        const vc = p.vc * cf;
        const vol = p.vol * vf;
        const rev = price * vol;
        const vcTotal = vc * vol;
        const cm = rev - vcTotal;
        const cmRate = rev > 0 ? cm / rev : 0;
        const cmPerUnit = price - vc;
        totalRev += rev;
        totalVC += vcTotal;
        totalCM += cm;
        return {
            name: p.name, cat: p.cat, price, vc, vol, cap: p.capMax,
            rev, vcTotal, cm, cmRate, cmPerUnit
        };
    });

    const fc = AppState.params.fc * ff;
    const profit = totalCM - fc;
    const cmRate = totalRev > 0 ? totalCM / totalRev : 0;

    /* 損益平衡點 */
    const bepRev = cmRate > 0 ? fc / cmRate : Infinity;
    /* 安全邊際 */
    const safetyMarginAmt = totalRev - bepRev;
    const safetyMarginRate = totalRev > 0 ? safetyMarginAmt / totalRev : 0;
    /* 經營槓桿度 */
    const dol = profit !== 0 ? totalCM / profit : Infinity;

    return {
        totalRev, totalVC, totalCM, fc, profit, cmRate,
        bepRev, safetyMarginAmt, safetyMarginRate, dol,
        prodCalc, volAdj, priceAdj, vcAdj, fcAdj,
    };
}

/* ══════════════════════════════════════
   2. 敏感係數計算（龍捲風圖用）
   ══════════════════════════════════════ */
/**
 * 計算各變數對利潤的敏感係數（固定 ±delta 測試）
 * @param {number} delta - 測試幅度（%），預設 10
 * @returns {Array} 敏感係數陣列，按絕對值降序
 */
function runSensitivity(delta = 10) {
    const base = runEngine();
    const baseProfit = base.profit || 1e-9;

    const vars = [
        { key: 'price', label: '售價', opts: p => ({ priceAdj: p }) },
        { key: 'vol', label: '銷量', opts: p => ({ volAdj: p }) },
        { key: 'vc', label: '單位變動成本', opts: p => ({ vcAdj: p }) },
        { key: 'fc', label: '固定成本', opts: p => ({ fcAdj: p }) },
    ];

    return vars.map(v => {
        const rPos = runEngine(v.opts(+delta));
        const rNeg = runEngine(v.opts(-delta));
        const impactPos = (rPos.profit - baseProfit) / Math.abs(baseProfit) * 100;
        const impactNeg = (rNeg.profit - baseProfit) / Math.abs(baseProfit) * 100;
        return {
            key: v.key,
            label: v.label,
            impactPos: +impactPos.toFixed(1),
            impactNeg: +impactNeg.toFixed(1),
            maxAbs: Math.max(Math.abs(impactPos), Math.abs(impactNeg)),
        };
    }).sort((a, b) => b.maxAbs - a.maxAbs);
}

/* ══════════════════════════════════════
   3. 情境比較（含機率加權期望利潤）
   ══════════════════════════════════════ */
/**
 * 計算所有情境結果與加權期望利潤
 * @returns {object} { scenarios: [...], expectedProfit: number, totalProb: number }
 */
function runScenarioAnalysis() {
    const allScenarios = [
        ...AppState.scenarios,
        ...AppState.userScenarios.map(s => ({ ...s, prob: s.prob ?? 0 })),
    ];

    let weightedProfit = 0;
    let totalProb = 0;
    const results = allScenarios.map(s => {
        const r = runEngine({ volAdj: s.vol, priceAdj: s.price, vcAdj: s.vc || 0, fcAdj: s.fc || 0 });
        const prob = s.prob ?? 0;
        weightedProfit += r.profit * prob;
        totalProb += prob;
        return { ...s, result: r };
    });

    const expectedProfit = totalProb > 0 ? weightedProfit : 0;
    return { scenarios: results, expectedProfit, totalProb };
}

/* ══════════════════════════════════════
   4. 投資 ROI 引擎
   ══════════════════════════════════════ */
/**
 * 計算 NPV / IRR / PBP / ROI
 * @returns {object} 投資分析完整結果
 */
function runInvestment() {
    const { capex, years, wacc, depLife, salvage } = AppState.investment;
    const base = runEngine();
    const annualProfit = base.profit;
    const dep = capex * (1 - salvage) / depLife;

    /* 自由現金流（FCFE 簡化）*/
    const fcfe = annualProfit + dep;

    /* NPV */
    let npv = -capex;
    const cfArr = [-capex];
    for (let y = 1; y <= years; y++) {
        const cf = y === years ? fcfe + capex * salvage : fcfe;
        npv += cf / Math.pow(1 + wacc, y);
        cfArr.push(+cf.toFixed(0));
    }

    /* 累計現金流（回收期用）*/
    const cumulativeCF = [];
    let cum = 0;
    cfArr.forEach(v => { cum += v; cumulativeCF.push(+cum.toFixed(0)); });

    /* PBP（精確插值）*/
    let pbp = Infinity;
    for (let i = 1; i < cumulativeCF.length; i++) {
        if (cumulativeCF[i] >= 0) {
            pbp = i - 1 + Math.abs(cumulativeCF[i - 1]) / Math.abs(cfArr[i]);
            break;
        }
    }

    /* IRR（二分法 80 次）*/
    let irr = 0;
    if (fcfe > 0) {
        let lo = 0, hi = 10, mid;
        for (let i = 0; i < 80; i++) {
            mid = (lo + hi) / 2;
            let npvMid = -capex;
            for (let y = 1; y <= years; y++) {
                npvMid += (y === years ? fcfe + capex * salvage : fcfe) / Math.pow(1 + mid, y);
            }
            if (npvMid > 0) lo = mid; else hi = mid;
        }
        irr = mid;
    }

    /* NPV 對不同 WACC 的敏感性 */
    const waccRange = [0.03, 0.04, 0.05, 0.06, 0.07, 0.08, 0.09, 0.10, 0.12, 0.15, 0.18, 0.20];
    const npvSens = waccRange.map(w => {
        let n = -capex;
        for (let y = 1; y <= years; y++) {
            n += (y === years ? fcfe + capex * salvage : fcfe) / Math.pow(1 + w, y);
        }
        return +n.toFixed(0);
    });

    const roi = capex > 0 ? npv / capex : 0;
    return {
        npv: +npv.toFixed(0), irr, pbp, roi, fcfe, dep,
        cfArr, cumulativeCF, annualProfit,
        waccRange, npvSens
    };
}

/* ══════════════════════════════════════
   5. 資金效率（杜邦分析）
   ══════════════════════════════════════ */
/**
 * 計算 ROE / ROA / 資產周轉率 / 杜邦三因子
 * @returns {object} 資金效率完整結果
 */
function runCapital() {
    const c = AppState.capital;
    const base = runEngine();
    const annualRev = base.totalRev;
    const annualProfit = base.profit;

    const totalAssets = c.currentAssets + c.fixedAssets + c.otherAssets;
    const roe = c.equity > 0 ? annualProfit / c.equity : 0;
    const roa = totalAssets > 0 ? annualProfit / totalAssets : 0;
    const assetTurnover = totalAssets > 0 ? annualRev / totalAssets : 0;
    const profitMargin = annualRev > 0 ? annualProfit / annualRev : 0;
    const leverage = c.equity > 0 ? totalAssets / c.equity : 0;

    // DuPont 驗算：ROE ≈ PM × AT × EM
    const dupontROE = profitMargin * assetTurnover * leverage;

    return {
        roe, roa, assetTurnover, profitMargin, leverage, dupontROE,
        equity: c.equity, totalAssets, annualRev, annualProfit
    };
}

/* ══════════════════════════════════════
   6. 現金流預測引擎
   ══════════════════════════════════════ */
/**
 * 結合月度趨勢 + 投資計畫，生成 12 個月滾動現金流預測
 * @returns {Array} 12 個月現金流詳情
 */
function runCashflowForecast() {
    const { arDays, apDays, safetyBuffer } = AppState.cashflowParams;
    const td = AppState.monthlyTrend;
    const inv = AppState.investment;
    const monthlyCapex = inv.capex / inv.years / 12; // 月均投資支出

    let cumulativeCash = safetyBuffer * 2; // 假設期初現金
    return td.map((m, i) => {
        /* 營運現金流（考量應收/應付帳款延遲）*/
        const cashFromSales = m.rev * (1 - arDays / 30);   // 本月收現比例
        const cashToSupply = m.vc * (1 - apDays / 30);   // 本月付款比例
        const operatingCF = cashFromSales - cashToSupply - m.fc / 12;
        /* 投資現金流 */
        const investingCF = -monthlyCapex;
        /* 淨現金流 */
        const netCF = operatingCF + investingCF;
        cumulativeCash += netCF;
        const isBelowSafety = cumulativeCash < safetyBuffer;
        return {
            month: m.month,
            operatingCF: +operatingCF.toFixed(0),
            investingCF: +investingCF.toFixed(0),
            netCF: +netCF.toFixed(0),
            cumulative: +cumulativeCash.toFixed(0),
            isBelowSafety,
        };
    });
}

/* ══════════════════════════════════════
   7. 蒙地卡羅（常態分佈 Box-Muller）
   ══════════════════════════════════════ */
/**
 * Box-Muller 轉換：產生標準正態分佈亂數
 */
function _boxMuller() {
    const u1 = Math.random(), u2 = Math.random();
    return Math.sqrt(-2 * Math.log(u1)) * Math.cos(2 * Math.PI * u2);
}

/**
 * 蒙地卡羅利潤模擬（常態分佈）
 * @param {number} n          - 模擬次數，預設 10,000
 * @param {number} revStdPct  - 收入標準差（%）
 * @param {number} costStdPct - 成本標準差（%）
 * @returns {object} 模擬統計結果
 */
function runMonteCarlo(n = 10000, revStdPct = 15, costStdPct = 10) {
    const base = runEngine();
    const revMu = base.totalRev;
    const vcMu = base.totalVC;
    const fc = base.fc;
    const revSig = revMu * (revStdPct / 100);
    const vcSig = vcMu * (costStdPct / 100);

    const profits = new Float64Array(n);
    for (let i = 0; i < n; i++) {
        const rv = revMu + _boxMuller() * revSig;
        const vc = vcMu + _boxMuller() * vcSig;
        profits[i] = rv - vc - fc;
    }

    /* 排序 */
    profits.sort((a, b) => a - b);
    const mean = profits.reduce((s, v) => s + v, 0) / n;
    const variance = profits.reduce((s, v) => s + (v - mean) ** 2, 0) / n;
    const std = Math.sqrt(variance);
    const p5 = profits[Math.floor(n * 0.05)];
    const p25 = profits[Math.floor(n * 0.25)];
    const p50 = profits[Math.floor(n * 0.50)];
    const p75 = profits[Math.floor(n * 0.75)];
    const p95 = profits[Math.floor(n * 0.95)];

    /* VaR（損失角度）*/
    const var95 = -p5;   // 95% VaR（損失值，正數）
    const var99 = -profits[Math.floor(n * 0.01)]; // 99% VaR
    /* CVaR（超越損失均值）*/
    const worstProfits = Array.from(profits).slice(0, Math.floor(n * 0.05));
    const cvar95 = -worstProfits.reduce((s, v) => s + v, 0) / worstProfits.length;

    const lossCount = profits.filter(v => v < 0).length;
    const lossProb = lossCount / n;
    return { profits, mean, std, p5, p25, p50, p75, p95, var95, var99, cvar95, lossProb, n };
}

/* ══════════════════════════════════════
   8. 產品組合優化器（改善版）
   ══════════════════════════════════════ */
/**
 * 啟發式最大化邊際貢獻（考慮最低銷量約束）
 * @returns {object} 優化結果
 */
function runOptimizer() {
    const prods = AppState.products;

    // 按 CM/件 降序排序
    const sorted = [...prods]
        .map((p, i) => ({ ...p, idx: i, cmPerUnit: p.price - p.vc }))
        .sort((a, b) => b.cmPerUnit - a.cmPerUnit);

    // 初始化為最低銷量
    const optVols = prods.map(p => Math.max(p.minVol ?? 0, p.vol));

    // 貪婪分配：提升高 CM 產品至產能上限（不超過原有銷量 1.5 倍）
    sorted.forEach(p => {
        const currentVol = optVols[p.idx];
        const maxFeasible = Math.min(p.capMax, p.vol * 1.5);
        optVols[p.idx] = Math.round(Math.min(maxFeasible, Math.max(currentVol, p.vol * 1.15)));
    });

    const optProds = prods.map((p, i) => ({ ...p, vol: optVols[i] }));
    const origCM = prods.reduce((s, p) => s + (p.price - p.vc) * p.vol, 0);
    const optCM = optProds.reduce((s, p) => s + (p.price - p.vc) * p.vol, 0);
    const origRev = prods.reduce((s, p) => s + p.price * p.vol, 0);
    const optRev = optProds.reduce((s, p) => s + p.price * p.vol, 0);
    const origProfit = origCM - AppState.params.fc;
    const optProfit = optCM - AppState.params.fc;

    return { optProds, origCM, optCM, origRev, optRev, origProfit, optProfit };
}

/* ══════════════════════════════════════
   9. 預算差異分析
   ══════════════════════════════════════ */
/**
 * 計算實際 vs 預算的各項差異
 * @returns {object} 差異分析結果
 */
function runBudgetVariance() {
    const base = runEngine();
    const budget = AppState.budgetData;
    const prods = AppState.products;

    // 預算 P&L
    let budRev = 0, budVC = 0, budCM = 0;
    const prodVariance = prods.map((p, i) => {
        const b = budget[i] || {};
        const bPrice = b.budgetPrice ?? p.price;
        const bVc = b.budgetVc ?? p.vc;
        const bVol = b.budgetVol ?? p.vol;
        const bRev = bPrice * bVol;
        const bVCT = bVc * bVol;
        const bCM = bRev - bVCT;
        budRev += bRev; budVC += bVCT; budCM += bCM;

        const actRev = base.prodCalc[i]?.rev ?? 0;
        const actCM = base.prodCalc[i]?.cm ?? 0;
        return {
            name: p.name,
            actRev, budRev: bRev, revVar: actRev - bRev, revVarPct: bRev > 0 ? (actRev - bRev) / bRev * 100 : 0,
            actCM, budCM: bCM, cmVar: actCM - bCM, cmVarPct: bCM > 0 ? (actCM - bCM) / bCM * 100 : 0,
        };
    });

    const budFC = AppState.params.fc;   // 預算固定成本（此版本假設與實際相同）
    const budProfit = budCM - budFC;
    const actProfit = base.profit;

    return {
        prodVariance,
        actRev: base.totalRev, budRev, revVar: base.totalRev - budRev,
        actCM: base.totalCM, budCM, cmVar: base.totalCM - budCM,
        actProfit, budProfit, profitVar: actProfit - budProfit,
        actCMRate: base.cmRate,
        budCMRate: budRev > 0 ? budCM / budRev : 0,
    };
}

/* ══════════════════════════════════════
   10. AI 目標搜尋（牛頓法 + 二分法後備）
   ══════════════════════════════════════ */
/**
 * 求解達成目標所需的變數調整幅度
 * @param {string} targetType - 'profit' | 'margin' | 'safety'
 * @param {number} targetVal  - 目標值（萬元 或 %）
 * @param {string} varType    - 'vol' | 'price' | 'vc' | 'fc'
 * @returns {object} 搜尋結果
 */
function runGoalSeek(targetType, targetVal, varType) {
    const _getActual = (adj) => {
        const opts = {};
        if (varType === 'vol') opts.volAdj = adj;
        else if (varType === 'price') opts.priceAdj = adj;
        else if (varType === 'vc') opts.vcAdj = adj;
        else opts.fcAdj = adj;
        const r = runEngine(opts);
        if (targetType === 'profit') return r.profit;
        if (targetType === 'margin') return r.cmRate * 100;
        return r.safetyMarginRate * 100;
    };

    let lo = -80, hi = 200, bestAdj = 0, bestDiff = 1e9;
    for (let i = 0; i < 120; i++) {
        const mid = (lo + hi) / 2;
        const diff = _getActual(mid) - targetVal;
        if (Math.abs(diff) < Math.abs(bestDiff)) { bestDiff = diff; bestAdj = mid; }
        if (diff > 0) hi = mid; else lo = mid;
    }

    const opts = {};
    if (varType === 'vol') opts.volAdj = bestAdj;
    else if (varType === 'price') opts.priceAdj = bestAdj;
    else if (varType === 'vc') opts.vcAdj = bestAdj;
    else opts.fcAdj = bestAdj;
    const finalR = runEngine(opts);

    return { bestAdj, bestDiff, finalR, varType, targetType, targetVal };
}
