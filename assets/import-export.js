/* ╔══════════════════════════════════════════════════════════╗
   ║  CVP V4.0 — Excel 匯入 / 匯出 / PDF                      ║
   ╚══════════════════════════════════════════════════════════╝ */

/* ══════════════════════════════════════
   Excel 匯入
   ══════════════════════════════════════ */

/** 處理 File Input 或 Drop 的統一入口 */
function readExcel(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
        try {
            const wb = XLSX.read(ev.target.result, { type: 'array' });
            const result = parseExcelData(wb);
            AppState.lastUpdate = new Date().toISOString();
            saveState();
            fullRender();
            updateHeader();
            const newTabs = result.newSheets.length > 0 ? `（新增：${result.newSheets.join(', ')}）` : '';
            showToast(`✅ Excel 匯入成功！${newTabs}`, 'ok');
            logChange(`匯入 Excel：${file.name}`);
        } catch (err) {
            showToast(`❌ 解析失敗：${err.message}`, 'err');
            console.error('[CVP V4] Excel 解析錯誤：', err);
        }
    };
    reader.onerror = () => showToast('❌ 檔案讀取失敗', 'err');
    reader.readAsArrayBuffer(file);
}

function handleExcelUpload(e) { readExcel(e.target.files[0]); }
function handleDrop(e) {
    e.preventDefault();
    document.getElementById('drop-zone')?.classList.remove('drag');
    readExcel(e.dataTransfer.files[0]);
}

/**
 * 解析 Excel Workbook 並更新 AppState
 * @param {object} wb - XLSX Workbook
 * @returns {object} { newSheets: string[] }（記錄本次新解析成功的工作表）
 */
function parseExcelData(wb) {
    const sheets = wb.SheetNames;
    const newSheets = [];

    /* ── Product_Mix ── */
    if (sheets.includes('Product_Mix')) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Product_Mix']);
        if (rows.length) {
            AppState.products = rows.map(r => ({
                name: r['產品名稱'] || '未命名',
                price: toNum(r['單價（元）']) / 10000,
                vc: toNum(r['單位變動成本（元）']) / 10000,
                vol: toNum(r['預估銷量（件）']),
                capMax: toNum(r['產能上限（件）']) || 9999,
                minVol: toNum(r['最低銷量（件）']) || 0,
                cat: r['產品類別'] || '其他',
            }));
            newSheets.push('Product_Mix');
        }
    }

    /* ── Parameters ── */
    if (sheets.includes('Parameters')) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Parameters']);
        rows.forEach(r => {
            const name = r['參數名稱'] || '', val = toNum(r['數值']);
            if (name.includes('固定成本')) AppState.params.fc = val / 10000;
            if (name.includes('目標利潤')) AppState.params.targetProfit = val / 10000;
            if (name.includes('安全邊際')) AppState.params.safetyWarn = val;
            if (name.includes('樂觀')) AppState.params.bullVol = val;
            if (name.includes('悲觀')) AppState.params.bearVol = val;
            if (name.includes('競爭')) AppState.params.compPrice = val;
            if (name.includes('行業平均邊際')) AppState.params.industryAvgCM = val;
            if (name.includes('公司名稱')) AppState.params.company = r['數值'] || AppState.params.company;
            if (name.includes('期間')) AppState.params.period = r['數值'] || AppState.params.period;
        });
        newSheets.push('Parameters');
    }

    /* ── Investment ── */
    if (sheets.includes('Investment')) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Investment']);
        rows.forEach(r => {
            const name = r['參數名稱'] || '', val = toNum(r['數值']);
            if (name.includes('資本支出') || name.includes('Capex')) AppState.investment.capex = val / 10000;
            if (name.includes('年期')) AppState.investment.years = val;
            if (name.includes('WACC') || name.includes('折現率')) AppState.investment.wacc = val;
            if (name.includes('折舊年限')) AppState.investment.depLife = val;
            if (name.includes('殘值')) AppState.investment.salvage = val;
        });
        newSheets.push('Investment');
    }

    /* ── Capital_Efficiency ── */
    if (sheets.includes('Capital_Efficiency')) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Capital_Efficiency']);
        rows.forEach(r => {
            const name = r['參數名稱'] || '', val = toNum(r['數值']);
            if (name.includes('流動資產')) AppState.capital.currentAssets = val / 10000;
            if (name.includes('固定資產')) AppState.capital.fixedAssets = val / 10000;
            if (name.includes('其他資產')) AppState.capital.otherAssets = val / 10000;
            if (name.includes('流動負債')) AppState.capital.currentLiab = val / 10000;
            if (name.includes('長期負債')) AppState.capital.longTermLiab = val / 10000;
            if (name.includes('股東權益')) AppState.capital.equity = val / 10000;
            if (name.includes('行業平均資')) AppState.capital.indAvgTurnover = val;
            if (name.includes('行業平均 ROE')) AppState.capital.indAvgROE = val;
            if (name.includes('行業平均 ROA')) AppState.capital.indAvgROA = val;
        });
        newSheets.push('Capital_Efficiency');
    }

    /* ── Monthly_Trend ── */
    if (sheets.includes('Monthly_Trend')) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Monthly_Trend']);
        if (rows.length >= 12) {
            AppState.monthlyTrend = rows.slice(0, 12).map(r => ({
                month: r['月份'] || '',
                rev: toNum(r['總收入（元）']) / 10000,
                vc: toNum(r['總變動成本（元）']) / 10000,
                fc: toNum(r['固定成本（元）']) / 10000,
                cm: toNum(r['邊際貢獻（元）']) / 10000,
                profit: toNum(r['營業利潤（元）']) / 10000,
                cmRate: toNum(r['邊際貢獻率']),
            }));
            newSheets.push('Monthly_Trend');
        }
    }

    /* ── Risk_Parameters ── */
    if (sheets.includes('Risk_Parameters')) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Risk_Parameters']);
        if (rows.length) {
            AppState.risks = rows.map(r => ({
                name: r['風險事件'] || '',
                prob: toNum(r['發生機率（0-1）']),
                impact: toNum(r['財務衝擊度（0-1）']),
                cat: r['分類'] || '其他',
                mitigation: r['應對措施'] || '',
            }));
            newSheets.push('Risk_Parameters');
        }
    }

    /* ── Budget_Data（V4 新增）── */
    if (sheets.includes('Budget_Data')) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Budget_Data']);
        if (rows.length) {
            AppState.budgetData = rows.map(r => ({
                name: r['產品名稱'] || '',
                budgetPrice: toNum(r['預算單價（元）']) / 10000,
                budgetVc: toNum(r['預算單位變動成本（元）']) / 10000,
                budgetVol: toNum(r['預算銷量（件）']),
                cat: r['產品類別'] || '其他',
            }));
            newSheets.push('Budget_Data');
        }
    }

    /* ── CashFlow_Params（V4 新增）── */
    if (sheets.includes('CashFlow_Params')) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['CashFlow_Params']);
        rows.forEach(r => {
            const name = r['參數名稱'] || '', val = toNum(r['數值']);
            if (name.includes('應收帳款天數')) AppState.cashflowParams.arDays = val;
            if (name.includes('應付帳款天數')) AppState.cashflowParams.apDays = val;
            if (name.includes('安全現金水位')) AppState.cashflowParams.safetyBuffer = val / 10000;
        });
        newSheets.push('CashFlow_Params');
    }

    /* ── YoY_Comparison（V4 新增）── */
    if (sheets.includes('YoY_Comparison')) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['YoY_Comparison']);
        if (rows.length) {
            const yearMap = {};
            rows.forEach(r => {
                const year = String(r['年度'] || '');
                if (!yearMap[year]) yearMap[year] = [];
                yearMap[year].push({
                    month: r['月份'] || '',
                    rev: toNum(r['總收入（元）']) / 10000,
                    cm: toNum(r['邊際貢獻（元）']) / 10000,
                    profit: toNum(r['營業利潤（元）']) / 10000,
                });
            });
            AppState.yoyData = Object.entries(yearMap).map(([year, months]) => ({ year, months }));
            newSheets.push('YoY_Comparison');
        }
    }

    /* ── Scenario_Weights（V4 新增）── */
    if (sheets.includes('Scenario_Weights')) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Scenario_Weights']);
        if (rows.length) {
            AppState.scenarios = rows.map(r => ({
                name: r['情境名稱'] || '',
                vol: toNum(r['銷量變動%']),
                price: toNum(r['單價變動%']),
                vc: toNum(r['變動成本變動%']) || 0,
                fc: toNum(r['固定成本變動%']) || 0,
                prob: toNum(r['發生機率（0-1）']),
            }));
            newSheets.push('Scenario_Weights');
        }
    }

    return { newSheets };
}

/* ══════════════════════════════════════
   Excel 匯出
   ══════════════════════════════════════ */

/**
 * 匯出 CVP 分析快照（Excel）
 */
function exportExcelSnapshot() {
    if (typeof XLSX === 'undefined') { showToast('❌ XLSX 套件未就緒', 'err'); return; }
    const wb2 = XLSX.utils.book_new();
    const base = runEngine();

    /* Sheet 1: CVP 摘要 */
    const summary = [
        ['項目', '數值', '單位'],
        ['總收入', +fmt2(base.totalRev), '萬元'],
        ['總變動成本', +fmt2(base.totalVC), '萬元'],
        ['邊際貢獻', +fmt2(base.totalCM), '萬元'],
        ['邊際貢獻率', +(base.cmRate * 100).toFixed(2), '%'],
        ['固定成本', +fmt2(base.fc), '萬元'],
        ['營業利潤', +fmt2(base.profit), '萬元'],
        ['損益平衡點', +fmt2(base.bepRev), '萬元'],
        ['安全邊際率', +(base.safetyMarginRate * 100).toFixed(2), '%'],
        ['經營槓桿度', isFinite(base.dol) ? +base.dol.toFixed(2) : '∞', '倍'],
    ];
    XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet(summary), 'CVP摘要');

    /* Sheet 2: 產品明細 */
    const prodHeader = ['產品名稱', '售價(萬)', '變動成本(萬)', 'CM/件(萬)', 'CM率(%)', '銷量', 'CM合計(萬)', '類別'];
    const prodRows = base.prodCalc.map(p => [
        p.name, +fmt2(p.price), +fmt2(p.vc), +fmt2(p.cmPerUnit),
        +(p.cmRate * 100).toFixed(2), +fmt2(p.vol), +fmt2(p.cm), p.cat,
    ]);
    XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet([prodHeader, ...prodRows]), '產品明細');

    /* Sheet 3: 月度趨勢 */
    const trendHeader = ['月份', '收入(萬)', '邊際貢獻(萬)', 'CM率(%)', '利潤(萬)'];
    const trendRows = AppState.monthlyTrend.map(m => [
        m.month, +fmt2(m.rev), +fmt2(m.cm), +(m.cmRate * 100).toFixed(2), +fmt2(m.profit),
    ]);
    XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet([trendHeader, ...trendRows]), '月度趨勢');

    /* Sheet 4: 投資分析 */
    const inv = runInvestment();
    const invData = [
        ['指標', '數值', '說明'],
        ['NPV', +fmt2(inv.npv), '萬元，>0 表示可行'],
        ['IRR', +(inv.irr * 100).toFixed(2), '%'],
        ['PBP', +inv.pbp.toFixed(2), '年'],
        ['ROI', +(inv.roi * 100).toFixed(2), '%，基於 NPV/投資額'],
        ['WACC', +(AppState.investment.wacc * 100).toFixed(1), '%'],
    ];
    XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet(invData), '投資分析');

    XLSX.writeFile(wb2, `CVP_V4_分析快照_${fmtDate(new Date())}.xlsx`);
    showToast('✅ Excel 快照已下載', 'ok');
    logChange('匯出 Excel 快照');
}
