/* ╔══════════════════════════════════════════════════════════╗
   ║  CVP V4.0 — 全域應用狀態 & localStorage 持久化           ║
   ╚══════════════════════════════════════════════════════════╝ */

const STORAGE_KEY = 'cvp_v4_state';

/* ─── 默認資料集 ─── */
const DEFAULT_PRODUCTS = [
  { name:'旗艦型伺服器 (A)', price:50,   vc:28,   vol:50,   capMax:80,   minVol:10, cat:'硬體產品' },
  { name:'高階工作站 (B)',   price:15,   vc:9.5,  vol:200,  capMax:300,  minVol:20, cat:'硬體產品' },
  { name:'商務筆電 (C)',     price:4.5,  vc:3.2,  vol:1000, capMax:1500, minVol:100,cat:'硬體產品' },
  { name:'入門型筆電 (D)',   price:2,    vc:1.5,  vol:2500, capMax:4000, minVol:500,cat:'硬體產品' },
  { name:'軟體授權 (E)',     price:0.8,  vc:0.12, vol:3000, capMax:9999, minVol:0,  cat:'軟體服務' },
  { name:'維護合約 (F)',     price:1.5,  vc:0.3,  vol:800,  capMax:9999, minVol:0,  cat:'軟體服務' },
];

const DEFAULT_PARAMS = {
  fc: 2500,           // 固定成本（萬元）
  targetProfit: 1500, // 目標利潤（萬元）
  safetyWarn: 0.15,   // 安全邊際警示門檻
  bullVol: 0.20,      // 樂觀情境銷量增幅
  bearVol: -0.15,     // 悲觀情境銷量減幅
  compPrice: -0.10,   // 競爭衝擊價格降幅
  cmGoodThreshold: 0.40,
  industryAvgCM: 0.35,
  company: '範例科技股份有限公司',
  period: '2026 年度',
};

const DEFAULT_INVESTMENT = {
  capex: 12000, // 資本支出（萬元）
  years: 5,
  wacc: 0.08,
  depLife: 5,
  salvage: 0.10,
};

const DEFAULT_CAPITAL = {
  totalAssets:   50000,
  currentAssets: 15000,
  fixedAssets:   28000,
  otherAssets:   7000,
  currentLiab:   8000,
  longTermLiab:  12000,
  equity:        30000,
  indAvgTurnover: 1.20,
  indAvgROE:      0.15,
  indAvgROA:      0.08,
};

const DEFAULT_CASHFLOW_PARAMS = {
  arDays: 45,      // 應收帳款天數
  apDays: 30,      // 應付帳款天數
  safetyBuffer: 500, // 最低現金水位（萬元）
};

const DEFAULT_RISKS = [
  { name:'主要原材料漲價 15%',  prob:0.45, impact:0.65, cat:'成本風險',   mitigation:'多供應商備份' },
  { name:'核心客戶流失',        prob:0.25, impact:0.80, cat:'收入風險',   mitigation:'客戶關係強化' },
  { name:'新競爭者市場進入',    prob:0.60, impact:0.50, cat:'競爭風險',   mitigation:'差異化加速' },
  { name:'匯率大幅波動(±10%)', prob:0.40, impact:0.40, cat:'財務風險',   mitigation:'外匯對沖策略' },
  { name:'關鍵人才離職',        prob:0.30, impact:0.35, cat:'營運風險',   mitigation:'人才保留計畫' },
  { name:'法規政策重大改變',    prob:0.20, impact:0.55, cat:'法規風險',   mitigation:'法規監控機制' },
  { name:'資訊系統崩潰',        prob:0.15, impact:0.45, cat:'技術風險',   mitigation:'DR 備援系統' },
  { name:'銀行信用額度收縮',    prob:0.20, impact:0.60, cat:'流動性風險', mitigation:'備用信貸額度' },
];

const DEFAULT_SCENARIOS_WEIGHTS = [
  { name:'基準情境', vol:0,   price:0,   vc:0, fc:0, prob:0.50 },
  { name:'樂觀情境', vol:20,  price:0,   vc:0, fc:0, prob:0.25 },
  { name:'悲觀情境', vol:-15, price:0,   vc:0, fc:0, prob:0.20 },
  { name:'競爭衝擊', vol:0,   price:-10, vc:0, fc:0, prob:0.05 },
];

/* ─── 生成默認月度趨勢 ─── */
function _genDefaultMonthly() {
  const base = 550;
  const season = [0.80,0.85,0.90,0.95,1.00,1.05,1.10,1.15,1.08,1.12,1.20,1.25];
  const months = [];
  for (let m = 0; m < 12; m++) {
    const rev = Math.round(base * season[m] * (1 + Math.random() * 0.06 - 0.03));
    const vc  = Math.round(rev * 0.60 * (1 + Math.random() * 0.04 - 0.02));
    const fc  = Math.round(2500 / 12);
    const cm  = rev - vc;
    months.push({ month:`${m+1}月`, rev, vc, fc, cm, profit: cm - fc, cmRate: +(cm/rev).toFixed(4) });
  }
  return months;
}

/* ─── 生成默認預算數據 ─── */
function _genDefaultBudget() {
  return DEFAULT_PRODUCTS.map(p => ({
    name: p.name,
    budgetPrice: p.price * 1.02,      // 預算單價（較實際微高）
    budgetVc:    p.vc * 0.98,         // 預算變動成本（較實際微低）
    budgetVol:   Math.round(p.vol * 0.95), // 預算銷量
    cat: p.cat,
  }));
}

/* ─── 全域狀態物件 ─── */
const AppState = {
  /* 產品 */
  products: JSON.parse(JSON.stringify(DEFAULT_PRODUCTS)),
  /* 參數 */
  params: { ...DEFAULT_PARAMS },
  /* 投資 */
  investment: { ...DEFAULT_INVESTMENT },
  /* 資本效率 */
  capital: { ...DEFAULT_CAPITAL },
  /* 現金流參數 */
  cashflowParams: { ...DEFAULT_CASHFLOW_PARAMS },
  /* 月度趨勢（12 筆）*/
  monthlyTrend: _genDefaultMonthly(),
  /* 預算對照 */
  budgetData: _genDefaultBudget(),
  /* 多年度（可選）*/
  yoyData: [],  // [ { year:'2024', months:[{...}] }, ... ]
  /* 風險事件 */
  risks: JSON.parse(JSON.stringify(DEFAULT_RISKS)),
  /* 情境（含機率加權）*/
  scenarios: JSON.parse(JSON.stringify(DEFAULT_SCENARIOS_WEIGHTS)),
  userScenarios: [],   // 使用者自訂情境

  /* 系統狀態 */
  changeLog: [],
  lastUpdate: null,
  theme: 'dark',       // 'dark' | 'light'
};

/* ──────────────────── localStorage 持久化 ──────────────────── */
/**
 * 將 AppState 序列化儲存至 localStorage
 * 排除大型圖表快取資料，只存必要資料
 */
function saveState() {
  try {
    const toSave = {
      products:       AppState.products,
      params:         AppState.params,
      investment:     AppState.investment,
      capital:        AppState.capital,
      cashflowParams: AppState.cashflowParams,
      monthlyTrend:   AppState.monthlyTrend,
      budgetData:     AppState.budgetData,
      yoyData:        AppState.yoyData,
      risks:          AppState.risks,
      scenarios:      AppState.scenarios,
      userScenarios:  AppState.userScenarios,
      theme:          AppState.theme,
      lastUpdate:     AppState.lastUpdate,
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(toSave));
  } catch (e) {
    console.warn('[CVP V4] localStorage 儲存失敗：', e.message);
  }
}

/**
 * 從 localStorage 恢復 AppState
 * 若無儲存資料，使用預設值
 */
function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const saved = JSON.parse(raw);
    Object.assign(AppState, saved);
    // 確保 changeLog 不從 localStorage 還原（每次重新開始）
    AppState.changeLog = [];
    return true;
  } catch (e) {
    console.warn('[CVP V4] localStorage 讀取失敗，使用預設值：', e.message);
    return false;
  }
}

/**
 * 清除 localStorage，還原預設資料
 */
function resetState() {
  localStorage.removeItem(STORAGE_KEY);
  AppState.products       = JSON.parse(JSON.stringify(DEFAULT_PRODUCTS));
  AppState.params         = { ...DEFAULT_PARAMS };
  AppState.investment     = { ...DEFAULT_INVESTMENT };
  AppState.capital        = { ...DEFAULT_CAPITAL };
  AppState.cashflowParams = { ...DEFAULT_CASHFLOW_PARAMS };
  AppState.monthlyTrend   = _genDefaultMonthly();
  AppState.budgetData     = _genDefaultBudget();
  AppState.yoyData        = [];
  AppState.risks          = JSON.parse(JSON.stringify(DEFAULT_RISKS));
  AppState.scenarios      = JSON.parse(JSON.stringify(DEFAULT_SCENARIOS_WEIGHTS));
  AppState.userScenarios  = [];
  AppState.changeLog      = [];
  AppState.lastUpdate     = null;
}

/**
 * 匯出完整狀態為 JSON
 */
function exportStateJSON() {
  const data = JSON.stringify({ state: AppState, version: 'V4.0', generated: new Date().toISOString() }, null, 2);
  const blob = new Blob([data], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `CVP_V4_狀態快照_${fmtDate(new Date())}.json`;
  a.click();
}

/**
 * 從 JSON 檔案載入狀態
 */
function importStateJSON(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const parsed = JSON.parse(ev.target.result);
        const src = parsed.state || parsed;
        // 謹慎合併，只更新有效欄位
        const validKeys = ['products','params','investment','capital','cashflowParams',
          'monthlyTrend','budgetData','yoyData','risks','scenarios','userScenarios'];
        validKeys.forEach(k => { if (src[k] !== undefined) AppState[k] = src[k]; });
        saveState();
        resolve();
      } catch (e) {
        reject(new Error('JSON 格式解析失敗：' + e.message));
      }
    };
    reader.readAsText(file);
  });
}
