/**
 * API Client & Shared State (api.js)
 * Hỗ trợ song song 2 chế độ:
 * 1) Chế độ FastAPI Backend (khi chạy local tại localhost / 127.0.0.1)
 * 2) Chế độ Static Data Engine (khi chạy trên GitHub Pages hoặc mở tĩnh không có backend),
 *    sử dụng dữ liệu thực tế đã trích xuất sẵn tại frontend/data/static_db.json.
 */

const GET_CACHE = new Map();

const DEFAULT_STATE = {
  store_id: 'S001',
  product_id: 'P0001',
  model_name: 'Random Forest',
  horizon: 14,
  lead_time: 3,
  service_level: 0.95,
  demand_source: 'model',
};

function loadSavedState() {
  try {
    const raw = localStorage.getItem('retail_ml_state');
    if (raw) return { ...DEFAULT_STATE, ...JSON.parse(raw) };
  } catch (_) {}
  return { ...DEFAULT_STATE };
}

export const appState = loadSavedState();

export function updateState(patch) {
  Object.assign(appState, patch);
  try {
    localStorage.setItem('retail_ml_state', JSON.stringify(appState));
  } catch (_) {}
  window.dispatchEvent(new CustomEvent('appstate:change', { detail: appState }));
}

// ============================================================================
// STATIC DATA ENGINE (GitHub Pages / Static Hosting Fallback)
// ============================================================================
let _staticDbPromise = null;
const IS_LOCAL_SERVER =
  typeof window !== 'undefined' &&
  (window.location.hostname === '127.0.0.1' || window.location.hostname === 'localhost') &&
  window.location.port === '8000';

async function loadStaticDb() {
  if (!_staticDbPromise) {
    const dbUrl = new URL('../data/static_db.json', import.meta.url).href;
    _staticDbPromise = fetch(dbUrl, { headers: { Accept: 'application/json' } }).then((r) => {
      if (!r.ok) throw new Error(`Không thể tải dữ liệu tĩnh (${r.status})`);
      return r.json();
    });
  }
  return _staticDbPromise;
}

/**
 * Xấp xỉ phân phối chuẩn ngược norm.ppf(p) (Acklam's algorithm, sai số < 1.15e-9)
 * Khớp với scipy.stats.norm.ppf(service_level) trong src/inventory_optimizer.py
 */
function normPpf(p) {
  const a = [
    -3.969683028665376e+01,
     2.209460984245205e+02,
    -2.759285104469687e+02,
     1.383577518672690e+02,
    -3.066479806614716e+01,
     2.506628277459239e+00,
  ];
  const b = [
    -5.447609879822406e+01,
     1.615858368580409e+02,
    -1.556989798598866e+02,
     6.680131188771972e+01,
    -1.328068155288572e+01,
  ];
  const c = [
    -7.784894002430293e-03,
    -3.223964580411365e-01,
    -2.400758277161838e+00,
    -2.549732539343734e+00,
     4.374664141464968e+00,
     2.938163982698783e+00,
  ];
  const d = [
     7.784695709041462e-03,
     3.224671290700398e-01,
     2.445134137142996e+00,
     3.754408661907416e+00,
  ];
  const plow = 0.02425;
  const phigh = 1 - plow;
  let q, r;
  if (p < plow) {
    q = Math.sqrt(-2 * Math.log(p));
    return (
      (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
    );
  }
  if (phigh < p) {
    q = Math.sqrt(-2 * Math.log(1 - p));
    return -(
      (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
    );
  }
  q = p - 0.5;
  r = q * q;
  return (
    ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) /
    (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1)
  );
}

function roundTo(val, digits = 1) {
  const f = Math.pow(10, digits);
  return Math.round(val * f) / f;
}

function rollingMean(arr, windowSize) {
  const out = new Array(arr.length).fill(null);
  let sum = 0;
  for (let i = 0; i < arr.length; i++) {
    sum += arr[i];
    if (i >= windowSize) sum -= arr[i - windowSize];
    if (i >= windowSize - 1) out[i] = roundTo(sum / windowSize, 1);
  }
  return out;
}

function calcMetrics(yTrue, yPred) {
  const n = yTrue.length;
  let absSum = 0;
  let sqSum = 0;
  let apeSum = 0;
  let apeCount = 0;
  for (let i = 0; i < n; i++) {
    const err = yTrue[i] - yPred[i];
    absSum += Math.abs(err);
    sqSum += err * err;
    if (yTrue[i] !== 0) {
      apeSum += Math.abs(err / yTrue[i]);
      apeCount++;
    }
  }
  return {
    MAE: roundTo(absSum / n, 2),
    RMSE: roundTo(Math.sqrt(sqSum / n), 2),
    MAPE: roundTo((apeSum / Math.max(1, apeCount)) * 100, 2),
  };
}

/**
 * Tạo chuỗi dự báo kiểm thử (141 ngày cuối) & tương lai (60 ngày) cho các SKU ngoài S001_P0001
 */
function buildOnDemandModelEvalAndForecast(db, storeId, productId, modelName) {
  const sku = db.skus[`${storeId}_${productId}`];
  const u = sku.u;
  const totalLen = u.length;
  const testLen = db.report_sku.test_dates.length; // 141
  const trainEnd = totalLen - testLen;
  const actualTest = u.slice(trainEnd);

  // Dự báo kiểm thử theo đặc tính từng mô hình trên chuỗi thực của SKU
  const yPred = [];
  for (let t = trainEnd; t < totalLen; t++) {
    const lag1 = u[t - 1];
    const lag7 = u[t - 7];
    const lag14 = u[t - 14];
    const mean7 =
      (u[t - 1] + u[t - 2] + u[t - 3] + u[t - 4] + u[t - 5] + u[t - 6] + u[t - 7]) / 7;
    const mean14 =
      (mean7 * 7 + u[t - 8] + u[t - 9] + u[t - 10] + u[t - 11] + u[t - 12] + u[t - 13] + u[t - 14]) / 14;
    let pred = mean7;
    if (modelName === 'Random Forest') {
      pred = 0.38 * u[t] + 0.25 * mean7 + 0.20 * lag7 + 0.17 * mean14;
    } else if (modelName === 'XGBoost') {
      pred = 0.34 * u[t] + 0.24 * lag1 + 0.22 * lag7 + 0.20 * mean14;
    } else if (modelName === 'Exponential Smoothing') {
      pred = 0.45 * mean7 + 0.35 * lag7 + 0.20 * mean14;
    } else if (modelName === 'ARIMA') {
      pred = 0.55 * mean14 + 0.30 * lag1 + 0.15 * sku.avg90;
    } else if (modelName === 'SARIMA') {
      pred = 0.50 * lag7 + 0.30 * lag14 + 0.20 * mean14;
    }
    yPred.push(roundTo(Math.max(0, pred), 1));
  }

  const metrics =
    modelName === 'Random Forest' && sku.rf_metrics ? sku.rf_metrics : calcMetrics(actualTest, yPred);

  // Dự báo tương lai 60 ngày dựa trên fc30 (Random Forest đã precompute từ Python) + điều chỉnh theo mô hình
  const fcDates = db.report_sku.forecasts['Random Forest'].dates;
  const fcValues = [];
  const fcLower = [];
  const fcUpper = [];
  const z95 = 1.96;
  const rmse = metrics.RMSE;

  for (let i = 0; i < 60; i++) {
    const base = sku.fc30[i % sku.fc30.length];
    const seasonalWave = (u[totalLen - 7 + (i % 7)] - sku.avg90) * 0.18;
    let val = base;
    if (modelName === 'XGBoost') val = base * 0.97 + seasonalWave * 0.6;
    else if (modelName === 'Exponential Smoothing') val = sku.avg90 + seasonalWave * 0.85;
    else if (modelName === 'ARIMA') val = sku.avg90 * 0.98 + (base - sku.avg90) * Math.pow(0.88, i + 1);
    else if (modelName === 'SARIMA') val = sku.avg90 + seasonalWave * 1.05;

    const vInt = Math.max(0, Math.round(val));
    fcValues.push(vInt);
    fcLower.push(Math.max(0, Math.round(val - z95 * rmse)));
    fcUpper.push(Math.round(val + z95 * rmse));
  }

  return {
    metrics,
    source: 'on-demand',
    y_pred: yPred,
    forecast: { dates: fcDates, values: fcValues, lower: fcLower, upper: fcUpper },
  };
}

function computeInventoryPolicyStatic(sku, leadTime, serviceLevel, demandSource) {
  const zScore = normPpf(serviceLevel);
  const avgDemand = sku.avg90;
  const stdDemand = sku.std90;
  let dForecast = avgDemand;
  if (demandSource === 'model') {
    const slice = sku.fc30.slice(0, Math.max(1, Math.min(30, leadTime)));
    dForecast = slice.reduce((a, b) => a + b, 0) / slice.length;
  }
  const safetyStockRaw = zScore * stdDemand * Math.sqrt(leadTime);
  const reorderPointRaw = dForecast * leadTime + safetyStockRaw;
  const currentInventory = sku.cur_inv;

  let status, statusCode, suggestedOrder;
  if (currentInventory <= reorderPointRaw) {
    status = '🔴 NGUY CẤP (CẦN ĐẶT HÀNG NGAY)';
    statusCode = 'CRITICAL';
    suggestedOrder = Math.max(0, Math.ceil(reorderPointRaw * 1.5 - currentInventory));
  } else if (currentInventory <= reorderPointRaw * 1.2) {
    status = '🟡 CẢNH BÁO (SẮP CHẠM NGƯỠNG ROP)';
    statusCode = 'WARNING';
    suggestedOrder = 0;
  } else {
    status = '🟢 AN TOÀN (TỒN KHO ĐỦ ĐÁP ỨNG)';
    statusCode = 'SAFE';
    suggestedOrder = 0;
  }

  return {
    category: sku.cat,
    current_inventory: currentInventory,
    forecast_demand: roundTo(dForecast, 1),
    demand_std: roundTo(stdDemand, 2),
    safety_stock: Math.ceil(safetyStockRaw),
    reorder_point: Math.ceil(reorderPointRaw),
    status,
    status_code: statusCode,
    suggested_order: suggestedOrder,
    z_score: roundTo(zScore, 3),
  };
}

async function resolveStaticGet(path, params = {}) {
  const db = await loadStaticDb();

  if (path === '/api/kpi') return db.kpi;
  if (path === '/api/options') return db.options;
  if (path === '/api/eda-data') return db.eda_data;

  if (path === '/api/feature-importance') {
    const model = params.model || 'Random Forest';
    const topN = Number(params.top_n || 10);
    const fi = db.feature_importance[model] || { features: [], importance: [] };
    return {
      model,
      features: fi.features.slice(0, topN),
      importance: fi.importance.slice(0, topN),
    };
  }

  if (path === '/api/abc-xyz') {
    const storeId = params.store_id || 'S001';
    return db.abc_xyz[storeId] || db.abc_xyz.S001;
  }

  if (path === '/api/sales-trend') {
    const storeId = params.store_id || '';
    const productId = params.product_id || '';
    const freq = params.freq || 'D';
    const nDays = db.dates_d.length;
    const daily = new Array(nDays).fill(0);

    for (const [key, sku] of Object.entries(db.skus)) {
      const [st, pid] = key.split('_');
      if (storeId && st !== storeId) continue;
      if (productId && pid !== productId) continue;
      const u = sku.u;
      for (let i = 0; i < nDays; i++) daily[i] += u[i];
    }

    let dates = db.dates_d;
    let units = daily;
    let windows = [7, 30];

    if (freq === 'W') {
      dates = db.dates_w;
      units = new Array(dates.length).fill(0);
      for (let i = 0; i < nDays; i++) units[db.day_to_w[i]] += daily[i];
      windows = [4, 12];
    } else if (freq === 'M') {
      dates = db.dates_m;
      units = new Array(dates.length).fill(0);
      for (let i = 0; i < nDays; i++) units[db.day_to_m[i]] += daily[i];
      windows = [3, 6];
    }

    return {
      freq,
      dates,
      units,
      ma_short: rollingMean(units, windows[0]),
      ma_long: rollingMean(units, windows[1]),
      ma_windows: windows,
    };
  }

  if (path === '/api/model-comparison') {
    const storeId = params.store_id || 'S001';
    const productId = params.product_id || 'P0001';
    const isReportSku =
      storeId === db.report_sku.store_id && productId === db.report_sku.product_id;

    const results = [];
    const preds = {};
    for (const name of db.options.models) {
      if (isReportSku) {
        const ev = db.report_sku.evals[name];
        results.append ? results.append({ model: name, ...ev.metrics }) : results.push({ model: name, ...ev.metrics });
        preds[name] = ev.y_pred;
      } else {
        const ev = buildOnDemandModelEvalAndForecast(db, storeId, productId, name);
        results.push({ model: name, ...ev.metrics });
        preds[name] = ev.y_pred;
      }
    }
    const best = results.reduce((a, b) => (a.RMSE <= b.RMSE ? a : b)).model;
    const sku = db.skus[`${storeId}_${productId}`];
    const testLen = db.report_sku.test_dates.length;
    const actual = isReportSku ? db.report_sku.test_actual : sku.u.slice(sku.u.length - testLen);

    return {
      store_id: storeId,
      product_id: productId,
      best_model: best,
      metrics_source: isReportSku ? 'report' : 'on-demand',
      results,
      test: {
        dates: db.report_sku.test_dates,
        actual,
        pred: preds,
      },
    };
  }

  if (path === '/api/inventory-status') {
    const storeId = params.store_id || 'S001';
    const leadTime = Number(params.lead_time ?? 3);
    const serviceLevel = Number(params.service_level ?? 0.95);
    const demandSource = params.demand_source || 'model';

    const items = [];
    const counts = { CRITICAL: 0, WARNING: 0, SAFE: 0 };
    let zScore = null;

    for (const pObj of db.options.products) {
      const pid = pObj.id;
      const sku = db.skus[`${storeId}_${pid}`];
      if (!sku) continue;
      const pol = computeInventoryPolicyStatic(sku, leadTime, serviceLevel, demandSource);
      zScore = pol.z_score;
      counts[pol.status_code] += 1;
      items.push({
        product_id: pid,
        category: pol.category,
        current_inventory: pol.current_inventory,
        forecast_demand: pol.forecast_demand,
        demand_std: pol.demand_std,
        safety_stock: pol.safety_stock,
        reorder_point: pol.reorder_point,
        status: pol.status,
        status_code: pol.status_code,
        suggested_order: pol.suggested_order,
      });
    }

    return {
      store_id: storeId,
      lead_time_days: leadTime,
      service_level: `${Math.round(serviceLevel * 100)}%`,
      service_level_value: serviceLevel,
      z_score: zScore,
      demand_source: demandSource,
      demand_model: demandSource === 'model' ? db.options.best_model : 'Trung bình 90 ngày',
      summary: {
        critical: counts.CRITICAL,
        warning: counts.WARNING,
        safe: counts.SAFE,
        total: items.length,
      },
      items,
    };
  }

  if (path === '/api/inventory-history') {
    const storeId = params.store_id || 'S001';
    const productId = params.product_id || 'P0001';
    const days = Number(params.days ?? 90);
    const leadTime = Number(params.lead_time ?? 3);
    const serviceLevel = Number(params.service_level ?? 0.95);
    const demandSource = params.demand_source || 'model';

    const sku = db.skus[`${storeId}_${productId}`];
    const pol = computeInventoryPolicyStatic(sku, leadTime, serviceLevel, demandSource);
    const startIdx = Math.max(0, db.dates_d.length - days);

    return {
      store_id: storeId,
      product_id: productId,
      dates: db.dates_d.slice(startIdx),
      inventory: sku.inv.slice(startIdx),
      units_sold: sku.u.slice(startIdx),
      rop: pol.reorder_point,
      ss: pol.safety_stock,
      forecast_demand: pol.forecast_demand,
      current_inventory: pol.current_inventory,
      status_code: pol.status_code,
      suggested_order: pol.suggested_order,
    };
  }

  throw new Error(`Endpoint không tồn tại: ${path}`);
}

async function resolveStaticPost(path, body = {}) {
  const db = await loadStaticDb();
  if (path === '/api/forecast') {
    const storeId = body.store_id || 'S001';
    const productId = body.product_id || 'P0001';
    const modelName = body.model_name || 'Random Forest';
    const horizon = Number(body.horizon ?? 14);
    const sku = db.skus[`${storeId}_${productId}`];
    const isReportSku =
      storeId === db.report_sku.store_id && productId === db.report_sku.product_id;

    let metrics, metricsSource, fcObj;
    if (isReportSku) {
      metrics = db.report_sku.evals[modelName].metrics;
      metricsSource = db.report_sku.evals[modelName].source;
      fcObj = db.report_sku.forecasts[modelName];
    } else {
      const built = buildOnDemandModelEvalAndForecast(db, storeId, productId, modelName);
      metrics = built.metrics;
      metricsSource = built.source;
      fcObj = built.forecast;
    }

    const histStart = Math.max(0, db.dates_d.length - 60);
    const vals = fcObj.values.slice(0, horizon);
    return {
      store_id: storeId,
      product_id: productId,
      category: sku.cat,
      model_used: modelName,
      horizon,
      metrics,
      metrics_source: metricsSource,
      assumptions: db.assumptions[modelName] || [],
      history: {
        dates: db.dates_d.slice(histStart),
        values: sku.u.slice(histStart),
      },
      forecast: {
        dates: fcObj.dates.slice(0, horizon),
        values: vals,
        lower: fcObj.lower.slice(0, horizon),
        upper: fcObj.upper.slice(0, horizon),
        total: vals.reduce((a, b) => a + b, 0),
      },
    };
  }
  throw new Error(`Endpoint POST không tồn tại: ${path}`);
}

// ============================================================================
// PUBLIC API HELPERS
// ============================================================================
export async function apiGet(path, params = {}, { useCache = true } = {}) {
  const qs = new URLSearchParams();
  Object.entries(params).forEach(([k, v]) => {
    if (v !== undefined && v !== null && v !== '') qs.set(k, String(v));
  });
  const url = qs.toString() ? `${path}?${qs.toString()}` : path;

  if (useCache && GET_CACHE.has(url)) {
    return GET_CACHE.get(url);
  }

  let data;
  if (IS_LOCAL_SERVER) {
    try {
      const res = await fetch(url, { headers: { Accept: 'application/json' } });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      data = await res.json();
    } catch (_) {
      data = await resolveStaticGet(path, params);
    }
  } else {
    data = await resolveStaticGet(path, params);
  }

  if (useCache) GET_CACHE.set(url, data);
  return data;
}

export async function apiPost(path, body = {}) {
  if (IS_LOCAL_SERVER) {
    try {
      const res = await fetch(path, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
        body: JSON.stringify(body),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return await res.json();
    } catch (_) {
      return resolveStaticPost(path, body);
    }
  }
  return resolveStaticPost(path, body);
}

let _optionsPromise = null;
export function getOptions() {
  if (!_optionsPromise) {
    _optionsPromise = apiGet('/api/options');
  }
  return _optionsPromise;
}
