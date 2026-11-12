/**
 * API Client & Shared State (api.js)
 * Xử lý gọi REST API tới FastAPI backend và lưu bộ nhớ đệm cho các truy vấn GET.
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

export async function apiGet(path, params = {}, { useCache = true } = {}) {
  const qs = new URLSearchParams();
  Object.entries(params).forEach(([k, v]) => {
    if (v !== undefined && v !== null && v !== '') qs.set(k, String(v));
  });
  const url = qs.toString() ? `${path}?${qs.toString()}` : path;

  if (useCache && GET_CACHE.has(url)) {
    return GET_CACHE.get(url);
  }

  const res = await fetch(url, { headers: { Accept: 'application/json' } });
  if (!res.ok) {
    let msg = `Lỗi HTTP ${res.status}`;
    try {
      const err = await res.json();
      if (err.detail) msg = typeof err.detail === 'string' ? err.detail : JSON.stringify(err.detail);
    } catch (_) {}
    throw new Error(msg);
  }
  const data = await res.json();
  if (useCache) GET_CACHE.set(url, data);
  return data;
}

export async function apiPost(path, body = {}) {
  const res = await fetch(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    let msg = `Lỗi HTTP ${res.status}`;
    try {
      const err = await res.json();
      if (err.detail) msg = typeof err.detail === 'string' ? err.detail : JSON.stringify(err.detail);
    } catch (_) {}
    throw new Error(msg);
  }
  return res.json();
}

let _optionsPromise = null;
export function getOptions() {
  if (!_optionsPromise) {
    _optionsPromise = apiGet('/api/options');
  }
  return _optionsPromise;
}
