/**
 * Hash Router & App Shell Controller (router.js)
 * Điều hướng giữa 6 trang, đồng bộ bộ lọc Store / SKU ở Topbar, và quản lý Dark / Light mode.
 */

import { appState, updateState, getOptions } from './api.js';
import { refreshIcons } from './ui.js';
import { destroyAllCharts } from './charts.js';
import { renderOverviewPage } from './pages/overview.js';
import { renderForecastPage } from './pages/forecast.js';
import { renderModelsPage } from './pages/models.js';
import { renderInventoryPage } from './pages/inventory.js';
import { renderAbcXyzPage } from './pages/abcxyz.js';
import { renderWhatIfPage } from './pages/whatif.js';

const ROUTES = {
  overview: { title: 'Tổng quan chuỗi & Khám phá dữ liệu (EDA)', render: renderOverviewPage },
  forecast: { title: 'Dự báo nhu cầu sản phẩm theo SKU', render: renderForecastPage },
  models: { title: 'So sánh 5 mô hình dự báo & Feature Importance', render: renderModelsPage },
  inventory: { title: 'Tồn kho an toàn (SS) & Cảnh báo điểm đặt hàng lại (ROP)', render: renderInventoryPage },
  abcxyz: { title: 'Ma trận phân loại tồn kho chiến lược ABC – XYZ', render: renderAbcXyzPage },
  whatif: { title: 'Mô phỏng kịch bản What-If chuỗi cung ứng', render: renderWhatIfPage },
};

function parseRoute() {
  const raw = window.location.hash.replace(/^#\/?/, '') || 'overview';
  const [pagePart, queryPart] = raw.split('?');
  const page = ROUTES[pagePart] ? pagePart : 'overview';
  if (queryPart) {
    const params = new URLSearchParams(queryPart);
    const patch = {};
    if (params.get('store')) patch.store_id = params.get('store');
    if (params.get('product')) patch.product_id = params.get('product');
    if (params.get('model')) patch.model_name = params.get('model');
    if (params.get('h')) patch.horizon = Number(params.get('h'));
    if (Object.keys(patch).length) updateState(patch);
  }
  return page;
}

async function navigate() {
  const page = parseRoute();
  const route = ROUTES[page];
  const view = document.getElementById('view');
  const heading = document.getElementById('page-heading');

  if (heading) heading.textContent = route.title;

  document.querySelectorAll('.nav-link').forEach((link) => {
    link.classList.toggle('active', link.dataset.page === page);
  });

  // Đóng sidebar trên mobile khi chuyển trang
  document.getElementById('sidebar')?.classList.remove('open');

  destroyAllCharts();
  view.innerHTML = `<div class="skeleton"></div>`;
  await route.render(view);
}

async function initGlobalControls() {
  refreshIcons();

  // Khởi tạo Dark / Light mode
  const savedTheme = localStorage.getItem('retail_ml_theme') || 'light';
  setTheme(savedTheme);

  document.getElementById('theme-toggle-btn')?.addEventListener('click', () => {
    const cur = document.documentElement.getAttribute('data-theme') || 'light';
    const next = cur === 'dark' ? 'light' : 'dark';
    setTheme(next);
    navigate();
  });

  // Nút mở menu trên mobile
  document.getElementById('mobile-menu-btn')?.addEventListener('click', () => {
    document.getElementById('sidebar')?.classList.toggle('open');
  });

  // Điền danh sách Store / SKU trên thanh Topbar
  try {
    const opts = await getOptions();
    const ctxStore = document.getElementById('ctx-store');
    const ctxProd = document.getElementById('ctx-product');

    if (ctxStore && ctxProd) {
      ctxStore.innerHTML = opts.stores
        .map((s) => `<option value="${s}" ${s === appState.store_id ? 'selected' : ''}>${s}</option>`)
        .join('');
      ctxProd.innerHTML = opts.products
        .map((p) => `<option value="${p.id}" ${p.id === appState.product_id ? 'selected' : ''}>${p.id}</option>`)
        .join('');

      ctxStore.addEventListener('change', () => {
        updateState({ store_id: ctxStore.value });
        navigate();
      });

      ctxProd.addEventListener('change', () => {
        updateState({ product_id: ctxProd.value });
        navigate();
      });

      window.addEventListener('appstate:change', (e) => {
        ctxStore.value = e.detail.store_id;
        ctxProd.value = e.detail.product_id;
      });
    }
  } catch (_) {}
}

function setTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('retail_ml_theme', theme);
  const icon = document.getElementById('theme-icon');
  if (icon) {
    icon.setAttribute('data-lucide', theme === 'dark' ? 'sun' : 'moon');
    refreshIcons();
  }
}

window.addEventListener('hashchange', navigate);
window.addEventListener('DOMContentLoaded', async () => {
  await initGlobalControls();
  await navigate();
});
