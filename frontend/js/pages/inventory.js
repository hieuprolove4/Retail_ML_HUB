/**
 * Page 4: Tồn kho & Cảnh báo Điểm đặt hàng lại ROP (inventory.js)
 * Kết nối trực tiếp nhu cầu dự báo từ mô hình tốt nhất vào công thức SS & ROP,
 * lọc theo trạng thái CRITICAL / WARNING / SAFE và vẽ biểu đồ diễn biến tồn kho 90 ngày.
 */

import { apiGet, appState, updateState, getOptions } from '../api.js';
import { fmtInt, fmtDec, statusBadge, refreshIcons, exportCSV, renderError } from '../ui.js';
import { baseChartOptions, getChartTheme, mountChart } from '../charts.js';

let currentItems = [];
let statusFilter = 'ALL';
let searchKeyword = '';

export async function renderInventoryPage(container) {
  const opts = await getOptions();

  container.innerHTML = `
    <!-- Bộ lọc cấu hình tồn kho -->
    <section class="card" aria-label="Bộ lọc chính sách tồn kho">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="warehouse" style="width:18px;height:18px"></i>
            <span>Kiểm tra tồn kho & Cảnh báo đặt hàng lại (ROP)</span>
          </div>
          <div class="card-subtitle">
            Công thức: <code style="font-family:var(--font-mono)">SS = Z × σ_D × √L</code> và
            <code style="font-family:var(--font-mono)">ROP = (D̂ × L) + SS</code> — Nhu cầu D̂ lấy từ dự báo mô hình học máy
          </div>
        </div>
        <span class="badge badge-INFO" id="inv-policy-badge">Đang tính toán...</span>
      </div>

      <div class="filter-bar">
        <div class="form-field">
          <label for="inv-store">Cửa hàng (Store ID)</label>
          <select id="inv-store" class="form-control">
            ${opts.stores.map((s) => `<option value="${s}" ${s === appState.store_id ? 'selected' : ''}>Cửa hàng ${s}</option>`).join('')}
          </select>
        </div>

        <div class="form-field">
          <label for="inv-lead">Thời gian giao hàng (Lead Time L)</label>
          <select id="inv-lead" class="form-control">
            ${[1, 2, 3, 5, 7, 10, 14].map((d) => `<option value="${d}" ${d === Number(appState.lead_time) ? 'selected' : ''}>${d} ngày</option>`).join('')}
          </select>
        </div>

        <div class="form-field">
          <label for="inv-sl">Mức độ dịch vụ (Service Level)</label>
          <select id="inv-sl" class="form-control">
            <option value="0.90" ${Number(appState.service_level) === 0.9 ? 'selected' : ''}>90% (Z = 1.28)</option>
            <option value="0.95" ${Number(appState.service_level) === 0.95 ? 'selected' : ''}>95% (Z = 1.645 — Chuẩn báo cáo)</option>
            <option value="0.99" ${Number(appState.service_level) === 0.99 ? 'selected' : ''}>99% (Z = 2.33)</option>
          </select>
        </div>

        <div class="form-field">
          <label for="inv-source">Nguồn nhu cầu dự kiến (D̂)</label>
          <select id="inv-source" class="form-control">
            <option value="model" ${appState.demand_source === 'model' ? 'selected' : ''}>Dự báo từ ${opts.best_model} (ML)</option>
            <option value="average" ${appState.demand_source === 'average' ? 'selected' : ''}>Trung bình lịch sử 90 ngày</option>
          </select>
        </div>

        <div class="form-field">
          <label for="inv-search">Tìm nhanh mã SKU / Ngành hàng</label>
          <input id="inv-search" type="search" class="form-control" placeholder="VD: P0001, Groceries..." />
        </div>
      </div>
    </section>

    <!-- 4 Thẻ đếm trạng thái (Click để lọc bảng) -->
    <section class="kpi-grid" aria-label="Thống kê trạng thái tồn kho">
      <div class="kpi-card clickable" data-filter="ALL">
        <div class="kpi-top">
          <span>Tổng số mã SKU</span>
          <div class="kpi-icon"><i data-lucide="layers" style="width:17px;height:17px"></i></div>
        </div>
        <div class="kpi-value" id="inv-count-total">—</div>
        <div class="kpi-meta">Nhấn để hiện tất cả</div>
      </div>

      <div class="kpi-card clickable" data-filter="CRITICAL" style="border-left:4px solid var(--color-critical)">
        <div class="kpi-top">
          <span>Nguy cấp (≤ ROP)</span>
          <div class="kpi-icon" style="background:var(--color-critical-bg);color:var(--color-critical)">
            <i data-lucide="alert-octagon" style="width:17px;height:17px"></i>
          </div>
        </div>
        <div class="kpi-value" id="inv-count-critical" style="color:var(--color-critical)">—</div>
        <div class="kpi-meta">Cần phát lệnh đặt hàng ngay</div>
      </div>

      <div class="kpi-card clickable" data-filter="WARNING" style="border-left:4px solid var(--color-warning)">
        <div class="kpi-top">
          <span>Cảnh báo (≤ 1.2 × ROP)</span>
          <div class="kpi-icon" style="background:var(--color-warning-bg);color:var(--color-warning)">
            <i data-lucide="alert-triangle" style="width:17px;height:17px"></i>
          </div>
        </div>
        <div class="kpi-value" id="inv-count-warning" style="color:var(--color-warning)">—</div>
        <div class="kpi-meta">Sắp chạm ngưỡng đặt hàng lại</div>
      </div>

      <div class="kpi-card clickable" data-filter="SAFE" style="border-left:4px solid var(--color-safe)">
        <div class="kpi-top">
          <span>An toàn (&gt; 1.2 × ROP)</span>
          <div class="kpi-icon" style="background:var(--color-safe-bg);color:var(--color-safe)">
            <i data-lucide="check-circle-2" style="width:17px;height:17px"></i>
          </div>
        </div>
        <div class="kpi-value" id="inv-count-safe" style="color:var(--color-safe)">—</div>
        <div class="kpi-meta">Tồn kho đủ đáp ứng</div>
      </div>
    </section>

    <!-- Biểu đồ mô phỏng tồn kho so với ROP & SS của SKU đang chọn -->
    <section class="card">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="activity" style="width:18px;height:18px"></i>
            <span id="inv-sim-title">Diễn biến tồn kho thực tế 90 ngày so với ngưỡng ROP & Safety Stock</span>
          </div>
          <div class="card-subtitle">Nhấn vào bất kỳ dòng sản phẩm nào trong bảng bên dưới để xem biểu đồ của SKU đó</div>
        </div>
        <a href="#/forecast" class="btn btn-outline">
          <i data-lucide="trending-up" style="width:16px;height:16px"></i>
          <span>Xem dự báo SKU này</span>
        </a>
      </div>
      <div id="inv-history-chart" style="min-height:310px"></div>
    </section>

    <!-- Bảng 20 sản phẩm -->
    <section class="card">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="clipboard-list" style="width:18px;height:18px"></i>
            <span>Bảng kiểm tra tồn kho & Đề xuất lượng đặt hàng (20 SKU)</span>
          </div>
          <div class="card-subtitle">Lượng đặt hàng đề xuất = làm tròn lên <code style="font-family:var(--font-mono)">max(0, 1.5 × ROP − Tồn hiện tại)</code> khi chạm ngưỡng ROP</div>
        </div>
        <button type="button" class="btn btn-outline" id="inv-export-btn">
          <i data-lucide="download" style="width:16px;height:16px"></i>
          <span>Xuất danh sách đặt hàng (CSV)</span>
        </button>
      </div>

      <div class="table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th>Mã SKU</th>
              <th>Ngành hàng</th>
              <th class="num">Tồn hiện tại</th>
              <th class="num">Nhu cầu D̂ / ngày</th>
              <th class="num">Độ lệch chuẩn σ_D</th>
              <th class="num">Safety Stock (SS)</th>
              <th class="num">Điểm đặt lại (ROP)</th>
              <th>Trạng thái</th>
              <th class="num">Đề xuất nhập (Đv)</th>
            </tr>
          </thead>
          <tbody id="inv-table-body"></tbody>
        </table>
      </div>
    </section>
  `;

  refreshIcons(container);

  const storeSel = container.querySelector('#inv-store');
  const leadSel = container.querySelector('#inv-lead');
  const slSel = container.querySelector('#inv-sl');
  const srcSel = container.querySelector('#inv-source');
  const searchInp = container.querySelector('#inv-search');

  const reload = async () => {
    updateState({
      store_id: storeSel.value,
      lead_time: Number(leadSel.value),
      service_level: Number(slSel.value),
      demand_source: srcSel.value,
    });
    await loadInventoryStatus(container);
  };

  storeSel.addEventListener('change', reload);
  leadSel.addEventListener('change', reload);
  slSel.addEventListener('change', reload);
  srcSel.addEventListener('change', reload);

  searchInp.addEventListener('input', () => {
    searchKeyword = searchInp.value.trim().toLowerCase();
    renderTableRows(container);
  });

  container.querySelectorAll('[data-filter]').forEach((card) => {
    card.addEventListener('click', () => {
      statusFilter = card.dataset.filter;
      renderTableRows(container);
    });
  });

  container.querySelector('#inv-export-btn').addEventListener('click', () => {
    if (!currentItems.length) return;
    const rows = currentItems.map((it) => [
      appState.store_id,
      it.product_id,
      it.category,
      it.current_inventory,
      it.forecast_demand,
      it.demand_std,
      it.safety_stock,
      it.reorder_point,
      it.status_code,
      it.suggested_order,
    ]);
    exportCSV(
      `ton_kho_rop_${appState.store_id}.csv`,
      ['Cửa hàng', 'Mã SKU', 'Ngành hàng', 'Tồn hiện tại', 'Nhu cầu dự báo/ngày', 'Độ lệch chuẩn', 'Safety Stock', 'ROP', 'Trạng thái', 'Đề xuất nhập'],
      rows
    );
  });

  await loadInventoryStatus(container);
}

async function loadInventoryStatus(container) {
  const badge = container.querySelector('#inv-policy-badge');
  badge.textContent = 'Đang tính toán 20 SKU...';

  try {
    const data = await apiGet('/api/inventory-status', {
      store_id: appState.store_id,
      lead_time: appState.lead_time,
      service_level: appState.service_level,
      demand_source: appState.demand_source,
    });

    currentItems = data.items;
    container.querySelector('#inv-count-total').textContent = fmtInt(data.summary.total);
    container.querySelector('#inv-count-critical').textContent = fmtInt(data.summary.critical);
    container.querySelector('#inv-count-warning').textContent = fmtInt(data.summary.warning);
    container.querySelector('#inv-count-safe').textContent = fmtInt(data.summary.safe);

    badge.textContent = `Store ${data.store_id} • L=${data.lead_time_days} ngày • SL=${data.service_level} (Z=${data.z_score}) • Nguồn D̂: ${data.demand_model}`;

    renderTableRows(container);
    await loadSkuHistoryChart(container, appState.product_id);
  } catch (err) {
    renderError(container.querySelector('#inv-table-body'), err, () => loadInventoryStatus(container));
  }
}

function renderTableRows(container) {
  const tbody = container.querySelector('#inv-table-body');
  const filtered = currentItems.filter((it) => {
    if (statusFilter !== 'ALL' && it.status_code !== statusFilter) return false;
    if (searchKeyword) {
      const hay = `${it.product_id} ${it.category}`.toLowerCase();
      if (!hay.includes(searchKeyword)) return false;
    }
    return true;
  });

  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;padding:20px;color:var(--color-muted-foreground)">Không có sản phẩm nào khớp bộ lọc.</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered
    .map(
      (it) => `
      <tr class="${it.product_id === appState.product_id ? 'row-selected' : ''}" data-sku="${it.product_id}" style="cursor:pointer">
        <td style="font-family:var(--font-mono);font-weight:700">${it.product_id}</td>
        <td>${it.category}</td>
        <td class="num">${fmtInt(it.current_inventory)}</td>
        <td class="num">${fmtDec(it.forecast_demand)}</td>
        <td class="num">${fmtDec(it.demand_std)}</td>
        <td class="num">${fmtInt(it.safety_stock)}</td>
        <td class="num"><strong>${fmtInt(it.reorder_point)}</strong></td>
        <td>${statusBadge(it.status_code)}</td>
        <td class="num" style="font-weight:700;color:${it.suggested_order > 0 ? 'var(--color-critical)' : 'var(--color-muted-foreground)'}">
          ${it.suggested_order > 0 ? `+${fmtInt(it.suggested_order)}` : '0'}
        </td>
      </tr>
    `
    )
    .join('');

  refreshIcons(tbody);

  tbody.querySelectorAll('tr[data-sku]').forEach((row) => {
    row.addEventListener('click', async () => {
      const sku = row.dataset.sku;
      updateState({ product_id: sku });
      tbody.querySelectorAll('tr').forEach((r) => r.classList.remove('row-selected'));
      row.classList.add('row-selected');
      await loadSkuHistoryChart(container, sku);
    });
  });
}

async function loadSkuHistoryChart(container, productId) {
  const titleEl = container.querySelector('#inv-sim-title');
  try {
    const h = await apiGet('/api/inventory-history', {
      store_id: appState.store_id,
      product_id: productId,
      days: 90,
      lead_time: appState.lead_time,
      service_level: appState.service_level,
      demand_source: appState.demand_source,
    });

    titleEl.textContent = `Diễn biến tồn kho 90 ngày — ${h.store_id} / ${h.product_id} (Tồn hiện tại: ${fmtInt(h.current_inventory)} | ROP: ${fmtInt(h.rop)} | SS: ${fmtInt(h.ss)})`;

    const t = getChartTheme();
    const ropLine = h.dates.map(() => h.rop);
    const ssLine = h.dates.map(() => h.ss);

    const options = baseChartOptions({
      chart: { type: 'line', height: 310 },
      series: [
        { name: 'Tồn kho thực tế', type: 'line', data: h.inventory },
        { name: `Điểm đặt hàng lại (ROP = ${h.rop})`, type: 'line', data: ropLine },
        { name: `Tồn kho an toàn (SS = ${h.ss})`, type: 'line', data: ssLine },
      ],
      colors: [t.actualColor, t.forecastColor, t.criticalColor],
      stroke: {
        width: [2.5, 2, 2],
        curve: 'smooth',
        dashArray: [0, 6, 3],
      },
      xaxis: {
        categories: h.dates,
        tickAmount: 10,
        labels: { rotate: -25, style: { colors: t.textColor, fontSize: '11px' } },
      },
      yaxis: {
        title: { text: 'Số lượng tồn kho (Đơn vị)', style: { color: t.textColor } },
        min: 0,
      },
      tooltip: {
        shared: true,
        intersect: false,
      },
    });

    mountChart('inv-history-chart', options);
  } catch (err) {
    renderError(container.querySelector('#inv-history-chart'), err);
  }
}
