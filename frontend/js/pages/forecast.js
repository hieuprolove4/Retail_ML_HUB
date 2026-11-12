/**
 * Page 2: Dự báo nhu cầu sản phẩm (forecast.js)
 * Hiển thị lịch sử 60 ngày + dự báo tương lai kèm khoảng tin cậy 95%, chỉ số đánh giá,
 * hộp giả định dự báo và bảng dữ liệu xuất CSV.
 */

import { apiPost, appState, updateState, getOptions } from '../api.js';
import { fmtInt, fmtDec, weekdayVN, refreshIcons, exportCSV, renderError } from '../ui.js';
import { baseChartOptions, getChartTheme, mountChart } from '../charts.js';

let lastForecastData = null;

export async function renderForecastPage(container) {
  const opts = await getOptions();

  container.innerHTML = `
    <!-- Bộ lọc cấu hình dự báo -->
    <section class="card" aria-label="Bộ lọc cấu hình dự báo">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="sliders" style="width:18px;height:18px"></i>
            <span>Cấu hình dự báo nhu cầu theo SKU</span>
          </div>
          <div class="card-subtitle">Chọn cửa hàng, sản phẩm, mô hình học máy và chân trời dự báo tương lai</div>
        </div>
        <span class="badge badge-INFO" id="fc-source-badge">Đang tải...</span>
      </div>

      <div class="filter-bar">
        <div class="form-field">
          <label for="fc-store">Cửa hàng (Store ID)</label>
          <select id="fc-store" class="form-control">
            ${opts.stores.map((s) => `<option value="${s}" ${s === appState.store_id ? 'selected' : ''}>Cửa hàng ${s}</option>`).join('')}
          </select>
        </div>

        <div class="form-field">
          <label for="fc-product">Sản phẩm (Product ID)</label>
          <select id="fc-product" class="form-control">
            ${opts.products.map((p) => `<option value="${p.id}" ${p.id === appState.product_id ? 'selected' : ''}>${p.id} — ${p.category}</option>`).join('')}
          </select>
        </div>

        <div class="form-field">
          <label for="fc-model">Mô hình dự báo</label>
          <select id="fc-model" class="form-control">
            ${opts.models.map((m) => `<option value="${m}" ${m === appState.model_name ? 'selected' : ''}>${m}${m === opts.best_model ? ' (Tốt nhất)' : ''}</option>`).join('')}
          </select>
        </div>

        <div class="form-field">
          <label>Chân trời dự báo (Horizon)</label>
          <div class="segmented" role="group" aria-label="Chọn số ngày dự báo">
            ${opts.horizons.map((h) => `<button type="button" class="seg-btn ${h === Number(appState.horizon) ? 'active' : ''}" data-horizon="${h}">${h} ngày</button>`).join('')}
          </div>
        </div>

        <div class="form-field">
          <button type="button" class="btn btn-primary" id="fc-run-btn">
            <i data-lucide="play" style="width:16px;height:16px"></i>
            <span>Chạy dự báo</span>
          </button>
        </div>
      </div>
    </section>

    <!-- 4 Thẻ chỉ số đánh giá & tổng nhu cầu dự kiến -->
    <section class="kpi-grid" aria-label="Chỉ số đánh giá mô hình và tổng hợp dự báo">
      <div class="kpi-card">
        <div class="kpi-top">
          <span>Tổng cầu dự kiến (<span id="fc-kpi-horizon">${appState.horizon}</span> ngày)</span>
          <div class="kpi-icon"><i data-lucide="shopping-bag" style="width:17px;height:17px"></i></div>
        </div>
        <div class="kpi-value" id="fc-kpi-total">—</div>
        <div class="kpi-meta" id="fc-kpi-avg">Bình quân: — đơn vị/ngày</div>
      </div>

      <div class="kpi-card">
        <div class="kpi-top">
          <span>RMSE (Sai số chuẩn)</span>
          <div class="kpi-icon"><i data-lucide="gauge" style="width:17px;height:17px"></i></div>
        </div>
        <div class="kpi-value" id="fc-kpi-rmse">—</div>
        <div class="kpi-meta">Đánh giá trên tập Test (20%)</div>
      </div>

      <div class="kpi-card">
        <div class="kpi-top">
          <span>MAE (Sai số tuyệt đối)</span>
          <div class="kpi-icon"><i data-lucide="activity" style="width:17px;height:17px"></i></div>
        </div>
        <div class="kpi-value" id="fc-kpi-mae">—</div>
        <div class="kpi-meta">Đơn vị sản phẩm</div>
      </div>

      <div class="kpi-card">
        <div class="kpi-top">
          <span>MAPE (%)</span>
          <div class="kpi-icon"><i data-lucide="percent" style="width:17px;height:17px"></i></div>
        </div>
        <div class="kpi-value" id="fc-kpi-mape">—</div>
        <div class="kpi-meta">Tỷ lệ phần trăm sai số tuyệt đối</div>
      </div>
    </section>

    <!-- Biểu đồ Lịch sử + Dự báo + Dải tin cậy 95% -->
    <section class="card">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="line-chart" style="width:18px;height:18px"></i>
            <span id="fc-chart-title">Biểu đồ Lịch sử 60 ngày & Dự báo tương lai</span>
          </div>
          <div class="card-subtitle">Đường nét liền xanh: Nhu cầu thực tế | Đường nét đứt cam: Dự báo | Vùng mờ: Khoảng tin cậy 95%</div>
        </div>
        <div style="display:flex;gap:8px">
          <a href="#/inventory" class="btn btn-outline" title="Xem trạng thái tồn kho của cửa hàng này">
            <i data-lucide="package-check" style="width:16px;height:16px"></i>
            <span>Kiểm tra tồn kho ROP</span>
          </a>
        </div>
      </div>
      <div id="fc-main-chart" style="min-height:370px"></div>
    </section>

    <!-- Hộp giả định & Bảng chi tiết dự báo -->
    <div class="grid-asymmetric">
      <section class="card">
        <div class="card-header">
          <div>
            <div class="card-title">
              <i data-lucide="table-2" style="width:18px;height:18px"></i>
              <span>Bảng chi tiết dự báo theo ngày</span>
            </div>
            <div class="card-subtitle">Số lượng dự báo điểm và cận dưới / cận trên (95% CI)</div>
          </div>
          <button type="button" class="btn btn-outline" id="fc-export-btn">
            <i data-lucide="download" style="width:16px;height:16px"></i>
            <span>Xuất CSV</span>
          </button>
        </div>

        <div class="table-wrap" style="max-height:340px;overflow-y:auto">
          <table class="data-table">
            <thead>
              <tr>
                <th>Ngày</th>
                <th>Thứ</th>
                <th class="num">Dự báo (Đơn vị)</th>
                <th class="num">Cận dưới 95%</th>
                <th class="num">Cận trên 95%</th>
              </tr>
            </thead>
            <tbody id="fc-table-body"></tbody>
          </table>
        </div>
      </section>

      <section class="card">
        <div class="card-header">
          <div>
            <div class="card-title">
              <i data-lucide="info" style="width:18px;height:18px"></i>
              <span>Cơ chế & Giả định dự báo</span>
            </div>
            <div class="card-subtitle">Giải thích minh bạch cách mô hình tạo dự báo tương lai</div>
          </div>
        </div>
        <div class="callout">
          <div class="callout-title">
            <i data-lucide="cpu" style="width:16px;height:16px"></i>
            <span id="fc-assumption-model">Mô hình</span>
          </div>
          <ul id="fc-assumptions-list"></ul>
        </div>
      </section>
    </div>
  `;

  refreshIcons(container);

  const storeSel = container.querySelector('#fc-store');
  const prodSel = container.querySelector('#fc-product');
  const modelSel = container.querySelector('#fc-model');
  const runBtn = container.querySelector('#fc-run-btn');
  const horizonBtns = container.querySelectorAll('[data-horizon]');

  const triggerForecast = async () => {
    updateState({
      store_id: storeSel.value,
      product_id: prodSel.value,
      model_name: modelSel.value,
    });
    await executeForecast(container);
  };

  storeSel.addEventListener('change', triggerForecast);
  prodSel.addEventListener('change', triggerForecast);
  modelSel.addEventListener('change', triggerForecast);
  runBtn.addEventListener('click', triggerForecast);

  horizonBtns.forEach((btn) => {
    btn.addEventListener('click', async () => {
      horizonBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      updateState({ horizon: Number(btn.dataset.horizon) });
      await executeForecast(container);
    });
  });

  container.querySelector('#fc-export-btn').addEventListener('click', () => {
    if (!lastForecastData) return;
    const { store_id, product_id, model_used, forecast } = lastForecastData;
    const rows = forecast.dates.map((d, i) => [
      d,
      weekdayVN(d),
      forecast.values[i],
      forecast.lower[i],
      forecast.upper[i],
    ]);
    exportCSV(
      `du_bao_${store_id}_${product_id}_${model_used.replace(/\s+/g, '_')}.csv`,
      ['Ngày', 'Thứ', 'Dự báo', 'Cận dưới 95%', 'Cận trên 95%'],
      rows
    );
  });

  await executeForecast(container);
}

async function executeForecast(container) {
  const runBtn = container.querySelector('#fc-run-btn');
  const chartBox = container.querySelector('#fc-main-chart');
  if (runBtn) runBtn.disabled = true;

  try {
    const data = await apiPost('/api/forecast', {
      store_id: appState.store_id,
      product_id: appState.product_id,
      model_name: appState.model_name,
      horizon: Number(appState.horizon),
    });
    lastForecastData = data;

    // Cập nhật KPI
    const total = data.forecast.total;
    const avg = total / data.horizon;
    container.querySelector('#fc-kpi-horizon').textContent = data.horizon;
    container.querySelector('#fc-kpi-total').textContent = `${fmtInt(total)} đv`;
    container.querySelector('#fc-kpi-avg').textContent = `Bình quân: ${fmtDec(avg)} đơn vị/ngày`;
    container.querySelector('#fc-kpi-rmse').textContent = fmtDec(data.metrics.RMSE);
    container.querySelector('#fc-kpi-mae').textContent = fmtDec(data.metrics.MAE);
    container.querySelector('#fc-kpi-mape').textContent = `${fmtDec(data.metrics['MAPE (%)'])}%`;

    const badge = container.querySelector('#fc-source-badge');
    if (data.metrics_source === 'report') {
      badge.className = 'badge badge-SAFE';
      badge.textContent = 'Khớp số liệu Báo cáo (S001 - P0001)';
    } else {
      badge.className = 'badge badge-INFO';
      badge.textContent = `Huấn luyện trực tiếp cho ${data.store_id} - ${data.product_id}`;
    }

    container.querySelector('#fc-chart-title').textContent =
      `Lịch sử 60 ngày & Dự báo ${data.horizon} ngày tới — ${data.store_id} / ${data.product_id} (${data.category}) — ${data.model_used}`;

    // Giả định
    container.querySelector('#fc-assumption-model').textContent = `Mô hình: ${data.model_used}`;
    container.querySelector('#fc-assumptions-list').innerHTML =
      data.assumptions.map((a) => `<li>${a}</li>`).join('');

    // Bảng dự báo
    const tbody = container.querySelector('#fc-table-body');
    tbody.innerHTML = data.forecast.dates
      .map(
        (d, i) => `
        <tr>
          <td style="font-family:var(--font-mono)">${d}</td>
          <td>${weekdayVN(d)}</td>
          <td class="num"><strong>${fmtInt(data.forecast.values[i])}</strong></td>
          <td class="num">${fmtInt(data.forecast.lower[i])}</td>
          <td class="num">${fmtInt(data.forecast.upper[i])}</td>
        </tr>
      `
      )
      .join('');

    // Vẽ biểu đồ
    renderForecastChart(data);
  } catch (err) {
    renderError(chartBox, err, () => executeForecast(container));
  } finally {
    if (runBtn) runBtn.disabled = false;
  }
}

function renderForecastChart(data) {
  const t = getChartTheme();
  const allDates = [...data.history.dates, ...data.forecast.dates];
  const histLen = data.history.dates.length;
  const fcLen = data.forecast.dates.length;

  // Nối điểm cuối lịch sử vào đầu chuỗi dự báo để đường biểu đồ liền mạch
  const lastHistVal = data.history.values[histLen - 1];
  const actualSeries = [...data.history.values, ...Array(fcLen).fill(null)];
  const forecastSeries = [
    ...Array(histLen - 1).fill(null),
    lastHistVal,
    ...data.forecast.values,
  ];
  const upperSeries = [
    ...Array(histLen - 1).fill(null),
    lastHistVal,
    ...data.forecast.upper,
  ];
  const lowerSeries = [
    ...Array(histLen - 1).fill(null),
    lastHistVal,
    ...data.forecast.lower,
  ];

  const boundaryDate = data.history.dates[histLen - 1];

  const options = baseChartOptions({
    chart: { type: 'line', height: 370 },
    series: [
      { name: 'Sản lượng thực tế', type: 'line', data: actualSeries },
      { name: `Dự báo (${data.model_used})`, type: 'line', data: forecastSeries },
      { name: 'Cận trên 95%', type: 'line', data: upperSeries },
      { name: 'Cận dưới 95%', type: 'line', data: lowerSeries },
    ],
    colors: [t.actualColor, t.forecastColor, t.forecastColor, t.forecastColor],
    stroke: {
      width: [2.4, 2.8, 1.2, 1.2],
      curve: 'smooth',
      dashArray: [0, 6, 3, 3],
    },
    xaxis: {
      categories: allDates,
      tickAmount: 10,
      labels: { rotate: -25, style: { colors: t.textColor, fontSize: '11px' } },
    },
    yaxis: {
      title: { text: 'Sản lượng bán (Đơn vị)', style: { color: t.textColor } },
      min: 0,
      labels: {
        style: { colors: t.textColor },
        formatter: (v) => (v === null || v === undefined ? '' : Math.round(v)),
      },
    },
    annotations: {
      xaxis: [
        {
          x: boundaryDate,
          borderColor: t.forecastColor,
          label: {
            text: 'Bắt đầu dự báo',
            style: { color: '#fff', background: t.forecastColor, fontWeight: 600 },
          },
        },
      ],
    },
    tooltip: {
      shared: true,
      intersect: false,
      y: {
        formatter: (v) => (v === null || v === undefined ? '—' : `${fmtInt(v)} đơn vị`),
      },
    },
  });

  mountChart('fc-main-chart', options);
}
