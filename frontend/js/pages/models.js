/**
 * Page 3: So sánh định lượng 5 Mô hình Dự báo & Feature Importance (models.js)
 * Tương ứng Bảng 4.2, Hình 4.4–4.10 trong báo cáo.
 */

import { apiGet, appState, updateState, getOptions } from '../api.js';
import { fmtDec, fmtInt, refreshIcons, renderError } from '../ui.js';
import { baseChartOptions, getChartTheme, mountChart } from '../charts.js';

export async function renderModelsPage(container) {
  const opts = await getOptions();

  container.innerHTML = `
    <!-- Bộ lọc chọn SKU để đánh giá 5 mô hình -->
    <section class="card" aria-label="Chọn đối tượng đánh giá mô hình">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="git-compare" style="width:18px;height:18px"></i>
            <span>Đánh giá & So sánh hiệu năng 5 Mô hình Dự báo (Train 80% / Test 20%)</span>
          </div>
          <div class="card-subtitle">Đối chiếu ARIMA, SARIMA, Holt-Winters, Random Forest và XGBoost theo chuẩn Bảng 4.2 trong báo cáo</div>
        </div>
        <span class="badge badge-INFO" id="mdl-source-badge">Đang tải...</span>
      </div>

      <div class="filter-bar">
        <div class="form-field">
          <label for="mdl-store">Cửa hàng (Store ID)</label>
          <select id="mdl-store" class="form-control">
            ${opts.stores.map((s) => `<option value="${s}" ${s === appState.store_id ? 'selected' : ''}>Cửa hàng ${s}</option>`).join('')}
          </select>
        </div>

        <div class="form-field">
          <label for="mdl-product">Sản phẩm (Product ID)</label>
          <select id="mdl-product" class="form-control">
            ${opts.products.map((p) => `<option value="${p.id}" ${p.id === appState.product_id ? 'selected' : ''}>${p.id} — ${p.category}</option>`).join('')}
          </select>
        </div>

        <div class="form-field">
          <button type="button" class="btn btn-primary" id="mdl-run-btn">
            <i data-lucide="refresh-cw" style="width:16px;height:16px"></i>
            <span>Đánh giá 5 mô hình</span>
          </button>
        </div>
      </div>
    </section>

    <!-- Bảng 4.2 + Biểu đồ cột so sánh RMSE / MAE / MAPE -->
    <div class="grid-2">
      <section class="card">
        <div class="card-header">
          <div>
            <div class="card-title">
              <i data-lucide="award" style="width:18px;height:18px"></i>
              <span>Bảng so sánh các chỉ số đánh giá (Table 4.2)</span>
            </div>
            <div class="card-subtitle">Xếp hạng từ tốt nhất đến thấp nhất theo chỉ số RMSE</div>
          </div>
        </div>

        <div class="table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>Hạng</th>
                <th>Mô hình (Model)</th>
                <th class="num">RMSE</th>
                <th class="num">MAE</th>
                <th class="num">MAPE (%)</th>
                <th>Đánh giá</th>
              </tr>
            </thead>
            <tbody id="mdl-table-body"></tbody>
          </table>
        </div>
      </section>

      <section class="card">
        <div class="card-header">
          <div>
            <div class="card-title">
              <i data-lucide="bar-chart-3" style="width:18px;height:18px"></i>
              <span>Biểu đồ so sánh RMSE & MAE (Hình 4.9)</span>
            </div>
            <div class="card-subtitle">Giá trị càng thấp chứng tỏ độ chính xác dự báo càng cao</div>
          </div>
        </div>
        <div id="mdl-metrics-chart" style="min-height:260px"></div>
      </section>
    </div>

    <!-- Biểu đồ Actual vs Predicted trên tập Test (20%) -->
    <section class="card">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="activity" style="width:18px;height:18px"></i>
            <span>Đối chiếu Thực tế vs Dự báo trên Tập kiểm thử 20% (Hình 4.4 – 4.8)</span>
          </div>
          <div class="card-subtitle">Nhấn vào tên mô hình ở chú giải (legend) phía trên biểu đồ để bật / tắt từng đường dự báo</div>
        </div>
      </div>
      <div id="mdl-test-chart" style="min-height:360px"></div>
    </section>

    <!-- Feature Importance (Hình 4.10) -->
    <section class="card">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="sparkles" style="width:18px;height:18px"></i>
            <span>Top 10 Đặc trưng quan trọng nhất (Feature Importance — Hình 4.10)</span>
          </div>
          <div class="card-subtitle">Trích xuất từ mô hình cây quyết định huấn luyện trên SKU chuẩn của báo cáo (S001 - P0001)</div>
        </div>

        <div class="segmented" role="group" aria-label="Chọn mô hình cây">
          <button type="button" class="seg-btn active" data-fimodel="Random Forest">Random Forest</button>
          <button type="button" class="seg-btn" data-fimodel="XGBoost">XGBoost</button>
        </div>
      </div>
      <div id="mdl-fi-chart" style="min-height:310px"></div>
    </section>
  `;

  refreshIcons(container);

  const storeSel = container.querySelector('#mdl-store');
  const prodSel = container.querySelector('#mdl-product');
  const runBtn = container.querySelector('#mdl-run-btn');
  const fiBtns = container.querySelectorAll('[data-fimodel]');

  const triggerEval = async () => {
    updateState({ store_id: storeSel.value, product_id: prodSel.value });
    await loadComparison(container);
  };

  storeSel.addEventListener('change', triggerEval);
  prodSel.addEventListener('change', triggerEval);
  runBtn.addEventListener('click', triggerEval);

  fiBtns.forEach((btn) => {
    btn.addEventListener('click', async () => {
      fiBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      await loadFeatureImportance(btn.dataset.fimodel);
    });
  });

  await Promise.all([loadComparison(container), loadFeatureImportance('Random Forest')]);
}

async function loadComparison(container) {
  const runBtn = container.querySelector('#mdl-run-btn');
  if (runBtn) runBtn.disabled = true;

  try {
    const data = await apiGet('/api/model-comparison', {
      store_id: appState.store_id,
      product_id: appState.product_id,
    });

    const badge = container.querySelector('#mdl-source-badge');
    if (data.metrics_source === 'report') {
      badge.className = 'badge badge-SAFE';
      badge.textContent = `Số liệu chuẩn Báo cáo (${data.store_id} - ${data.product_id}) • Tốt nhất: ${data.best_model}`;
    } else {
      badge.className = 'badge badge-INFO';
      badge.textContent = `Đánh giá trực tiếp (${data.store_id} - ${data.product_id}) • Tốt nhất: ${data.best_model}`;
    }

    const sorted = [...data.results].sort((a, b) => a.RMSE - b.RMSE);
    const tbody = container.querySelector('#mdl-table-body');
    tbody.innerHTML = sorted
      .map(
        (r, idx) => `
        <tr class="${idx === 0 ? 'row-selected' : ''}">
          <td class="num"><strong>#${idx + 1}</strong></td>
          <td><strong>${r.model}</strong></td>
          <td class="num">${fmtDec(r.RMSE)}</td>
          <td class="num">${fmtDec(r.MAE)}</td>
          <td class="num">${fmtDec(r['MAPE (%)'])}%</td>
          <td>
            ${
              idx === 0
                ? '<span class="badge badge-SAFE">Tối ưu nhất</span>'
                : '<span class="badge badge-INFO">Đối chứng</span>'
            }
          </td>
        </tr>
      `
      )
      .join('');

    const t = getChartTheme();

    // Biểu đồ cột RMSE & MAE
    mountChart(
      'mdl-metrics-chart',
      baseChartOptions({
        chart: { type: 'bar', height: 260 },
        series: [
          { name: 'RMSE', data: sorted.map((r) => r.RMSE) },
          { name: 'MAE', data: sorted.map((r) => r.MAE) },
        ],
        colors: [t.actualColor, t.safeColor],
        plotOptions: { bar: { borderRadius: 4, columnWidth: '58%' } },
        xaxis: { categories: sorted.map((r) => r.model), labels: { style: { colors: t.textColor } } },
      })
    );

    // Biểu đồ đường trên tập Test (lấy 90 ngày cuối của tập test cho rõ)
    const sliceN = 90;
    const dates = data.test.dates.slice(-sliceN);
    const actual = data.test.actual.slice(-sliceN);
    const series = [
      { name: 'Thực tế (Actual Test)', data: actual },
      ...sorted.map((r) => ({
        name: r.model,
        data: (data.test.pred[r.model] || []).slice(-sliceN),
      })),
    ];

    mountChart(
      'mdl-test-chart',
      baseChartOptions({
        chart: { type: 'line', height: 360 },
        series,
        colors: ['#0F172A', '#16A34A', '#2563EB', '#D97706', '#9333EA', '#DC2626'],
        stroke: {
          width: [2.6, 2.2, 2.0, 1.8, 1.8, 1.8],
          curve: 'smooth',
          dashArray: [0, 0, 4, 3, 2, 2],
        },
        xaxis: {
          categories: dates,
          tickAmount: 10,
          labels: { rotate: -20, style: { colors: t.textColor, fontSize: '11px' } },
        },
        yaxis: {
          title: { text: 'Sản lượng bán (Đơn vị)', style: { color: t.textColor } },
          labels: { formatter: (v) => fmtInt(Math.round(v)), style: { colors: t.textColor } },
        },
      })
    );
  } catch (err) {
    renderError(container.querySelector('#mdl-test-chart'), err, () => loadComparison(container));
  } finally {
    if (runBtn) runBtn.disabled = false;
  }
}

async function loadFeatureImportance(modelName) {
  const data = await apiGet('/api/feature-importance', { model: modelName, top_n: 10 });
  const t = getChartTheme();

  mountChart(
    'mdl-fi-chart',
    baseChartOptions({
      chart: { type: 'bar', height: 310 },
      series: [{ name: 'Mức độ quan trọng (Importance)', data: data.importance }],
      colors: [t.actualColor],
      plotOptions: { bar: { horizontal: true, borderRadius: 4, barHeight: '62%' } },
      xaxis: {
        categories: data.features,
        labels: { style: { colors: t.textColor } },
      },
      tooltip: {
        y: { formatter: (v) => `${(v * 100).toFixed(2)}%` },
      },
    })
  );
}
