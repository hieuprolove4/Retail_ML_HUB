/**
 * Page 1: Tổng quan chuỗi cung ứng & Khám phá dữ liệu EDA (overview.js)
 */

import { apiGet, getOptions } from '../api.js';
import { fmtInt, fmtDec, fmtUSD, refreshIcons, renderError } from '../ui.js';
import { baseChartOptions, getChartTheme, mountChart } from '../charts.js';

export async function renderOverviewPage(container) {
  try {
    const [kpi, eda, opts] = await Promise.all([
      apiGet('/api/kpi'),
      apiGet('/api/eda-data'),
      getOptions(),
    ]);

    container.innerHTML = `
      <!-- 5 Thẻ KPI tổng quan -->
      <section class="kpi-grid" aria-label="Chỉ số tổng quan hệ thống">
        <div class="kpi-card">
          <div class="kpi-top">
            <span>Tổng bản ghi giao dịch</span>
            <div class="kpi-icon"><i data-lucide="database" style="width:17px;height:17px"></i></div>
          </div>
          <div class="kpi-value">${fmtInt(kpi.total_records)}</div>
          <div class="kpi-meta">${kpi.date_range}</div>
        </div>

        <div class="kpi-card">
          <div class="kpi-top">
            <span>Quy mô chuỗi</span>
            <div class="kpi-icon"><i data-lucide="store" style="width:17px;height:17px"></i></div>
          </div>
          <div class="kpi-value">${kpi.total_stores} Store × ${kpi.total_products} SKU</div>
          <div class="kpi-meta">${kpi.total_categories} ngành hàng bán lẻ</div>
        </div>

        <div class="kpi-card">
          <div class="kpi-top">
            <span>Tổng sản lượng tiêu thụ</span>
            <div class="kpi-icon"><i data-lucide="shopping-cart" style="width:17px;height:17px"></i></div>
          </div>
          <div class="kpi-value">${fmtInt(kpi.total_units_sold)}</div>
          <div class="kpi-meta">Đơn vị sản phẩm (2 năm)</div>
        </div>

        <div class="kpi-card">
          <div class="kpi-top">
            <span>Tổng doanh thu thuần</span>
            <div class="kpi-icon"><i data-lucide="dollar-sign" style="width:17px;height:17px"></i></div>
          </div>
          <div class="kpi-value">${fmtUSD(kpi.total_revenue)}</div>
          <div class="kpi-meta">Đã trừ chiết khấu khuyến mãi</div>
        </div>

        <div class="kpi-card">
          <div class="kpi-top">
            <span>Tồn kho trung bình / SKU</span>
            <div class="kpi-icon"><i data-lucide="package" style="width:17px;height:17px"></i></div>
          </div>
          <div class="kpi-value">${fmtDec(kpi.avg_inventory_per_store)}</div>
          <div class="kpi-meta">Đơn vị tồn mỗi ngày</div>
        </div>
      </section>

      <!-- Biểu đồ xu hướng sản lượng bán theo thời gian (D/W/M + MA) -->
      <section class="card">
        <div class="card-header">
          <div>
            <div class="card-title">
              <i data-lucide="trending-up" style="width:18px;height:18px"></i>
              <span>Xu hướng sản lượng bán & Đường trung bình trượt (Moving Average)</span>
            </div>
            <div class="card-subtitle">Hỗ trợ lọc theo cửa hàng, sản phẩm và tần suất Ngày / Tuần / Tháng</div>
          </div>

          <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
            <select id="ov-trend-store" class="form-control" style="width:auto" aria-label="Lọc theo cửa hàng">
              <option value="">Toàn chuỗi (5 cửa hàng)</option>
              ${opts.stores.map((s) => `<option value="${s}">Cửa hàng ${s}</option>`).join('')}
            </select>

            <select id="ov-trend-product" class="form-control" style="width:auto" aria-label="Lọc theo sản phẩm">
              <option value="">Tất cả sản phẩm (20 SKU)</option>
              ${opts.products.map((p) => `<option value="${p.id}">${p.id} — ${p.category}</option>`).join('')}
            </select>

            <div class="segmented" role="group" aria-label="Tần suất thời gian">
              <button type="button" class="seg-btn" data-freq="D">Ngày</button>
              <button type="button" class="seg-btn active" data-freq="W">Tuần</button>
              <button type="button" class="seg-btn" data-freq="M">Tháng</button>
            </div>
          </div>
        </div>
        <div id="ov-trend-chart" style="min-height:330px"></div>
      </section>

      <!-- Lưới 2x2 biểu đồ EDA -->
      <div class="grid-2">
        <section class="card">
          <div class="card-header">
            <div>
              <div class="card-title">
                <i data-lucide="pie-chart" style="width:18px;height:18px"></i>
                <span>Tổng sản lượng theo Ngành hàng</span>
              </div>
              <div class="card-subtitle">Phân bổ nhu cầu giữa 5 nhóm ngành hàng</div>
            </div>
          </div>
          <div id="ov-cat-chart" style="min-height:280px"></div>
        </section>

        <section class="card">
          <div class="card-header">
            <div>
              <div class="card-title">
                <i data-lucide="building-2" style="width:18px;height:18px"></i>
                <span>Tổng sản lượng theo Cửa hàng</span>
              </div>
              <div class="card-subtitle">So sánh quy mô tiêu thụ giữa 5 cửa hàng S001 – S005</div>
            </div>
          </div>
          <div id="ov-store-chart" style="min-height:280px"></div>
        </section>

        <section class="card">
          <div class="card-header">
            <div>
              <div class="card-title">
                <i data-lucide="cloud-sun" style="width:18px;height:18px"></i>
                <span>Sản lượng bán trung bình theo Thời tiết</span>
              </div>
              <div class="card-subtitle">Ảnh hưởng của điều kiện thời tiết tới sức mua hàng ngày</div>
            </div>
          </div>
          <div id="ov-weather-chart" style="min-height:280px"></div>
        </section>

        <section class="card">
          <div class="card-header">
            <div>
              <div class="card-title">
                <i data-lucide="calendar-range" style="width:18px;height:18px"></i>
                <span>Mùa vụ & Hiệu ứng Khuyến mãi</span>
              </div>
              <div class="card-subtitle">Sản lượng trung bình/ngày theo 4 mùa và chương trình khuyến mãi</div>
            </div>
          </div>
          <div id="ov-season-chart" style="min-height:280px"></div>
        </section>
      </div>
    `;

    refreshIcons(container);
    renderEdaCharts(eda);

    let currentFreq = 'W';
    const storeSel = container.querySelector('#ov-trend-store');
    const prodSel = container.querySelector('#ov-trend-product');
    const freqBtns = container.querySelectorAll('[data-freq]');

    const reloadTrend = () => loadSalesTrend(storeSel.value, prodSel.value, currentFreq);
    storeSel.addEventListener('change', reloadTrend);
    prodSel.addEventListener('change', reloadTrend);
    freqBtns.forEach((btn) => {
      btn.addEventListener('click', () => {
        freqBtns.forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        currentFreq = btn.dataset.freq;
        reloadTrend();
      });
    });

    await reloadTrend();
  } catch (err) {
    renderError(container, err, () => renderOverviewPage(container));
  }
}

async function loadSalesTrend(store_id, product_id, freq) {
  const data = await apiGet('/api/sales-trend', { store_id, product_id, freq });
  const t = getChartTheme();
  const [w1, w2] = data.ma_windows;
  const unitLabel = { D: 'ngày', W: 'tuần', M: 'tháng' }[freq];

  const options = baseChartOptions({
    chart: { type: 'line', height: 330 },
    series: [
      { name: `Sản lượng theo ${unitLabel}`, type: 'area', data: data.units },
      { name: `TB trượt ${w1} ${unitLabel}`, type: 'line', data: data.ma_short },
      { name: `TB trượt ${w2} ${unitLabel}`, type: 'line', data: data.ma_long },
    ],
    colors: [t.actualColor, t.forecastColor, t.safeColor],
    stroke: { width: [2, 2.6, 2.6], curve: 'smooth', dashArray: [0, 0, 4] },
    fill: { opacity: [0.22, 1, 1] },
    xaxis: {
      categories: data.dates,
      tickAmount: 10,
      labels: { rotate: -20, style: { colors: t.textColor, fontSize: '11px' } },
    },
    yaxis: {
      labels: {
        style: { colors: t.textColor },
        formatter: (v) => (v === null || v === undefined ? '' : fmtInt(Math.round(v))),
      },
    },
  });

  mountChart('ov-trend-chart', options);
}

function renderEdaCharts(eda) {
  const t = getChartTheme();

  // 1. Category
  mountChart(
    'ov-cat-chart',
    baseChartOptions({
      chart: { type: 'bar', height: 280 },
      series: [{ name: 'Tổng sản lượng', data: eda.category_sales.series }],
      colors: [t.actualColor],
      plotOptions: { bar: { borderRadius: 5, columnWidth: '52%' } },
      xaxis: { categories: eda.category_sales.categories, labels: { style: { colors: t.textColor } } },
      yaxis: { labels: { formatter: (v) => fmtInt(v), style: { colors: t.textColor } } },
    })
  );

  // 2. Store
  mountChart(
    'ov-store-chart',
    baseChartOptions({
      chart: { type: 'bar', height: 280 },
      series: [{ name: 'Tổng sản lượng', data: eda.store_sales.series }],
      colors: [t.secondaryColor],
      plotOptions: { bar: { borderRadius: 5, columnWidth: '48%' } },
      xaxis: { categories: eda.store_sales.stores, labels: { style: { colors: t.textColor } } },
      yaxis: { labels: { formatter: (v) => fmtInt(v), style: { colors: t.textColor } } },
    })
  );

  // 3. Weather
  mountChart(
    'ov-weather-chart',
    baseChartOptions({
      chart: { type: 'bar', height: 280 },
      series: [{ name: 'Sản lượng TB / ngày', data: eda.weather_sales.series }],
      colors: [t.forecastColor],
      plotOptions: { bar: { borderRadius: 5, horizontal: true, barHeight: '55%' } },
      xaxis: { categories: eda.weather_sales.conditions, labels: { style: { colors: t.textColor } } },
    })
  );

  // 4. Seasonality + Promo
  const labels = [...eda.seasonality_sales.seasons, ...eda.promo_effect.labels];
  const values = [...eda.seasonality_sales.series, ...eda.promo_effect.series];
  mountChart(
    'ov-season-chart',
    baseChartOptions({
      chart: { type: 'bar', height: 280 },
      series: [{ name: 'Sản lượng TB / ngày', data: values }],
      colors: [t.safeColor],
      plotOptions: { bar: { borderRadius: 5, columnWidth: '55%' } },
      xaxis: { categories: labels, labels: { rotate: -15, style: { colors: t.textColor, fontSize: '11px' } } },
    })
  );
}
