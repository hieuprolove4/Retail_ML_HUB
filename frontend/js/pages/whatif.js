/**
 * Page 6: Mô phỏng Kịch bản What-If Chuỗi cung ứng (whatif.js)
 */

import { apiGet, appState, updateState, getOptions } from '../api.js';
import { fmtInt, statusBadge, refreshIcons, renderError } from '../ui.js';
import { baseChartOptions, getChartTheme, mountChart } from '../charts.js';

export async function renderWhatIfPage(container) {
  const opts = await getOptions();

  container.innerHTML = `
    <section class="card">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="sliders-horizontal" style="width:18px;height:18px"></i>
            <span>Mô phỏng Kịch bản What-If (Độ nhạy Thời gian giao hàng & Mức độ dịch vụ)</span>
          </div>
          <div class="card-subtitle">
            Đối chiếu chính sách mặc định (L = 3 ngày, SL = 95%) với kịch bản tùy chỉnh khi thời gian giao hàng thay đổi hoặc nâng mức dịch vụ
          </div>
        </div>
      </div>

      <div class="filter-bar">
        <div class="form-field">
          <label for="wi-store">Cửa hàng (Store ID)</label>
          <select id="wi-store" class="form-control">
            ${opts.stores.map((s) => `<option value="${s}" ${s === appState.store_id ? 'selected' : ''}>Cửa hàng ${s}</option>`).join('')}
          </select>
        </div>

        <div class="form-field">
          <label for="wi-product">Sản phẩm trọng tâm (SKU)</label>
          <select id="wi-product" class="form-control">
            ${opts.products.map((p) => `<option value="${p.id}" ${p.id === appState.product_id ? 'selected' : ''}>${p.id} — ${p.category}</option>`).join('')}
          </select>
        </div>

        <div class="form-field range-wrap">
          <label for="wi-lead">Thời gian giao hàng giả định: <strong id="wi-lead-val">5</strong> ngày (Mặc định: 3 ngày)</label>
          <input id="wi-lead" type="range" min="1" max="14" step="1" value="5" />
        </div>

        <div class="form-field">
          <label for="wi-sl">Mức độ dịch vụ giả định (Service Level)</label>
          <select id="wi-sl" class="form-control">
            <option value="0.90">90% (Tiết kiệm vốn lưu kho)</option>
            <option value="0.95">95% (Cân bằng mặc định)</option>
            <option value="0.98">98% (Ưu tiên chống đứt gãy)</option>
            <option value="0.99" selected>99% (An toàn tối đa)</option>
          </select>
        </div>

        <div class="form-field">
          <label for="wi-source">Nguồn dữ liệu nhu cầu</label>
          <select id="wi-source" class="form-control">
            <option value="model" selected>Dự báo ML (${opts.best_model})</option>
            <option value="average">Trung bình 90 ngày</option>
          </select>
        </div>
      </div>
    </section>

    <!-- Thẻ so sánh trước / sau cho SKU đang chọn -->
    <section class="kpi-grid" aria-label="Kết quả đối chiếu SKU">
      <div class="kpi-card">
        <div class="kpi-top">
          <span>Tồn kho an toàn (Safety Stock)</span>
          <div class="kpi-icon"><i data-lucide="shield-check" style="width:17px;height:17px"></i></div>
        </div>
        <div class="kpi-value" id="wi-ss-diff">—</div>
        <div class="kpi-meta" id="wi-ss-sub">Mặc định: — → Kịch bản: —</div>
      </div>

      <div class="kpi-card">
        <div class="kpi-top">
          <span>Điểm đặt hàng lại (ROP)</span>
          <div class="kpi-icon"><i data-lucide="flag" style="width:17px;height:17px"></i></div>
        </div>
        <div class="kpi-value" id="wi-rop-diff">—</div>
        <div class="kpi-meta" id="wi-rop-sub">Mặc định: — → Kịch bản: —</div>
      </div>

      <div class="kpi-card">
        <div class="kpi-top">
          <span>Lượng đề xuất nhập của SKU</span>
          <div class="kpi-icon"><i data-lucide="truck" style="width:17px;height:17px"></i></div>
        </div>
        <div class="kpi-value" id="wi-order-diff">—</div>
        <div class="kpi-meta" id="wi-order-sub">Mặc định: — → Kịch bản: —</div>
      </div>

      <div class="kpi-card">
        <div class="kpi-top">
          <span>Tổng lượng cần nhập toàn cửa hàng (20 SKU)</span>
          <div class="kpi-icon"><i data-lucide="boxes" style="width:17px;height:17px"></i></div>
        </div>
        <div class="kpi-value" id="wi-store-order">—</div>
        <div class="kpi-meta" id="wi-store-order-sub">So với mặc định: —</div>
      </div>
    </section>

    <!-- Biểu đồ so sánh & Bảng chênh lệch -->
    <div class="grid-2">
      <section class="card">
        <div class="card-header">
          <div>
            <div class="card-title">
              <i data-lucide="bar-chart-2" style="width:18px;height:18px"></i>
              <span id="wi-chart-title">So sánh Chính sách mặc định vs Kịch bản What-If</span>
            </div>
            <div class="card-subtitle">Đối chiếu Tồn hiện tại, Safety Stock, ROP và Lượng đặt hàng đề xuất</div>
          </div>
        </div>
        <div id="wi-sku-chart" style="min-height:290px"></div>
      </section>

      <section class="card">
        <div class="card-header">
          <div>
            <div class="card-title">
              <i data-lucide="list-checks" style="width:18px;height:18px"></i>
              <span>Bảng chênh lệch ROP & Lượng đặt hàng</span>
            </div>
            <div class="card-subtitle">Tác động của kịch bản lên các sản phẩm trong cửa hàng</div>
          </div>
        </div>
        <div class="table-wrap" style="max-height:290px;overflow-y:auto">
          <table class="data-table">
            <thead>
              <tr>
                <th>SKU</th>
                <th class="num">ROP mặc định</th>
                <th class="num">ROP kịch bản</th>
                <th class="num">Đặt mặc định</th>
                <th class="num">Đặt kịch bản</th>
                <th>Trạng thái mới</th>
              </tr>
            </thead>
            <tbody id="wi-table-body"></tbody>
          </table>
        </div>
      </section>
    </div>
  `;

  refreshIcons(container);

  const storeSel = container.querySelector('#wi-store');
  const prodSel = container.querySelector('#wi-product');
  const leadInp = container.querySelector('#wi-lead');
  const leadVal = container.querySelector('#wi-lead-val');
  const slSel = container.querySelector('#wi-sl');
  const srcSel = container.querySelector('#wi-source');

  const runSim = async () => {
    leadVal.textContent = leadInp.value;
    updateState({ store_id: storeSel.value, product_id: prodSel.value });
    await executeSimulation(container, {
      store_id: storeSel.value,
      product_id: prodSel.value,
      lead_time: Number(leadInp.value),
      service_level: Number(slSel.value),
      demand_source: srcSel.value,
    });
  };

  storeSel.addEventListener('change', runSim);
  prodSel.addEventListener('change', runSim);
  leadInp.addEventListener('input', () => {
    leadVal.textContent = leadInp.value;
  });
  leadInp.addEventListener('change', runSim);
  slSel.addEventListener('change', runSim);
  srcSel.addEventListener('change', runSim);

  await runSim();
}

async function executeSimulation(container, params) {
  try {
    const [baseData, simData] = await Promise.all([
      apiGet('/api/inventory-status', {
        store_id: params.store_id,
        lead_time: 3,
        service_level: 0.95,
        demand_source: 'model',
      }),
      apiGet('/api/inventory-status', {
        store_id: params.store_id,
        lead_time: params.lead_time,
        service_level: params.service_level,
        demand_source: params.demand_source,
      }),
    ]);

    const baseItem = baseData.items.find((i) => i.product_id === params.product_id) || baseData.items[0];
    const simItem = simData.items.find((i) => i.product_id === params.product_id) || simData.items[0];

    const fmtDelta = (a, b) => {
      const d = b - a;
      const sign = d > 0 ? '+' : '';
      return `${fmtInt(b)} (${sign}${fmtInt(d)})`;
    };

    container.querySelector('#wi-ss-diff').textContent = fmtDelta(baseItem.safety_stock, simItem.safety_stock);
    container.querySelector('#wi-ss-sub').textContent = `Mặc định: ${fmtInt(baseItem.safety_stock)} → Kịch bản: ${fmtInt(simItem.safety_stock)}`;

    container.querySelector('#wi-rop-diff').textContent = fmtDelta(baseItem.reorder_point, simItem.reorder_point);
    container.querySelector('#wi-rop-sub').textContent = `Mặc định: ${fmtInt(baseItem.reorder_point)} → Kịch bản: ${fmtInt(simItem.reorder_point)}`;

    container.querySelector('#wi-order-diff').textContent = fmtDelta(baseItem.suggested_order, simItem.suggested_order);
    container.querySelector('#wi-order-sub').textContent = `Mặc định: ${fmtInt(baseItem.suggested_order)} → Kịch bản: ${fmtInt(simItem.suggested_order)}`;

    const totalBase = baseData.items.reduce((s, x) => s + x.suggested_order, 0);
    const totalSim = simData.items.reduce((s, x) => s + x.suggested_order, 0);
    container.querySelector('#wi-store-order').textContent = `${fmtInt(totalSim)} đv`;
    container.querySelector('#wi-store-order-sub').textContent = `Mặc định (L=3, 95%): ${fmtInt(totalBase)} đv (Chênh: ${totalSim - totalBase >= 0 ? '+' : ''}${fmtInt(totalSim - totalBase)})`;

    container.querySelector('#wi-chart-title').textContent =
      `So sánh SKU ${simItem.product_id} (${simItem.category}) — Mặc định (L=3, 95%) vs Kịch bản (L=${params.lead_time}, ${Math.round(params.service_level * 100)}%)`;

    const t = getChartTheme();
    mountChart(
      'wi-sku-chart',
      baseChartOptions({
        chart: { type: 'bar', height: 290 },
        series: [
          {
            name: 'Chính sách mặc định (L=3, SL=95%)',
            data: [baseItem.current_inventory, baseItem.safety_stock, baseItem.reorder_point, baseItem.suggested_order],
          },
          {
            name: `Kịch bản What-If (L=${params.lead_time}, SL=${Math.round(params.service_level * 100)}%)`,
            data: [simItem.current_inventory, simItem.safety_stock, simItem.reorder_point, simItem.suggested_order],
          },
        ],
        colors: [t.actualColor, t.forecastColor],
        plotOptions: { bar: { borderRadius: 5, columnWidth: '52%' } },
        xaxis: {
          categories: ['Tồn hiện tại', 'Safety Stock (SS)', 'Điểm đặt lại (ROP)', 'Đề xuất nhập'],
          labels: { style: { colors: t.textColor } },
        },
      })
    );

    const baseMap = Object.fromEntries(baseData.items.map((x) => [x.product_id, x]));
    const tbody = container.querySelector('#wi-table-body');
    tbody.innerHTML = simData.items
      .slice(0, 12)
      .map((it) => {
        const b = baseMap[it.product_id];
        return `
          <tr>
            <td style="font-family:var(--font-mono);font-weight:700">${it.product_id}</td>
            <td class="num">${fmtInt(b.reorder_point)}</td>
            <td class="num"><strong>${fmtInt(it.reorder_point)}</strong></td>
            <td class="num">${fmtInt(b.suggested_order)}</td>
            <td class="num"><strong>${fmtInt(it.suggested_order)}</strong></td>
            <td>${statusBadge(it.status_code)}</td>
          </tr>
        `;
      })
      .join('');
    refreshIcons(tbody);
  } catch (err) {
    renderError(container, err);
  }
}
