/**
 * Page 5: Ma trận Phân loại Tồn kho Chiến lược ABC - XYZ (abcxyz.js)
 * Phân tích đúng 20 sản phẩm của cửa hàng theo tỷ trọng doanh thu (ABC) và hệ số biến thiên CV (XYZ).
 */

import { apiGet, appState, updateState, getOptions } from '../api.js';
import { fmtDec, fmtUSD, refreshIcons, exportCSV, renderError } from '../ui.js';
import { baseChartOptions, getChartTheme, mountChart } from '../charts.js';

let abcData = null;
let selectedCell = 'ALL';

export async function renderAbcXyzPage(container) {
  const opts = await getOptions();

  container.innerHTML = `
    <section class="card">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="grid-3x3" style="width:18px;height:18px"></i>
            <span>Ma trận Phân loại Tồn kho Chiến lược ABC – XYZ (Hình 4.12)</span>
          </div>
          <div class="card-subtitle">
            ABC theo doanh thu tích lũy (A: ≤80%, B: 80–95%, C: &gt;95%) • XYZ theo hệ số biến thiên nhu cầu CV = σ/μ (X: ≤0.40, Y: 0.40–0.70, Z: &gt;0.70)
          </div>
        </div>

        <div style="display:flex;gap:10px;align-items:center">
          <select id="abc-store" class="form-control" style="width:auto" aria-label="Chọn cửa hàng">
            ${opts.stores.map((s) => `<option value="${s}" ${s === appState.store_id ? 'selected' : ''}>Cửa hàng ${s}</option>`).join('')}
          </select>
          <button type="button" class="btn btn-outline" id="abc-reset-btn">Hiện tất cả (20 SKU)</button>
        </div>
      </div>

      <div class="grid-asymmetric">
        <!-- Ma trận 3x3 tương tác -->
        <div>
          <div class="matrix-grid" id="abc-matrix-grid"></div>
          <div class="callout" style="margin-top:12px">
            <div class="callout-title">
              <i data-lucide="compass" style="width:16px;height:16px"></i>
              <span id="abc-strategy-title">Khuyến nghị chiến lược quản trị tồn kho</span>
            </div>
            <div id="abc-strategy-body">Nhấn vào một ô trong ma trận 3×3 để xem khuyến nghị chiến lược và lọc danh sách sản phẩm tương ứng.</div>
          </div>
        </div>

        <!-- Biểu đồ phân tán Doanh thu vs CV -->
        <div>
          <div class="card-title" style="margin-bottom:6px">
            <i data-lucide="scatter-chart" style="width:17px;height:17px"></i>
            <span>Phân bố Doanh thu (USD) & Hệ số biến thiên (CV)</span>
          </div>
          <div id="abc-scatter-chart" style="min-height:290px"></div>
        </div>
      </div>
    </section>

    <!-- Bảng 20 sản phẩm -->
    <section class="card">
      <div class="card-header">
        <div>
          <div class="card-title">
            <i data-lucide="table" style="width:18px;height:18px"></i>
            <span id="abc-table-title">Danh sách phân loại 20 mã sản phẩm</span>
          </div>
          <div class="card-subtitle">Đã khắc phục lỗi trùng lặp ngành hàng: tổng hợp đúng 20 Product ID duy nhất</div>
        </div>
        <button type="button" class="btn btn-outline" id="abc-export-btn">
          <i data-lucide="download" style="width:16px;height:16px"></i>
          <span>Xuất CSV</span>
        </button>
      </div>

      <div class="table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th>Mã SKU</th>
              <th>Ngành hàng chính</th>
              <th class="num">Tổng doanh thu</th>
              <th class="num">Tỷ trọng DT (%)</th>
              <th class="num">Cầu TB / ngày</th>
              <th class="num">Hệ số biến thiên (CV)</th>
              <th>Nhóm ABC</th>
              <th>Nhóm XYZ</th>
              <th>Phân loại ABC-XYZ</th>
            </tr>
          </thead>
          <tbody id="abc-table-body"></tbody>
        </table>
      </div>
    </section>
  `;

  refreshIcons(container);

  const storeSel = container.querySelector('#abc-store');
  storeSel.addEventListener('change', async () => {
    updateState({ store_id: storeSel.value });
    selectedCell = 'ALL';
    await loadAbcXyz(container);
  });

  container.querySelector('#abc-reset-btn').addEventListener('click', () => {
    selectedCell = 'ALL';
    renderMatrixAndTable(container);
  });

  container.querySelector('#abc-export-btn').addEventListener('click', () => {
    if (!abcData) return;
    const rows = abcData.items.map((it) => [
      abcData.store_id,
      it.product_id,
      it.category,
      it.total_revenue,
      it.revenue_share,
      it.mean_demand,
      it.cv,
      it.abc,
      it.xyz,
      it.group,
    ]);
    exportCSV(
      `abc_xyz_${abcData.store_id}.csv`,
      ['Cửa hàng', 'Mã SKU', 'Ngành hàng', 'Tổng doanh thu', 'Tỷ trọng (%)', 'Nhu cầu TB', 'CV', 'ABC', 'XYZ', 'Nhóm'],
      rows
    );
  });

  await loadAbcXyz(container);
}

async function loadAbcXyz(container) {
  try {
    abcData = await apiGet('/api/abc-xyz', { store_id: appState.store_id });
    renderMatrixAndTable(container);
    renderScatter(abcData);
  } catch (err) {
    renderError(container, err, () => loadAbcXyz(container));
  }
}

function renderMatrixAndTable(container) {
  const grid = container.querySelector('#abc-matrix-grid');
  const rows = ['A', 'B', 'C'];
  const cols = ['X', 'Y', 'Z'];
  const rowDesc = { A: 'A (80% DT)', B: 'B (15% DT)', C: 'C (5% DT)' };

  const cellTone = (r, c) => {
    if (r === 'A' && c === 'Z') return 'var(--color-critical-bg)';
    if (r === 'A') return 'var(--color-info-bg)';
    if (r === 'B') return 'var(--color-warning-bg)';
    return 'var(--color-surface-elevated)';
  };

  let html = `
    <div class="matrix-header">ABC \\ XYZ</div>
    <div class="matrix-header">X (Ổn định)<br/><small>CV ≤ 0.40</small></div>
    <div class="matrix-header">Y (Dao động)<br/><small>0.40 &lt; CV ≤ 0.70</small></div>
    <div class="matrix-header">Z (Biến động)<br/><small>CV &gt; 0.70</small></div>
  `;

  for (const r of rows) {
    html += `<div class="matrix-header">${rowDesc[r]}</div>`;
    for (const c of cols) {
      const code = `${r}${c}`;
      const count = abcData.matrix[r][c];
      const active = selectedCell === code ? 'active' : '';
      html += `
        <div class="matrix-cell ${active}" data-cell="${code}" style="background:${cellTone(r, c)}">
          <div class="cell-code">Nhóm ${code}</div>
          <div class="cell-count">${count}</div>
          <div class="cell-sub">sản phẩm</div>
        </div>
      `;
    }
  }
  grid.innerHTML = html;

  grid.querySelectorAll('.matrix-cell').forEach((cell) => {
    cell.addEventListener('click', () => {
      const code = cell.dataset.cell;
      selectedCell = selectedCell === code ? 'ALL' : code;
      renderMatrixAndTable(container);
    });
  });

  // Cập nhật hộp chiến lược
  const stTitle = container.querySelector('#abc-strategy-title');
  const stBody = container.querySelector('#abc-strategy-body');
  if (selectedCell === 'ALL') {
    stTitle.textContent = 'Tổng quan phân loại ABC – XYZ (Tất cả 20 sản phẩm)';
    stBody.textContent = 'Nhấn vào một ô bất kỳ trong ma trận để xem chiến lược tồn kho khuyến nghị cho nhóm đó.';
  } else {
    stTitle.textContent = `Chiến lược cho Nhóm ${selectedCell}`;
    stBody.textContent = abcData.strategy[selectedCell] || '';
  }

  // Bảng dữ liệu
  const items =
    selectedCell === 'ALL'
      ? abcData.items
      : abcData.items.filter((it) => it.group === selectedCell);

  container.querySelector('#abc-table-title').textContent =
    selectedCell === 'ALL'
      ? `Danh sách phân loại 20 mã sản phẩm — Cửa hàng ${abcData.store_id}`
      : `Danh sách sản phẩm Nhóm ${selectedCell} (${items.length} SKU)`;

  const tbody = container.querySelector('#abc-table-body');
  if (!items.length) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;padding:18px;color:var(--color-muted-foreground)">Không có sản phẩm nào thuộc nhóm ${selectedCell}.</td></tr>`;
    return;
  }

  tbody.innerHTML = items
    .map(
      (it) => `
      <tr>
        <td style="font-family:var(--font-mono);font-weight:700">${it.product_id}</td>
        <td>${it.category}</td>
        <td class="num">${fmtUSD(it.total_revenue)}</td>
        <td class="num">${fmtDec(it.revenue_share)}%</td>
        <td class="num">${fmtDec(it.mean_demand)}</td>
        <td class="num">${it.cv.toFixed(3)}</td>
        <td><span class="badge ${it.abc === 'A' ? 'badge-INFO' : it.abc === 'B' ? 'badge-WARNING' : 'badge-SAFE'}"> Hạng ${it.abc}</span></td>
        <td><span class="badge badge-CRITICAL">Nhóm ${it.xyz}</span></td>
        <td><strong style="font-family:var(--font-mono)">${it.group}</strong></td>
      </tr>
    `
    )
    .join('');
}

function renderScatter(data) {
  const t = getChartTheme();
  const byClass = { A: [], B: [], C: [] };
  data.items.forEach((it) => {
    byClass[it.abc].push({ x: Number((it.total_revenue / 1000).toFixed(1)), y: it.cv, sku: it.product_id });
  });

  mountChart(
    'abc-scatter-chart',
    baseChartOptions({
      chart: { type: 'scatter', height: 290, zoom: { enabled: true, type: 'xy' } },
      series: [
        { name: 'Nhóm A (80% DT)', data: byClass.A.map((p) => [p.x, p.y]) },
        { name: 'Nhóm B (15% DT)', data: byClass.B.map((p) => [p.x, p.y]) },
        { name: 'Nhóm C (5% DT)', data: byClass.C.map((p) => [p.x, p.y]) },
      ],
      colors: [t.actualColor, t.forecastColor, t.safeColor],
      markers: { size: 7 },
      xaxis: {
        title: { text: 'Tổng doanh thu (Nghìn USD)', style: { color: t.textColor } },
        tickAmount: 6,
      },
      yaxis: {
        title: { text: 'Hệ số biến thiên nhu cầu (CV)', style: { color: t.textColor } },
      },
    })
  );
}
