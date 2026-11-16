/**
 * ApexCharts Wrapper (charts.js)
 * Quản lý vòng đời biểu đồ và tự động đồng bộ màu theo Light / Dark mode.
 */

const ACTIVE_CHARTS = new Map();

export function getChartTheme() {
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  return {
    isDark,
    mode: isDark ? 'dark' : 'light',
    textColor: isDark ? '#94A3B8' : '#475569',
    gridColor: isDark ? '#1E293B' : '#E2E8F0',
    actualColor: isDark ? '#60A5FA' : '#2563EB',
    forecastColor: isDark ? '#F59E0B' : '#D97706',
    safeColor: isDark ? '#4ADE80' : '#16A34A',
    criticalColor: isDark ? '#F87171' : '#DC2626',
    secondaryColor: isDark ? '#818CF8' : '#4F46E5',
    bgSurface: isDark ? '#111827' : '#FFFFFF',
  };
}

export function baseChartOptions(overrides = {}) {
  const t = getChartTheme();
  const {
    chart: ovChart,
    grid: ovGrid,
    dataLabels: ovDataLabels,
    tooltip: ovTooltip,
    legend: ovLegend,
    ...rest
  } = overrides;

  return {
    ...rest,
    chart: {
      fontFamily: "'Inter', sans-serif",
      background: 'transparent',
      foreColor: t.textColor,
      toolbar: {
        show: true,
        tools: { download: true, selection: false, zoom: true, zoomin: true, zoomout: true, pan: false, reset: true },
      },
      animations: { enabled: true, easing: 'easeinout', speed: 320 },
      ...(ovChart || {}),
    },
    theme: { mode: t.mode },
    grid: {
      borderColor: t.gridColor,
      strokeDashArray: 3,
      ...(ovGrid || {}),
    },
    dataLabels: { enabled: false, ...(ovDataLabels || {}) },
    tooltip: {
      theme: t.mode,
      ...(ovTooltip || {}),
    },
    legend: {
      position: 'top',
      horizontalAlign: 'right',
      labels: { colors: t.textColor },
      ...(ovLegend || {}),
    },
  };
}

export function mountChart(elementId, options) {
  const el = document.getElementById(elementId);
  if (!el) return null;
  if (ACTIVE_CHARTS.has(elementId)) {
    try {
      ACTIVE_CHARTS.get(elementId).destroy();
    } catch (_) {}
    ACTIVE_CHARTS.delete(elementId);
  }
  el.innerHTML = '';
  if (typeof window.ApexCharts === 'undefined') {
    el.innerHTML = `<div class="error-banner">Thư viện biểu đồ ApexCharts chưa được tải. Hãy nhấn Ctrl + F5 để làm mới trang.</div>`;
    return null;
  }
  try {
    const chart = new window.ApexCharts(el, options);
    const p = chart.render();
    if (p && typeof p.catch === 'function') {
      p.catch((err) => {
        console.error('ApexCharts render error:', err);
        el.innerHTML = `<div class="error-banner">Lỗi hiển thị biểu đồ: ${err.message || err}</div>`;
      });
    }
    ACTIVE_CHARTS.set(elementId, chart);
    return chart;
  } catch (err) {
    console.error('ApexCharts init error:', err);
    el.innerHTML = `<div class="error-banner">Lỗi khởi tạo biểu đồ: ${err.message || err}</div>`;
    return null;
  }
}

export function destroyAllCharts() {
  for (const [, c] of ACTIVE_CHARTS.entries()) {
    try {
      c.destroy();
    } catch (_) {}
  }
  ACTIVE_CHARTS.clear();
}
