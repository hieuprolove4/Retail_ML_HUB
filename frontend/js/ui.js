/**
 * UI Helpers (ui.js): Định dạng số vi-VN, tạo badge trạng thái, xuất CSV, thông báo toast
 */

const nfInt = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 0 });
const nfDec = new Intl.NumberFormat('vi-VN', { minimumFractionDigits: 1, maximumFractionDigits: 2 });
const nfCur = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });

export function fmtInt(n) {
  if (n === null || n === undefined || Number.isNaN(Number(n))) return '—';
  return nfInt.format(Number(n));
}

export function fmtDec(n) {
  if (n === null || n === undefined || Number.isNaN(Number(n))) return '—';
  return nfDec.format(Number(n));
}

export function fmtUSD(n) {
  if (n === null || n === undefined || Number.isNaN(Number(n))) return '—';
  return nfCur.format(Number(n));
}

export function weekdayVN(dateStr) {
  const d = new Date(`${dateStr}T00:00:00`);
  const names = ['Chủ nhật', 'Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7'];
  return names[d.getDay()] || '';
}

export function statusBadge(statusCode) {
  const map = {
    CRITICAL: { label: 'Nguy cấp (Cần đặt ngay)', icon: 'alert-octagon', cls: 'badge-CRITICAL' },
    WARNING: { label: 'Cảnh báo (Sắp chạm ROP)', icon: 'alert-triangle', cls: 'badge-WARNING' },
    SAFE: { label: 'An toàn (Đủ đáp ứng)', icon: 'check-circle-2', cls: 'badge-SAFE' },
  };
  const info = map[statusCode] || { label: statusCode, icon: 'info', cls: 'badge-INFO' };
  return `<span class="badge ${info.cls}"><i data-lucide="${info.icon}" style="width:13px;height:13px"></i><span>${info.label}</span></span>`;
}

export function refreshIcons(root = document) {
  if (window.lucide && typeof window.lucide.createIcons === 'function') {
    window.lucide.createIcons({ attrs: { 'stroke-width': 2 }, nameAttr: 'data-lucide', root });
  }
}

export function showToast(message, duration = 3200) {
  const container = document.getElementById('toast-container');
  if (!container) return;
  const el = document.createElement('div');
  el.className = 'toast';
  el.textContent = message;
  container.appendChild(el);
  setTimeout(() => {
    el.remove();
  }, duration);
}

export function exportCSV(filename, headers, rows) {
  const escapeCell = (v) => {
    const s = v === null || v === undefined ? '' : String(v);
    if (s.includes(',') || s.includes('"') || s.includes('\n')) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  };
  const lines = [
    headers.map(escapeCell).join(','),
    ...rows.map((r) => r.map(escapeCell).join(',')),
  ];
  const blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  showToast(`Đã xuất file ${filename}`);
}

export function renderError(container, err, retryFn) {
  container.innerHTML = `
    <div class="error-banner" role="alert">
      <div><strong>Không thể tải dữ liệu:</strong> ${err.message || err}</div>
      ${retryFn ? '<button class="btn btn-outline" id="retry-btn">Thử lại</button>' : ''}
    </div>
  `;
  if (retryFn) {
    container.querySelector('#retry-btn')?.addEventListener('click', retryFn);
  }
}
