"""
Script: time_series_diagnostics.py
Mô tả: Kiểm định tính dừng (Stationarity: ADF, KPSS), Hàm tự tương quan (ACF/PACF),
       và Phân rã chuỗi thời gian (Decomposition).
Phục vụ: Mục 4.1.2 & 4.1.3 trong Báo cáo (theo chuẩn bài mẫu UEL của Thầy Trần Duy Thanh)
"""

import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from statsmodels.tsa.stattools import adfuller, kpss
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf
from statsmodels.tsa.seasonal import seasonal_decompose

# Fix Unicode stdout
if sys.platform.startswith('win'):
    sys.stdout.reconfigure(encoding='utf-8')

# Style biểu đồ và hỗ trợ tiếng Việt
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['font.sans-serif'] = ['Segoe UI', 'Arial', 'DejaVu Sans']
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.size'] = 11
plt.rcParams['axes.titlesize'] = 13
plt.rcParams['axes.titleweight'] = 'bold'
plt.rcParams['axes.labelsize'] = 11
plt.rcParams['figure.dpi'] = 300

DATA_PATH = os.path.join("Dataset", "retail_store_inventory.csv")
OUTPUT_DIR = os.path.join("reports", "figures", "diagnostics")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def load_sku_series(store_id="S001", product_id="P0001"):
    """Lấy chuỗi thời gian Units Sold theo ngày của 1 SKU cụ thể"""
    df = pd.read_csv(DATA_PATH)
    df['Date'] = pd.to_datetime(df['Date'])
    
    sku_df = df[(df['Store ID'] == store_id) & (df['Product ID'] == product_id)].copy()
    sku_df = sku_df.sort_values(by='Date').set_index('Date')
    
    # Đảm bảo tần suất là hàng ngày ('D')
    ts = sku_df['Units Sold'].asfreq('D')
    ts = ts.interpolate(method='linear')
    return ts, sku_df


def plot_rolling_statistics(ts, window=7):
    """Hình 4.3: Rolling Mean & Standard Deviation (Tương tự Figure 4.3 bài mẫu UEL)"""
    print(f"--> Đang vẽ Hình 4.3 [Chương 4]: Giá trị Trung bình & Độ lệch chuẩn trượt (window={window})...")
    rolling_mean = ts.rolling(window=window).mean()
    rolling_std = ts.rolling(window=window).std()

    fig, ax = plt.subplots(figsize=(14, 6))
    ax.plot(ts.index, ts.values, color='#3b82f6', alpha=0.6, linewidth=1.2, label='Nhu cầu gốc (Sản lượng bán thực tế)')
    ax.plot(ts.index, rolling_mean, color='#dc2626', linewidth=2, label=f'Trung bình trượt ({window} ngày)')
    ax.plot(ts.index, rolling_std, color='#1e293b', linewidth=2, linestyle='--', label=f'Độ lệch chuẩn trượt ({window} ngày)')

    ax.set_title(f'Hình 4.3 [Chương 4 - Kiểm định Chuỗi]: Giá trị Trung bình Trượt và Độ lệch Chuẩn Trượt (Cửa hàng S001 - Sản phẩm P0001)', pad=15)
    ax.set_xlabel('Ngày')
    ax.set_ylabel('Sản lượng bán (Đơn vị)')
    ax.legend(frameon=True, facecolor='white', loc='upper right')
    plt.tight_layout()

    save_path_vn = os.path.join(OUTPUT_DIR, 'chuong_4_hinh_4_3_trung_binh_do_lech_truot.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.close()
    print(f"    Đã lưu: {save_path_vn}")


def run_stationarity_tests(ts):
    """Thực hiện ADF test và KPSS test (Tương tự phần ADF and KPSS test trang 28 bài mẫu UEL)"""
    print("\n" + "=" * 60)
    print("   KẾT QUẢ KIỂM ĐỊNH TÍNH DỪNG (STATIONARITY TESTS)")
    print("=" * 60)

    # 1. ADF Test (Null hypothesis: Chuỗi có unit root -> phi dừng)
    adf_res = adfuller(ts.dropna())
    print(f"[1] Augmented Dickey-Fuller (ADF) Test:")
    print(f"    - ADF Statistic : {adf_res[0]:.4f}")
    print(f"    - p-value       : {adf_res[1]:.4e}")
    print(f"    - Lags Used     : {adf_res[2]}")
    print(f"    - Observations  : {adf_res[3]}")
    print(f"    - Critical Values:")
    for key, val in adf_res[4].items():
        print(f"        {key}: {val:.4f}")
    
    # 2. KPSS Test (Null hypothesis: Chuỗi có tính dừng quanh mức trung bình)
    kpss_res = kpss(ts.dropna(), regression='c', nlags="auto")

    # Kết luận & Lưu kết quả ra CSV phục vụ Chương 4
    adf_conclusion = (
        "Chuỗi có tính dừng (Stationary, p-value < 0.05 -> bác bỏ H0)"
        if adf_res[1] < 0.05
        else "Chuỗi có tính phi dừng (Non-stationary, p-value >= 0.05 -> chưa bác bỏ H0)"
    )
    kpss_conclusion = (
        "Chuỗi là dừng (Stationary, p-value >= 0.05 -> không bác bỏ H0)"
        if kpss_res[1] >= 0.05
        else "Chuỗi là phi dừng (Non-stationary, p-value < 0.05 -> bác bỏ H0)"
    )

    if adf_res[1] < 0.05:
        print("    => Kết luận ADF: p-value < 0.05, bác bỏ giả thuyết H0. Chuỗi có tính dừng (Stationary)!")
    else:
        print("    => Kết luận ADF: p-value >= 0.05, chưa đủ cơ sở bác bỏ H0. Chuỗi có tính phi dừng (Non-stationary)!")

    print(f"\n[2] Kwiatkowski-Phillips-Schmidt-Shin (KPSS) Test:")
    print(f"    - KPSS Statistic: {kpss_res[0]:.4f}")
    print(f"    - p-value       : {kpss_res[1]:.4f}")
    print(f"    - Truncation Lag: {kpss_res[2]}")
    print(f"    - Critical Values:")
    for key, val in kpss_res[3].items():
        print(f"        {key}: {val:.4f}")

    if kpss_res[1] >= 0.05:
        print("    => Kết luận KPSS: p-value >= 0.05, không thể bác bỏ H0. Chuỗi là dừng (Stationary)!")
    else:
        print("    => Kết luận KPSS: p-value < 0.05, bác bỏ H0. Chuỗi là phi dừng (Non-stationary)!")

    rows = [
        {
            "Test": "ADF (Augmented Dickey-Fuller)",
            "Statistic": round(float(adf_res[0]), 4),
            "p_value": float(adf_res[1]),
            "Lags_Used": int(adf_res[2]),
            "Observations": int(adf_res[3]),
            "Critical_1pct": round(float(adf_res[4]["1%"]), 4),
            "Critical_5pct": round(float(adf_res[4]["5%"]), 4),
            "Critical_10pct": round(float(adf_res[4]["10%"]), 4),
            "Conclusion": adf_conclusion,
        },
        {
            "Test": "KPSS (Kwiatkowski-Phillips-Schmidt-Shin)",
            "Statistic": round(float(kpss_res[0]), 4),
            "p_value": round(float(kpss_res[1]), 4),
            "Lags_Used": int(kpss_res[2]),
            "Observations": int(len(ts.dropna())),
            "Critical_1pct": round(float(kpss_res[3]["1%"]), 4),
            "Critical_5pct": round(float(kpss_res[3]["5%"]), 4),
            "Critical_10pct": round(float(kpss_res[3]["10%"]), 4),
            "Conclusion": kpss_conclusion,
        },
    ]

    reports_dir = "reports"
    os.makedirs(reports_dir, exist_ok=True)
    csv_path = os.path.join(reports_dir, "stationarity_tests_s001_p0001.csv")
    pd.DataFrame(rows).to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"--> Đã xuất bảng kết quả kiểm định ra: {csv_path}")
    print("=" * 60 + "\n")

    return {"adf": rows[0], "kpss": rows[1], "csv_path": csv_path}


def plot_acf_pacf(ts):
    """Hình 4.2: Hàm tự tương quan ACF và Tự tương quan riêng phần PACF"""
    print("--> Đang vẽ Hình 4.2 [Chương 4]: Phân tích ACF & PACF...")
    fig, axes = plt.subplots(1, 2, figsize=(15, 5))

    plot_acf(ts.dropna(), lags=30, ax=axes[0], color='#2563eb', vlines_kwargs={'colors': '#2563eb'})
    axes[0].set_title('Hàm tự tương quan (ACF)', pad=12)
    axes[0].set_xlabel('Độ trễ (Lag)')
    axes[0].set_ylabel('Hệ số tự tương quan')

    plot_pacf(ts.dropna(), lags=30, ax=axes[1], color='#dc2626', vlines_kwargs={'colors': '#dc2626'}, method='ywm')
    axes[1].set_title('Hàm tự tương quan riêng phần (PACF)', pad=12)
    axes[1].set_xlabel('Độ trễ (Lag)')
    axes[1].set_ylabel('Hệ số tự tương quan riêng phần')

    plt.suptitle('Hình 4.2 [Chương 4 - Kiểm định Chuỗi]: Phân tích Tự tương quan (ACF) và Tự tương quan Riêng phần (PACF)', fontsize=14, fontweight='bold', y=1.02)
    plt.tight_layout()

    save_path_vn = os.path.join(OUTPUT_DIR, 'chuong_4_hinh_4_2_tu_tuong_quan_acf_pacf.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.close()
    print(f"    Đã lưu: {save_path_vn}")


def plot_seasonal_decomposition(ts):
    """Hình 4.1: Phân rã chuỗi thời gian (Observed, Trend, Seasonal, Residual)"""
    print("--> Đang vẽ Hình 4.1 [Chương 4]: Phân rã Chuỗi Thời gian (STL)...")
    decomp = seasonal_decompose(ts.dropna(), model='additive', period=7)

    fig, axes = plt.subplots(4, 1, figsize=(14, 9), sharex=True)
    decomp.observed.plot(ax=axes[0], color='#1e293b', linewidth=1.2)
    axes[0].set_ylabel('Dữ liệu gốc\n(Observed)')
    axes[0].set_title('Hình 4.1 [Chương 4 - Kiểm định Chuỗi]: Phân rã Chuỗi Thời gian Nhu cầu Hàng ngày (Chu kỳ 7 ngày)', pad=10)

    decomp.trend.plot(ax=axes[1], color='#2563eb', linewidth=2)
    axes[1].set_ylabel('Xu hướng\n(Trend)')

    decomp.seasonal.plot(ax=axes[2], color='#16a34a', linewidth=1.2)
    axes[2].set_ylabel('Mùa vụ (7 ngày)\n(Seasonal)')

    decomp.resid.plot(ax=axes[3], color='#dc2626', marker='o', markersize=2, linestyle='none')
    axes[3].axhline(0, color='black', linestyle='--', linewidth=0.8)
    axes[3].set_ylabel('Phần dư\n(Residual)')
    axes[3].set_xlabel('Ngày')

    plt.tight_layout()
    save_path_vn = os.path.join(OUTPUT_DIR, 'chuong_4_hinh_4_1_phan_ra_chuoi_thoi_gian.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.close()
    print(f"    Đã lưu: {save_path_vn}")


def main():
    print("=" * 60)
    print("   BẮT ĐẦU KIỂM ĐỊNH CHUỖI THỜI GIAN & TÍNH DỪNG (CHAPTER 4)")
    print("=" * 60)
    ts, _ = load_sku_series(store_id="S001", product_id="P0001")
    plot_rolling_statistics(ts, window=7)
    run_stationarity_tests(ts)
    plot_acf_pacf(ts)
    plot_seasonal_decomposition(ts)
    print("=" * 60)
    print(f"--> HOÀN THÀNH KIỂM ĐỊNH! Các biểu đồ chẩn đoán đã lưu tại: {OUTPUT_DIR}")
    print("=" * 60)


if __name__ == '__main__':
    main()
