"""
Inventory Optimization Engine: inventory_optimizer.py
Chịu trách nhiệm:
1. Tính toán Tồn kho an toàn (Safety Stock - SS)
2. Điểm đặt hàng lại (Reorder Point - ROP)
3. Số lượng đề xuất nhập hàng (Suggested Order Quantity)
4. Phân loại ma trận ABC - XYZ cho 20 sản phẩm
5. Vẽ biểu đồ mô phỏng tồn kho và ma trận ABC-XYZ
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from src.config import LEAD_TIME_DAYS, Z_SCORE, SERVICE_LEVEL, FIGURES_INVENTORY_DIR

# Style đồ họa và hỗ trợ tiếng Việt
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['font.sans-serif'] = ['Segoe UI', 'Arial', 'DejaVu Sans']
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['figure.dpi'] = 300


def _mode(series):
    """Lấy giá trị xuất hiện nhiều nhất (mode) của 1 cột danh mục"""
    m = series.mode()
    return m.iloc[0] if len(m) else series.iloc[-1]


def z_from_service_level(service_level):
    """Chuyển Service Level (vd: 0.95) sang Z-score tương ứng của phân phối chuẩn"""
    from scipy.stats import norm
    return float(norm.ppf(service_level))


def calculate_inventory_policy(sku_df, forecast_daily_demand=None, lead_time=LEAD_TIME_DAYS, z_score=Z_SCORE,
                               service_level=None):
    """
    Tính toán các chỉ số tồn kho tối ưu cho 1 SKU:
    - Safety Stock: SS = Z * sigma_D * sqrt(L)
    - Reorder Point: ROP = (D_forecast * L) + SS
    Nếu truyền service_level (vd: 0.90/0.95/0.99) thì Z được tính lại từ service_level.
    """
    if service_level is not None:
        z_score = z_from_service_level(service_level)
    else:
        service_level = SERVICE_LEVEL
    recent_demand = sku_df['Units Sold'].tail(90) # Dữ liệu 90 ngày gần nhất
    avg_demand = recent_demand.mean()
    std_demand = recent_demand.std()
    
    # Nhu cầu dự báo (nếu không có thì lấy trung bình gần đây)
    d_forecast = forecast_daily_demand if forecast_daily_demand is not None else avg_demand
    
    # 1. Tồn kho an toàn (Safety Stock)
    safety_stock = z_score * std_demand * np.sqrt(lead_time)
    
    # 2. Điểm đặt hàng lại (Reorder Point)
    lead_time_demand = d_forecast * lead_time
    reorder_point = lead_time_demand + safety_stock
    
    # Mức tồn kho hiện tại (ngày gần nhất)
    current_inventory = float(sku_df['Inventory Level'].iloc[-1])
    
    rop_int = int(np.ceil(reorder_point))

    # 3. Đánh giá trạng thái và Đề xuất lượng đặt hàng
    # Quy tắc thực nghiệm (s, S): khi chưa có chi phí đặt hàng K và chi phí lưu kho h để giải EOQ,
    # đặt mức tồn kho mục tiêu S = 1.5 * ROP (tạo vùng đệm bổ sung 50% phía trên ngưỡng đặt hàng lại).
    if current_inventory <= reorder_point:
        status = "🔴 NGUY CẤP (CẦN ĐẶT HÀNG NGAY)"
        status_code = "CRITICAL"
        target_inventory = reorder_point * 1.5
        suggested_order = max(0, int(np.ceil(target_inventory - current_inventory)))
        suggested_order_rounded = max(0, int(np.ceil(rop_int * 1.5 - current_inventory)))
    elif current_inventory <= reorder_point * 1.2:
        status = "🟡 CẢNH BÁO (SẮP CHẠM NGƯỠNG ROP)"
        status_code = "WARNING"
        suggested_order = 0
        suggested_order_rounded = 0
    else:
        status = "🟢 AN TOÀN (TỒN KHO ĐỦ ĐÁP ỨNG)"
        status_code = "SAFE"
        suggested_order = 0
        suggested_order_rounded = 0

    return {
        'Store ID': sku_df['Store ID'].iloc[-1],
        'Product ID': sku_df['Product ID'].iloc[-1],
        # Category trong dataset thay đổi theo từng dòng -> lấy giá trị phổ biến nhất
        'Category': _mode(sku_df['Category']),
        'Forecast Daily Demand (D_hat)': round(float(d_forecast), 1),
        'Demand Std Dev (sigma_D)': round(float(std_demand), 2),
        'Lead Time (Days)': lead_time,
        'Service Level': f"{round(service_level * 100)}%",
        'Z Score': round(float(z_score), 3),
        'Safety Stock (SS)': int(np.ceil(safety_stock)),
        'Reorder Point (ROP)': rop_int,
        'Reorder Point Unrounded': round(float(reorder_point), 2),
        'Current Inventory': int(current_inventory),
        'Status': status,
        'Status Code': status_code,
        'Suggested Order Qty': suggested_order,
        'Suggested Order Qty (Rounded ROP)': suggested_order_rounded,
    }


def compute_abc_xyz_analysis(df, store_id="S001"):
    """
    Phân tích toàn bộ 20 sản phẩm của 1 cửa hàng theo ma trận ABC - XYZ:
    - ABC: Dựa trên tổng doanh thu (A: 80%, B: 15%, C: 5%)
    - XYZ: Dựa trên Hệ số biến thiên nhu cầu (CV = std / mean)
           (X: CV <= 0.40 ổn định, Y: 0.40 < CV <= 0.70 dao động, Z: CV > 0.70 biến động mạnh)
    """
    store_df = df[df['Store ID'] == store_id].sort_values(by='Date').copy()
    store_df['Revenue'] = store_df['Units Sold'] * store_df['Price'] * (1 - store_df['Discount'] / 100.0)

    # Lưu ý: trong dataset, cột Category của cùng 1 Product ID thay đổi theo từng dòng,
    # nên chỉ group theo Product ID (đúng 20 sản phẩm) và lấy Category phổ biến nhất (mode).
    summary = store_df.groupby('Product ID').agg(
        Category=('Category', _mode),
        Total_Revenue=('Revenue', 'sum'),
        Mean_Demand=('Units Sold', 'mean'),
        Std_Demand=('Units Sold', 'std'),
        Total_Units_Sold=('Units Sold', 'sum'),
        Current_Inventory=('Inventory Level', 'last')
    ).reset_index()

    # 1. Phân loại ABC theo Doanh thu
    summary = summary.sort_values(by='Total_Revenue', ascending=False)
    summary['Revenue_Share'] = summary['Total_Revenue'] / summary['Total_Revenue'].sum()
    summary['Cumulative_Share'] = summary['Revenue_Share'].cumsum()

    def classify_abc(cum_share):
        if cum_share <= 0.80:
            return 'A'
        elif cum_share <= 0.95:
            return 'B'
        else:
            return 'C'

    summary['ABC_Class'] = summary['Cumulative_Share'].apply(classify_abc)

    # 2. Phân loại XYZ theo Hệ số biến thiên (CV)
    # Lưu ý: với dữ liệu mô phỏng Kaggle (CV của 20 SKU tại S001 nằm trong khoảng 0,758–0,834),
    # ngưỡng cố định CV > 0,70 khiến cả 20 mã dồn vào cột Z (suy biến); trong y văn (Syntetos và cộng sự, 2016)
    # ngưỡng biến động mạnh thường đặt ở CV > 1,0 hoặc hiệu chỉnh theo phân phối CV thực tế.
    summary['CV'] = summary['Std_Demand'] / (summary['Mean_Demand'] + 1e-6)

    def classify_xyz(cv):
        if cv <= 0.40:
            return 'X'
        elif cv <= 0.70:
            return 'Y'
        else:
            return 'Z'

    summary['XYZ_Class'] = summary['CV'].apply(classify_xyz)
    summary['ABC_XYZ'] = summary['ABC_Class'] + summary['XYZ_Class']

    return summary


def plot_inventory_simulation(sku_df, policy, window_days=60):
    """
    Vẽ biểu đồ mô phỏng diễn biến tồn kho thực tế đối chiếu với ROP và Safety Stock
    """
    recent_df = sku_df.tail(window_days).copy()
    rop = policy['Reorder Point (ROP)']
    ss = policy['Safety Stock (SS)']

    fig, ax = plt.subplots(figsize=(14, 6))

    ax.plot(recent_df.index, recent_df['Inventory Level'], color='#2563eb', linewidth=2, label='Mức tồn kho thực tế')
    ax.axhline(rop, color='#f59e0b', linestyle='--', linewidth=2, label=f'Điểm đặt hàng lại (ROP = {rop})')
    ax.axhline(ss, color='#dc2626', linestyle=':', linewidth=2.2, label=f'Tồn kho an toàn (SS = {ss})')

    # Vùng nguy cấp (Dưới ROP)
    ax.fill_between(recent_df.index, 0, rop, color='#fee2e2', alpha=0.5, label='Vùng rủi ro thiếu hàng (Dưới ngưỡng ROP)')

    ax.set_title(f'Hình 4.11 [Chương 4 - Tối ưu Tồn kho]: Mô phỏng Biến động Tồn kho Thực tế so với Điểm Đặt hàng lại ROP ({policy["Store ID"]} - {policy["Product ID"]})', fontsize=13, fontweight='bold', pad=15)
    ax.set_xlabel('Ngày')
    ax.set_ylabel('Số lượng tồn kho (Đơn vị)')
    ax.legend(frameon=True, facecolor='white', loc='upper right')
    plt.tight_layout()

    save_path_vn = os.path.join(FIGURES_INVENTORY_DIR, 'chuong_4_hinh_4_11_mo_phong_ton_kho_rop.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.close()
    return save_path_vn


def plot_abc_xyz_matrix(abc_xyz_df):
    """
    Vẽ ma trận nhiệt phân bổ sản phẩm theo 9 nhóm ABC - XYZ
    """
    matrix_counts = pd.crosstab(abc_xyz_df['ABC_Class'], abc_xyz_df['XYZ_Class'], dropna=False)
    # Đảm bảo đủ các nhãn
    for row in ['A', 'B', 'C']:
        if row not in matrix_counts.index:
            matrix_counts.loc[row] = 0
    for col in ['X', 'Y', 'Z']:
        if col not in matrix_counts.columns:
            matrix_counts[col] = 0
    matrix_counts = matrix_counts.loc[['A', 'B', 'C'], ['X', 'Y', 'Z']]

    fig, ax = plt.subplots(figsize=(8, 6))
    sns.heatmap(matrix_counts, annot=True, fmt='d', cmap='YlGnBu', cbar=False, ax=ax,
                linewidths=1, annot_kws={'size': 14, 'weight': 'bold'})

    ax.set_title('Hình 4.12 [Chương 4 - Tối ưu Tồn kho]: Ma trận Phân loại Tồn kho Chiến lược ABC - XYZ (20 Sản phẩm)', fontsize=13, fontweight='bold', pad=15)
    ax.set_xlabel('Nhóm Biến thiên Nhu cầu (X: Ổn định, Y: Dao động, Z: Biến động mạnh)', fontsize=11)
    ax.set_ylabel('Nhóm Đóng góp Doanh thu (A: 80%, B: 15%, C: 5%)', fontsize=11)
    plt.tight_layout()

    save_path_vn = os.path.join(FIGURES_INVENTORY_DIR, 'chuong_4_hinh_4_12_ma_tran_abc_xyz.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.close()
    return save_path_vn
