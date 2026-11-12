"""
Script: eda_analysis.py
Mô tả: Phân tích Khám phá Dữ liệu (EDA) cho tập dữ liệu Bán lẻ & Tồn kho (retail_store_inventory.csv)
Phục vụ: Chương 1 - Data Understanding & Exploratory Data Analysis (theo chuẩn bài mẫu UEL)
"""

import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Fix Unicode print in Windows PowerShell
if sys.platform.startswith('win'):
    sys.stdout.reconfigure(encoding='utf-8')

# Thiết lập phong cách đồ họa chuyên nghiệp, học thuật và hỗ trợ tiếng Việt
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['font.sans-serif'] = ['Segoe UI', 'Arial', 'DejaVu Sans']
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.size'] = 11
plt.rcParams['axes.titlesize'] = 13
plt.rcParams['axes.titleweight'] = 'bold'
plt.rcParams['axes.labelsize'] = 11
plt.rcParams['figure.dpi'] = 300

# Đường dẫn thư mục
DATA_PATH = os.path.join("Dataset", "retail_store_inventory.csv")
OUTPUT_DIR = os.path.join("reports", "figures", "eda")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def load_and_inspect_data():
    print(f"--> Đang đọc dữ liệu từ: {DATA_PATH}...")
    df = pd.read_csv(DATA_PATH)
    df['Date'] = pd.to_datetime(df['Date'])
    
    print(f"Tổng số dòng: {len(df):,}")
    print(f"Tổng số cột: {len(df.columns)}")
    print(f"Khoảng thời gian: {df['Date'].min().strftime('%Y-%m-%d')} đến {df['Date'].max().strftime('%Y-%m-%d')}")
    print(f"Số lượng cửa hàng: {df['Store ID'].nunique()} ({list(df['Store ID'].unique())})")
    print(f"Số lượng sản phẩm: {df['Product ID'].nunique()} ({list(df['Product ID'].unique())})")
    print(f"Các ngành hàng: {list(df['Category'].unique())}")
    
    # Kiểm tra missing values
    missing = df.isnull().sum().sum()
    print(f"Số lượng giá trị khuyết thiếu (Null/NaN): {missing}")
    return df


def plot_figure_1_daily_sales_trend(df):
    """Hình 1.1: Tổng lượng bán theo ngày qua 2 năm kèm đường trung bình trượt"""
    print("--> Đang vẽ Hình 1.1 [Chương 1]: Xu hướng và Tổng Sản lượng Bán Hàng ngày...")
    daily_sales = df.groupby('Date')['Units Sold'].sum().reset_index()
    daily_sales['Rolling_7'] = daily_sales['Units Sold'].rolling(window=7).mean()
    daily_sales['Rolling_30'] = daily_sales['Units Sold'].rolling(window=30).mean()

    fig, ax = plt.subplots(figsize=(14, 6))
    ax.plot(daily_sales['Date'], daily_sales['Units Sold'], color='#94a3b8', alpha=0.5, linewidth=1, label='Sản lượng bán hàng ngày')
    ax.plot(daily_sales['Date'], daily_sales['Rolling_7'], color='#2563eb', linewidth=2, label='Đường trung bình trượt 7 ngày (Tuần)')
    ax.plot(daily_sales['Date'], daily_sales['Rolling_30'], color='#dc2626', linewidth=2.5, linestyle='--', label='Đường xu hướng 30 ngày (Tháng)')

    ax.set_title('Hình 1.1 [Chương 1 - Khám phá Dữ liệu]: Phân tích Xu hướng và Tổng Sản lượng Bán Hàng ngày (2022 - 2024)', pad=15)
    ax.set_xlabel('Ngày')
    ax.set_ylabel('Tổng sản lượng bán (Đơn vị)')
    ax.legend(frameon=True, facecolor='white', loc='upper right')
    plt.tight_layout()
    
    save_path_vn = os.path.join(OUTPUT_DIR, 'chuong_1_hinh_1_1_xu_huong_ban_hang_ngay.png')
    save_path_orig = os.path.join(OUTPUT_DIR, 'figure_1_1_daily_sales_trend.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.savefig(save_path_orig, bbox_inches='tight')
    plt.close()
    print(f"    Đã lưu: {save_path_vn}")


def plot_figure_2_sales_by_category(df):
    """Hình 1.2: Tổng số lượng bán và doanh thu theo 5 ngành hàng"""
    print("--> Đang vẽ Hình 1.2 [Chương 1]: Tổng Sản lượng Bán theo Từng Ngành hàng...")
    df['Revenue'] = df['Units Sold'] * df['Price'] * (1 - df['Discount'] / 100.0)
    cat_summary = df.groupby('Category').agg({
        'Units Sold': 'sum',
        'Revenue': 'sum'
    }).reset_index().sort_values(by='Units Sold', ascending=False)

    colors = ['#1e40af', '#3b82f6', '#60a5fa', '#93c5fd', '#bfdbfe']
    fig, ax1 = plt.subplots(figsize=(10, 5.5))

    bars = ax1.bar(cat_summary['Category'], cat_summary['Units Sold'], color=colors, width=0.55, edgecolor='black', linewidth=0.5)
    ax1.set_title('Hình 1.2 [Chương 1 - Khám phá Dữ liệu]: Tổng Sản lượng Bán theo Từng Danh mục Ngành hàng', pad=15)
    ax1.set_xlabel('Danh mục ngành hàng')
    ax1.set_ylabel('Tổng sản lượng bán (Đơn vị)')

    # Thêm nhãn số lượng trên cột
    for bar in bars:
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                 f"{int(height):,}", ha='center', va='bottom', fontsize=10, fontweight='bold')

    plt.tight_layout()
    save_path_vn = os.path.join(OUTPUT_DIR, 'chuong_1_hinh_1_2_san_luong_theo_nganh_hang.png')
    save_path_orig = os.path.join(OUTPUT_DIR, 'figure_1_2_sales_by_category.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.savefig(save_path_orig, bbox_inches='tight')
    plt.close()
    print(f"    Đã lưu: {save_path_vn}")


def plot_figure_3_sales_by_store_region(df):
    """Hình 1.3: So sánh lượng bán giữa 5 cửa hàng và 4 khu vực"""
    print("--> Đang vẽ Hình 1.3 [Chương 1]: So sánh Sản lượng theo Cửa hàng và Khu vực...")
    store_region = df.groupby(['Store ID', 'Region'])['Units Sold'].sum().reset_index()

    fig, ax = plt.subplots(figsize=(11, 5.5))
    sns.barplot(data=store_region, x='Store ID', y='Units Sold', hue='Region', palette='Blues_d', ax=ax, edgecolor='black', linewidth=0.5)
    
    ax.set_title('Hình 1.3 [Chương 1 - Khám phá Dữ liệu]: So sánh Hiệu quả Bán hàng giữa các Cửa hàng và Khu vực', pad=15)
    ax.set_xlabel('Mã cửa hàng (Store ID)')
    ax.set_ylabel('Tổng sản lượng bán (Đơn vị)')
    ax.legend(title='Khu vực (Region)', frameon=True, facecolor='white')
    plt.tight_layout()

    save_path_vn = os.path.join(OUTPUT_DIR, 'chuong_1_hinh_1_3_san_luong_cua_hang_khu_vuc.png')
    save_path_orig = os.path.join(OUTPUT_DIR, 'figure_1_3_sales_by_store_region.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.savefig(save_path_orig, bbox_inches='tight')
    plt.close()
    print(f"    Đã lưu: {save_path_vn}")


def plot_figure_4_weather_impact(df):
    """Hình 1.4: Tác động của điều kiện thời tiết tới lượng bán"""
    print("--> Đang vẽ Hình 1.4 [Chương 1]: Tác động của Thời tiết đến Sản lượng...")
    df_plot = df.copy()
    weather_labels = {'Sunny': 'Nắng (Sunny)', 'Rainy': 'Mưa (Rainy)', 'Cloudy': 'Nhiều mây (Cloudy)', 'Snowy': 'Có tuyết (Snowy)'}
    df_plot['Weather_VN'] = df_plot['Weather Condition'].map(weather_labels)

    fig, ax = plt.subplots(figsize=(9, 5.5))
    weather_palette = {'Nắng (Sunny)': '#f59e0b', 'Mưa (Rainy)': '#3b82f6', 'Nhiều mây (Cloudy)': '#94a3b8', 'Có tuyết (Snowy)': '#06b6d4'}

    sns.boxplot(data=df_plot, x='Weather_VN', y='Units Sold', hue='Weather_VN', palette=weather_palette, ax=ax, width=0.5,
                legend=False, flierprops={'marker': 'o', 'markersize': 3, 'alpha': 0.3})
    
    ax.set_title('Hình 1.4 [Chương 1 - Khám phá Dữ liệu]: Tác động của Điều kiện Thời tiết đến Sản lượng Bán Hàng ngày', pad=15)
    ax.set_xlabel('Điều kiện thời tiết')
    ax.set_ylabel('Sản lượng bán hàng ngày theo SKU (Đơn vị)')
    plt.tight_layout()

    save_path_vn = os.path.join(OUTPUT_DIR, 'chuong_1_hinh_1_4_tac_dong_thoi_tiet.png')
    save_path_orig = os.path.join(OUTPUT_DIR, 'figure_1_4_weather_impact.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.savefig(save_path_orig, bbox_inches='tight')
    plt.close()
    print(f"    Đã lưu: {save_path_vn}")


def plot_figure_5_seasonality_promotion(df):
    """Hình 1.5: Tác động của Mùa vụ và Khuyến mãi / Lễ"""
    print("--> Đang vẽ Hình 1.5 [Chương 1]: Mùa vụ và Khuyến mãi...")
    season_map = {'Spring': 'Mùa Xuân', 'Summer': 'Mùa Hè', 'Autumn': 'Mùa Thu', 'Winter': 'Mùa Đông'}
    df_plot = df.copy()
    df_plot['Season_VN'] = df_plot['Seasonality'].map(lambda x: season_map.get(x, x))
    
    season_promo = df_plot.groupby(['Season_VN', 'Holiday/Promotion'])['Units Sold'].mean().reset_index()
    season_promo['Trạng thái sự kiện'] = season_promo['Holiday/Promotion'].map({0: 'Ngày bình thường', 1: 'Ngày lễ / Khuyến mãi'})

    fig, ax = plt.subplots(figsize=(10, 5.5))
    sns.barplot(data=season_promo, x='Season_VN', y='Units Sold', hue='Trạng thái sự kiện', palette=['#93c5fd', '#1d4ed8'], ax=ax, edgecolor='black', linewidth=0.5)

    ax.set_title('Hình 1.5 [Chương 1 - Khám phá Dữ liệu]: Sản lượng Bán Trung bình theo Mùa vụ và Chương trình Khuyến mãi', pad=15)
    ax.set_xlabel('Mùa vụ trong năm')
    ax.set_ylabel('Sản lượng bán trung bình mỗi ngày (Đơn vị)')
    ax.legend(title='Trạng thái sự kiện', frameon=True, facecolor='white')
    plt.tight_layout()

    save_path_vn = os.path.join(OUTPUT_DIR, 'chuong_1_hinh_1_5_mua_vu_va_khuyen_mai.png')
    save_path_orig = os.path.join(OUTPUT_DIR, 'figure_1_5_seasonality_promotion.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.savefig(save_path_orig, bbox_inches='tight')
    plt.close()
    print(f"    Đã lưu: {save_path_vn}")


def plot_figure_6_price_elasticity(df):
    """Hình 1.6: Độ co giãn giá - Tương quan giữa Tỷ lệ Giá (Price / Competitor Pricing) và Lượng bán"""
    print("--> Đang vẽ Hình 1.6 [Chương 1]: Độ co giãn giá...")
    sample_df = df.sample(n=2500, random_state=42).copy()
    sample_df['Price_Ratio'] = sample_df['Price'] / sample_df['Competitor Pricing']

    fig, ax = plt.subplots(figsize=(10, 5.5))
    sns.regplot(data=sample_df, x='Price_Ratio', y='Units Sold',
                scatter_kws={'alpha': 0.25, 'color': '#2563eb', 's': 20},
                line_kws={'color': '#dc2626', 'linewidth': 2}, ax=ax)

    ax.set_title('Hình 1.6 [Chương 1 - Khám phá Dữ liệu]: Độ co giãn Giá - Tỷ lệ Giá bán (Cửa hàng / Đối thủ) vs Sản lượng bán', pad=15)
    ax.set_xlabel('Tỷ lệ giá bán (Giá cửa hàng / Giá đối thủ cạnh tranh)')
    ax.set_ylabel('Sản lượng bán (Đơn vị)')
    plt.tight_layout()

    save_path_vn = os.path.join(OUTPUT_DIR, 'chuong_1_hinh_1_6_do_co_gian_gia.png')
    save_path_orig = os.path.join(OUTPUT_DIR, 'figure_1_6_price_elasticity.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.savefig(save_path_orig, bbox_inches='tight')
    plt.close()
    print(f"    Đã lưu: {save_path_vn}")


def plot_figure_7_correlation_heatmap(df):
    """Hình 1.7: Ma trận tương quan giữa các biến số định lượng"""
    print("--> Đang vẽ Hình 1.7 [Chương 1]: Ma trận Tương quan...")
    numeric_cols = [
        'Units Sold', 'Inventory Level', 'Units Ordered',
        'Demand Forecast', 'Price', 'Discount', 'Competitor Pricing'
    ]
    col_rename = {
        'Units Sold': 'Sản lượng bán',
        'Inventory Level': 'Mức tồn kho',
        'Units Ordered': 'Lượng đặt hàng',
        'Demand Forecast': 'Dự báo nhu cầu',
        'Price': 'Giá bán',
        'Discount': 'Chiết khấu (%)',
        'Competitor Pricing': 'Giá đối thủ'
    }
    corr_df = df[numeric_cols].rename(columns=col_rename)
    corr_matrix = corr_df.corr()

    fig, ax = plt.subplots(figsize=(9.5, 8))
    sns.heatmap(corr_matrix, annot=True, fmt='.2f', cmap='Blues', cbar=True, square=True,
                linewidths=0.5, ax=ax, annot_kws={'size': 10, 'weight': 'bold'})

    ax.set_title('Hình 1.7 [Chương 1 - Khám phá Dữ liệu]: Ma trận Tương quan giữa các Biến Định lượng', pad=15)
    plt.tight_layout()

    save_path_vn = os.path.join(OUTPUT_DIR, 'chuong_1_hinh_1_7_ma_tran_tuong_quan.png')
    save_path_orig = os.path.join(OUTPUT_DIR, 'figure_1_7_correlation_heatmap.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.savefig(save_path_orig, bbox_inches='tight')
    plt.close()
    print(f"    Đã lưu: {save_path_vn}")


def main():
    print("=" * 60)
    print("   BẮT ĐẦU QUY TRÌNH PHÂN TÍCH KHÁM PHÁ DỮ LIỆU (EDA)")
    print("=" * 60)
    df = load_and_inspect_data()
    plot_figure_1_daily_sales_trend(df)
    plot_figure_2_sales_by_category(df)
    plot_figure_3_sales_by_store_region(df)
    plot_figure_4_weather_impact(df)
    plot_figure_5_seasonality_promotion(df)
    plot_figure_6_price_elasticity(df)
    plot_figure_7_correlation_heatmap(df)
    print("=" * 60)
    print(f"--> HOÀN THÀNH TẤT CẢ 7 BIỂU ĐỒ! Tất cả ảnh đã được lưu tại: {OUTPUT_DIR}")
    print("=" * 60)


if __name__ == '__main__':
    main()
