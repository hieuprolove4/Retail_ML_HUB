"""
Data Loader & Feature Engineering Module: data_loader.py
Chịu trách nhiệm nạp dữ liệu, tiền xử lý và trích xuất đặc trưng chuỗi thời gian (Feature Engineering).
Phục vụ: Chương 4 - Data Preparation
"""

import pandas as pd
import numpy as np
from src.config import DATA_PATH, TRAIN_RATIO, DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID


def load_raw_data(data_path=DATA_PATH):
    """Đọc dữ liệu thô từ file CSV và chuẩn hóa trường ngày tháng"""
    df = pd.read_csv(data_path)
    df['Date'] = pd.to_datetime(df['Date'])
    df = df.sort_values(by=['Store ID', 'Product ID', 'Date']).reset_index(drop=True)
    return df


def get_sku_data(df, store_id=DEFAULT_STORE_ID, product_id=DEFAULT_PRODUCT_ID):
    """Lọc dữ liệu cho 1 cặp Cửa hàng (Store ID) và Sản phẩm (Product ID) cụ thể"""
    sku_df = df[(df['Store ID'] == store_id) & (df['Product ID'] == product_id)].copy()
    sku_df = sku_df.sort_values(by='Date').set_index('Date')
    
    # Đảm bảo tần suất là hàng ngày liên tục
    sku_df = sku_df.asfreq('D')
    
    # Điền giá trị nội suy cho các trường số nếu có ngày gián đoạn
    numeric_cols = sku_df.select_dtypes(include=[np.number]).columns
    sku_df[numeric_cols] = sku_df[numeric_cols].interpolate(method='linear')
    
    # Điền forward-fill cho các biến danh mục
    cat_cols = sku_df.select_dtypes(include=['object']).columns
    sku_df[cat_cols] = sku_df[cat_cols].ffill().bfill()
    
    return sku_df


def engineer_features(sku_df):
    """
    Tạo các đặc trưng chuỗi thời gian (Feature Engineering) cho Machine Learning:
    - Biến lịch (Calendar features)
    - Biến độ trễ (Lag features)
    - Biến thống kê trượt (Rolling features)
    - Biến tương tác giá & khuyến mãi
    """
    df = sku_df.copy()
    
    # 1. Biến lịch thời gian
    df['dayofweek'] = df.index.dayofweek
    df['day'] = df.index.day
    df['month'] = df.index.month
    df['quarter'] = df.index.quarter
    df['is_weekend'] = df['dayofweek'].isin([5, 6]).astype(int)

    # 2. Biến độ trễ (Lag features) dựa trên nhu cầu thực tế (Units Sold)
    for lag in [1, 2, 7, 14, 21, 30]:
        df[f'lag_{lag}'] = df['Units Sold'].shift(lag)

    # 3. Biến thống kê trượt (Rolling Statistics)
    df['rolling_mean_7'] = df['Units Sold'].shift(1).rolling(window=7).mean()
    df['rolling_std_7'] = df['Units Sold'].shift(1).rolling(window=7).std()
    df['rolling_mean_14'] = df['Units Sold'].shift(1).rolling(window=14).mean()
    df['rolling_std_14'] = df['Units Sold'].shift(1).rolling(window=14).std()

    # 4. Biến giá & cạnh tranh
    df['price_ratio'] = df['Price'] / (df['Competitor Pricing'] + 1e-6)
    df['effective_price'] = df['Price'] * (1 - df['Discount'] / 100.0)

    # 5. Mã hóa biến ngoại cảnh (One-Hot Encoding cho Weather & Seasonality)
    df = pd.get_dummies(df, columns=['Weather Condition', 'Seasonality'], drop_first=True)

    # Loại bỏ các dòng đầu tiên bị NaN do dịch chuyển độ trễ (Lag 30)
    df_clean = df.dropna().copy()
    
    return df_clean


def prepare_ml_data(sku_df, train_ratio=TRAIN_RATIO):
    """
    Chuẩn bị ma trận X, y và phân chia Train / Test theo chuỗi thời gian (không xáo trộn)
    """
    featured_df = engineer_features(sku_df)
    
    target_col = 'Units Sold'
    # Các cột không dùng làm features huấn luyện
    ignore_cols = ['Store ID', 'Product ID', 'Category', 'Region', 'Units Ordered', 'Demand Forecast', target_col]
    feature_cols = [col for col in featured_df.columns if col not in ignore_cols]

    X = featured_df[feature_cols]
    y = featured_df[target_col]

    # Chia train/test theo thứ tự thời gian
    split_idx = int(len(featured_df) * train_ratio)
    
    X_train, X_test = X.iloc[:split_idx], X.iloc[split_idx:]
    y_train, y_test = y.iloc[:split_idx], y.iloc[split_idx:]
    dates_train, dates_test = featured_df.index[:split_idx], featured_df.index[split_idx:]

    return {
        'X_train': X_train, 'X_test': X_test,
        'y_train': y_train, 'y_test': y_test,
        'dates_train': dates_train, 'dates_test': dates_test,
        'feature_names': feature_cols,
        'full_df': featured_df
    }
