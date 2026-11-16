"""
Forecasting Engine: forecasting_models.py
Huấn luyện, đánh giá và so sánh 5 mô hình dự báo chuỗi thời gian:
1. ARIMA
2. SARIMA
3. Holt-Winters Exponential Smoothing
4. Random Forest
5. XGBoost
Phục vụ: Mục 4.1.4 trong Báo cáo (Chuẩn bài mẫu UEL của Thầy Trần Duy Thanh)
"""

import os
import joblib
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import mean_squared_error, mean_absolute_error
from sklearn.ensemble import RandomForestRegressor
from xgboost import XGBRegressor
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.statespace.sarimax import SARIMAX
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from src.config import FIGURES_MODELS_DIR, SAVED_MODELS_DIR

# Style đồ họa và hỗ trợ tiếng Việt
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['font.sans-serif'] = ['Segoe UI', 'Arial', 'DejaVu Sans']
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['figure.dpi'] = 300


def calculate_metrics(y_true, y_pred):
    """Tính toán bộ 3 chỉ số chuẩn: RMSE, MAE, MAPE (%)"""
    y_true, y_pred = np.array(y_true), np.array(y_pred)
    rmse = np.sqrt(mean_squared_error(y_true, y_pred))
    mae = mean_absolute_error(y_true, y_pred)
    # Tránh chia cho 0 khi tính MAPE
    non_zero_mask = y_true != 0
    mape = np.mean(np.abs((y_true[non_zero_mask] - y_pred[non_zero_mask]) / y_true[non_zero_mask])) * 100.0
    return {
        'RMSE': round(float(rmse), 2),
        'MAE': round(float(mae), 2),
        'MAPE (%)': round(float(mape), 2)
    }


def plot_model_result(dates_train, y_train, dates_test, y_test, y_pred, model_name, fig_number):
    """Vẽ biểu đồ Training result của từng mô hình (Tương tự Figure 4.4 - 4.8 bài mẫu UEL)"""
    fig, ax = plt.subplots(figsize=(14, 5.5))
    ax.plot(dates_train, y_train, label='Tập huấn luyện (Train 80%)', color='#2563eb', linewidth=1.2)
    ax.plot(dates_test, y_test, label='Tập kiểm thử thực tế (Actual 20%)', color='#f59e0b', linewidth=1.5)
    ax.plot(dates_test, y_pred, label=f'Dự báo mô hình ({model_name})', color='#16a34a', linestyle='--', linewidth=1.8)

    ax.set_title(f'{fig_number} [Chương 4 - Đánh giá Mô hình]: Kết quả Huấn luyện và Dự báo của Mô hình {model_name}', fontsize=13, fontweight='bold', pad=15)
    ax.set_xlabel('Ngày')
    ax.set_ylabel('Sản lượng bán (Đơn vị)')
    ax.legend(frameon=True, facecolor='white', loc='upper right')
    plt.tight_layout()

    # Tạo tên file tiếng Việt có chú thích chương rõ ràng
    clean_num = fig_number.lower().replace('hình', '').replace('figure', '').replace(':', '').strip().replace('.', '_').replace(' ', '_')
    model_slug = model_name.lower().replace(' ', '_')
    filename_vn = f"chuong_4_hinh_{clean_num}_{model_slug}.png"
    save_path_vn = os.path.join(FIGURES_MODELS_DIR, filename_vn)
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.close()
    return save_path_vn


def train_arima(y_train, y_test, order=(1, 1, 1)):
    """Mô hình 1: ARIMA"""
    print("--> Đang huấn luyện Mô hình 1: ARIMA...")
    model = ARIMA(y_train, order=order)
    model_fit = model.fit()
    y_pred = model_fit.forecast(steps=len(y_test))
    # Chuyển về index dạng DatetimeIndex tương ứng
    y_pred = pd.Series(y_pred.values, index=y_test.index)
    metrics = calculate_metrics(y_test, y_pred)
    return model_fit, y_pred, metrics


def train_sarima(y_train, y_test, order=(1, 1, 1), seasonal_order=(1, 0, 1, 7)):
    """Mô hình 2: SARIMA"""
    print("--> Đang huấn luyện Mô hình 2: SARIMA...")
    model = SARIMAX(y_train, order=order, seasonal_order=seasonal_order, enforce_stationarity=False, enforce_invertibility=False)
    model_fit = model.fit(disp=False)
    y_pred = model_fit.forecast(steps=len(y_test))
    y_pred = pd.Series(y_pred.values, index=y_test.index)
    metrics = calculate_metrics(y_test, y_pred)
    return model_fit, y_pred, metrics


def train_exponential_smoothing(y_train, y_test, seasonal_periods=7):
    """Mô hình 3: Holt-Winters Exponential Smoothing"""
    print("--> Đang huấn luyện Mô hình 3: Exponential Smoothing (Holt-Winters)...")
    # Thay thế giá trị <= 0 để dùng được multiplicative nếu cần, ở đây dùng additive
    model = ExponentialSmoothing(y_train, trend='add', seasonal='add', seasonal_periods=seasonal_periods)
    model_fit = model.fit()
    y_pred = model_fit.forecast(steps=len(y_test))
    y_pred = pd.Series(y_pred.values, index=y_test.index)
    metrics = calculate_metrics(y_test, y_pred)
    return model_fit, y_pred, metrics


def train_random_forest(X_train, y_train, X_test, y_test):
    """Mô hình 4: Random Forest Regressor"""
    print("--> Đang huấn luyện Mô hình 4: Random Forest Regressor...")
    rf = RandomForestRegressor(n_estimators=100, max_depth=12, random_state=42, n_jobs=-1)
    rf.fit(X_train, y_train)
    y_pred = rf.predict(X_test)
    y_pred = pd.Series(y_pred, index=y_test.index)
    metrics = calculate_metrics(y_test, y_pred)
    return rf, y_pred, metrics


def train_xgboost(X_train, y_train, X_test, y_test):
    """Mô hình 5: XGBoost Regressor"""
    print("--> Đang huấn luyện Mô hình 5: XGBoost Regressor...")
    xgb = XGBRegressor(n_estimators=150, learning_rate=0.05, max_depth=6, subsample=0.8, colsample_bytree=0.8, random_state=42)
    xgb.fit(X_train, y_train)
    y_pred = xgb.predict(X_test)
    y_pred = pd.Series(y_pred, index=y_test.index)
    metrics = calculate_metrics(y_test, y_pred)
    return xgb, y_pred, metrics


def plot_metrics_comparison(results_df):
    """Vẽ biểu đồ cột so sánh định lượng RMSE, MAE, MAPE giữa 5 mô hình"""
    fig, axes = plt.subplots(1, 3, figsize=(16, 5))

    metrics_info = [
        ('RMSE', 'Chỉ số RMSE (Sai số chuẩn)', '#2563eb'),
        ('MAE', 'Chỉ số MAE (Sai số tuyệt đối)', '#16a34a'),
        ('MAPE (%)', 'Chỉ số MAPE (%) (Tỷ lệ sai số)', '#dc2626')
    ]

    for i, (col, title, color) in enumerate(metrics_info):
        sns.barplot(data=results_df, x='Model', y=col, ax=axes[i], color=color, edgecolor='black', linewidth=0.5)
        axes[i].set_title(title, fontsize=12, fontweight='bold', pad=10)
        axes[i].set_xlabel('Mô hình')
        axes[i].tick_params(axis='x', rotation=25)
        # Thêm nhãn số trên đỉnh cột
        for p in axes[i].patches:
            axes[i].annotate(f"{p.get_height():.2f}", (p.get_x() + p.get_width() / 2., p.get_height()),
                             ha='center', va='bottom', fontsize=9, fontweight='bold', xytext=(0, 3), textcoords='offset points')

    plt.suptitle('Hình 4.9 [Chương 4 - Đánh giá Mô hình]: So sánh Định lượng Hiệu năng 5 Mô hình Dự báo (RMSE, MAE, MAPE)', fontsize=14, fontweight='bold', y=1.03)
    plt.tight_layout()
    save_path_vn = os.path.join(FIGURES_MODELS_DIR, 'chuong_4_hinh_4_9_so_sanh_hieu_nang_mo_hinh.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.close()
    return save_path_vn


def plot_feature_importance(model, feature_names, top_n=10):
    """Vẽ biểu đồ các đặc trưng quan trọng nhất (Feature Importance) của mô hình cây"""
    if not hasattr(model, 'feature_importances_'):
        return None
    
    importances = model.feature_importances_
    feat_df = pd.DataFrame({'Feature': feature_names, 'Importance': importances})
    feat_df = feat_df.sort_values(by='Importance', ascending=False).head(top_n)

    fig, ax = plt.subplots(figsize=(10, 5.5))
    sns.barplot(data=feat_df, x='Importance', y='Feature', hue='Feature', palette='Blues_r', ax=ax, edgecolor='black', linewidth=0.5, legend=False)
    ax.set_title(f'Hình 4.10 [Chương 4 - Đánh giá Mô hình]: Top {top_n} Đặc trưng Quan trọng nhất trong Dự báo Nhu cầu', fontsize=13, fontweight='bold', pad=12)
    ax.set_xlabel('Điểm số tầm quan trọng tương đối (Importance Score)')
    ax.set_ylabel('Đặc trưng (Features)')
    plt.tight_layout()

    save_path_vn = os.path.join(FIGURES_MODELS_DIR, 'chuong_4_hinh_4_10_tam_quan_trong_dac_trung.png')
    plt.savefig(save_path_vn, bbox_inches='tight')
    plt.close()
    return save_path_vn
