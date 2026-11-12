"""
Master Script: main_train_and_evaluate.py
Thực thi toàn bộ quy trình:
1. Nạp dữ liệu và kiểm định chuỗi thời gian (ADF, KPSS, ACF/PACF, Rolling Stats)
2. Huấn luyện & đánh giá 5 mô hình dự báo (ARIMA, SARIMA, Holt-Winters, Random Forest, XGBoost)
3. Xuất toàn bộ biểu đồ Thực tế vs Dự báo (Figures 4.4 - 4.8) và Bảng so sánh (Table 4.2)
4. Lưu mô hình xuất sắc nhất ra file .pkl
5. Tính toán Tồn kho an toàn (SS), Điểm đặt hàng lại (ROP), Ma trận ABC-XYZ
6. Xuất báo cáo kết quả chi tiết
"""

import os
import sys
import json
import joblib
import pandas as pd
import numpy as np

# Fix Unicode console output on Windows
if sys.platform.startswith('win'):
    sys.stdout.reconfigure(encoding='utf-8')

from src.config import (
    DATA_PATH, SAVED_MODELS_DIR, REPORTS_DIR,
    DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID, LEAD_TIME_DAYS, Z_SCORE
)
from src.data_loader import load_raw_data, get_sku_data, prepare_ml_data
from src.time_series_diagnostics import (
    plot_rolling_statistics, run_stationarity_tests,
    plot_acf_pacf, plot_seasonal_decomposition
)
from src.forecasting_models import (
    train_arima, train_sarima, train_exponential_smoothing,
    train_random_forest, train_xgboost,
    plot_model_result, plot_metrics_comparison, plot_feature_importance
)
from src.inventory_optimizer import (
    calculate_inventory_policy, compute_abc_xyz_analysis,
    plot_inventory_simulation, plot_abc_xyz_matrix
)


def main():
    print("=" * 75)
    print("   HỆ THỐNG DỰ BÁO NHU CẦU & TỐI ƯU TỒN KHO BÁN LẺ (RETAIL ML HUB)")
    print("   Quy trình Huấn luyện Mô hình & Đánh giá Hiệu năng chuẩn đồ án UEL")
    print("=" * 75)

    # BƯỚC 1: NẠP DỮ LIỆU GỐC
    print("\n[BƯỚC 1] Nạp dữ liệu và chuẩn bị chuỗi thời gian SKU...")
    raw_df = load_raw_data(DATA_PATH)
    sku_df = get_sku_data(raw_df, store_id=DEFAULT_STORE_ID, product_id=DEFAULT_PRODUCT_ID)
    ts = sku_df['Units Sold']
    print(f"    Đã nạp SKU: Cửa hàng = {DEFAULT_STORE_ID}, Sản phẩm = {DEFAULT_PRODUCT_ID}")
    print(f"    Khoảng thời gian: {sku_df.index.min().strftime('%Y-%m-%d')} đến {sku_df.index.max().strftime('%Y-%m-%d')} ({len(sku_df)} ngày)")

    # BƯỚC 2: CHẨN ĐOÁN VÀ KIỂM ĐỊNH CHUỖI THỜI GIAN
    print("\n[BƯỚC 2] Kiểm định tính dừng & Chẩn đoán chuỗi thời gian (Mục 4.1.2)...")
    plot_rolling_statistics(ts, window=7)
    run_stationarity_tests(ts)
    plot_acf_pacf(ts)
    plot_seasonal_decomposition(ts)

    # BƯỚC 3: TIỀN XỬ LÝ & TRÍCH XUẤT ĐẶC TRƯNG CHO MACHINE LEARNING
    print("\n[BƯỚC 3] Kỹ thuật trích xuất đặc trưng (Feature Engineering) & Chia Train/Test...")
    ml_data = prepare_ml_data(sku_df)
    X_train, X_test = ml_data['X_train'], ml_data['X_test']
    y_train, y_test = ml_data['y_train'], ml_data['y_test']
    dates_train, dates_test = ml_data['dates_train'], ml_data['dates_test']
    print(f"    Số lượng mẫu Train (80%): {len(X_train)} ngày ({dates_train.min().strftime('%Y-%m-%d')} -> {dates_train.max().strftime('%Y-%m-%d')})")
    print(f"    Số lượng mẫu Test  (20%): {len(X_test)} ngày ({dates_test.min().strftime('%Y-%m-%d')} -> {dates_test.max().strftime('%Y-%m-%d')})")
    print(f"    Số lượng đặc trưng (Features): {len(ml_data['feature_names'])}")

    # BƯỚC 4: HUẤN LUYỆN VÀ ĐÁNH GIÁ 5 MÔ HÌNH DỰ BÁO
    print("\n[BƯỚC 4] Huấn luyện tuần tự 5 mô hình và tính toán chỉ số RMSE, MAE, MAPE...")
    results = []
    models_dict = {}

    # 1. ARIMA
    m1, pred1, met1 = train_arima(y_train, y_test, order=(2, 1, 1))
    plot_model_result(dates_train, y_train, dates_test, y_test, pred1, "ARIMA", "Hình 4.4")
    results.append({'Model': 'ARIMA', **met1})
    models_dict['ARIMA'] = {'model': m1, 'pred': pred1, 'metrics': met1}

    # 2. Holt-Winters Exponential Smoothing
    m2, pred2, met2 = train_exponential_smoothing(y_train, y_test, seasonal_periods=7)
    plot_model_result(dates_train, y_train, dates_test, y_test, pred2, "Exponential Smoothing", "Hình 4.5")
    results.append({'Model': 'Exponential Smoothing', **met2})
    models_dict['Exponential Smoothing'] = {'model': m2, 'pred': pred2, 'metrics': met2}

    # 3. SARIMA
    m3, pred3, met3 = train_sarima(y_train, y_test, order=(1, 1, 1), seasonal_order=(1, 0, 1, 7))
    plot_model_result(dates_train, y_train, dates_test, y_test, pred3, "SARIMA", "Hình 4.6")
    results.append({'Model': 'SARIMA', **met3})
    models_dict['SARIMA'] = {'model': m3, 'pred': pred3, 'metrics': met3}

    # 4. Random Forest
    m4, pred4, met4 = train_random_forest(X_train, y_train, X_test, y_test)
    plot_model_result(dates_train, y_train, dates_test, y_test, pred4, "Random Forest", "Hình 4.7")
    results.append({'Model': 'Random Forest', **met4})
    models_dict['Random Forest'] = {'model': m4, 'pred': pred4, 'metrics': met4}

    # 5. XGBoost
    m5, pred5, met5 = train_xgboost(X_train, y_train, X_test, y_test)
    plot_model_result(dates_train, y_train, dates_test, y_test, pred5, "XGBoost", "Hình 4.8")
    results.append({'Model': 'XGBoost', **met5})
    models_dict['XGBoost'] = {'model': m5, 'pred': pred5, 'metrics': met5}

    # BƯỚC 5: TỔNG HỢP VÀ SO SÁNH ĐỊNH LƯỢNG (TABLE 4.2)
    results_df = pd.DataFrame(results)
    print("\n" + "=" * 60)
    print("   TABLE 4.2: SO SÁNH CÁC CHỈ SỐ ĐÁNH GIÁ MÔ HÌNH (EVALUATION METRICS)")
    print("=" * 60)
    print(results_df.to_string(index=False))
    print("=" * 60)

    # Xuất đồ thị so sánh định lượng và feature importance
    plot_metrics_comparison(results_df)
    plot_feature_importance(m5, ml_data['feature_names'], top_n=10)

    # Xác định mô hình tốt nhất (dựa trên RMSE thấp nhất)
    best_row = results_df.sort_values(by='RMSE').iloc[0]
    best_model_name = best_row['Model']
    print(f"\n--> MÔ HÌNH TỐI ƯU NHẤT LÀ: {best_model_name} (RMSE: {best_row['RMSE']}, MAE: {best_row['MAE']}, MAPE: {best_row['MAPE (%)']}%)")

    # Lưu toàn bộ 5 mô hình ra file .pkl để Web App có thể tùy chọn load từng model
    print("\n--> Đang lưu toàn bộ 5 mô hình vào thư mục saved_models/...")
    model_filenames = {
        'ARIMA': 'arima_model.pkl',
        'SARIMA': 'sarima_model.pkl',
        'Exponential Smoothing': 'exponential_smoothing_model.pkl',
        'Random Forest': 'random_forest_model.pkl',
        'XGBoost': 'xgboost_model.pkl'
    }
    for m_name, m_info in models_dict.items():
        fname = model_filenames.get(m_name, f"{m_name.lower().replace(' ', '_')}.pkl")
        m_path = os.path.join(SAVED_MODELS_DIR, fname)
        joblib.dump(m_info['model'], m_path)
        print(f"    * Đã lưu mô hình {m_name:25} -> {fname}")

    # Đồng thời lưu bản sao Best Model
    best_model_obj = models_dict[best_model_name]['model']
    best_model_path = os.path.join(SAVED_MODELS_DIR, "best_forecasting_model.pkl")
    joblib.dump(best_model_obj, best_model_path)
    print(f"    * Đã lưu Best Model ({best_model_name}) -> best_forecasting_model.pkl")
    
    # Lưu metadata của mô hình
    metadata = {
        'best_model': best_model_name,
        'metrics': dict(best_row),
        'store_id': DEFAULT_STORE_ID,
        'product_id': DEFAULT_PRODUCT_ID,
        'features': ml_data['feature_names'],
        'all_results': results
    }
    with open(os.path.join(SAVED_MODELS_DIR, "model_metadata.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=4, ensure_ascii=False)
    print(f"    Đã lưu mô hình tốt nhất tại: {best_model_path}")

    # BƯỚC 6: TỐI ƯU HÓA TỒN KHO & CẢNH BÁO (INVENTORY OPTIMIZATION)
    print("\n[BƯỚC 6] Tính toán Tối ưu hóa Tồn kho (Safety Stock, ROP, ABC-XYZ)...")
    # Lấy giá trị dự báo trung bình của mô hình tốt nhất trên tập test
    forecast_avg = float(models_dict[best_model_name]['pred'].mean())
    policy = calculate_inventory_policy(sku_df, forecast_daily_demand=forecast_avg, lead_time=LEAD_TIME_DAYS, z_score=Z_SCORE)
    
    print("\n--- BÁO CÁO THÔNG SỐ TỒN KHO CHO SKU ---")
    for k, v in policy.items():
        print(f"  * {k:35}: {v}")
    
    plot_inventory_simulation(sku_df, policy, window_days=60)

    # Phân tích ma trận ABC - XYZ cho toàn bộ 20 sản phẩm
    print("\n--- PHÂN TÍCH MA TRẬN TỒN KHO ABC - XYZ (20 SẢN PHẨM) ---")
    abc_xyz_df = compute_abc_xyz_analysis(raw_df, store_id=DEFAULT_STORE_ID)
    plot_abc_xyz_matrix(abc_xyz_df)
    
    abc_xyz_csv_path = os.path.join(REPORTS_DIR, "abc_xyz_analysis_s001.csv")
    try:
        abc_xyz_df.to_csv(abc_xyz_csv_path, index=False)
        print(f"    Đã lưu bảng phân tích ABC-XYZ tại: {abc_xyz_csv_path}")
    except PermissionError:
        backup_csv_path = os.path.join(REPORTS_DIR, "abc_xyz_analysis_s001_latest.csv")
        abc_xyz_df.to_csv(backup_csv_path, index=False)
        print(f"    [Thông báo] File {abc_xyz_csv_path} đang mở trong ứng dụng khác. Đã lưu tại: {backup_csv_path}")

    print("\n" + "=" * 75)
    print("   HOÀN THÀNH TOÀN BỘ QUY TRÌNH ML & TỐI ƯU TỒN KHO!")
    print("   Tất cả biểu đồ đã được lưu vào thư mục reports/figures/")
    print("=" * 75)


if __name__ == '__main__':
    main()
