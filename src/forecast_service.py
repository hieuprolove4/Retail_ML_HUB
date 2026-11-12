"""
Forecast Service: forecast_service.py
Dịch vụ dự báo thực tế phục vụ Web API (backend/api.py):
- Đánh giá (Train 80% / Test 20%) và dự báo tương lai cho TỪNG SKU (Store x Product)
- Random Forest / XGBoost: dự báo đệ quy nhiều bước (recursive multi-step) dựa trên lag & rolling features
- ARIMA / SARIMA / Holt-Winters: fit lại trên toàn bộ chuỗi của SKU rồi dự báo tiếp từ ngày cuối dữ liệu
- Kết quả được lưu đệm (cache) trong RAM, KHÔNG ghi đè các file .pkl trong saved_models/

Với SKU dùng trong báo cáo (S001 - P0001), mô hình Random Forest / XGBoost được nạp từ file .pkl đã huấn luyện,
nên chỉ số đánh giá khớp với Bảng 4.2 trong báo cáo.
"""

import os
import json
import threading
import warnings

import matplotlib
matplotlib.use('Agg')  # Không mở cửa sổ đồ họa khi chạy trong server

import joblib
import numpy as np
import pandas as pd
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.statespace.sarimax import SARIMAX
from statsmodels.tsa.holtwinters import ExponentialSmoothing

from src.config import SAVED_MODELS_DIR, DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID
from src.data_loader import get_sku_data, engineer_features, prepare_ml_data
from src.forecasting_models import (
    calculate_metrics, train_arima, train_sarima, train_exponential_smoothing,
    train_random_forest, train_xgboost
)

warnings.filterwarnings('ignore')

# Danh sách mô hình hỗ trợ và tham số giống hệt quy trình huấn luyện trong báo cáo
MODEL_NAMES = ['Random Forest', 'XGBoost', 'SARIMA', 'ARIMA', 'Exponential Smoothing']
ML_MODELS = {'Random Forest', 'XGBoost'}
ARIMA_ORDER = (2, 1, 1)
SARIMA_ORDER = (1, 1, 1)
SARIMA_SEASONAL_ORDER = (1, 0, 1, 7)
HW_SEASONAL_PERIODS = 7
REPORT_SKU = (DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID)

MODEL_FILES = {
    'Random Forest': 'random_forest_model.pkl',
    'XGBoost': 'xgboost_model.pkl',
}

Z_95 = 1.96  # Hệ số cho khoảng tin cậy 95%


class ForecastService:
    """Quản lý dữ liệu SKU, đánh giá mô hình và dự báo tương lai có bộ nhớ đệm"""

    def __init__(self, raw_df):
        self.raw_df = raw_df
        self._sku_cache = {}
        self._ml_data_cache = {}
        self._eval_cache = {}
        self._forecast_cache = {}
        self._locks = {}
        self._global_lock = threading.Lock()
        self.report_metadata = self._load_report_metadata()

    # ------------------------------------------------------------------ #
    # Tiện ích nội bộ
    # ------------------------------------------------------------------ #
    def _lock_for(self, key):
        with self._global_lock:
            if key not in self._locks:
                self._locks[key] = threading.Lock()
            return self._locks[key]

    @staticmethod
    def _load_report_metadata():
        path = os.path.join(SAVED_MODELS_DIR, 'model_metadata.json')
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}

    def best_model_name(self):
        return self.report_metadata.get('best_model', 'Random Forest')

    def get_sku(self, store_id, product_id):
        key = (store_id, product_id)
        if key not in self._sku_cache:
            self._sku_cache[key] = get_sku_data(self.raw_df, store_id=store_id, product_id=product_id)
        return self._sku_cache[key]

    def get_ml_data(self, store_id, product_id):
        key = (store_id, product_id)
        if key not in self._ml_data_cache:
            self._ml_data_cache[key] = prepare_ml_data(self.get_sku(store_id, product_id))
        return self._ml_data_cache[key]

    def _load_report_ml_model(self, model_name, feature_names):
        """Nạp mô hình .pkl của báo cáo; trả về None nếu không dùng được"""
        fname = MODEL_FILES.get(model_name)
        if not fname:
            return None
        path = os.path.join(SAVED_MODELS_DIR, fname)
        if not os.path.exists(path):
            return None
        try:
            model = joblib.load(path)
            saved_features = list(getattr(model, 'feature_names_in_', feature_names))
            if saved_features != list(feature_names):
                return None
            return model
        except Exception:
            return None

    # ------------------------------------------------------------------ #
    # 1. ĐÁNH GIÁ MÔ HÌNH (Train 80% / Test 20%)
    # ------------------------------------------------------------------ #
    def evaluate(self, store_id, product_id, model_name):
        """
        Trả về dict: model (fit trên tập train), metrics, test_dates, y_test, y_pred, source
        source = 'report' (dùng mô hình/kết quả của báo cáo) hoặc 'on-demand' (huấn luyện ngay cho SKU này)
        """
        key = (store_id, product_id, model_name)
        if key in self._eval_cache:
            return self._eval_cache[key]

        with self._lock_for(('eval',) + key):
            if key in self._eval_cache:
                return self._eval_cache[key]

            ml = self.get_ml_data(store_id, product_id)
            X_train, X_test = ml['X_train'], ml['X_test']
            y_train, y_test = ml['y_train'], ml['y_test']
            is_report_sku = (store_id, product_id) == REPORT_SKU
            source = 'on-demand'

            if model_name in ML_MODELS:
                model = self._load_report_ml_model(model_name, ml['feature_names']) if is_report_sku else None
                if model is not None:
                    y_pred = pd.Series(model.predict(X_test), index=y_test.index)
                    metrics = calculate_metrics(y_test, y_pred)
                    source = 'report'
                elif model_name == 'Random Forest':
                    model, y_pred, metrics = train_random_forest(X_train, y_train, X_test, y_test)
                else:
                    model, y_pred, metrics = train_xgboost(X_train, y_train, X_test, y_test)
            elif model_name == 'ARIMA':
                model, y_pred, metrics = train_arima(y_train, y_test, order=ARIMA_ORDER)
            elif model_name == 'SARIMA':
                model, y_pred, metrics = train_sarima(y_train, y_test, order=SARIMA_ORDER,
                                                      seasonal_order=SARIMA_SEASONAL_ORDER)
            elif model_name == 'Exponential Smoothing':
                model, y_pred, metrics = train_exponential_smoothing(y_train, y_test,
                                                                     seasonal_periods=HW_SEASONAL_PERIODS)
            else:
                raise ValueError(f"Mô hình không hỗ trợ: {model_name}")

            # SKU của báo cáo: các mô hình thống kê được huấn luyện lại với cùng tham số -> cùng kết quả báo cáo
            if is_report_sku and model_name not in ML_MODELS:
                source = 'report'

            result = {
                'model': model,
                'metrics': metrics,
                'test_dates': list(y_test.index),
                'y_test': [float(v) for v in y_test.values],
                'y_pred': [float(v) for v in np.asarray(y_pred)],
                'source': source,
                'feature_names': ml['feature_names'],
                'feature_dtypes': X_train.dtypes.to_dict(),
            }
            self._eval_cache[key] = result
            return result

    # ------------------------------------------------------------------ #
    # 2. DỰ BÁO TƯƠNG LAI
    # ------------------------------------------------------------------ #
    @staticmethod
    def _future_exogenous(sku_df):
        """
        Giả định giá trị các biến ngoại sinh trong tương lai (chưa biết trước):
        - Price, Discount, Competitor Pricing, Inventory Level: trung bình 14 ngày gần nhất
        - Holiday/Promotion: 0 (không có khuyến mãi)
        - Weather Condition, Seasonality: giá trị phổ biến nhất 30 ngày gần nhất
        """
        r14, r30 = sku_df.tail(14), sku_df.tail(30)
        return {
            'Price': float(r14['Price'].mean()),
            'Discount': float(r14['Discount'].mean()),
            'Competitor Pricing': float(r14['Competitor Pricing'].mean()),
            'Inventory Level': float(r14['Inventory Level'].mean()),
            'Holiday/Promotion': 0,
            'Weather Condition': r30['Weather Condition'].mode().iloc[0],
            'Seasonality': r30['Seasonality'].mode().iloc[0],
        }

    @staticmethod
    def assumptions_text(model_name):
        if model_name in ML_MODELS:
            return [
                "Dự báo đệ quy: giá trị dự báo của ngày trước được dùng làm biến trễ (lag) cho ngày sau.",
                "Giá bán, chiết khấu, giá đối thủ và mức tồn kho tương lai = trung bình 14 ngày gần nhất.",
                "Không giả định khuyến mãi/ngày lễ trong kỳ dự báo.",
                "Thời tiết và mùa vụ = giá trị phổ biến nhất trong 30 ngày gần nhất.",
                "Khoảng tin cậy 95% ước lượng bằng ±1.96 × RMSE trên tập kiểm thử.",
            ]
        if model_name == 'Exponential Smoothing':
            return [
                "Mô hình Holt-Winters (xu hướng cộng, mùa vụ cộng, chu kỳ 7 ngày) được fit lại trên toàn bộ lịch sử của SKU.",
                "Khoảng tin cậy 95% ước lượng bằng ±1.96 × RMSE trên tập kiểm thử.",
            ]
        return [
            f"Mô hình {model_name} được fit lại trên toàn bộ lịch sử của SKU (cùng tham số trong báo cáo).",
            "Khoảng tin cậy 95% lấy trực tiếp từ phân phối dự báo của mô hình.",
        ]

    def _forecast_ml(self, store_id, product_id, model_name, horizon):
        ev = self.evaluate(store_id, product_id, model_name)
        model = ev['model']
        feature_names = ev['feature_names']
        dtypes = ev['feature_dtypes']

        hist = self.get_sku(store_id, product_id).copy()
        last_date = hist.index.max()
        exo = self._future_exogenous(hist)

        preds = []
        for i in range(1, horizon + 1):
            new_date = last_date + pd.Timedelta(days=i)
            new_row = hist.iloc[[-1]].copy()
            new_row.index = pd.DatetimeIndex([new_date])
            for col, val in exo.items():
                new_row[col] = val
            # Giá trị tạm, không ảnh hưởng features của chính ngày này (lag/rolling đều dùng shift >= 1)
            new_row['Units Sold'] = 0.0
            hist = pd.concat([hist, new_row])

            feats = engineer_features(hist)
            x = feats.iloc[[-1]].reindex(columns=feature_names, fill_value=0)
            x = x.astype(dtypes)
            y_hat = max(0.0, float(model.predict(x)[0]))
            hist.loc[new_date, 'Units Sold'] = y_hat
            preds.append(y_hat)

        rmse = ev['metrics']['RMSE']
        lower = [max(0.0, p - Z_95 * rmse) for p in preds]
        upper = [p + Z_95 * rmse for p in preds]
        return preds, lower, upper

    def _forecast_ts(self, store_id, product_id, model_name, horizon):
        ts = self.get_sku(store_id, product_id)['Units Sold'].astype(float)
        if model_name == 'ARIMA':
            res = ARIMA(ts, order=ARIMA_ORDER).fit()
        elif model_name == 'SARIMA':
            res = SARIMAX(ts, order=SARIMA_ORDER, seasonal_order=SARIMA_SEASONAL_ORDER,
                          enforce_stationarity=False, enforce_invertibility=False).fit(disp=False)
        else:
            res = ExponentialSmoothing(ts, trend='add', seasonal='add',
                                       seasonal_periods=HW_SEASONAL_PERIODS).fit()

        if model_name in ('ARIMA', 'SARIMA'):
            fc = res.get_forecast(steps=horizon)
            preds = [float(v) for v in fc.predicted_mean.values]
            ci = fc.conf_int(alpha=0.05).values
            lower = [max(0.0, float(v)) for v in ci[:, 0]]
            upper = [float(v) for v in ci[:, 1]]
        else:
            preds = [float(v) for v in res.forecast(horizon).values]
            rmse = self.evaluate(store_id, product_id, model_name)['metrics']['RMSE']
            lower = [max(0.0, p - Z_95 * rmse) for p in preds]
            upper = [p + Z_95 * rmse for p in preds]
        preds = [max(0.0, p) for p in preds]
        return preds, lower, upper

    def forecast(self, store_id, product_id, model_name, horizon):
        """Dự báo `horizon` ngày tiếp theo sau ngày cuối cùng của dữ liệu"""
        key = (store_id, product_id, model_name, horizon)
        if key in self._forecast_cache:
            return self._forecast_cache[key]

        with self._lock_for(('fc',) + key):
            if key in self._forecast_cache:
                return self._forecast_cache[key]

            if model_name in ML_MODELS:
                preds, lower, upper = self._forecast_ml(store_id, product_id, model_name, horizon)
            else:
                preds, lower, upper = self._forecast_ts(store_id, product_id, model_name, horizon)

            last_date = self.get_sku(store_id, product_id).index.max()
            dates = [last_date + pd.Timedelta(days=i) for i in range(1, horizon + 1)]
            result = {'dates': dates, 'values': preds, 'lower': lower, 'upper': upper}
            self._forecast_cache[key] = result
            return result

    def forecast_daily_demand(self, store_id, product_id, days, model_name=None):
        """Nhu cầu bình quân/ngày trong `days` ngày tới (dùng cho tính ROP)"""
        model_name = model_name or self.best_model_name()
        fc = self.forecast(store_id, product_id, model_name, max(1, int(days)))
        return float(np.mean(fc['values']))

    # ------------------------------------------------------------------ #
    # 3. FEATURE IMPORTANCE (mô hình của báo cáo - SKU S001-P0001)
    # ------------------------------------------------------------------ #
    def feature_importance(self, model_name, top_n=10):
        ev = self.evaluate(REPORT_SKU[0], REPORT_SKU[1], model_name)
        model = ev['model']
        if not hasattr(model, 'feature_importances_'):
            return {'features': [], 'importance': []}
        imp = pd.Series(model.feature_importances_, index=ev['feature_names'])
        imp = imp.sort_values(ascending=False).head(top_n)
        return {'features': list(imp.index), 'importance': [round(float(v), 4) for v in imp.values]}
