r"""
extra_diagnostics.py
Script độc lập bổ sung các kiểm định, mô hình cơ sở và đối chứng loại bỏ đặc trưng / đa bước
(KHÔNG sửa pipeline huấn luyện chính trong src/forecasting_models.py hay src/forecast_service.py):
  1. Baseline Naive (y_train[-1]) và Seasonal Naive (chu kỳ 7 ngày cuối train) trên 141 ngày test S001–P0001.
  2. Đối chứng loại bỏ biến Inventory Level (27 đặc trưng) cho Random Forest và XGBoost.
  3. Đối chứng dự báo đệ quy thuần 141 bước (không dùng biến ngoại sinh tương lai) cho RF và XGBoost.
  4. Kiểm định Ljung-Box (lag 7, 14) trên phần dư của ARIMA(2,1,1) và SARIMA(1,1,1)(1,0,1,7).
  5. Kiểm định Diebold-Mariano (hiệu chỉnh Harvey, sai số MSE và MAE) giữa các cặp mô hình.
  6. Đối chiếu công thức tồn kho SS, ROP (777 / 721 / 707) và Q (943 vs 942).
Xuất kết quả ra reports/extra_diagnostics_s001_p0001.json và .csv.
"""

import json
import os
import sys

import numpy as np
import pandas as pd
from scipy import stats
from statsmodels.stats.diagnostic import acorr_ljungbox
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.statespace.sarimax import SARIMAX

if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.config import DEFAULT_PRODUCT_ID, DEFAULT_STORE_ID, LEAD_TIME_DAYS, REPORTS_DIR, SERVICE_LEVEL
from src.data_loader import get_sku_data, load_raw_data, prepare_ml_data
from src.forecast_service import ForecastService
from src.forecasting_models import calculate_metrics, train_random_forest, train_xgboost
from src.inventory_optimizer import calculate_inventory_policy, compute_abc_xyz_analysis


def diebold_mariano_test(e1, e2, h=1, power=2):
    """Kiểm định Diebold-Mariano với hiệu chỉnh Harvey-Leybourne-Newbold (1997)."""
    e1 = np.asarray(e1, dtype=float)
    e2 = np.asarray(e2, dtype=float)
    d = (np.abs(e1) ** power) - (np.abs(e2) ** power)
    n = len(d)
    d_bar = float(np.mean(d))
    gamma0 = float(np.var(d, ddof=0))
    gamma_sum = gamma0
    for k in range(1, h):
        cov_k = float(np.sum((d[k:] - d_bar) * (d[:-k] - d_bar)) / n)
        gamma_sum += 2.0 * cov_k
    var_d = max(gamma_sum / n, 1e-12)
    dm_stat = d_bar / np.sqrt(var_d)
    harvey_adj = np.sqrt((n + 1.0 - 2.0 * h + (h * (h - 1.0)) / n) / n)
    dm_adj = dm_stat * harvey_adj
    p_value = 2.0 * stats.t.sf(np.abs(dm_adj), df=n - 1)
    return {"dm_stat": round(float(dm_adj), 4), "p_value": round(float(p_value), 6)}


def _pure_multistep_ml_eval(model, train_sku_df, X_train, X_test, y_train, y_test):
    """Dự báo đệ quy thuần 141 bước từ điểm cắt 13/08/2023 (không dùng ngoại sinh thực của tập test)."""
    exo = ForecastService._future_exogenous(train_sku_df)
    u_hist = list(y_train.values.astype(float))
    preds = []
    dtypes = X_train.dtypes.to_dict()
    cols = list(X_train.columns)
    for idx in range(len(X_test)):
        row = X_test.iloc[idx].copy()
        row["Inventory Level"] = exo["Inventory Level"]
        row["Price"] = exo["Price"]
        row["Discount"] = exo["Discount"]
        row["Competitor Pricing"] = exo["Competitor Pricing"]
        row["Holiday/Promotion"] = 0
        row["price_ratio"] = exo["Price"] / (exo["Competitor Pricing"] + 1e-6)
        row["effective_price"] = exo["Price"] * (1.0 - exo["Discount"] / 100.0)
        for c in cols:
            if c.startswith("Weather Condition_"):
                row[c] = 1 if c == "Weather Condition_" + exo["Weather Condition"] else 0
            elif c.startswith("Seasonality_"):
                row[c] = 1 if c == "Seasonality_" + exo["Seasonality"] else 0
        row["lag_1"], row["lag_2"], row["lag_7"] = u_hist[-1], u_hist[-2], u_hist[-7]
        row["lag_14"], row["lag_21"], row["lag_30"] = u_hist[-14], u_hist[-21], u_hist[-30]
        w7, w14 = np.array(u_hist[-7:], dtype=float), np.array(u_hist[-14:], dtype=float)
        row["rolling_mean_7"], row["rolling_std_7"] = float(np.mean(w7)), float(np.std(w7, ddof=1))
        row["rolling_mean_14"], row["rolling_std_14"] = float(np.mean(w14)), float(np.std(w14, ddof=1))
        x_df = pd.DataFrame([row.values], columns=cols).astype(dtypes)
        y_hat = max(0.0, float(model.predict(x_df)[0]))
        u_hist.append(y_hat)
        preds.append(y_hat)
    pred_series = pd.Series(preds, index=y_test.index)
    return pred_series, calculate_metrics(y_test, pred_series)


def main():
    raw_df = load_raw_data()
    sku_df = get_sku_data(raw_df, DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID)
    ml = prepare_ml_data(sku_df)
    X_train, X_test = ml["X_train"], ml["X_test"]
    y_train, y_test = ml["y_train"], ml["y_test"]
    train_y = y_train.values.astype(float)
    test_y = y_test.values.astype(float)
    n = len(test_y)

    # 1. Baseline Naive và Seasonal Naive
    naive_pred = np.repeat(train_y[-1], n)
    snaive_pred = np.tile(train_y[-7:], n // 7 + 1)[:n]
    naive_metrics = calculate_metrics(pd.Series(test_y), pd.Series(naive_pred))
    snaive_metrics = calculate_metrics(pd.Series(test_y), pd.Series(snaive_pred))

    # 2. Ablation: Loại bỏ Inventory Level (27 đặc trưng, one-step-ahead)
    X_tr_no_inv = X_train.drop(columns=["Inventory Level"])
    X_te_no_inv = X_test.drop(columns=["Inventory Level"])
    rf_no_inv, pred_rf_no_inv, m_rf_no_inv = train_random_forest(X_tr_no_inv, y_train, X_te_no_inv, y_test)
    xgb_no_inv, pred_xgb_no_inv, m_xgb_no_inv = train_xgboost(X_tr_no_inv, y_train, X_te_no_inv, y_test)
    top5_rf_no_inv = {
        k: round(float(v), 4)
        for k, v in pd.Series(rf_no_inv.feature_importances_, index=X_tr_no_inv.columns)
        .sort_values(ascending=False)
        .head(5)
        .items()
    }

    # 3. Đệ quy thuần 141 bước (không biết ngoại sinh tương lai)
    svc = ForecastService(raw_df)
    rf_28_ev = svc.evaluate(DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID, "Random Forest")
    xgb_28_ev = svc.evaluate(DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID, "XGBoost")
    train_sku_df = sku_df.iloc[: len(sku_df) - n].copy()
    pred_rf_pure, m_rf_pure = _pure_multistep_ml_eval(rf_28_ev["model"], train_sku_df, X_train, X_test, y_train, y_test)
    pred_xgb_pure, m_xgb_pure = _pure_multistep_ml_eval(xgb_28_ev["model"], train_sku_df, X_train, X_test, y_train, y_test)

    # 4. Kiểm định Ljung-Box trên phần dư ARIMA & SARIMA (tập huấn luyện)
    arima_res = ARIMA(y_train, order=(2, 1, 1)).fit()
    sarima_res = SARIMAX(
        y_train,
        order=(1, 1, 1),
        seasonal_order=(1, 0, 1, 7),
        enforce_stationarity=False,
        enforce_invertibility=False,
    ).fit(disp=False)

    lb_arima = acorr_ljungbox(arima_res.resid, lags=[7, 14], return_df=True)
    lb_sarima = acorr_ljungbox(sarima_res.resid, lags=[7, 14], return_df=True)

    ljung_box = {
        "ARIMA": {
            f"lag_{int(lag)}": {"lb_stat": round(float(row["lb_stat"]), 4), "lb_pvalue": round(float(row["lb_pvalue"]), 4)}
            for lag, row in lb_arima.iterrows()
        },
        "SARIMA": {
            f"lag_{int(lag)}": {"lb_stat": round(float(row["lb_stat"]), 4), "lb_pvalue": round(float(row["lb_pvalue"]), 4)}
            for lag, row in lb_sarima.iterrows()
        },
    }

    # 5. Kiểm định Diebold-Mariano giữa các cặp mô hình trên 141 ngày test
    preds = {
        m: np.asarray(svc.evaluate(DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID, m)["y_pred"], dtype=float)
        for m in ["Random Forest", "XGBoost", "SARIMA", "ARIMA", "Exponential Smoothing"]
    }
    preds["Naive"] = naive_pred
    preds["Seasonal Naive"] = snaive_pred
    preds["RF (no Inv)"] = np.asarray(pred_rf_no_inv, dtype=float)
    preds["RF (Pure 141-step)"] = np.asarray(pred_rf_pure, dtype=float)

    pairs = [
        ("Random Forest", "SARIMA"),
        ("Random Forest", "ARIMA"),
        ("Random Forest", "Naive"),
        ("Random Forest", "Seasonal Naive"),
        ("Random Forest", "XGBoost"),
        ("Random Forest", "RF (no Inv)"),
        ("RF (no Inv)", "SARIMA"),
        ("RF (Pure 141-step)", "SARIMA"),
        ("XGBoost", "SARIMA"),
        ("ARIMA", "SARIMA"),
    ]
    dm_results = []
    for m1, m2 in pairs:
        e1 = test_y - preds[m1]
        e2 = test_y - preds[m2]
        mse_res = diebold_mariano_test(e1, e2, h=1, power=2)
        mae_res = diebold_mariano_test(e1, e2, h=1, power=1)
        dm_results.append(
            {
                "pair": f"{m1} vs {m2}",
                "model_1": m1,
                "model_2": m2,
                "dm_mse": mse_res["dm_stat"],
                "p_mse": mse_res["p_value"],
                "dm_mae": mae_res["dm_stat"],
                "p_mae": mae_res["p_value"],
            }
        )

    # 6. Đối chiếu tồn kho SS / ROP / Q và ABC-XYZ CV
    rf_test_mean = float(np.mean(preds["Random Forest"]))
    pol_img411 = calculate_inventory_policy(sku_df, forecast_daily_demand=rf_test_mean, lead_time=LEAD_TIME_DAYS, service_level=None)
    pol_web90 = calculate_inventory_policy(sku_df, forecast_daily_demand=None, lead_time=LEAD_TIME_DAYS, service_level=SERVICE_LEVEL)
    abc_df = compute_abc_xyz_analysis(raw_df, store_id=DEFAULT_STORE_ID)

    payload = {
        "sku": f"{DEFAULT_STORE_ID}-{DEFAULT_PRODUCT_ID}",
        "train_n": int(len(train_y)),
        "test_n": int(n),
        "baselines": {
            "Naive": {"last_train_value": float(train_y[-1]), **naive_metrics},
            "Seasonal Naive": {"last_7_train_values": [float(v) for v in train_y[-7:]], **snaive_metrics},
        },
        "ablation_no_inventory": {
            "Random Forest (27 feats)": {**m_rf_no_inv, "top5_features": top5_rf_no_inv},
            "XGBoost (27 feats)": m_xgb_no_inv,
        },
        "pure_multistep_141": {
            "Random Forest (Pure 141-step)": m_rf_pure,
            "XGBoost (Pure 141-step)": m_xgb_pure,
        },
        "ljung_box": ljung_box,
        "diebold_mariano": dm_results,
        "inventory_reconciliation": {
            "rf_test_141_mean": round(rf_test_mean, 2),
            "hinh_4_11_static": pol_img411,
            "web_90d_average": pol_web90,
            "abc_xyz_cv_summary": {
                "min": round(float(abc_df["CV"].min()), 3),
                "median": round(float(abc_df["CV"].median()), 3),
                "max": round(float(abc_df["CV"].max()), 3),
                "count_le_0_80": int((abc_df["CV"] <= 0.80).sum()),
                "count_gt_0_80": int((abc_df["CV"] > 0.80).sum()),
            },
        },
    }

    os.makedirs(REPORTS_DIR, exist_ok=True)
    json_path = os.path.join(REPORTS_DIR, "extra_diagnostics_s001_p0001.json")
    csv_path = os.path.join(REPORTS_DIR, "extra_diagnostics_s001_p0001.csv")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    pd.DataFrame(dm_results).to_csv(csv_path, index=False, encoding="utf-8-sig")

    print("=" * 76)
    print("BỔ SUNG BASELINES, ABLATION INVENTORY LEVEL, ĐA BƯỚC 141 NGÀY & KIỂM ĐỊNH")
    print("=" * 76)
    print(f"Naive (y_last={train_y[-1]:.0f})               -> {naive_metrics}")
    print(f"Seasonal Naive (7 ngày)         -> {snaive_metrics}")
    print(f"RF (27 feats, bỏ Inv Level)     -> {m_rf_no_inv}")
    print(f"XGB (27 feats, bỏ Inv Level)    -> {m_xgb_no_inv}")
    print(f"RF (Đệ quy thuần 141 bước)      -> {m_rf_pure}")
    print(f"XGB (Đệ quy thuần 141 bước)     -> {m_xgb_pure}")
    print("\nTop 5 RF (bỏ Inventory Level):", top5_rf_no_inv)
    print("Ljung-Box ARIMA :", ljung_box["ARIMA"])
    print("Ljung-Box SARIMA:", ljung_box["SARIMA"])
    print("\nDiebold-Mariano Tests (Harvey-corrected):")
    print(pd.DataFrame(dm_results).to_string(index=False))
    print(f"\nĐã lưu: {json_path}\nĐã lưu: {csv_path}")


if __name__ == "__main__":
    main()
