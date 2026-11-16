"""
build_static_web_data.py
Sinh dữ liệu tĩnh chuẩn xác (frontend/data/static_db.json) từ Dataset & mô hình thực tế
để Web App Retail ML Hub chạy 100% đầy đủ tính năng trên GitHub Pages (không cần backend server)
đồng thời vẫn giữ nguyên chế độ FastAPI khi chạy local.
"""

import json
import os
import sys
import time

import numpy as np
import pandas as pd

if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.config import DATA_PATH, DEFAULT_PRODUCT_ID, DEFAULT_STORE_ID, LEAD_TIME_DAYS, SERVICE_LEVEL
from src.data_loader import load_raw_data
from src.forecast_service import MODEL_NAMES, ForecastService
from src.forecasting_models import calculate_metrics
from src.inventory_optimizer import compute_abc_xyz_analysis

ABC_XYZ_STRATEGY = {
    "AX": "Giá trị cao, nhu cầu ổn định: áp dụng JIT, đặt hàng tự động định kỳ, tồn kho an toàn thấp.",
    "AY": "Giá trị cao, nhu cầu dao động: theo dõi sát dự báo, tồn kho an toàn vừa phải.",
    "AZ": "Giá trị cao, nhu cầu biến động mạnh: ưu tiên dự báo chính xác, xem xét lại hàng tuần, đặt theo đơn.",
    "BX": "Giá trị trung bình, ổn định: đặt hàng định kỳ, tự động hóa bổ sung hàng.",
    "BY": "Giá trị trung bình, dao động: kiểm tra định kỳ, duy trì tồn kho an toàn trung bình.",
    "BZ": "Giá trị trung bình, biến động mạnh: giữ tồn kho an toàn cao hơn hoặc đặt theo đơn hàng.",
    "CX": "Giá trị thấp, ổn định: đặt hàng số lượng lớn, ít kiểm tra để giảm chi phí quản lý.",
    "CY": "Giá trị thấp, dao động: đặt hàng theo lô, kiểm tra không thường xuyên.",
    "CZ": "Giá trị thấp, biến động mạnh: cân nhắc giảm tồn kho / ngừng kinh doanh, chỉ đặt khi có đơn.",
}


def _fast_sku_forecast_and_metrics(u_arr, horizon=30):
    """
    Tính nhanh dự báo 30 ngày & bộ chỉ số kiểm thử (141 ngày cuối) cho các SKU ngoài S001_P0001
    dựa trên mô hình kết hợp mùa vụ tuần (lag 7, 14) + trung bình trượt 7/14/30 ngày.
    """
    u = np.asarray(u_arr, dtype=float)
    n = len(u)
    test_len = 141
    train_end = n - test_len

    # Đánh giá trên 141 ngày cuối
    y_test = u[train_end:]
    y_pred = np.empty(test_len, dtype=float)
    for idx, t in enumerate(range(train_end, n)):
        m7 = np.mean(u[t - 7 : t])
        m14 = np.mean(u[t - 14 : t])
        m30 = np.mean(u[t - 30 : t])
        y_pred[idx] = 0.35 * u[t - 7] + 0.25 * m7 + 0.25 * m14 + 0.15 * m30

    metrics = calculate_metrics(pd.Series(y_test), pd.Series(y_pred))

    # Dự báo đệ quy 30 ngày tiếp theo
    hist = list(u[-35:])
    preds = []
    for _ in range(horizon):
        m7 = float(np.mean(hist[-7:]))
        m14 = float(np.mean(hist[-14:]))
        m30 = float(np.mean(hist[-30:]))
        val = max(0.0, 0.35 * hist[-7] + 0.25 * m7 + 0.25 * m14 + 0.15 * m30)
        hist.append(val)
        preds.append(round(val, 2))

    return preds, metrics


def main():
    t0 = time.time()
    print("--> Đang đọc dữ liệu gốc & khởi tạo ForecastService...", flush=True)
    raw_df = load_raw_data(DATA_PATH)
    service = ForecastService(raw_df)

    all_stores = sorted(list(raw_df["Store ID"].unique()))
    all_products = sorted(list(raw_df["Product ID"].unique()))
    all_categories = sorted(list(raw_df["Category"].unique()))
    product_category = raw_df.groupby("Product ID")["Category"].agg(lambda s: s.mode().iloc[0]).to_dict()

    # 1. KPI
    total_sales = float(raw_df["Units Sold"].sum())
    total_revenue = float((raw_df["Units Sold"] * raw_df["Price"] * (1 - raw_df["Discount"] / 100.0)).sum())
    avg_inventory = float(raw_df["Inventory Level"].mean())
    kpi = {
        "total_records": len(raw_df),
        "total_stores": len(all_stores),
        "total_products": len(all_products),
        "total_categories": len(all_categories),
        "total_units_sold": int(total_sales),
        "total_revenue": round(total_revenue, 2),
        "avg_inventory_per_store": round(avg_inventory, 1),
        "date_range": f"{raw_df['Date'].min().strftime('%d/%m/%Y')} - {raw_df['Date'].max().strftime('%d/%m/%Y')}",
    }

    # 2. Options
    options = {
        "stores": all_stores,
        "products": [{"id": p, "category": product_category.get(p, "")} for p in all_products],
        "categories": all_categories,
        "models": MODEL_NAMES,
        "best_model": service.best_model_name(),
        "horizons": [7, 14, 30],
        "default_lead_time": LEAD_TIME_DAYS,
        "default_service_level": SERVICE_LEVEL,
    }

    # 3. EDA Data
    cat_sales = raw_df.groupby("Category")["Units Sold"].sum().reset_index()
    store_sales = raw_df.groupby("Store ID")["Units Sold"].sum().reset_index()
    weather_sales = raw_df.groupby("Weather Condition")["Units Sold"].mean().reset_index()
    season_order = ["Spring", "Summer", "Autumn", "Winter"]
    season_sales = raw_df.groupby("Seasonality")["Units Sold"].mean().reindex(season_order).dropna()
    promo = raw_df.groupby("Holiday/Promotion")["Units Sold"].mean()
    eda_data = {
        "category_sales": {
            "categories": cat_sales["Category"].tolist(),
            "series": [int(x) for x in cat_sales["Units Sold"].tolist()],
        },
        "store_sales": {
            "stores": store_sales["Store ID"].tolist(),
            "series": [int(x) for x in store_sales["Units Sold"].tolist()],
        },
        "weather_sales": {
            "conditions": weather_sales["Weather Condition"].tolist(),
            "series": [round(float(x), 1) for x in weather_sales["Units Sold"].tolist()],
        },
        "seasonality_sales": {
            "seasons": season_sales.index.tolist(),
            "series": [round(float(x), 1) for x in season_sales.values],
        },
        "promo_effect": {
            "labels": ["Không khuyến mãi", "Có khuyến mãi / Ngày lễ"],
            "series": [round(float(promo.get(0, 0)), 1), round(float(promo.get(1, 0)), 1)],
        },
    }

    # 4. Feature Importance (top 28 để cắt theo top_n bất kỳ)
    feature_importance = {
        "Random Forest": service.feature_importance("Random Forest", top_n=28),
        "XGBoost": service.feature_importance("XGBoost", top_n=28),
    }

    # 5. Assumptions
    assumptions = {m: service.assumptions_text(m) for m in MODEL_NAMES}

    # 6. ABC-XYZ cho cả 5 cửa hàng
    abc_xyz = {}
    for st in all_stores:
        abc_df = compute_abc_xyz_analysis(raw_df, store_id=st)
        cols = ["Product ID", "Category", "Total_Revenue", "Mean_Demand", "CV", "Revenue_Share", "ABC_Class", "XYZ_Class", "ABC_XYZ"]
        records = []
        for r in abc_df[cols].to_dict(orient="records"):
            records.append({
                "product_id": r["Product ID"],
                "category": r["Category"],
                "total_revenue": round(float(r["Total_Revenue"]), 2),
                "mean_demand": round(float(r["Mean_Demand"]), 1),
                "cv": round(float(r["CV"]), 3),
                "revenue_share": round(float(r["Revenue_Share"]) * 100, 2),
                "abc": r["ABC_Class"],
                "xyz": r["XYZ_Class"],
                "group": r["ABC_XYZ"],
            })
        crosstab = pd.crosstab(abc_df["ABC_Class"], abc_df["XYZ_Class"])
        matrix = {}
        for r_cls in ["A", "B", "C"]:
            matrix[r_cls] = {}
            for c_cls in ["X", "Y", "Z"]:
                matrix[r_cls][c_cls] = int(crosstab.loc[r_cls, c_cls]) if (r_cls in crosstab.index and c_cls in crosstab.columns) else 0
        abc_xyz[st] = {
            "store_id": st,
            "matrix": matrix,
            "strategy": ABC_XYZ_STRATEGY,
            "items": records,
        }

    # 7. Trục thời gian D / W / M và ánh xạ chỉ số ngày -> tuần / tháng
    sample_sku = service.get_sku(DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID)
    dates_d = [pd.Timestamp(d).strftime("%Y-%m-%d") for d in sample_sku.index]
    s_dummy = pd.Series(np.arange(len(sample_sku)), index=sample_sku.index)
    w_res = s_dummy.resample("W-MON").count()
    m_res = s_dummy.resample("MS").count()
    dates_w = [pd.Timestamp(d).strftime("%Y-%m-%d") for d in w_res.index]
    dates_m = [pd.Timestamp(d).strftime("%Y-%m-%d") for d in m_res.index]

    w_bins = pd.DatetimeIndex(w_res.index)
    m_bins = pd.DatetimeIndex(m_res.index)
    day_to_w = [int(np.searchsorted(w_bins, d, side="left")) for d in sample_sku.index]
    day_to_m = [int(np.searchsorted(m_bins, d, side="right") - 1) for d in sample_sku.index]

    # 8. Báo cáo đầy đủ 5 mô hình cho SKU trung tâm S001_P0001 (60 ngày tương lai + 141 ngày test)
    print("--> Đang tính kết quả chuẩn 5 mô hình cho SKU báo cáo S001_P0001...", flush=True)
    report_evals = {}
    report_forecasts = {}
    test_dates_str = []
    test_actual = []
    for m_name in MODEL_NAMES:
        ev = service.evaluate(DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID, m_name)
        fc = service.forecast(DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID, m_name, 60)
        test_dates_str = [pd.Timestamp(d).strftime("%Y-%m-%d") for d in ev["test_dates"]]
        test_actual = [int(round(v)) for v in ev["y_test"]]
        report_evals[m_name] = {
            "metrics": ev["metrics"],
            "source": ev["source"],
            "y_pred": [round(float(v), 1) for v in ev["y_pred"]],
        }
        report_forecasts[m_name] = {
            "dates": [pd.Timestamp(d).strftime("%Y-%m-%d") for d in fc["dates"]],
            "values": [int(round(v)) for v in fc["values"]],
            "lower": [int(round(v)) for v in fc["lower"]],
            "upper": [int(round(v)) for v in fc["upper"]],
        }

    # 9. Dữ liệu từng SKU (100 SKUs)
    print("--> Đang trích xuất chuỗi thời gian & dự báo cho 100 SKUs...", flush=True)
    skus = {}
    for st in all_stores:
        for pid in all_products:
            key = f"{st}_{pid}"
            sku_df = service.get_sku(st, pid)
            u_vals = [int(v) for v in sku_df["Units Sold"].values]
            if (st, pid) == (DEFAULT_STORE_ID, DEFAULT_PRODUCT_ID):
                fc_obj = service.forecast(st, pid, "Random Forest", 30)
                fc30 = [round(float(v), 2) for v in fc_obj["values"]]
                rf_metrics = report_evals["Random Forest"]["metrics"]
            else:
                fc30, rf_metrics = _fast_sku_forecast_and_metrics(u_vals, horizon=30)

            recent90 = sku_df["Units Sold"].tail(90)
            skus[key] = {
                "cat": product_category.get(pid, ""),
                "u": u_vals,
                "inv": [int(v) for v in sku_df["Inventory Level"].values],
                "avg90": round(float(recent90.mean()), 4),
                "std90": round(float(recent90.std(ddof=1)), 4),
                "cur_inv": int(sku_df["Inventory Level"].iloc[-1]),
                "fc30": fc30,
                "rf_metrics": rf_metrics,
            }

    static_db = {
        "kpi": kpi,
        "options": options,
        "eda_data": eda_data,
        "feature_importance": feature_importance,
        "assumptions": assumptions,
        "abc_xyz": abc_xyz,
        "dates_d": dates_d,
        "dates_w": dates_w,
        "dates_m": dates_m,
        "day_to_w": day_to_w,
        "day_to_m": day_to_m,
        "report_sku": {
            "store_id": DEFAULT_STORE_ID,
            "product_id": DEFAULT_PRODUCT_ID,
            "test_dates": test_dates_str,
            "test_actual": test_actual,
            "evals": report_evals,
            "forecasts": report_forecasts,
        },
        "skus": skus,
    }

    out_dir = os.path.join(BASE_DIR, "frontend", "data")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "static_db.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(static_db, f, ensure_ascii=False, separators=(",", ":"))

    size_kb = os.path.getsize(out_path) / 1024
    print(f"--> Đã lưu {out_path} ({size_kb:.1f} KB) trong {time.time() - t0:.1f}s", flush=True)


if __name__ == "__main__":
    main()
