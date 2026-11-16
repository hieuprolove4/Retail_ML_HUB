"""
FastAPI Backend: api.py
Cung cấp toàn bộ RESTful API cho Web Dashboard:
- /api/kpi: Thẻ số liệu tổng quan
- /api/eda-data: Dữ liệu vẽ biểu đồ kinh doanh
- /api/sales-trend: Xu hướng sản lượng bán theo ngày / tuần / tháng
- /api/options: Danh sách Store ID và Product ID
- /api/forecast: Chạy dự báo nhu cầu tương lai bằng 1 trong 5 mô hình ML
- /api/model-comparison: So sánh chỉ số và kết quả dự báo trên tập kiểm thử của 5 mô hình
- /api/feature-importance: Mức độ quan trọng của đặc trưng (Random Forest / XGBoost)
- /api/inventory-status: Bảng kiểm tra mức tồn kho, tính ROP, SS và cảnh báo
- /api/inventory-history: Diễn biến tồn kho thực tế của 1 SKU so với ROP / SS
- /api/abc-xyz: Ma trận phân loại tồn kho chiến lược

Chạy từ thư mục gốc dự án:  uvicorn backend.api:app --reload
"""

import os
import threading
from contextlib import asynccontextmanager

import pandas as pd
import numpy as np
from fastapi import FastAPI, Query, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from src.config import DATA_PATH, LEAD_TIME_DAYS, SERVICE_LEVEL, DEFAULT_STORE_ID, BASE_DIR
from src.data_loader import load_raw_data
from src.inventory_optimizer import calculate_inventory_policy, compute_abc_xyz_analysis
from src.forecast_service import ForecastService, MODEL_NAMES


# Load dữ liệu vào RAM một lần khi khởi động để phản hồi cực nhanh
print("--> Đang khởi tạo bộ nhớ đệm dữ liệu (In-memory Cache)...")
RAW_DF = load_raw_data(DATA_PATH)
ALL_STORES = sorted(list(RAW_DF['Store ID'].unique()))
ALL_PRODUCTS = sorted(list(RAW_DF['Product ID'].unique()))
ALL_CATEGORIES = sorted(list(RAW_DF['Category'].unique()))
# Category của 1 Product ID thay đổi theo từng dòng trong dataset -> dùng giá trị phổ biến nhất
PRODUCT_CATEGORY = RAW_DF.groupby('Product ID')['Category'].agg(lambda s: s.mode().iloc[0]).to_dict()
print(f"    Sẵn sàng: {len(RAW_DF):,} dòng, {len(ALL_STORES)} stores, {len(ALL_PRODUCTS)} products.")

SERVICE = ForecastService(RAW_DF)
_INVENTORY_CACHE = {}


def _warm_up():
    """Huấn luyện trước mô hình tốt nhất cho cửa hàng mặc định để trang Tồn kho phản hồi nhanh khi demo"""
    try:
        print(f"--> [Warm-up] Đang chuẩn bị dự báo tồn kho cho cửa hàng {DEFAULT_STORE_ID}...")
        _compute_inventory_status(DEFAULT_STORE_ID, LEAD_TIME_DAYS, SERVICE_LEVEL, 'model')
        print("--> [Warm-up] Hoàn tất.")
    except Exception as exc:  # Không để lỗi warm-up làm dừng server
        print(f"--> [Warm-up] Bỏ qua do lỗi: {exc}")


@asynccontextmanager
async def lifespan(_app):
    threading.Thread(target=_warm_up, daemon=True).start()
    yield


app = FastAPI(title="Retail Demand Forecasting & Inventory Optimization API", lifespan=lifespan)

# Cấu hình CORS để frontend gọi API mượt mà
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------- #
# Tiện ích kiểm tra đầu vào
# ---------------------------------------------------------------------- #
def _check_store(store_id):
    if store_id not in ALL_STORES:
        raise HTTPException(status_code=404, detail=f"Không tìm thấy cửa hàng '{store_id}'")


def _check_product(product_id):
    if product_id not in ALL_PRODUCTS:
        raise HTTPException(status_code=404, detail=f"Không tìm thấy sản phẩm '{product_id}'")


def _check_model(model_name):
    if model_name not in MODEL_NAMES:
        raise HTTPException(status_code=422, detail=f"Mô hình không hợp lệ. Chọn một trong: {', '.join(MODEL_NAMES)}")


def _fmt_dates(dates):
    return [pd.Timestamp(d).strftime('%Y-%m-%d') for d in dates]


# Request Model cho API Dự báo
class ForecastRequest(BaseModel):
    store_id: str = "S001"
    product_id: str = "P0001"
    model_name: str = "Random Forest"
    horizon: int = Field(14, ge=1, le=60)


@app.get("/api/kpi")
def get_kpi_overview():
    """Lấy 4 thẻ chỉ số KPI tổng quan toàn chuỗi"""
    total_sales = float(RAW_DF['Units Sold'].sum())
    total_revenue = float((RAW_DF['Units Sold'] * RAW_DF['Price'] * (1 - RAW_DF['Discount'] / 100.0)).sum())
    avg_inventory = float(RAW_DF['Inventory Level'].mean())
    return {
        "total_records": len(RAW_DF),
        "total_stores": len(ALL_STORES),
        "total_products": len(ALL_PRODUCTS),
        "total_categories": len(ALL_CATEGORIES),
        "total_units_sold": int(total_sales),
        "total_revenue": round(total_revenue, 2),
        "avg_inventory_per_store": round(avg_inventory, 1),
        "date_range": f"{RAW_DF['Date'].min().strftime('%d/%m/%Y')} - {RAW_DF['Date'].max().strftime('%d/%m/%Y')}"
    }


@app.get("/api/options")
def get_filter_options():
    """Danh sách lựa chọn cho bộ lọc trên giao diện"""
    return {
        "stores": ALL_STORES,
        "products": [{"id": p, "category": PRODUCT_CATEGORY.get(p, "")} for p in ALL_PRODUCTS],
        "categories": ALL_CATEGORIES,
        "models": MODEL_NAMES,
        "best_model": SERVICE.best_model_name(),
        "horizons": [7, 14, 30],
        "default_lead_time": LEAD_TIME_DAYS,
        "default_service_level": SERVICE_LEVEL,
    }


@app.get("/api/eda-data")
def get_eda_chart_data():
    """Dữ liệu tổng hợp phục vụ Tab 1: Biểu đồ khám phá dữ liệu"""
    # 1. Doanh số theo ngành hàng
    cat_sales = RAW_DF.groupby('Category')['Units Sold'].sum().reset_index()
    cat_data = {
        "categories": cat_sales['Category'].tolist(),
        "series": [int(x) for x in cat_sales['Units Sold'].tolist()]
    }

    # 2. Doanh số theo Cửa hàng
    store_sales = RAW_DF.groupby('Store ID')['Units Sold'].sum().reset_index()
    store_data = {
        "stores": store_sales['Store ID'].tolist(),
        "series": [int(x) for x in store_sales['Units Sold'].tolist()]
    }

    # 3. Doanh số trung bình theo Thời tiết
    weather_sales = RAW_DF.groupby('Weather Condition')['Units Sold'].mean().reset_index()
    weather_data = {
        "conditions": weather_sales['Weather Condition'].tolist(),
        "series": [round(float(x), 1) for x in weather_sales['Units Sold'].tolist()]
    }

    # 4. Doanh số trung bình theo Mùa vụ
    season_order = ['Spring', 'Summer', 'Autumn', 'Winter']
    season_sales = RAW_DF.groupby('Seasonality')['Units Sold'].mean().reindex(season_order).dropna()
    season_data = {
        "seasons": season_sales.index.tolist(),
        "series": [round(float(x), 1) for x in season_sales.values]
    }

    # 5. Tác động của Khuyến mãi / Ngày lễ
    promo = RAW_DF.groupby('Holiday/Promotion')['Units Sold'].mean()
    promo_data = {
        "labels": ["Không khuyến mãi", "Có khuyến mãi / Ngày lễ"],
        "series": [round(float(promo.get(0, 0)), 1), round(float(promo.get(1, 0)), 1)]
    }

    return {
        "category_sales": cat_data,
        "store_sales": store_data,
        "weather_sales": weather_data,
        "seasonality_sales": season_data,
        "promo_effect": promo_data
    }


@app.get("/api/sales-trend")
def get_sales_trend(store_id: str = "", product_id: str = "", freq: str = Query("D", pattern="^[DWM]$")):
    """Xu hướng sản lượng bán theo ngày (D) / tuần (W) / tháng (M) kèm đường trung bình trượt"""
    df = RAW_DF
    if store_id:
        _check_store(store_id)
        df = df[df['Store ID'] == store_id]
    if product_id:
        _check_product(product_id)
        df = df[df['Product ID'] == product_id]

    daily = df.groupby('Date')['Units Sold'].sum()
    rule = {'D': 'D', 'W': 'W-MON', 'M': 'MS'}[freq]
    series = daily.resample(rule).sum() if freq != 'D' else daily
    windows = {'D': (7, 30), 'W': (4, 12), 'M': (3, 6)}[freq]
    ma_short = series.rolling(windows[0]).mean()
    ma_long = series.rolling(windows[1]).mean()

    def clean(s):
        return [None if pd.isna(v) else round(float(v), 1) for v in s.values]

    return {
        "freq": freq,
        "dates": _fmt_dates(series.index),
        "units": [int(v) for v in series.values],
        "ma_short": clean(ma_short),
        "ma_long": clean(ma_long),
        "ma_windows": list(windows),
    }


@app.post("/api/forecast")
def run_forecast(req: ForecastRequest):
    """
    Chạy dự báo nhu cầu tương lai (7, 14, 30 ngày) cho 1 SKU cụ thể bằng mô hình đã chọn
    """
    _check_store(req.store_id)
    _check_product(req.product_id)
    _check_model(req.model_name)

    sku_df = SERVICE.get_sku(req.store_id, req.product_id)

    # Lấy 60 ngày lịch sử gần nhất để hiển thị
    recent_history = sku_df.tail(60)
    hist_dates = _fmt_dates(recent_history.index)
    hist_actual = [int(v) for v in recent_history['Units Sold'].tolist()]

    try:
        evaluation = SERVICE.evaluate(req.store_id, req.product_id, req.model_name)
        fc = SERVICE.forecast(req.store_id, req.product_id, req.model_name, req.horizon)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Không thể chạy mô hình {req.model_name}: {exc}")

    return {
        "store_id": req.store_id,
        "product_id": req.product_id,
        "category": PRODUCT_CATEGORY.get(req.product_id, ""),
        "model_used": req.model_name,
        "horizon": req.horizon,
        "metrics": evaluation['metrics'],
        "metrics_source": evaluation['source'],
        "assumptions": SERVICE.assumptions_text(req.model_name),
        "history": {
            "dates": hist_dates,
            "values": hist_actual
        },
        "forecast": {
            "dates": _fmt_dates(fc['dates']),
            "values": [int(round(v)) for v in fc['values']],
            "lower": [int(round(v)) for v in fc['lower']],
            "upper": [int(round(v)) for v in fc['upper']],
            "total": int(round(sum(fc['values'])))
        }
    }


@app.get("/api/model-comparison")
def get_model_comparison(store_id: str = DEFAULT_STORE_ID, product_id: str = "P0001"):
    """So sánh RMSE / MAE / MAPE và đường dự báo trên tập kiểm thử (20%) của 5 mô hình cho 1 SKU"""
    _check_store(store_id)
    _check_product(product_id)

    results, preds = [], {}
    test_dates, actual, sources = [], [], set()
    for name in MODEL_NAMES:
        try:
            ev = SERVICE.evaluate(store_id, product_id, name)
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Không thể đánh giá mô hình {name}: {exc}")
        results.append({"model": name, **ev['metrics']})
        preds[name] = [round(v, 1) for v in ev['y_pred']]
        test_dates, actual = ev['test_dates'], ev['y_test']
        sources.add(ev['source'])

    best = min(results, key=lambda r: r['RMSE'])['model']
    return {
        "store_id": store_id,
        "product_id": product_id,
        "best_model": best,
        "metrics_source": "report" if sources == {"report"} else "on-demand",
        "results": results,
        "test": {
            "dates": _fmt_dates(test_dates),
            "actual": [int(v) for v in actual],
            "pred": preds
        }
    }


@app.get("/api/feature-importance")
def get_feature_importance(model: str = "Random Forest", top_n: int = Query(10, ge=1, le=30)):
    """Top đặc trưng quan trọng nhất của mô hình cây (huấn luyện trên SKU của báo cáo S001 - P0001)"""
    if model not in ('Random Forest', 'XGBoost'):
        raise HTTPException(status_code=422, detail="Chỉ hỗ trợ 'Random Forest' hoặc 'XGBoost'")
    return {"model": model, **SERVICE.feature_importance(model, top_n)}


def _compute_inventory_status(store_id, lead_time, service_level, demand_source):
    """Tính SS, ROP, trạng thái cho toàn bộ sản phẩm của 1 cửa hàng (có cache)"""
    key = (store_id, lead_time, round(service_level, 4), demand_source)
    if key in _INVENTORY_CACHE:
        return _INVENTORY_CACHE[key]

    store_df = RAW_DF[RAW_DF['Store ID'] == store_id]
    products = sorted(list(store_df['Product ID'].unique()))

    items = []
    counts = {'CRITICAL': 0, 'WARNING': 0, 'SAFE': 0}
    for pid in products:
        sku_df = store_df[store_df['Product ID'] == pid].sort_values(by='Date')
        forecast_demand = None
        if demand_source == 'model':
            forecast_demand = SERVICE.forecast_daily_demand(store_id, pid, days=lead_time)
        policy = calculate_inventory_policy(sku_df, forecast_daily_demand=forecast_demand,
                                            lead_time=lead_time, service_level=service_level)
        counts[policy['Status Code']] += 1

        items.append({
            "product_id": pid,
            "category": policy['Category'],
            "current_inventory": policy['Current Inventory'],
            "forecast_demand": policy['Forecast Daily Demand (D_hat)'],
            "demand_std": policy['Demand Std Dev (sigma_D)'],
            "safety_stock": policy['Safety Stock (SS)'],
            "reorder_point": policy['Reorder Point (ROP)'],
            "status": policy['Status'],
            "status_code": policy['Status Code'],
            "suggested_order": policy['Suggested Order Qty']
        })

    result = {
        "store_id": store_id,
        "lead_time_days": lead_time,
        "service_level": f"{round(service_level * 100)}%",
        "service_level_value": service_level,
        "z_score": round(float(policy['Z Score']), 3) if products else None,
        "demand_source": demand_source,
        "demand_model": SERVICE.best_model_name() if demand_source == 'model' else "Trung bình 90 ngày",
        "summary": {
            "critical": counts['CRITICAL'],
            "warning": counts['WARNING'],
            "safe": counts['SAFE'],
            "total": len(products)
        },
        "items": items
    }
    _INVENTORY_CACHE[key] = result
    return result


@app.get("/api/inventory-status")
def get_inventory_status(store_id: str = "S001",
                         lead_time: int = Query(LEAD_TIME_DAYS, ge=1, le=30),
                         service_level: float = Query(SERVICE_LEVEL, ge=0.5, le=0.999),
                         demand_source: str = Query("model", pattern="^(model|average)$")):
    """
    Tính toán Tồn kho an toàn (SS), Điểm đặt hàng lại (ROP)
    và trạng thái cảnh báo cho toàn bộ 20 sản phẩm của cửa hàng.
    demand_source = 'model': nhu cầu dự báo bởi mô hình tốt nhất; 'average': trung bình 90 ngày gần nhất
    """
    _check_store(store_id)
    try:
        return _compute_inventory_status(store_id, lead_time, service_level, demand_source)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Không thể tính toán tồn kho: {exc}")


@app.get("/api/inventory-history")
def get_inventory_history(store_id: str = "S001", product_id: str = "P0001",
                          days: int = Query(90, ge=14, le=731),
                          lead_time: int = Query(LEAD_TIME_DAYS, ge=1, le=30),
                          service_level: float = Query(SERVICE_LEVEL, ge=0.5, le=0.999),
                          demand_source: str = Query("model", pattern="^(model|average)$")):
    """Diễn biến tồn kho & sản lượng bán của 1 SKU trong `days` ngày gần nhất, kèm ngưỡng ROP và SS"""
    _check_store(store_id)
    _check_product(product_id)
    sku_df = RAW_DF[(RAW_DF['Store ID'] == store_id) & (RAW_DF['Product ID'] == product_id)].sort_values(by='Date')
    forecast_demand = SERVICE.forecast_daily_demand(store_id, product_id, days=lead_time) \
        if demand_source == 'model' else None
    policy = calculate_inventory_policy(sku_df, forecast_daily_demand=forecast_demand,
                                        lead_time=lead_time, service_level=service_level)
    recent = sku_df.tail(days)
    return {
        "store_id": store_id,
        "product_id": product_id,
        "dates": _fmt_dates(recent['Date']),
        "inventory": [int(v) for v in recent['Inventory Level']],
        "units_sold": [int(v) for v in recent['Units Sold']],
        "rop": policy['Reorder Point (ROP)'],
        "ss": policy['Safety Stock (SS)'],
        "forecast_demand": policy['Forecast Daily Demand (D_hat)'],
        "current_inventory": policy['Current Inventory'],
        "status_code": policy['Status Code'],
        "suggested_order": policy['Suggested Order Qty']
    }


# Gợi ý chiến lược quản trị tồn kho cho 9 nhóm ABC - XYZ
ABC_XYZ_STRATEGY = {
    'AX': "Giá trị cao, nhu cầu ổn định: áp dụng JIT, đặt hàng tự động định kỳ, tồn kho an toàn thấp.",
    'AY': "Giá trị cao, nhu cầu dao động: theo dõi sát dự báo, tồn kho an toàn vừa phải.",
    'AZ': "Giá trị cao, nhu cầu biến động mạnh: ưu tiên dự báo chính xác, xem xét lại hàng tuần, đặt theo đơn.",
    'BX': "Giá trị trung bình, ổn định: đặt hàng định kỳ, tự động hóa bổ sung hàng.",
    'BY': "Giá trị trung bình, dao động: kiểm tra định kỳ, duy trì tồn kho an toàn trung bình.",
    'BZ': "Giá trị trung bình, biến động mạnh: giữ tồn kho an toàn cao hơn hoặc đặt theo đơn hàng.",
    'CX': "Giá trị thấp, ổn định: đặt hàng số lượng lớn, ít kiểm tra để giảm chi phí quản lý.",
    'CY': "Giá trị thấp, dao động: đặt hàng theo lô, kiểm tra không thường xuyên.",
    'CZ': "Giá trị thấp, biến động mạnh: cân nhắc giảm tồn kho / ngừng kinh doanh, chỉ đặt khi có đơn.",
}


@app.get("/api/abc-xyz")
def get_abc_xyz_matrix(store_id: str = "S001"):
    """Lấy dữ liệu ma trận phân loại ABC - XYZ của 20 sản phẩm"""
    _check_store(store_id)
    abc_df = compute_abc_xyz_analysis(RAW_DF, store_id=store_id)
    cols = ['Product ID', 'Category', 'Total_Revenue', 'Mean_Demand', 'CV', 'Revenue_Share', 'ABC_Class', 'XYZ_Class', 'ABC_XYZ']
    records = []
    for r in abc_df[cols].to_dict(orient='records'):
        records.append({
            "product_id": r['Product ID'],
            "category": r['Category'],
            "total_revenue": round(float(r['Total_Revenue']), 2),
            "mean_demand": round(float(r['Mean_Demand']), 1),
            "cv": round(float(r['CV']), 3),
            "revenue_share": round(float(r['Revenue_Share']) * 100, 2),
            "abc": r['ABC_Class'],
            "xyz": r['XYZ_Class'],
            "group": r['ABC_XYZ'],
        })

    # Đếm số lượng theo 9 ô ma trận
    crosstab = pd.crosstab(abc_df['ABC_Class'], abc_df['XYZ_Class'])
    matrix = {}
    for r in ['A', 'B', 'C']:
        matrix[r] = {}
        for c in ['X', 'Y', 'Z']:
            matrix[r][c] = int(crosstab.loc[r, c]) if (r in crosstab.index and c in crosstab.columns) else 0

    return {
        "store_id": store_id,
        "matrix": matrix,
        "strategy": ABC_XYZ_STRATEGY,
        "items": records
    }


# Mount thư mục giao diện Web Frontend tĩnh
FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")
if os.path.exists(FRONTEND_DIR):
    app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")
    app.mount("/frontend", StaticFiles(directory=FRONTEND_DIR), name="frontend")


@app.get("/")
def serve_index():
    index_file = os.path.join(FRONTEND_DIR, "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    return {"message": "Frontend is building..."}
