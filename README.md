# Retail ML Hub — Dự báo nhu cầu & Tối ưu tồn kho bán lẻ

Đồ án Machine Learning: xây dựng **pipeline dự báo nhu cầu theo SKU** (cặp *Store ID* × *Product ID*) trên dữ liệu bán lẻ chuỗi thời gian, so sánh 5 mô hình dự báo, rồi chuyển kết quả dự báo thành **quyết định tồn kho** (Safety Stock, Reorder Point, phân loại ABC–XYZ). Toàn bộ pipeline được đóng gói thành một **web app** (FastAPI backend + dashboard JavaScript) để demo tương tác.

> Ngôn ngữ: Python (phân tích/dự báo) + JavaScript thuần ES Modules (frontend). Tổng mã nguồn Python khoảng 2.000 dòng.

---

## 1. Bài toán

**Mục tiêu:** dự báo nhu cầu bán hàng ngày của một SKU trong 7 / 14 / 30 ngày tới, và từ dự báo đó tính ra các thông số vận hành tồn kho nhằm trả lời câu hỏi: *"Khi nào phải đặt hàng lại và đặt bao nhiêu?"*

**Đầu vào:** chuỗi thời gian `Units Sold` theo ngày của một SKU, kèm các biến ngoại cảnh: giá, khuyến mãi, đối thủ cạnh tranh, thời tiết, mùa vụ, tồn kho, ngày lễ.

**Đầu ra:**

1. Dự báo nhu cầu tương lai (7/14/30 ngày) theo từng mô hình.
2. Bảng so sánh định lượng 5 mô hình (RMSE, MAE, MAPE).
3. Chỉ số tồn kho cho SKU: `Safety Stock`, `Reorder Point`, trạng thái tồn kho và lượng đặt hàng đề xuất.
4. Ma trận phân loại ABC–XYZ cho danh mục sản phẩm của một cửa hàng.
5. Biểu đồ EDA / chẩn đoán chuỗi thời gian / kết quả mô hình xuất ra thư mục `reports/figures/`.

**Dữ liệu:** `Dataset/retail_store_inventory.csv` — file CSV 15 cột: `Date, Store ID, Product ID, Category, Region, Inventory Level, Units Sold, Units Ordered, Demand Forecast, Price, Discount, Weather Condition, Holiday/Promotion, Competitor Pricing, Seasonality`. File dữ liệu **không được commit vào repo** (xem mục 6).

---

## 2. Kiến trúc luồng xử lý

```
┌────────────────────────────────────────────────────────────────────────┐
│ GIAI ĐOẠN HUẤN LUYỆN (offline)                                         │
│                                                                        │
│  main_train_and_evaluate.py                                            │
│    └─ src/data_loader.py            nạp dữ liệu, tách SKU, feature eng. │
│    └─ src/time_series_diagnostics.py ADF/KPSS, ACF/PACF, rolling, phân rã│
│    └─ src/forecasting_models.py     huấn luyện 5 mô hình, tính metrics  │
│    └─ src/inventory_optimizer.py    SS, ROP, ABC-XYZ, mô phỏng tồn kho  │
│                                                                        │
│         │                                                              │
│         ├──► saved_models/*.pkl              (5 mô hình + best model)  │
│         ├──► saved_models/model_metadata.json (chỉ số đánh giá thật)   │
│         └──► reports/figures/*.png + reports/abc_xyz_analysis_s001.csv │
└────────────────────────────────────────────────────────────────────────┘
                              │  (nạp .pkl + metadata)
                              ▼
┌────────────────────────────────────────────────────────────────────────┐
│ GIAI ĐOẠN PHỤC VỤ (online)                                             │
│                                                                        │
│  run_app.py ──► uvicorn ──► backend/api.py (FastAPI, REST /api/*)      │
│                                 └─ src/forecast_service.py (cache RAM) │
│                                 ▲                                      │
│  frontend/ (HTML/CSS/JS) ───────┘  fetch JSON, ApexCharts, hash router │
└────────────────────────────────────────────────────────────────────────┘
```

**Luồng rút gọn:** *huấn luyện → lưu model `.pkl` + metadata → backend FastAPI nạp model và phục vụ REST API → frontend dashboard gọi API và hiển thị.*

### 2.1. Các bước trong `main_train_and_evaluate.py`

| Bước | Nội dung | Module |
|------|----------|--------|
| 1 | Nạp dữ liệu gốc, tách chuỗi `Units Sold` của SKU mặc định (S001 – P0001) | `src/data_loader.py` |
| 2 | Kiểm định tính dừng & chẩn đoán: rolling mean/std, ADF, KPSS, ACF/PACF, seasonal decomposition | `src/time_series_diagnostics.py` |
| 3 | Feature engineering + chia Train 80% / Test 20% **theo thứ tự thời gian** | `src/data_loader.py` |
| 4 | Huấn luyện & đánh giá 5 mô hình (ARIMA, Exponential Smoothing, SARIMA, Random Forest, XGBoost) | `src/forecasting_models.py` |
| 5 | Tổng hợp bảng so sánh, xác định mô hình tốt nhất theo **RMSE thấp nhất**, lưu `.pkl` + `model_metadata.json` | `src/forecasting_models.py` |
| 6 | Tính `Safety Stock`, `Reorder Point`, mô phỏng tồn kho, ma trận ABC–XYZ | `src/inventory_optimizer.py` |

### 2.2. REST API (`backend/api.py`)

| Endpoint | Chức năng |
|----------|-----------|
| `GET /` | Trả về giao diện dashboard (`frontend/index.html`) |
| `GET /api/kpi` | Thẻ số liệu tổng quan |
| `GET /api/options` | Danh sách `Store ID`, `Product ID`, danh mục |
| `GET /api/eda-data` | Dữ liệu vẽ biểu đồ EDA kinh doanh |
| `GET /api/sales-trend` | Xu hướng sản lượng bán theo ngày/tuần/tháng |
| `POST /api/forecast` | Chạy dự báo nhu cầu tương lai bằng 1 trong 5 mô hình |
| `GET /api/model-comparison` | So sánh chỉ số + kết quả dự báo trên tập test của 5 mô hình |
| `GET /api/feature-importance` | Mức độ quan trọng của đặc trưng (Random Forest / XGBoost) |
| `GET /api/inventory-status` | Bảng tồn kho, ROP, SS và cảnh báo |
| `GET /api/inventory-history` | Diễn biến tồn kho thực tế so với ROP/SS |
| `GET /api/abc-xyz` | Ma trận phân loại ABC–XYZ |

Backend mount thư mục `frontend/` tại `/static` và cache dữ liệu trong RAM khi khởi động.

### 2.3. Các trang của dashboard (`frontend/`)

| Route | Trang |
|-------|-------|
| `#/overview` | Tổng quan chuỗi & khám phá dữ liệu (EDA) |
| `#/forecast` | Dự báo nhu cầu sản phẩm theo SKU |
| `#/models` | So sánh 5 mô hình dự báo & Feature Importance |
| `#/inventory` | Tồn kho an toàn (SS) & cảnh báo điểm đặt hàng lại (ROP) |
| `#/abcxyz` | Ma trận phân loại tồn kho chiến lược ABC–XYZ |
| `#/whatif` | Mô phỏng kịch bản What-If chuỗi cung ứng |

---

## 3. Năm mô hình và kết quả đánh giá thật

Chỉ số dưới đây **đọc trực tiếp** từ `saved_models/model_metadata.json` (huấn luyện trên SKU `S001` – `P0001`, chia Train 80% / Test 20% theo thời gian). Định nghĩa chỉ số nằm ở `src/forecasting_models.py:33-46` (RMSE, MAE, MAPE %).

| Mô hình | RMSE | MAE | MAPE (%) | Nguồn |
|---------|-----:|----:|---------:|-------|
| ARIMA | 107.07 | 90.04 | 221.29 | `saved_models/model_metadata.json:43-47` |
| Exponential Smoothing (Holt-Winters) | 108.21 | 90.23 | 219.37 | `saved_models/model_metadata.json:48-53` |
| SARIMA | 107.03 | 90.03 | 221.23 | `saved_models/model_metadata.json:54-59` |
| **Random Forest** | **92.42** | **73.71** | **173.29** | `saved_models/model_metadata.json:60-65` |
| XGBoost | 95.73 | 76.24 | 181.99 | `saved_models/model_metadata.json:66-71` |

**Mô hình tốt nhất: `Random Forest`** (RMSE 92.42 — thấp nhất), khai báo tại `saved_models/model_metadata.json:2`, chỉ số tại `saved_models/model_metadata.json:3-8`.

Ghi chú kỹ thuật:

- Mô hình tốt nhất được chọn theo **RMSE thấp nhất** (`main_train_and_evaluate.py`, bước 5).
- Cấu hình tham số: ARIMA `(2, 1, 1)`; SARIMA `(1, 1, 1)(1, 0, 1, 7)`; Holt-Winters chu kỳ mùa `7`.
- Random Forest và XGBoost dùng 28 đặc trưng (`saved_models/model_metadata.json:11-40`): biến lịch (`dayofweek`, `day`, `month`, `quarter`, `is_weekend`), lag (`lag_1`, `lag_2`, `lag_7`, `lag_14`, `lag_21`, `lag_30`), rolling (`rolling_mean_7`, `rolling_std_7`, `rolling_mean_14`, `rolling_std_14`), biến giá (`price_ratio`, `effective_price`) và biến one-hot cho thời tiết/mùa vụ.
- Thông số tồn kho mặc định (`src/config.py`): lead time 3 ngày, service level 95% (Z = 1.65), Train/Test = 80/20.
- Công thức tồn kho (`src/inventory_optimizer.py:37-100`): `Safety Stock = Z × σ_D × √L` và `Reorder Point = D̂ × L + SS`.

---

## 4. Cấu trúc thư mục

```
ML_Project/
├── main_train_and_evaluate.py   # Script chính: chạy toàn bộ pipeline huấn luyện & đánh giá
├── run_app.py                   # Khởi động FastAPI server + tự mở trình duyệt
├── start_app.bat                # Wrapper cho Windows (tự dùng .venv nếu có)
├── requirements.txt             # Danh sách thư viện Python
├── src/                         # Lõi xử lý
│   ├── config.py                # Đường dẫn, tham số (lead time, service level, tỉ lệ train)
│   ├── data_loader.py           # Nạp dữ liệu, tách SKU, feature engineering, chia train/test
│   ├── eda_analysis.py          # Phân tích khám phá dữ liệu (7 biểu đồ EDA)
│   ├── time_series_diagnostics.py # Kiểm định ADF/KPSS, ACF/PACF, rolling, phân rã mùa vụ
│   ├── forecasting_models.py    # 5 mô hình dự báo + hàm tính RMSE/MAE/MAPE + vẽ biểu đồ
│   ├── forecast_service.py      # Dịch vụ dự báo có cache cho API (đánh giá & dự báo tương lai)
│   └── inventory_optimizer.py   # Safety Stock, ROP, ABC-XYZ, mô phỏng tồn kho
├── backend/
│   └── api.py                   # FastAPI: REST API + phục vụ frontend tĩnh
├── frontend/                    # Dashboard (HTML/CSS/JS thuần, ES Modules)
│   ├── index.html
│   ├── css/                     # tokens.css, app.css
│   ├── js/                      # router, api client, ui, charts, pages/*
│   └── vendor/                  # ApexCharts & Lucide (bản local, chạy offline)
└── saved_models/
    └── model_metadata.json      # Chỉ số & metadata của 5 mô hình (ĐƯỢC commit)
                                 # (các file *.pkl KHÔNG commit — xem mục 6)
```

---

## 5. Cách chạy

### 5.1. Cài đặt

```bash
# (khuyến nghị) tạo môi trường ảo
python -m venv .venv
# Windows: .venv\Scripts\activate    |  macOS/Linux: source .venv/bin/activate

pip install -r requirements.txt
```

### 5.2. Huấn luyện lại mô hình & sinh báo cáo

```bash
# Yêu cầu: có file Dataset/retail_store_inventory.csv (xem mục 6)
python main_train_and_evaluate.py
```

Script này sẽ: kiểm định chuỗi thời gian → huấn luyện 5 mô hình → in bảng so sánh chỉ số → lưu `saved_models/*.pkl` và `saved_models/model_metadata.json` → tính SS/ROP/ABC-XYZ → xuất biểu đồ ra `reports/figures/`.

### 5.3. Chạy web app

```bash
python run_app.py
```

Server chạy tại `http://127.0.0.1:8000` và trình duyệt tự mở. Trên Windows có thể nhấp đúp `start_app.bat`.

> Có thể chạy trực tiếp backend bằng: `uvicorn backend.api:app --reload`

---

## 6. Dữ liệu và file mô hình KHÔNG nằm trong repo

Repo được cố ý giữ nhẹ, các thành phần sau bị `.gitignore` (không commit):

| Thành phần | Lý do |
|------------|-------|
| `Dataset/retail_store_inventory.csv` | File dữ liệu lớn (~6 MB) |
| `saved_models/*.pkl` | File nhị phân ~22 MB (tái tạo được bằng `python main_train_and_evaluate.py`) |
| `reports/` | Biểu đồ/CSV sinh tự động (~20 MB) |
| `ML_Project.zip` | Gói nén dự án (~29 MB) |
| `.venv/`, `__pycache__/`, `*.pyc`, `*.log` | Môi trường ảo và file tạm |

**Riêng `saved_models/model_metadata.json` được commit** vì là bằng chứng kết quả huấn luyện (chỉ số thật của 5 mô hình).

Để chạy lại từ đầu: đặt file dữ liệu vào đúng đường dẫn `Dataset/retail_store_inventory.csv` (đường dẫn khai báo ở `src/config.py`), sau đó chạy `python main_train_and_evaluate.py`.

---

## 7. Hạn chế đã biết

Đây là các điểm yếu thật của đồ án, ghi lại để người đọc đánh giá đúng — **không tô hồng kết quả**:

1. **MAPE rất cao (173% – 221% ở cả 5 mô hình)** — đọc từ `saved_models/model_metadata.json:43-71`. Nguyên nhân khả dĩ:
   - MAPE chia cho nhu cầu thực tế từng ngày, mà nhu cầu của một SKU đơn lẻ biến động mạnh và có những ngày giá trị rất nhỏ, khiến tỷ lệ sai số bị khuếch đại; chỉ cần vài ngày nhu cầu thấp là MAPE trung bình vọt lên hàng trăm phần trăm.
   - Hàm tính MAPE trong `src/forecasting_models.py:39-40` đã **loại các mẫu có nhu cầu thực tế bằng 0** (`non_zero_mask = y_true != 0`) để tránh chia cho 0, nhưng điều này chỉ xử lý đúng mẫu bằng 0 tuyệt đối — các mẫu nhu cầu rất nhỏ (khác 0) vẫn làm mẫu số nhỏ và MAPE phình to.
   - Nói ngắn gọn: **MAPE không phù hợp với chuỗi nhu cầu nhỏ, nhiều nhiễu**. Với đồ án này nên dùng **RMSE / MAE**, và nếu muốn dùng chỉ số phần trăm thì nên dùng **WAPE** (tổng sai số tuyệt đối / tổng nhu cầu thực tế) — ổn định hơn nhiều so với MAPE.
2. **Hai mô hình họ cây không vượt trội rõ rệt** so với các mô hình thống kê: RMSE tốt nhất là 92.42 (Random Forest) so với ~107 của nhóm ARIMA/SARIMA (chênh lệch không lớn so với độ biến động của chuỗi). Với dữ liệu nhiều nhiễu, chênh lệch này chưa đủ để kết luận chắc chắn.
3. **Chỉ đánh giá trên 1 SKU (S001 – P0001)**: kết quả trong `model_metadata.json` không đại diện cho toàn bộ 5 cửa hàng × 20 sản phẩm. Khi xem SKU khác, backend sẽ huấn luyện lại tại thời điểm gọi API (`src/forecast_service.py`), nên chỉ số có thể khác.
4. **Dự báo nhiều bước của mô hình học máy là dự báo đệ quy** (recursive multi-step): dùng dự báo của bước trước làm `lag` cho bước sau, nên sai số tích lũy dần theo horizon — càng dự báo xa (14, 30 ngày) càng kém tin cậy.
5. **Giả định của công thức tồn kho khá đơn giản**: `SS = Z × σ_D × √L` giả định nhu cầu phân phối chuẩn và lead time cố định (3 ngày). Thực tế lead time biến động sẽ làm SS/ROP bị đánh giá thấp.
6. **Bộ dữ liệu nhiều khả năng là dữ liệu mô phỏng** (các cột như `Demand Forecast`, `Units Ordered` không được dùng làm đặc trưng), nên kết luận không nên suy rộng sang dữ liệu bán lẻ thật.

---

## 8. Ghi chú thêm

- `frontend/vendor/` chứa bản local của ApexCharts và Lucide để dashboard chạy được cả khi offline / khi trình duyệt chặn CDN.
