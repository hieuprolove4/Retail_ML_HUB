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

## 3. Các mô hình, đối chứng bổ sung và kết quả đánh giá

Chỉ số dưới đây **đọc trực tiếp** từ `saved_models/model_metadata.json` và `reports/extra_diagnostics_s001_p0001.json` (SKU `S001–P0001`, Train 560 ngày / Test 141 ngày theo thời gian). Định nghĩa chỉ số nằm ở `src/forecasting_models.py:33-46`.

| Mô hình / Thí nghiệm | Giao thức | RMSE | MAE | MAPE (%) |
|---|---|---|---|---|
| **Random Forest (28 biến)** | 1 bước (`Inventory Level` đã biết) | **92,42** | **73,71** | 173,29 |
| XGBoost (28 biến) | 1 bước (`Inventory Level` đã biết) | 95,73 | 76,24 | 181,99 |
| SARIMA `(1,1,1)(1,0,1)7` | Đa bước 141 ngày | 107,03 | 90,03 | 221,23 |
| ARIMA `(2,1,1)` | Đa bước 141 ngày | 107,07 | 90,04 | 221,29 |
| Naive (`y_last = 145`) | Đa bước 141 ngày (cơ sở) | 107,52 | 91,48 | 231,07 |
| Exponential Smoothing | Đa bước 141 ngày | 108,21 | 90,23 | 219,37 |
| RF (27 biến, **bỏ Inventory Level**) | 1 bước (ablation) | 113,19 | 93,75 | 229,77 |
| XGBoost (27 biến, **bỏ Inventory Level**) | 1 bước (ablation) | 114,29 | 92,47 | 216,71 |
| XGBoost (Đệ quy thuần 141 bước) | Đa bước 141 ngày (cùng chuẩn Web) | 116,85 | 101,89 | 291,11 |
| Random Forest (Đệ quy thuần 141 bước) | Đa bước 141 ngày (cùng chuẩn Web) | 119,09 | 104,91 | 304,03 |
| Seasonal Naive (`s = 7`) | Đa bước 141 ngày (cơ sở) | 197,64 | 154,55 | 344,00 |

Ghi chú kỹ thuật & kiểm định (`src/extra_diagnostics.py`):

- **Ljung-Box** trên phần dư huấn luyện: ARIMA `Q(7)=3,02 (p=0,8827)`, SARIMA `Q(7)=3,03 (p=0,8818)` ($\alpha = 0,05$: không bác bỏ giả thuyết không về tính độc lập của phần dư).
- **Diebold-Mariano (hiệu chỉnh Harvey)**:
  - `RF (28 biến, 1 bước)` vs `RF (27 biến, bỏ Inv)`: `DM = -4,03` ($p < 0,0001$) — khẳng định mô hình phụ thuộc mạnh vào `Inventory Level` (chiếm `42,33%` tầm quan trọng).
  - `RF (27 biến)` vs `SARIMA`: `DM = +2,99` ($p = 0,0033$) — bỏ tồn kho làm RF thua SARIMA có ý nghĩa.
  - `RF (Đệ quy thuần 141 bước)` vs `SARIMA`: `DM = +3,42` ($p = 0,0008$) — ở bài toán đa bước thuần, mô hình đơn biến ít tham số ổn định hơn.
  - `RF (1 bước)` vs `XGBoost (1 bước)`: `DM = -1,78` ($p = 0,0769$) và `ARIMA` vs `SARIMA`: `DM = 0,47` ($p = 0,6365$) — không có ý nghĩa thống kê ở mức 5%.
- **Thông số tồn kho (`src/inventory_optimizer.py`)**: Lead time $L = 3$ ngày, Service Level $95\%$ ($Z = 1,645$), mục tiêu đặt bù $S = 1,5 \times \text{ROP}$ (quy tắc $(s,S)$). Với $\hat{D} = 155,1$ (dự báo 3 ngày đệ quy RF trên Web): `SS = 311`, `ROP = 777` (`776,27` chưa làm tròn), nên `Q = ⌈1,5 × 777 − 223⌉ = 943` (hoặc `942` theo `ROP` chưa làm tròn `776,27`). Nguồn tham chiếu 90 ngày trên Web (`D = 136,50`) cho `ROP = 721`; biểu đồ tĩnh `Hình 4.11` (`main_train_and_evaluate.py`, `D_test = 131,50`, `Z = 1,65`) cho `SS = 312, ROP = 707`.

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
│   ├── extra_diagnostics.py     # Baselines Naive/SNaive, ablation bỏ Inventory Level, đệ quy 141 bước, Ljung-Box & DM
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

Đây là các giới hạn phương pháp luận được trình bày thẳng thắn trong báo cáo Word (Mục 4.3.5 & Chương 6):

1. **Lệch giao thức đánh giá giữa ML (1 bước) và Thống kê (141 bước)**: trong `src/data_loader.py`, `lag_1..lag_30` và biến `rolling` trên 141 ngày kiểm thử của Random Forest/XGBoost được tính từ `Units Sold` thực tế (one-step-ahead), còn ARIMA/SARIMA/Holt-Winters/Naive dự báo liên tục 141 bước không biết trung gian. Khi kiểm thử bằng **dự báo đệ quy thuần 141 bước** (cơ chế tạo $\hat{D} = 155,1$ cho tồn kho), Random Forest có `RMSE = 119,09` (kém `SARIMA 107,03`, `DM = +3,42`, `p = 0,0008`).
2. **Phụ thuộc mạnh vào biến `Inventory Level` (42,33% tầm quan trọng)**: cột tồn kho được ghi nhận đầu ngày (trước khi bán), nhưng trong bộ dữ liệu mô phỏng Kaggle sản lượng bán bị chặn trên bởi tồn kho. Khi bỏ cột này (27 biến), `RF RMSE` tăng lên `113,19` (`DM = -4,03`, `p < 0,0001`).
3. **MAPE rất cao (173% – 344%)**: do có nhiều ngày nhu cầu gần bằng 0 ở mẫu số; kết luận chính dựa vào **RMSE / MAE** (tương lai nên bổ sung `WAPE` hoặc `sMAPE`).
4. **Chỉ đánh giá sâu 1 SKU (`S001–P0001`) trên 1 lần chia `560/141` ngày**: chưa áp dụng `rolling-origin backtesting` trên toàn bộ `100` cặp `Store × SKU`.
5. **Ma trận ABC–XYZ suy biến và chưa mô phỏng tồn kho đóng vòng**: dưới ngưỡng cố định `CV > 0,70`, toàn bộ `20` SKU tại `S001` (`CV` `0,758–0,834`) dồn hết vào cột `Z`; chính sách `SS`/`ROP`/`1,5 × ROP` mới dừng ở công thức tĩnh, chưa đo lường chi phí lưu kho hay tỷ lệ thiếu hàng thực tế qua thời gian.
6. **Bộ dữ liệu mô phỏng công khai (Kaggle — `Retail Store Inventory Forecasting Dataset`)**: cột `Inventory Level` sinh trong đoạn `[50, 500]` (1 ngày) nên nằm thấp hơn ngưỡng `ROP` (`707–777`) thiết kế cho Lead Time $L = 3$ ngày.

---

## 8. Ghi chú thêm

- `frontend/vendor/` chứa bản local của ApexCharts và Lucide để dashboard chạy được cả khi offline / khi trình duyệt chặn CDN.
