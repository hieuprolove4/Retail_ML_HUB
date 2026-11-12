"""
Configuration file: config.py
Chứa toàn bộ cấu hình, tham số mặc định và đường dẫn thư mục cho dự án.
"""

import os

# Thư mục gốc dự án
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Đường dẫn dữ liệu
DATA_PATH = os.path.join(BASE_DIR, "Dataset", "retail_store_inventory.csv")

# Thư mục lưu kết quả và biểu đồ báo cáo
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
FIGURES_EDA_DIR = os.path.join(REPORTS_DIR, "figures", "eda")
FIGURES_DIAGNOSTICS_DIR = os.path.join(REPORTS_DIR, "figures", "diagnostics")
FIGURES_MODELS_DIR = os.path.join(REPORTS_DIR, "figures", "models")
FIGURES_INVENTORY_DIR = os.path.join(REPORTS_DIR, "figures", "inventory")

# Thư mục lưu mô hình đã huấn luyện
SAVED_MODELS_DIR = os.path.join(BASE_DIR, "saved_models")

# Tạo sẵn các thư mục cần thiết
for d in [REPORTS_DIR, FIGURES_EDA_DIR, FIGURES_DIAGNOSTICS_DIR, FIGURES_MODELS_DIR, FIGURES_INVENTORY_DIR, SAVED_MODELS_DIR]:
    os.makedirs(d, exist_ok=True)

# Tham số mặc định cho bài toán Quản trị Tồn kho (Supply Chain & Inventory Management)
LEAD_TIME_DAYS = 3            # Thời gian giao hàng L = 3 ngày
SERVICE_LEVEL = 0.95          # Mức độ dịch vụ 95%
Z_SCORE = 1.65                # Z-score tương ứng với Service Level 95%
TRAIN_RATIO = 0.80            # Tỷ lệ phân chia tập Train / Test = 80 / 20

# Cấu hình dự báo mặc định
DEFAULT_STORE_ID = "S001"
DEFAULT_PRODUCT_ID = "P0001"
FORECAST_HORIZONS = [7, 14, 30] # Số ngày dự báo tương lai
