"""
Flowchart & BPMN Generator: generate_flowcharts.py
Xuất đồng thời bộ Flowchart chuẩn BPMN Swimlane 3 làn dọc (học theo mẫu BPMN_Order_Payment_Compact.drawio):
1. 5 file .drawio riêng biệt (Chương 1, Chương 4 và Hình 5.1–5.3 Chương 5) + 1 file .drawio gộp 5 tab tại: reports/figures/flowcharts/*.drawio
2. 5 file ảnh .png độ phân giải 300 DPI phục vụ chèn vào Báo cáo Word & Excel tại: reports/figures/flowcharts/*.png
"""

import base64
import io
import os
import sys
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Ellipse, FancyArrowPatch, FancyBboxPatch, Polygon, Rectangle

if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")

plt.rcParams["font.sans-serif"] = ["Segoe UI", "Arial", "DejaVu Sans"]
plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["figure.dpi"] = 300

OUTPUT_DIR = os.path.join("reports", "figures", "flowcharts")
ICON_DIR = os.path.join(OUTPUT_DIR, "icons")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(ICON_DIR, exist_ok=True)

_ICON_DATA_CACHE = {}


# ---------------------------------------------------------------------------
# ĐỊNH NGHĨA BỐ CỤC CHUẨN BPMN SWIMLANE 3 LÀN DỌC (SINGLE SOURCE OF TRUTH)
# ---------------------------------------------------------------------------
def build_diagrams():
    """Trả về đặc tả của cả 5 sơ đồ BPMN Swimlane (Chương 1, Chương 4 và 3 sơ đồ Chương 5) dùng chung cho .drawio và .png"""

    # === SƠ ĐỒ CHƯƠNG 1: KIẾN TRÚC KHÁM PHÁ DỮ LIỆU ĐA CHIỀU (8 BƯỚC ĐẶC THÙ BÀI TOÁN) ===
    d_ch1 = {
        "id": "bpmn_ch1",
        "layout": "pipeline",
        "tab_name": "Ch1. Dữ liệu & EDA",
        "drawio_file": "BPMN_0a_Chuong_1_Quy_Trinh_Kham_Pha_Du_Lieu_EDA.drawio",
        "png_file": "chuong_1_flowchart_quy_trinh_kham_pha_du_lieu_eda.png",
        "png_title": "Sơ đồ Quy trình Thu thập, Làm sạch & Khám phá Dữ liệu Đa chiều (Chương 1 — EDA)",
        "pool_w": 920,
        "pool_h": 760,
        "groups": [
            {"id": "g_src", "x": 30, "y": 35, "w": 230, "h": 460, "fill": "#DAE8FC", "stroke": "#3A546D", "title": "NGUỒN DỮ LIỆU\nBÁN LẺ", "title_pos": "top_center"},
            {"id": "g_prep", "x": 295, "y": 35, "w": 595, "h": 175, "fill": "#FFF2CC", "stroke": "#3A546D", "title": "TIỀN XỬ LÝ & LÀM SẠCH DỮ LIỆU", "title_pos": "top_center"},
            {"id": "g_eda", "x": 295, "y": 240, "w": 595, "h": 255, "fill": "#E2F0D9", "stroke": "#3A546D", "title": "PHÂN TÍCH KHÁM PHÁ ĐA CHIỀU — EDA", "title_pos": "top_left"},
            {"id": "g_out", "x": 30, "y": 530, "w": 860, "h": 190, "fill": "#FBE5D6", "stroke": "#3A546D", "title": "TỔNG HỢP ĐẶC TRƯNG & KẾT LUẬN EDA", "title_pos": "bottom_center"},
        ],
        "nodes": [
            {"id": "n1", "num": "1", "x": 50, "y": 95, "w": 190, "h": 145, "fill": "#FFFFFF", "style": "vertical", "icon": "system",
             "title": "Hệ thống Cửa hàng", "desc": "Hoạt động bán hàng\n& quản lý tồn kho"},
            {"id": "n2", "num": "2", "x": 50, "y": 295, "w": 190, "h": 150, "fill": "#FFFFFF", "style": "vertical", "icon": "database",
             "title": "Bộ dữ liệu Bán lẻ", "desc": "Dữ liệu giao dịch\ntheo ngày & SKU"},
            {"id": "n3", "num": "3", "x": 320, "y": 80, "w": 255, "h": 110, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "clean",
             "title": "Làm sạch &\nChuẩn hóa", "desc": "Ép kiểu thời gian,\nkiểm tra giá trị thiếu"},
            {"id": "n4", "num": "4", "x": 610, "y": 80, "w": 255, "h": 110, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "schema",
             "title": "Chuẩn hóa\n15 Thuộc tính", "desc": "Xác định biến mục tiêu\n& loại cột rò rỉ"},
            {"id": "n5", "num": "5", "x": 312, "y": 300, "w": 176, "h": 155, "fill": "#FFFFFF", "style": "vertical", "icon": "eval",
             "title": "Xu hướng & Quy mô", "desc": "Sản lượng theo\nThời gian & Cửa hàng"},
            {"id": "n6", "num": "6", "x": 504, "y": 300, "w": 176, "h": 155, "fill": "#FFFFFF", "style": "vertical", "icon": "weather",
             "title": "Mùa vụ & Ngoại cảnh", "desc": "Tác động Thời tiết,\nLễ Tết & Khuyến mãi"},
            {"id": "n7", "num": "7", "x": 696, "y": 300, "w": 176, "h": 155, "fill": "#FFFFFF", "style": "vertical", "icon": "price",
             "title": "Phân tích Giá & Kho", "desc": "Độ co giãn giá bán\n& Giá đối thủ"},
            {"id": "n8", "num": "8", "x": 185, "y": 560, "w": 550, "h": 110, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "team",
             "title": "Ma trận Tương quan & Định hướng Biến Đặc trưng", "desc": "Tổng hợp bộ biểu đồ EDA & chọn biến đầu vào cho mô hình dự báo (Chương 4)"},
        ],
        "arrows": [
            {"id": "a1", "src": "n1", "tgt": "n2", "pts": [(145, 240), (145, 295)], "color": "#1F4E79"},
            {"id": "a2", "src": "n2", "tgt": "n3", "pts": [(240, 370), (278, 370), (278, 135), (320, 135)], "color": "#1F4E79"},
            {"id": "a3", "src": "n3", "tgt": "n4", "pts": [(575, 135), (610, 135)], "color": "#1F4E79"},
            {"id": "a4", "src": "n4", "tgt": "n5", "pts": [(738, 190), (738, 280), (400, 280), (400, 300)], "color": "#7030A0"},
            {"id": "a5", "src": "n4", "tgt": "n6", "pts": [(738, 190), (738, 280), (592, 280), (592, 300)], "color": "#7030A0"},
            {"id": "a6", "src": "n4", "tgt": "n7", "pts": [(738, 190), (738, 280), (784, 280), (784, 300)], "color": "#7030A0"},
            {"id": "a7", "src": "n5", "tgt": "n8", "pts": [(400, 455), (400, 560)], "color": "#1F4E79"},
            {"id": "a8", "src": "n6", "tgt": "n8", "pts": [(592, 455), (592, 560)], "color": "#1F4E79"},
            {"id": "a9", "src": "n7", "tgt": "n8", "pts": [(784, 455), (784, 512), (685, 512), (685, 560)], "color": "#1F4E79"},
            {"id": "a10", "src": "n8", "tgt": "n1", "pts": [(185, 615), (10, 615), (10, 168), (50, 168)], "color": "#7030A0", "ls": ":"},
        ],
    }

    # === SƠ ĐỒ CHƯƠNG 4: KIẾN TRÚC ĐỐI SÁNH 2 NHÓM MÔ HÌNH & TỐI ƯU TỒN KHO (10 BƯỚC) ===
    d_ch4 = {
        "id": "bpmn_ch4",
        "layout": "pipeline",
        "tab_name": "Ch4. ML & Tồn kho",
        "drawio_file": "BPMN_0b_Chuong_4_Quy_Trinh_Huan_Luyen_5_Mo_Hinh.drawio",
        "png_file": "chuong_4_flowchart_quy_trinh_huan_luyen_5_mo_hinh.png",
        "png_title": "Sơ đồ Quy trình Đối sánh 2 Nhóm Mô hình Dự báo & Tối ưu Tồn kho (Chương 4)",
        "pool_w": 920,
        "pool_h": 875,
        "groups": [
            {"id": "g_prep", "x": 30, "y": 30, "w": 860, "h": 175, "fill": "#FFF2CC", "stroke": "#3A546D", "title": "CHUẨN BỊ CHUỖI THỜI GIAN & KỸ THUẬT ĐẶC TRƯNG", "title_pos": "top_center"},
            {"id": "g_model", "x": 30, "y": 235, "w": 860, "h": 405, "fill": "#E2F0D9", "stroke": "#3A546D", "title": "HUẤN LUYỆN ĐỐI SÁNH 2 NHÓM MÔ HÌNH & ĐÁNH GIÁ", "title_pos": "top_center"},
            {"id": "g_inv", "x": 30, "y": 670, "w": 860, "h": 175, "fill": "#FBE5D6", "stroke": "#3A546D", "title": "TỐI ƯU HÓA TỒN KHO & HỖ TRỢ QUYẾT ĐỊNH", "title_pos": "bottom_center"},
        ],
        "nodes": [
            {"id": "n1", "num": "1", "x": 50, "y": 72, "w": 245, "h": 112, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "database",
             "title": "Trích xuất\nChuỗi SKU", "desc": "Đồng bộ chuỗi thời gian\nliên tục theo ngày"},
            {"id": "n2", "num": "2", "x": 338, "y": 72, "w": 245, "h": 112, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "clean",
             "title": "Kiểm định\nTính dừng", "desc": "Phân rã mùa vụ, ACF/PACF\n& kiểm định ADF / KPSS"},
            {"id": "n3", "num": "3", "x": 625, "y": 72, "w": 245, "h": 112, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "schema",
             "title": "Tạo Đặc trưng\n& Chia tập", "desc": "Biến trễ, trượt, Giá\n& chia tập Train / Test"},
            {"id": "n4", "num": "4", "x": 65, "y": 290, "w": 340, "h": 104, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "gears",
             "title": "Nhóm Thống kê Chuỗi thời gian", "desc": "ARIMA, SARIMA & Holt-Winters\n(Mô hình chuỗi đơn biến)"},
            {"id": "n5", "num": "5", "x": 515, "y": 290, "w": 340, "h": 104, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "brain",
             "title": "Nhóm Học máy Đa biến", "desc": "Random Forest & XGBoost\n(Học đa biến ngoại sinh)"},
            {"id": "n6", "num": "6", "x": 270, "y": 424, "w": 380, "h": 90, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "eval",
             "title": "Đánh giá & So sánh Tập Kiểm thử", "desc": "Đối chiếu sai số RMSE, MAE, MAPE giữa các mô hình"},
            {"id": "n7", "num": "7", "x": 270, "y": 552, "w": 380, "h": 74, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "save",
             "title": "Chọn Mô hình Tốt nhất & Lưu (.pkl)", "desc": "Trích xuất đặc trưng quan trọng & lưu Random Forest"},
            {"id": "n8", "num": "8", "x": 50, "y": 698, "w": 245, "h": 112, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "load",
             "title": "Dự báo Nhu cầu\nTương lai", "desc": "Nạp mô hình tốt nhất dự báo\nnhu cầu ngắn hạn theo SKU"},
            {"id": "n9", "num": "9", "x": 338, "y": 698, "w": 245, "h": 112, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "filter",
             "title": "Tính Tồn kho\nAn toàn & ROP", "desc": "Xác định ngưỡng SS, điểm đặt\nhàng ROP & phân loại ABC–XYZ"},
            {"id": "n10", "num": "10", "x": 625, "y": 698, "w": 245, "h": 112, "fill": "#FFFFFF", "style": "compact_horiz", "icon": "team",
             "title": "Cảnh báo &\nĐề xuất Đặt hàng", "desc": "Tích hợp lên Web Studio\n& xuất danh sách đặt hàng"},
        ],
        "arrows": [
            {"id": "a1", "src": "n1", "tgt": "n2", "pts": [(295, 128), (338, 128)], "color": "#1F4E79"},
            {"id": "a2", "src": "n2", "tgt": "n3", "pts": [(583, 128), (625, 128)], "color": "#1F4E79"},
            {"id": "a3", "src": "n3", "tgt": "n4", "pts": [(748, 184), (748, 220), (195, 220), (195, 290)], "color": "#1F4E79"},
            {"id": "a4", "src": "n3", "tgt": "n5", "pts": [(748, 184), (748, 220), (710, 220), (710, 290)], "color": "#1F4E79"},
            {"id": "a5", "src": "n4", "tgt": "n6", "pts": [(340, 394), (340, 424)], "color": "#1F4E79"},
            {"id": "a6", "src": "n5", "tgt": "n6", "pts": [(580, 394), (580, 424)], "color": "#1F4E79"},
            {"id": "a7", "src": "n6", "tgt": "n4", "pts": [(270, 469), (150, 394)], "color": "#7030A0", "rad": -0.34, "wp": [(150, 469)]},
            {"id": "a8", "src": "n6", "tgt": "n5", "pts": [(650, 469), (770, 394)], "color": "#7030A0", "rad": 0.34, "wp": [(770, 469)]},
            {"id": "a9", "src": "n6", "tgt": "n7", "pts": [(460, 514), (460, 552)], "color": "#7030A0", "ls": ":"},
            {"id": "a10", "src": "n7", "tgt": "n8", "pts": [(460, 626), (460, 654), (172, 654), (172, 698)], "color": "#1F4E79"},
            {"id": "a11", "src": "n8", "tgt": "n9", "pts": [(295, 754), (338, 754)], "color": "#1F4E79"},
            {"id": "a12", "src": "n9", "tgt": "n10", "pts": [(583, 754), (625, 754)], "color": "#1F4E79"},
        ],
    }

    # === SƠ ĐỒ 1: KHỞI ĐỘNG HỆ THỐNG, TỔNG QUAN & EDA ===
    d1 = {
        "id": "bpmn_1",
        "tab_name": "1. Tổng quan & EDA",
        "drawio_file": "BPMN_1_Tong_Quan_Va_EDA.drawio",
        "png_file": "chuong_5_hinh_5_1_flowchart_tong_quan_he_thong.png",
        "png_title": "Hình 5.1: Sơ đồ BPMN Quy trình Khởi động Hệ thống & Phân tích Tổng quan / EDA",
        "pool_title": "QUY TRÌNH KHỞI ĐỘNG HỆ THỐNG & PHÂN TÍCH TỔNG QUAN / EDA (RETAIL ML HUB)",
        "pool_w": 1020,
        "pool_h": 700,
        "lanes": [
            {"id": "lane_user", "title": "Quản lý / Người dùng", "x": 30, "w": 310, "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "lane_web", "title": "Hệ thống Web & FastAPI Backend", "x": 340, "w": 400, "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "lane_data", "title": "Kho Dữ liệu CSV & RAM Cache", "x": 740, "w": 280, "fill": "#f8cecc", "stroke": "#b85450"},
        ],
        "nodes": [
            {"id": "start", "parent": "lane_user", "kind": "start", "x": 125, "y": 48, "w": 60, "h": 42, "label": "Bắt đầu", "fill": "#d5e8d4", "stroke": "#82b366", "bold": True},
            {"id": "u_open", "parent": "lane_user", "kind": "task", "x": 60, "y": 118, "w": 190, "h": 52, "label": "Truy cập Web Dashboard\n(http://127.0.0.1:8000)", "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "s_api_kpi", "parent": "lane_web", "kind": "task", "x": 105, "y": 118, "w": 190, "h": 52, "label": "Gửi request /api/kpi,\n/api/eda-data & /api/options", "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "d_ram_load", "parent": "lane_data", "kind": "db", "x": 35, "y": 108, "w": 210, "h": 72, "label": "retail_store_inventory.csv\nNạp 73.100 dòng vào RAM Cache\n& Warm-up cửa hàng S001", "fill": "#f8cecc", "stroke": "#b85450"},
            {"id": "s_render_eda", "parent": "lane_web", "kind": "task", "x": 105, "y": 224, "w": 190, "h": 54, "label": "Tổng hợp 5 KPI & Render\n5 biểu đồ EDA (ApexCharts)", "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "u_view_eda", "parent": "lane_user", "kind": "task", "x": 60, "y": 224, "w": 190, "h": 54, "label": "Xem 5 thẻ KPI & 5 biểu đồ\nEDA tương tác", "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "u_filter", "parent": "lane_user", "kind": "task", "x": 60, "y": 324, "w": 190, "h": 54, "label": "Chọn bộ lọc Store, SKU\n& Chu kỳ (Ngày / Tuần / Tháng)", "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "s_api_trend", "parent": "lane_web", "kind": "task", "x": 105, "y": 324, "w": 190, "h": 54, "label": "Gọi /api/sales-trend & Tính\nđường TB trượt (MA 4/12 kỳ)", "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "d_filter_ram", "parent": "lane_data", "kind": "db", "x": 35, "y": 315, "w": 210, "h": 72, "label": "In-Memory DataFrame\nLọc chuỗi bán hàng theo\nStore / SKU / Chu kỳ", "fill": "#f8cecc", "stroke": "#b85450"},
            {"id": "gw_valid", "parent": "lane_web", "kind": "gateway", "x": 155, "y": 424, "w": 90, "h": 72, "label": "Dữ liệu\nhợp lệ?", "fill": "#fff2cc", "stroke": "#d6b656", "bold": True},
            {"id": "s_empty_err", "parent": "lane_web", "kind": "task", "x": 16, "y": 435, "w": 112, "h": 50, "label": "Báo bộ lọc trống\n(Yêu cầu chọn lại)", "fill": "#f8cecc", "stroke": "#b85450"},
            {"id": "s_update_chart", "parent": "lane_web", "kind": "task", "x": 105, "y": 532, "w": 190, "h": 52, "label": "Cập nhật lại biểu đồ\nxu hướng tức thời", "fill": "#d5e8d4", "stroke": "#82b366", "bold": True},
            {"id": "u_done_view", "parent": "lane_user", "kind": "task", "x": 60, "y": 532, "w": 190, "h": 52, "label": "Quan sát xu hướng & Xuất\nảnh biểu đồ báo cáo (nếu cần)", "fill": "#d5e8d4", "stroke": "#82b366", "bold": True},
            {"id": "end", "parent": "lane_user", "kind": "end", "x": 125, "y": 622, "w": 60, "h": 42, "label": "Kết thúc", "fill": "#f8cecc", "stroke": "#b85450", "bold": True},
        ],
        "edges": [
            {"id": "e1", "parent": "lane_user", "source": "start", "target": "u_open", "pts": [(185, 90), (185, 118)]},
            {"id": "e2", "parent": "pool", "source": "u_open", "target": "s_api_kpi", "pts": [(280, 144), (445, 144)]},
            {"id": "e3", "parent": "pool", "source": "s_api_kpi", "target": "d_ram_load", "pts": [(635, 144), (775, 144)]},
            {"id": "e4", "parent": "pool", "source": "d_ram_load", "target": "s_render_eda", "pts": [(880, 180), (880, 251), (635, 251)]},
            {"id": "e5", "parent": "pool", "source": "s_render_eda", "target": "u_view_eda", "pts": [(445, 251), (280, 251)]},
            {"id": "e6", "parent": "lane_user", "source": "u_view_eda", "target": "u_filter", "pts": [(185, 278), (185, 324)]},
            {"id": "e7", "parent": "pool", "source": "u_filter", "target": "s_api_trend", "pts": [(280, 351), (445, 351)]},
            {"id": "e8", "parent": "pool", "source": "s_api_trend", "target": "d_filter_ram", "pts": [(635, 351), (775, 351)]},
            {"id": "e9", "parent": "pool", "source": "d_filter_ram", "target": "gw_valid", "pts": [(880, 387), (880, 460), (585, 460)]},
            {"id": "e10", "parent": "lane_web", "source": "gw_valid", "target": "s_empty_err", "label": "Không", "lcolor": "#b85450", "lpos": (482, 444), "pts": [(495, 460), (468, 460)]},
            {"id": "e11", "parent": "pool", "source": "s_empty_err", "target": "u_filter", "label": "Chọn lại", "lcolor": "#b85450", "lpos": (310, 412), "pts": [(356, 460), (310, 460), (310, 366), (280, 366)]},
            {"id": "e12", "parent": "lane_web", "source": "gw_valid", "target": "s_update_chart", "label": "Có", "lcolor": "#006600", "lpos": (558, 512), "pts": [(540, 496), (540, 532)]},
            {"id": "e13", "parent": "pool", "source": "s_update_chart", "target": "u_done_view", "pts": [(445, 558), (280, 558)]},
            {"id": "e14", "parent": "lane_user", "source": "u_done_view", "target": "end", "pts": [(185, 584), (185, 622)]},
        ],
    }

    # === SƠ ĐỒ 2: DỰ BÁO NHU CẦU SKU & SO SÁNH 5 MÔ HÌNH ===
    d2 = {
        "id": "bpmn_2",
        "tab_name": "2. Dự báo & 5 Mô hình",
        "drawio_file": "BPMN_2_Du_Bao_Nhu_Cau_Va_So_Sanh_Mo_Hinh.drawio",
        "png_file": "chuong_5_hinh_5_2_flowchart_du_bao_va_so_sanh_mo_hinh.png",
        "png_title": "Hình 5.2: Sơ đồ BPMN Quy trình Dự báo Nhu cầu SKU & So sánh 5 Mô hình Học máy",
        "pool_title": "QUY TRÌNH DỰ BÁO NHU CẦU SKU & SO SÁNH 5 MÔ HÌNH HỌC MÁY",
        "pool_w": 1060,
        "pool_h": 910,
        "lanes": [
            {"id": "lane_user", "title": "Quản lý / Người dùng", "x": 30, "w": 300, "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "lane_web", "title": "Hệ thống Web & FastAPI Backend", "x": 330, "w": 440, "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "lane_ml", "title": "Động cơ ML & Kho Mô hình", "x": 770, "w": 290, "fill": "#f8cecc", "stroke": "#b85450"},
        ],
        "nodes": [
            {"id": "start", "parent": "lane_user", "kind": "start", "x": 120, "y": 46, "w": 60, "h": 42, "label": "Bắt đầu", "fill": "#d5e8d4", "stroke": "#82b366", "bold": True},
            {"id": "u_choose", "parent": "lane_user", "kind": "task", "x": 50, "y": 114, "w": 200, "h": 54, "label": "Chọn Store ID, Product ID,\nMô hình & Kỳ hạn (7/14/30 ngày)", "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "u_run", "parent": "lane_user", "kind": "task", "x": 50, "y": 196, "w": 200, "h": 50, "label": "Nhấn \"Chạy dự báo\" hoặc\n\"Đánh giá 5 mô hình\"", "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "s_recv", "parent": "lane_web", "kind": "task", "x": 115, "y": 196, "w": 210, "h": 50, "label": "Nhận POST /api/forecast hoặc\nGET /api/model-comparison", "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "ml_prep", "parent": "lane_ml", "kind": "db", "x": 38, "y": 185, "w": 214, "h": 72, "label": "data_loader.py\nChuẩn hóa chuỗi ngày ('D')\n& Tạo 28 biến đặc trưng", "fill": "#f8cecc", "stroke": "#b85450"},
            {"id": "gw_pkl", "parent": "lane_web", "kind": "gateway", "x": 153, "y": 291, "w": 134, "h": 78, "label": "SKU S001–P0001\ncó sẵn .pkl?", "fill": "#fff2cc", "stroke": "#d6b656", "bold": True},
            {"id": "ml_load_pkl", "parent": "lane_ml", "kind": "db", "x": 38, "y": 294, "w": 214, "h": 72, "label": "saved_models/*.pkl\nNạp mô hình đã lưu &\nmodel_metadata.json", "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "ml_train_live", "parent": "lane_ml", "kind": "task", "x": 38, "y": 392, "w": 214, "h": 58, "label": "forecasting_models.py\nHuấn luyện trực tiếp SKU\n(Chia Train 80% / Test 20%)", "fill": "#f8cecc", "stroke": "#b85450"},
            {"id": "gw_type", "parent": "lane_web", "kind": "gateway", "x": 165, "y": 462, "w": 110, "h": 76, "label": "Loại mô hình\ndự báo?", "fill": "#fff2cc", "stroke": "#d6b656", "bold": True},
            {"id": "s_rec_ml", "parent": "lane_web", "kind": "task", "x": 18, "y": 572, "w": 192, "h": 64, "label": "Nhóm ML (RF / XGBoost)\nDự báo đệ quy từng ngày\n(Lấy dự báo t làm lag t+1)", "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "s_fit_ts", "parent": "lane_web", "kind": "task", "x": 230, "y": 572, "w": 192, "h": 64, "label": "Nhóm Thống kê\n(ARIMA / SARIMA / ES)\nFit lịch sử & Dự báo H bước", "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "s_metrics", "parent": "lane_web", "kind": "task", "x": 100, "y": 672, "w": 240, "h": 56, "label": "Tính Khoảng tin cậy 95%, Xếp hạng\nRMSE / MAE / MAPE & Feature Importance", "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "s_show", "parent": "lane_web", "kind": "task", "x": 100, "y": 760, "w": 240, "h": 54, "label": "Hiển thị KPI, Biểu đồ 60 ngày\n+ Dự báo H ngày & Bảng chi tiết", "fill": "#d5e8d4", "stroke": "#82b366", "bold": True},
            {"id": "u_view", "parent": "lane_user", "kind": "task", "x": 50, "y": 760, "w": 200, "h": 54, "label": "Xem Biểu đồ dự báo, Bảng\nxếp hạng 5 mô hình & Xuất CSV", "fill": "#d5e8d4", "stroke": "#82b366", "bold": True},
            {"id": "end", "parent": "lane_user", "kind": "end", "x": 120, "y": 844, "w": 60, "h": 42, "label": "Kết thúc", "fill": "#f8cecc", "stroke": "#b85450", "bold": True},
        ],
        "edges": [
            {"id": "e1", "parent": "lane_user", "source": "start", "target": "u_choose", "pts": [(180, 88), (180, 114)]},
            {"id": "e2", "parent": "lane_user", "source": "u_choose", "target": "u_run", "pts": [(180, 168), (180, 196)]},
            {"id": "e3", "parent": "pool", "source": "u_run", "target": "s_recv", "pts": [(280, 221), (445, 221)]},
            {"id": "e4", "parent": "pool", "source": "s_recv", "target": "ml_prep", "pts": [(655, 221), (808, 221)]},
            {"id": "e5", "parent": "pool", "source": "ml_prep", "target": "gw_pkl", "pts": [(915, 257), (915, 275), (550, 275), (550, 292)]},
            {"id": "e6", "parent": "pool", "source": "gw_pkl", "target": "ml_load_pkl", "label": "Có", "lcolor": "#006600", "lpos": (700, 316), "pts": [(605, 330), (808, 330)]},
            {"id": "e7", "parent": "pool", "source": "gw_pkl", "target": "ml_train_live", "label": "Không", "lcolor": "#b85450", "lpos": (700, 407), "pts": [(550, 368), (550, 421), (808, 421)]},
            {"id": "e8", "parent": "pool", "source": "ml_load_pkl", "target": "gw_type", "pts": [(1022, 330), (1042, 330), (1042, 500), (605, 500)]},
            {"id": "e9", "parent": "pool", "source": "ml_train_live", "target": "gw_type", "pts": [(915, 450), (915, 500), (605, 500)]},
            {"id": "e10", "parent": "lane_web", "source": "gw_type", "target": "s_rec_ml", "label": "ML (RF / XGB)", "lcolor": "#006600", "lpos": (440, 530), "pts": [(495, 500), (444, 500), (444, 572)]},
            {"id": "e11", "parent": "lane_web", "source": "gw_type", "target": "s_fit_ts", "label": "Chuỗi thời gian", "lcolor": "#1E40AF", "lpos": (660, 530), "pts": [(605, 500), (656, 500), (656, 572)]},
            {"id": "e12", "parent": "lane_web", "source": "s_rec_ml", "target": "s_metrics", "pts": [(444, 636), (444, 654), (550, 654), (550, 672)]},
            {"id": "e13", "parent": "lane_web", "source": "s_fit_ts", "target": "s_metrics", "pts": [(656, 636), (656, 654), (550, 654), (550, 672)]},
            {"id": "e14", "parent": "lane_web", "source": "s_metrics", "target": "s_show", "pts": [(550, 728), (550, 760)]},
            {"id": "e15", "parent": "pool", "source": "s_show", "target": "u_view", "pts": [(430, 787), (280, 787)]},
            {"id": "e16", "parent": "lane_user", "source": "u_view", "target": "end", "pts": [(180, 814), (180, 844)]},
        ],
    }

    # === SƠ ĐỒ 3: TỐI ƯU TỒN KHO (SS/ROP), MA TRẬN ABC–XYZ & MÔ PHỎNG WHAT-IF ===
    d3 = {
        "id": "bpmn_3",
        "tab_name": "3. Tồn kho ROP & ABC-XYZ",
        "drawio_file": "BPMN_3_Toi_Uu_Ton_Kho_ROP_ABCXYZ_WhatIf.drawio",
        "png_file": "chuong_5_hinh_5_3_flowchart_ton_kho_rop_abcxyz_whatif.png",
        "png_title": "Hình 5.3: Sơ đồ BPMN Quy trình Tối ưu Tồn kho (SS/ROP), Ma trận ABC–XYZ & What-If",
        "pool_title": "QUY TRÌNH TỐI ƯU TỒN KHO (SS/ROP), MA TRẬN ABC–XYZ & MÔ PHỎNG WHAT-IF",
        "pool_w": 1090,
        "pool_h": 910,
        "lanes": [
            {"id": "lane_user", "title": "Quản lý Kho / Mua hàng", "x": 30, "w": 290, "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "lane_web", "title": "Hệ thống Web & FastAPI Backend", "x": 320, "w": 480, "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "lane_opt", "title": "Động cơ Tối ưu & Dữ liệu Dự báo", "x": 800, "w": 290, "fill": "#f8cecc", "stroke": "#b85450"},
        ],
        "nodes": [
            {"id": "start", "parent": "lane_user", "kind": "start", "x": 115, "y": 46, "w": 60, "h": 42, "label": "Bắt đầu", "fill": "#d5e8d4", "stroke": "#82b366", "bold": True},
            {"id": "u_params", "parent": "lane_user", "kind": "task", "x": 45, "y": 116, "w": 200, "h": 56, "label": "Chọn Cửa hàng, Lead Time L,\nService Level SL & Nguồn cầu D_hat", "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "s_req", "parent": "lane_web", "kind": "task", "x": 130, "y": 116, "w": 220, "h": 56, "label": "Gửi yêu cầu /api/inventory-status\n& /api/abc-xyz", "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "opt_demand", "parent": "lane_opt", "kind": "db", "x": 38, "y": 108, "w": 214, "h": 72, "label": "ForecastService\nTrích xuất cầu dự báo D_hat (ML)\n& độ lệch chuẩn σ_D 90 ngày", "fill": "#f8cecc", "stroke": "#b85450"},
            {"id": "s_calc_rop", "parent": "lane_web", "kind": "task", "x": 120, "y": 216, "w": 240, "h": 56, "label": "Tính SS = ceil(Z × σ_D × √L)\n& ROP = ceil(D_hat × L + SS)", "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "gw_rop", "parent": "lane_web", "kind": "gateway", "x": 185, "y": 304, "w": 110, "h": 76, "label": "So sánh I_curr\nvới ROP?", "fill": "#fff2cc", "stroke": "#d6b656", "bold": True},
            {"id": "s_crit", "parent": "lane_web", "kind": "task", "x": 14, "y": 428, "w": 142, "h": 80, "label": "NGUY CẤP (CRITICAL)\nĐặt hàng ngay:\nQ = max(0, ceil(\n1.5×ROP - I_curr))", "fill": "#f8cecc", "stroke": "#b85450", "bold": True},
            {"id": "s_warn", "parent": "lane_web", "kind": "task", "x": 169, "y": 430, "w": 142, "h": 76, "label": "CẢNH BÁO (WARNING)\nSắp chạm ngưỡng ROP\n(Đề xuất đặt Q = 0)", "fill": "#fff2cc", "stroke": "#d6b656", "bold": True},
            {"id": "s_safe", "parent": "lane_web", "kind": "task", "x": 324, "y": 430, "w": 142, "h": 76, "label": "AN TOÀN (SAFE)\nTồn kho đủ đáp ứng\n(Đề xuất đặt Q = 0)", "fill": "#d5e8d4", "stroke": "#82b366", "bold": True},
            {"id": "opt_abc", "parent": "lane_opt", "kind": "db", "x": 38, "y": 544, "w": 214, "h": 76, "label": "compute_abc_xyz_analysis()\nPhân loại 20 SKU theo % Doanh thu\n(ABC) & Hệ số biến thiên CV (XYZ)", "fill": "#f8cecc", "stroke": "#b85450"},
            {"id": "u_whatif", "parent": "lane_user", "kind": "task", "x": 45, "y": 646, "w": 200, "h": 56, "label": "Lọc ô Ma trận ABC–XYZ & Kéo\nthanh trượt What-If (L', SL')", "fill": "#d5e8d4", "stroke": "#82b366"},
            {"id": "s_whatif", "parent": "lane_web", "kind": "task", "x": 120, "y": 646, "w": 240, "h": 56, "label": "Tính độ nhạy What-If (ΔSS, ΔROP, ΔQ)\n& Tổng hợp kết quả hiển thị", "fill": "#fff2cc", "stroke": "#d6b656"},
            {"id": "u_export", "parent": "lane_user", "kind": "task", "x": 45, "y": 754, "w": 200, "h": 56, "label": "Xem Bảng cảnh báo, Biểu đồ 90 ngày\n& Xuất danh sách đặt hàng CSV", "fill": "#d5e8d4", "stroke": "#82b366", "bold": True},
            {"id": "end", "parent": "lane_user", "kind": "end", "x": 115, "y": 844, "w": 60, "h": 42, "label": "Kết thúc", "fill": "#f8cecc", "stroke": "#b85450", "bold": True},
        ],
        "edges": [
            {"id": "e1", "parent": "lane_user", "source": "start", "target": "u_params", "pts": [(175, 88), (175, 116)]},
            {"id": "e2", "parent": "pool", "source": "u_params", "target": "s_req", "pts": [(275, 144), (450, 144)]},
            {"id": "e3", "parent": "pool", "source": "s_req", "target": "opt_demand", "pts": [(670, 144), (838, 144)]},
            {"id": "e4", "parent": "pool", "source": "opt_demand", "target": "s_calc_rop", "pts": [(945, 180), (945, 244), (680, 244)]},
            {"id": "e5", "parent": "lane_web", "source": "s_calc_rop", "target": "gw_rop", "pts": [(560, 272), (560, 304)]},
            {"id": "e6", "parent": "lane_web", "source": "gw_rop", "target": "s_crit", "label": "I_curr ≤ ROP", "lcolor": "#b85450", "lpos": (415, 392), "pts": [(505, 342), (405, 342), (405, 430)]},
            {"id": "e7", "parent": "lane_web", "source": "gw_rop", "target": "s_warn", "label": "ROP < I_curr ≤ 1.2×ROP", "lcolor": "#996600", "lpos": (560, 405), "pts": [(560, 380), (560, 430)]},
            {"id": "e8", "parent": "lane_web", "source": "gw_rop", "target": "s_safe", "label": "I_curr > 1.2×ROP", "lcolor": "#006600", "lpos": (708, 392), "pts": [(615, 342), (715, 342), (715, 430)]},
            {"id": "e9", "parent": "pool", "source": "s_crit", "target": "opt_abc", "pts": [(405, 506), (405, 526), (945, 526), (945, 544)]},
            {"id": "e10", "parent": "pool", "source": "s_warn", "target": "opt_abc", "pts": [(560, 506), (560, 526), (945, 526), (945, 544)]},
            {"id": "e11", "parent": "pool", "source": "s_safe", "target": "opt_abc", "pts": [(715, 506), (715, 526), (945, 526), (945, 544)]},
            {"id": "e12", "parent": "pool", "source": "opt_abc", "target": "u_whatif", "pts": [(945, 620), (945, 632), (175, 632), (175, 646)]},
            {"id": "e13", "parent": "pool", "source": "u_whatif", "target": "s_whatif", "pts": [(275, 674), (440, 674)]},
            {"id": "e14", "parent": "pool", "source": "s_whatif", "target": "u_export", "pts": [(560, 702), (560, 782), (275, 782)]},
            {"id": "e15", "parent": "lane_user", "source": "u_export", "target": "end", "pts": [(175, 810), (175, 844)]},
        ],
    }

    return [d_ch1, d_ch4, d1, d2, d3]


# ---------------------------------------------------------------------------
# PHẦN 1: XUẤT FILE .DRAWIO XML CHUẨN (GIỐNG MẪU BPMN_Order_Payment_Compact)
# ---------------------------------------------------------------------------
def _xml_val(text):
    """Escape HTML/XML đặc biệt và chuyển newline thành &#xa; cho thuộc tính value trong Draw.io"""
    return escape(text, {'"': "&quot;"}).replace("\n", "&#xa;")


def _lane_map(diagram):
    return {lane["id"]: lane for lane in diagram["lanes"]}


def _node_style(node):
    base_font = ";fontStyle=1;" if node.get("bold") else ";"
    fill = node["fill"]
    stroke = node["stroke"]
    kind = node["kind"]
    if kind == "start":
        return f"ellipse;whiteSpace=wrap;html=1;aspect=fixed;fillColor={fill};strokeColor={stroke};fontSize=10{base_font}"
    if kind == "end":
        return f"ellipse;whiteSpace=wrap;html=1;aspect=fixed;fillColor={fill};strokeColor={stroke};strokeWidth=2;fontSize=10{base_font}"
    if kind == "gateway":
        return f"rhombus;whiteSpace=wrap;html=1;fillColor={fill};strokeColor={stroke};fontSize=9{base_font}"
    if kind == "db":
        return f"shape=cylinder3;whiteSpace=wrap;html=1;boundedLbl=1;backgroundOutline=1;size=12;fillColor={fill};strokeColor={stroke};fontSize=10{base_font}"
    return f"rounded=1;whiteSpace=wrap;html=1;fillColor={fill};strokeColor={stroke};fontSize=10{base_font}"


def _build_pipeline_mxgraph_xml(diagram, indent=""):
    """Sinh block <mxGraphModel> cho sơ đồ kiểu Pipeline đặc thù từng chương"""
    pw, ph = diagram["pool_w"], diagram["pool_h"]
    lines = [
        f'{indent}<mxGraphModel dx="1050" dy="920" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="{pw + 40}" pageHeight="{ph + 40}" math="0" shadow="1">',
        f"{indent}  <root>",
        f'{indent}    <mxCell id="0" />',
        f'{indent}    <mxCell id="1" parent="0" />',
        "",
    ]

    for g in diagram["groups"]:
        v_pad = "spacingBottom=8" if g["title_pos"] == "bottom_center" else "spacingTop=6"
        style = (
            f"rounded=1;arcSize=12;whiteSpace=wrap;html=1;shadow=1;fillColor={g['fill']};"
            f"strokeColor={g['stroke']};strokeWidth=1.8;fontStyle=1;fontSize=12;fontColor=#7F6000;"
            f"fontFamily=Times New Roman;verticalAlign={'bottom' if 'bottom' in g['title_pos'] else 'top'};"
            f"align={'left' if 'left' in g['title_pos'] else 'center'};{v_pad};"
            + ("spacingLeft=18;" if "left" in g["title_pos"] else "")
        )
        lines.extend([
            f'{indent}    <mxCell id="{g["id"]}" value="{_xml_val(g["title"])}" style="{style}" vertex="1" parent="1">',
            f'{indent}      <mxGeometry x="{g["x"]}" y="{g["y"]}" width="{g["w"]}" height="{g["h"]}" as="geometry" />',
            f"{indent}    </mxCell>",
        ])

    for node in diagram["nodes"]:
        key = node["id"]
        title_html = "<br>".join(escape(line) for line in node["title"].split("\n"))
        desc_html = "<br>".join(escape(line) for line in node["desc"].split("\n"))
        html_val = f'<b>{title_html}</b><br><font style="font-size:10px;color:#334155">{desc_html}</font>'
        is_vert = node["style"] == "vertical"
        v_align = "verticalAlign=bottom;spacingBottom=10;" if is_vert else "verticalAlign=middle;spacingTop=6;spacingLeft=10;spacingRight=10;"
        n_style = (
            f"rounded=1;arcSize=18;whiteSpace=wrap;html=1;shadow=1;fillColor={node['fill']};"
            f"strokeColor=#4A6583;strokeWidth=1.5;fontSize=11;fontFamily=Times New Roman;fontColor=#0F172A;{v_align}"
        )
        bx = node["x"] + 22
        by = node["y"] + 22
        b64_str, iw, ih = _get_icon_data_uri(node["icon"])
        if is_vert:
            ix = node["x"] + (node["w"] - iw) // 2 + 10
            iy = node["y"] + 12
        else:
            ix = node["x"] + node["w"] - iw - 10
            iy = node["y"] + 8
        icon_style = f"shape=image;html=1;verticalLabelPosition=bottom;verticalAlign=top;imageAspect=0;aspect=fixed;image=data:image/png,{b64_str};"
        escaped_html = _xml_val(html_val)
        lines.extend([
            f'{indent}    <mxCell id="{key}" value="{escaped_html}" style="{n_style}" vertex="1" parent="1">',
            f'{indent}      <mxGeometry x="{node["x"]}" y="{node["y"]}" width="{node["w"]}" height="{node["h"]}" as="geometry" />',
            f"{indent}    </mxCell>",
            f'{indent}    <mxCell id="{key}_icon" value="" style="{icon_style}" vertex="1" parent="1">',
            f'{indent}      <mxGeometry x="{ix}" y="{iy}" width="{iw}" height="{ih}" as="geometry" />',
            f"{indent}    </mxCell>",
            f'{indent}    <mxCell id="{key}_num" value="{node["num"]}" style="ellipse;whiteSpace=wrap;html=1;aspect=fixed;fillColor=#FFE599;strokeColor=#333333;strokeWidth=1.4;fontStyle=1;fontSize=12;fontFamily=Times New Roman;fontColor=#000000;" vertex="1" parent="1">',
            f'{indent}      <mxGeometry x="{bx - 14}" y="{by - 14}" width="28" height="28" as="geometry" />',
            f"{indent}    </mxCell>",
        ])

    for arr in diagram["arrows"]:
        is_curved = abs(arr.get("rad", 0.0)) > 1e-6
        is_dashed = arr.get("ls", "-") != "-"
        base = "curved=1;" if is_curved else "edgeStyle=orthogonalEdgeStyle;rounded=0;"
        dash = "dashed=1;dashPattern=2 3;" if is_dashed else ""
        st = f"{base}{dash}html=1;strokeColor={arr['color']};strokeWidth=2.2;endArrow=block;endFill=1;"
        pts = arr.get("wp") if arr.get("wp") else (arr["pts"][1:-1] if len(arr["pts"]) > 2 else [])
        lines.append(f'{indent}    <mxCell id="{arr["id"]}" style="{st}" edge="1" source="{arr["src"]}" target="{arr["tgt"]}" parent="1">')
        if pts:
            lines.append(f'{indent}      <mxGeometry relative="1" as="geometry"><Array as="points">')
            for wx, wy in pts:
                lines.append(f'{indent}        <mxPoint x="{wx}" y="{wy}" />')
            lines.append(f'{indent}      </Array></mxGeometry>')
        else:
            lines.append(f'{indent}      <mxGeometry relative="1" as="geometry" />')
        lines.append(f"{indent}    </mxCell>")

    lines.extend([
        f"{indent}  </root>",
        f"{indent}</mxGraphModel>",
    ])
    return "\n".join(lines)


def _build_mxgraph_xml(diagram, indent=""):
    """Sinh block <mxGraphModel>...</mxGraphModel> cho một sơ đồ"""
    if diagram.get("layout") == "pipeline":
        return _build_pipeline_mxgraph_xml(diagram, indent=indent)
    pw = diagram["pool_w"]
    ph = diagram["pool_h"]
    lanes = _lane_map(diagram)
    lines = [
        f'{indent}<mxGraphModel dx="1100" dy="850" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="{pw + 80}" pageHeight="{ph + 80}" math="0" shadow="0">',
        f"{indent}  <root>",
        f'{indent}    <mxCell id="0" />',
        f'{indent}    <mxCell id="1" parent="0" />',
        "",
        f"{indent}    <!-- ===== POOL DỌC (COMPACT VERTICAL BPMN) ===== -->",
        f'{indent}    <mxCell id="pool" value="{_xml_val(diagram["pool_title"])}"',
        f'{indent}      style="swimlane;html=1;childLayout=stackLayout;resizeParent=1;resizeParentMax=0;horizontal=0;startSize=30;horizontalStack=1;fillColor=#dae8fc;strokeColor=#6c8ebf;fontStyle=1;fontSize=11;"',
        f'{indent}      vertex="1" parent="1">',
        f'{indent}      <mxGeometry x="30" y="30" width="{pw}" height="{ph}" as="geometry" />',
        f"{indent}    </mxCell>",
        "",
    ]

    for lane in diagram["lanes"]:
        lines.extend([
            f'{indent}    <mxCell id="{lane["id"]}" value="{_xml_val(lane["title"])}"',
            f'{indent}      style="swimlane;html=1;startSize=30;fillColor={lane["fill"]};strokeColor={lane["stroke"]};fontStyle=1;fontSize=11;"',
            f'{indent}      vertex="1" parent="pool">',
            f'{indent}      <mxGeometry x="{lane["x"]}" y="0" width="{lane["w"]}" height="{ph}" as="geometry" />',
            f"{indent}    </mxCell>",
        ])
    lines.append("")

    for node in diagram["nodes"]:
        lines.extend([
            f'{indent}    <mxCell id="{node["id"]}" value="{_xml_val(node["label"])}" style="{_node_style(node)}" vertex="1" parent="{node["parent"]}">',
            f'{indent}      <mxGeometry x="{node["x"]}" y="{node["y"]}" width="{node["w"]}" height="{node["h"]}" as="geometry" />',
            f"{indent}    </mxCell>",
        ])
    lines.append("")

    for edge in diagram["edges"]:
        style = "edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;fontSize=9;"
        if edge.get("lcolor"):
            style += f"fontColor={edge['lcolor']};fontStyle=1;"
        val_attr = f'value="{_xml_val(edge["label"])}" ' if edge.get("label") else ""

        pts = edge["pts"]
        # Với edge có >= 3 điểm (đường gấp khúc), xuất các điểm bẻ góc trung gian theo hệ tọa độ của edge['parent']
        if len(pts) > 2:
            offset_x = 0 if edge["parent"] == "pool" else lanes[edge["parent"]]["x"]
            lines.append(
                f'{indent}    <mxCell id="{edge["id"]}" {val_attr}style="{style}" edge="1" source="{edge["source"]}" target="{edge["target"]}" parent="{edge["parent"]}">'
            )
            lines.append(f'{indent}      <mxGeometry relative="1" as="geometry">')
            lines.append(f'{indent}        <Array as="points">')
            for px, py in pts[1:-1]:
                lines.append(f'{indent}          <mxPoint x="{px - offset_x}" y="{py}" />')
            lines.append(f"{indent}        </Array>")
            lines.append(f"{indent}      </mxGeometry>")
            lines.append(f"{indent}    </mxCell>")
        else:
            lines.append(
                f'{indent}    <mxCell id="{edge["id"]}" {val_attr}style="{style}" edge="1" source="{edge["source"]}" target="{edge["target"]}" parent="{edge["parent"]}">'
            )
            lines.append(f'{indent}      <mxGeometry relative="1" as="geometry" />')
            lines.append(f"{indent}    </mxCell>")

    lines.extend([
        f"{indent}  </root>",
        f"{indent}</mxGraphModel>",
    ])
    return "\n".join(lines)


def export_drawio_files(diagrams):
    """Xuất các file .drawio riêng biệt + 1 file .drawio đa tab và kiểm tra hợp lệ XML"""
    created = []
    for d in diagrams:
        xml_content = '<?xml version="1.0" encoding="UTF-8"?>\n' + _build_mxgraph_xml(d)
        path = os.path.join(OUTPUT_DIR, d["drawio_file"])
        with open(path, "w", encoding="utf-8") as f:
            f.write(xml_content)
        ET.parse(path)
        created.append(path)
        print(f"Đã tạo & kiểm tra XML: {path}")

    # File All-in-one gộp toàn bộ tab (Chương 1, 4, 5)
    all_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<mxfile host="app.diagrams.net" modified="2026-11-15T19:20:00.000Z" agent="Retail ML Hub" version="24.0.0">',
    ]
    for d in diagrams:
        all_lines.append(f'  <diagram id="{d["id"]}" name="{_xml_val(d["tab_name"])}">')
        all_lines.append(_build_mxgraph_xml(d, indent="    "))
        all_lines.append("  </diagram>")
    all_lines.append("</mxfile>")

    all_path = os.path.join(OUTPUT_DIR, "BPMN_Retail_ML_Hub_All_In_One.drawio")
    with open(all_path, "w", encoding="utf-8") as f:
        f.write("\n".join(all_lines))
    ET.parse(all_path)
    created.append(all_path)
    print(f"Đã tạo & kiểm tra XML ({len(diagrams)} tabs): {all_path}")
    return created


# ---------------------------------------------------------------------------
# PHẦN 2: VẼ ẢNH .PNG 300 DPI ĐỒNG BỘ 100% VỚI FILE .DRAWIO
# ---------------------------------------------------------------------------
def _draw_cylinder(ax, x, y_top, w, h, fill, stroke, label, H):
    """Vẽ khối cơ sở dữ liệu Cylinder3 (y_top theo hệ top-down của Pool)"""
    y_bot = H - (y_top + h)
    cap_h = 10.0

    # Thân trụ bo nhẹ ở đáy
    body = FancyBboxPatch(
        (x, y_bot),
        w,
        h - cap_h / 2,
        boxstyle="round,pad=0,rounding_size=6",
        facecolor=fill,
        edgecolor=stroke,
        linewidth=1.6,
        zorder=3,
    )
    ax.add_patch(body)

    # Nắp oval trên đỉnh
    top_oval = Ellipse(
        (x + w / 2, y_bot + h - cap_h / 2),
        width=w,
        height=cap_h,
        facecolor=fill,
        edgecolor=stroke,
        linewidth=1.6,
        zorder=4,
    )
    ax.add_patch(top_oval)

    ax.text(
        x + w / 2,
        y_bot + (h - cap_h) / 2,
        label,
        ha="center",
        va="center",
        fontsize=8.3,
        color="#0F172A",
        linespacing=1.25,
        zorder=5,
    )


def _draw_pipeline_icon(ax, kind, cx, cy):
    """Vẽ icon vector phẳng sạch đẹp cho từng bước (0–11) của sơ đồ Pipeline"""
    import math
    dark = "#1E293B"
    if kind == "consumers":
        # Hai hình người đơn giản gọn gàng
        ax.add_patch(Circle((cx - 14, cy + 16), 9, facecolor="#FDE68A", edgecolor=dark, lw=1.2, zorder=5))
        ax.add_patch(FancyBboxPatch((cx - 24, cy - 16), 20, 22, boxstyle="round,pad=0,rounding_size=4", facecolor="#60A5FA", edgecolor=dark, lw=1.2, zorder=5))
        ax.add_patch(Circle((cx + 15, cy + 14), 8.5, facecolor="#FDBA74", edgecolor=dark, lw=1.2, zorder=5))
        ax.add_patch(FancyBboxPatch((cx + 6, cy - 16), 18, 21, boxstyle="round,pad=0,rounding_size=4", facecolor="#F87171", edgecolor=dark, lw=1.2, zorder=5))
    elif kind == "system":
        # Điện thoại POS + xe đẩy + $, thẻ user 4 sao, quả địa cầu có gian hàng
        ax.add_patch(FancyBboxPatch((cx - 56, cy - 20), 28, 42, boxstyle="round,pad=0,rounding_size=4", facecolor="#E2E8F0", edgecolor=dark, lw=1.2, zorder=5))
        ax.add_patch(Rectangle((cx - 52, cy - 12), 20, 26, facecolor="#FFFFFF", edgecolor=dark, lw=0.8, zorder=6))
        # Mini shopping cart inside phone screen
        ax.plot([cx - 50, cx - 47.5, cx - 45.5, cx - 37.5, cx - 36], [cy + 5, cy + 5, cy - 1, cy - 1, cy + 3], color=dark, lw=1.0, zorder=7)
        ax.add_patch(Polygon([(cx - 46, cy + 3), (cx - 36.5, cy + 3), (cx - 38, cy - 1), (cx - 45, cy - 1)], closed=True, facecolor="#FB7185", edgecolor=dark, lw=0.7, zorder=7))
        ax.add_patch(Circle((cx - 44.5, cy - 3.5), 1.4, facecolor=dark, zorder=7))
        ax.add_patch(Circle((cx - 38.5, cy - 3.5), 1.4, facecolor=dark, zorder=7))
        ax.add_patch(Circle((cx - 32, cy - 2), 8, facecolor="#22C55E", edgecolor=dark, lw=0.9, zorder=8))
        ax.text(cx - 32, cy - 2, "$", ha="center", va="center", fontsize=8.5, fontweight="bold", color="#FFFFFF", zorder=9)

        # Thẻ người dùng có thanh xanh 4 sao ở trên
        ax.add_patch(FancyBboxPatch((cx - 14, cy - 16), 28, 32, boxstyle="round,pad=0,rounding_size=4", facecolor="#FDE047", edgecolor=dark, lw=1.1, zorder=5))
        ax.add_patch(Rectangle((cx - 12, cy + 9), 24, 5, facecolor="#38BDF8", edgecolor=dark, lw=0.7, zorder=6))
        for sx in [-8, -2.7, 2.7, 8]:
            ax.add_patch(Circle((cx + sx, cy + 11.5), 1.1, facecolor="#FDE047", edgecolor="none", zorder=7))
        ax.add_patch(Circle((cx, cy + 1), 5.2, facecolor="#FFFFFF", edgecolor=dark, lw=0.9, zorder=6))
        ax.add_patch(FancyBboxPatch((cx - 7, cy - 12), 14, 7, boxstyle="round,pad=0,rounding_size=2", facecolor="#FFFFFF", edgecolor=dark, lw=0.9, zorder=6))

        # Quả địa cầu + gian hàng bán lẻ nhỏ ở góc dưới trái
        ax.add_patch(Circle((cx + 42, cy), 19, facecolor="#3B82F6", edgecolor=dark, lw=1.2, zorder=5))
        ax.add_patch(Ellipse((cx + 42, cy), width=16, height=36, facecolor="none", edgecolor="#DBEAFE", lw=0.9, zorder=6))
        ax.plot([cx + 24, cx + 60], [cy, cy], color="#DBEAFE", lw=0.9, zorder=6)
        ax.add_patch(FancyBboxPatch((cx + 22, cy - 17), 16, 11, boxstyle="round,pad=0,rounding_size=1.5", facecolor="#FEF3C7", edgecolor=dark, lw=0.9, zorder=7))
        ax.add_patch(Polygon([(cx + 21, cy - 6), (cx + 39, cy - 6), (cx + 37, cy - 2), (cx + 23, cy - 2)], closed=True, facecolor="#EF4444", edgecolor=dark, lw=0.8, zorder=8))
        ax.add_patch(Rectangle((cx + 27, cy - 17), 6, 6, facecolor="#60A5FA", edgecolor=dark, lw=0.7, zorder=8))
    elif kind == "database":
        # Màn hình có giỏ hàng + hình trụ CSDL + quả địa cầu nhỏ (thu gọn chiều ngang)
        ax.add_patch(FancyBboxPatch((cx - 30, cy - 7), 24, 18, boxstyle="round,pad=0,rounding_size=2", facecolor="#E2E8F0", edgecolor=dark, lw=1.1, zorder=5))
        ax.add_patch(Rectangle((cx - 27.5, cy - 4), 19, 12, facecolor="#93C5FD", edgecolor=dark, lw=0.7, zorder=6))
        # Mini shopping cart inside monitor screen
        ax.plot([cx - 25.5, cx - 23.5, cx - 22, cx - 13, cx - 11.5], [cy + 4.5, cy + 4.5, cy - 0.5, cy - 0.5, cy + 3], color=dark, lw=0.9, zorder=7)
        ax.add_patch(Polygon([(cx - 22.5, cy + 3), (cx - 12, cy + 3), (cx - 13.5, cy - 0.5), (cx - 21.5, cy - 0.5)], closed=True, facecolor="#F59E0B", edgecolor=dark, lw=0.6, zorder=7))
        ax.add_patch(Circle((cx - 21, cy - 2.3), 1.2, facecolor=dark, zorder=7))
        ax.add_patch(Circle((cx - 14.5, cy - 2.3), 1.2, facecolor=dark, zorder=7))
        ax.plot([cx - 18, cx - 18], [cy - 7, cy - 13], color=dark, lw=1.4, zorder=5)
        ax.plot([cx - 24, cx - 12], [cy - 13, cy - 13], color=dark, lw=1.4, zorder=5)

        ax.add_patch(FancyBboxPatch((cx, cy - 15), 24, 26, boxstyle="round,pad=0,rounding_size=4", facecolor="#CBD5E1", edgecolor=dark, lw=1.1, zorder=5))
        ax.add_patch(Ellipse((cx + 12, cy + 10), 24, 8, facecolor="#E2E8F0", edgecolor=dark, lw=1.1, zorder=6))
        ax.add_patch(Circle((cx + 22, cy - 8), 8.5, facecolor="#38BDF8", edgecolor=dark, lw=1.1, zorder=7))
    elif kind == "clean":
        # Hình trụ dữ liệu bạc + dấu tick xanh
        ax.add_patch(FancyBboxPatch((cx - 16, cy - 15), 26, 28, boxstyle="round,pad=0,rounding_size=4", facecolor="#CBD5E1", edgecolor=dark, lw=1.1, zorder=5))
        ax.add_patch(Ellipse((cx - 3, cy + 12), 26, 8, facecolor="#F1F5F9", edgecolor=dark, lw=1.1, zorder=6))
        ax.plot([cx - 15, cx + 9], [cy + 2, cy + 2], color="#64748B", lw=0.9, zorder=6)
        ax.plot([cx - 15, cx + 9], [cy - 6, cy - 6], color="#64748B", lw=0.9, zorder=6)
        ax.add_patch(Circle((cx + 12, cy - 8), 8.5, facecolor="#22C55E", edgecolor="#166534", lw=1.1, zorder=7))
        ax.plot([cx + 8, cx + 11.5, cx + 16.5], [cy - 8, cy - 11.5, cy - 4.5], color="#FFFFFF", lw=1.8, zorder=8)
    elif kind == "schema":
        # 3 lớp bảng dữ liệu xếp tầng + bánh răng nhỏ
        for idx, col in enumerate(["#FBBF24", "#34D399", "#38BDF8"]):
            yy = cy - 10 + idx * 9
            pts = [(cx - 22, yy), (cx + 10, yy), (cx + 18, yy + 6), (cx - 14, yy + 6)]
            ax.add_patch(Polygon(pts, closed=True, facecolor=col, edgecolor=dark, lw=1.0, zorder=5 + idx))
        ax.add_patch(Circle((cx + 16, cy - 8), 7, facecolor="#CBD5E1", edgecolor=dark, lw=1.1, zorder=9))
        ax.add_patch(Circle((cx + 16, cy - 8), 2.5, facecolor="#FFFFFF", edgecolor=dark, lw=0.9, zorder=10))
    elif kind == "filter":
        # Phễu lọc dữ liệu màu xanh ngọc + điểm dữ liệu vàng
        ax.add_patch(Circle((cx, cy + 14), 5, facecolor="#FACC15", edgecolor=dark, lw=1.0, zorder=5))
        pts = [(cx - 15, cy + 8), (cx + 15, cy + 8), (cx + 4, cy - 4), (cx + 4, cy - 16), (cx - 4, cy - 12), (cx - 4, cy - 4)]
        ax.add_patch(Polygon(pts, closed=True, facecolor="#38BDF8", edgecolor="#0369A1", lw=1.2, zorder=6))
    elif kind == "gears":
        # Hai bánh răng cơ khí chọn thuật toán
        for gx, gy, r, col in [(cx - 5, cy - 4, 10, "#C4B5FD"), (cx + 9, cy + 7, 7, "#FDBA74")]:
            teeth = []
            n_t = 8
            for k in range(n_t * 2):
                ang = k * math.pi / n_t
                rad = r + (2.5 if k % 2 == 0 else 0)
                teeth.append((gx + rad * math.cos(ang), gy + rad * math.sin(ang)))
            ax.add_patch(Polygon(teeth, closed=True, facecolor=col, edgecolor=dark, lw=1.0, zorder=5))
            ax.add_patch(Circle((gx, gy), r * 0.38, facecolor="#FFFFFF", edgecolor=dark, lw=0.9, zorder=6))
    elif kind == "brain":
        # Biểu tượng bộ não ML (hai bán cầu)
        ax.add_patch(Ellipse((cx - 6, cy), 16, 26, facecolor="#FDA4AF", edgecolor="#9F1239", lw=1.2, zorder=5))
        ax.add_patch(Ellipse((cx + 6, cy), 16, 26, facecolor="#FB7185", edgecolor="#9F1239", lw=1.2, zorder=5))
        ax.plot([cx, cx], [cy - 12, cy + 12], color="#9F1239", lw=1.2, zorder=6)
        ax.plot([cx - 9, cx - 3], [cy + 5, cy + 2], color="#9F1239", lw=0.9, zorder=6)
        ax.plot([cx + 3, cx + 9], [cy - 2, cy - 5], color="#9F1239", lw=0.9, zorder=6)
    elif kind == "eval":
        # Màn hình biểu đồ đánh giá & cải tiến
        ax.add_patch(FancyBboxPatch((cx - 17, cy - 11), 34, 24, boxstyle="round,pad=0,rounding_size=3", facecolor="#FFFFFF", edgecolor=dark, lw=1.2, zorder=5))
        ax.add_patch(Rectangle((cx - 17, cy + 8), 34, 5, facecolor="#5EEAD4", edgecolor=dark, lw=0.9, zorder=6))
        ax.plot([cx - 12, cx - 4, cx + 2, cx + 11], [cy - 6, cy + 1, cy - 3, cy + 5], color="#EC4899", lw=1.8, zorder=7)
        ax.plot([cx, cx - 8], [cy - 11, cy - 18], color=dark, lw=1.2, zorder=4)
        ax.plot([cx, cx + 8], [cy - 11, cy - 18], color=dark, lw=1.2, zorder=4)
    elif kind == "save":
        # Khay lưu mô hình + mũi tên xanh đi xuống
        ax.add_patch(FancyBboxPatch((cx - 15, cy - 14), 30, 14, boxstyle="round,pad=0,rounding_size=2", facecolor="#CBD5E1", edgecolor=dark, lw=1.1, zorder=5))
        ax.add_patch(Rectangle((cx - 9, cy - 10), 18, 4, facecolor="#FFFFFF", edgecolor=dark, lw=0.7, zorder=6))
        pts = [(cx - 5, cy + 14), (cx + 5, cy + 14), (cx + 5, cy + 4), (cx + 11, cy + 4), (cx, cy - 6), (cx - 11, cy + 4), (cx - 5, cy + 4)]
        ax.add_patch(Polygon(pts, closed=True, facecolor="#3B82F6", edgecolor="#1D4ED8", lw=1.1, zorder=7))
    elif kind == "load":
        # Vòng tròn cam + mũi tên xanh nạp mô hình
        ax.add_patch(Circle((cx, cy - 2), 14, facecolor="#FB923C", edgecolor=dark, lw=1.1, zorder=5))
        ax.add_patch(Circle((cx, cy - 2), 9, facecolor="#FBE5D6", edgecolor=dark, lw=0.9, zorder=6))
        pts = [(cx - 5, cy - 8), (cx + 5, cy - 8), (cx + 5, cy + 4), (cx + 10, cy + 4), (cx, cy + 15), (cx - 10, cy + 4), (cx - 5, cy + 4)]
        ax.add_patch(Polygon(pts, closed=True, facecolor="#22C55E", edgecolor="#166534", lw=1.1, zorder=7))
    elif kind == "team":
        # Áo tím + huy chương + biểu đồ cột + nhóm 3 người (khớp bước 11 mẫu)
        shirt_pts = [
            (cx - 23, cy + 12), (cx - 18, cy + 12), (cx - 16, cy + 9), (cx - 14, cy + 12), (cx - 9, cy + 12),
            (cx - 5, cy + 6), (cx - 8, cy + 2), (cx - 10, cy + 5), (cx - 10, cy - 10),
            (cx - 22, cy - 10), (cx - 22, cy + 5), (cx - 24, cy + 2), (cx - 27, cy + 6),
        ]
        ax.add_patch(Polygon(shirt_pts, closed=True, facecolor="#A78BFA", edgecolor=dark, lw=0.95, zorder=5))
        ax.add_patch(Polygon([(cx - 24, cy - 8), (cx - 26, cy - 14), (cx - 22.5, cy - 11.5), (cx - 19, cy - 14), (cx - 21, cy - 8)],
                             closed=True, facecolor="#EC4899", edgecolor=dark, lw=0.7, zorder=7))
        ax.add_patch(Circle((cx - 22.5, cy - 5.5), 4.5, facecolor="#FACC15", edgecolor=dark, lw=0.9, zorder=8))
        ax.add_patch(Circle((cx - 22.5, cy - 5.5), 2.1, facecolor="#4ADE80", edgecolor="none", zorder=9))

        for bx, bh, col in [(1, 11, "#2DD4BF"), (8, 17, "#38BDF8"), (15, 14, "#FBBF24"), (22, 19, "#60A5FA")]:
            ax.add_patch(Rectangle((cx + bx, cy), 4.8, bh, facecolor=col, edgecolor=dark, lw=0.75, zorder=5))
        for ox in [5, 14, 23]:
            ax.add_patch(Circle((cx + ox, cy + 4), 4.0, facecolor="#FDE68A", edgecolor=dark, lw=0.85, zorder=6))
            ax.add_patch(FancyBboxPatch((cx + ox - 3.6, cy - 11), 7.2, 10.5, boxstyle="round,pad=0,rounding_size=2", facecolor="#93C5FD", edgecolor=dark, lw=0.85, zorder=6))
    elif kind == "weather":
        # Mặt trời + đám mây xanh nhạt
        ax.add_patch(Circle((cx + 8, cy + 7), 8, facecolor="#FACC15", edgecolor=dark, lw=1.0, zorder=5))
        ax.add_patch(Ellipse((cx - 4, cy - 3), 22, 13, facecolor="#E0F2FE", edgecolor=dark, lw=1.0, zorder=6))
        ax.add_patch(Circle((cx - 9, cy + 1), 6, facecolor="#E0F2FE", edgecolor=dark, lw=1.0, zorder=6))
        ax.add_patch(Circle((cx + 1, cy + 2), 6.5, facecolor="#E0F2FE", edgecolor=dark, lw=1.0, zorder=6))
        for rx in [-9, -2, 5]:
            ax.plot([cx + rx, cx + rx - 2], [cy - 11, cy - 15], color="#0284C7", lw=1.4, zorder=6)
    elif kind == "price":
        # Thẻ giảm giá xanh + đồng xu vàng
        pts = [(cx - 18, cy + 8), (cx + 2, cy + 8), (cx + 12, cy - 2), (cx - 2, cy - 15), (cx - 18, cy - 2)]
        ax.add_patch(Polygon(pts, closed=True, facecolor="#34D399", edgecolor=dark, lw=1.1, zorder=5))
        ax.text(cx - 5, cy - 2, "%", ha="center", va="center", fontsize=9.5, fontweight="bold", color="#FFFFFF", zorder=6)
        ax.add_patch(Circle((cx + 12, cy - 6), 8.5, facecolor="#FBBF24", edgecolor=dark, lw=1.1, zorder=7))
        ax.text(cx + 12, cy - 6, "$", ha="center", va="center", fontsize=8.5, fontweight="bold", color=dark, zorder=8)


def _get_icon_data_uri(kind):
    """Render icon thành PNG nền trong suốt (lưu ra icons/*.png) và trả về (base64_str, w_px, h_px) cho Draw.io"""
    if kind in _ICON_DATA_CACHE:
        return _ICON_DATA_CACHE[kind]

    if kind == "system":
        xlim, ylim, size_px = (-66, 66), (-26, 26), (105, 42)
    elif kind in ("database", "team"):
        xlim, ylim, size_px = (-36, 36), (-22, 22), (58, 36)
    else:
        xlim, ylim, size_px = (-26, 26), (-22, 22), (42, 36)

    w_span = xlim[1] - xlim[0]
    h_span = ylim[1] - ylim[0]
    fig = plt.figure(figsize=(w_span / 24.0, h_span / 24.0), dpi=150)
    fig.patch.set_alpha(0.0)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.patch.set_alpha(0.0)
    _draw_pipeline_icon(ax, kind, 0, 0)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", transparent=True, dpi=150)
    icon_path = os.path.join(ICON_DIR, f"{kind}.png")
    fig.savefig(icon_path, format="png", transparent=True, dpi=150)
    plt.close(fig)

    b64_str = base64.b64encode(buf.getvalue()).decode("ascii")
    _ICON_DATA_CACHE[kind] = (b64_str, size_px[0], size_px[1])
    return _ICON_DATA_CACHE[kind]


def _render_pipeline_png(diagram):
    """Vẽ ảnh PNG 300 DPI theo phong cách Pipeline đặc thù cho từng chương"""
    pw, ph = diagram["pool_w"], diagram["pool_h"]
    font_fam = "Times New Roman"
    fig_w = 11.4
    fig_h = fig_w * (ph + 58) / (pw + 24)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_xlim(-14, pw + 10)
    ax.set_ylim(-6, ph + 52)
    ax.axis("off")
    fig.patch.set_facecolor("#FFFFFF")
    ax.set_facecolor("#FFFFFF")

    ax.text(
        pw / 2,
        ph + 28,
        diagram["png_title"],
        ha="center",
        va="center",
        fontsize=12.8,
        fontweight="bold",
        fontfamily=font_fam,
        color="#1E3A8A",
    )

    for g in diagram["groups"]:
        gx, gy_bot, gw, gh = g["x"], ph - (g["y"] + g["h"]), g["w"], g["h"]
        ax.add_patch(FancyBboxPatch(
            (gx + 4, gy_bot - 4), gw, gh,
            boxstyle="round,pad=0,rounding_size=20", facecolor="#CBD5E1", edgecolor="none", alpha=0.55, zorder=1
        ))
        ax.add_patch(FancyBboxPatch(
            (gx, gy_bot), gw, gh,
            boxstyle="round,pad=0,rounding_size=20", facecolor=g["fill"], edgecolor=g["stroke"], linewidth=1.8, zorder=2
        ))
        if g["title_pos"] == "top_center":
            line_cnt = g["title"].count("\n") + 1
            ty = gy_bot + gh - (22 if line_cnt > 1 else 17)
            ax.text(gx + gw / 2, ty, g["title"], ha="center", va="center",
                    fontsize=10.6, fontweight="bold", fontfamily=font_fam, color="#7F6000", linespacing=1.18, zorder=4)
        elif g["title_pos"] == "top_left":
            ax.text(gx + 18, gy_bot + gh - 18, g["title"], ha="left", va="center",
                    fontsize=10.4, fontweight="bold", fontfamily=font_fam, color="#7F6000", zorder=4)
        else:
            ax.text(gx + gw / 2, gy_bot + 16, g["title"], ha="center", va="center",
                    fontsize=10.6, fontweight="bold", fontfamily=font_fam, color="#7F6000", zorder=4)

    for node in diagram["nodes"]:
        gx, gy_bot, gw, gh = node["x"], ph - (node["y"] + node["h"]), node["w"], node["h"]
        cx, cy = gx + gw / 2, gy_bot + gh / 2

        ax.add_patch(FancyBboxPatch(
            (gx + 3, gy_bot - 3), gw, gh,
            boxstyle="round,pad=0,rounding_size=14", facecolor="#CBD5E1", edgecolor="none", alpha=0.45, zorder=3
        ))
        ax.add_patch(FancyBboxPatch(
            (gx, gy_bot), gw, gh,
            boxstyle="round,pad=0,rounding_size=14", facecolor=node["fill"], edgecolor="#4A6583", linewidth=1.45, zorder=4
        ))

        if node["style"] == "vertical":
            bx, by = gx + 24, gy_bot + gh - 24
            _draw_pipeline_icon(ax, node["icon"], cx + 14, gy_bot + gh - 34)
            ax.text(cx, gy_bot + 62, node["title"], ha="center", va="center",
                    fontsize=10.2, fontweight="bold", fontfamily=font_fam, color="#0F172A", zorder=6)
            ax.text(cx, gy_bot + 26, node["desc"], ha="center", va="center",
                    fontsize=9.0, fontfamily=font_fam, color="#334155", linespacing=1.18, zorder=6)
        else:
            bx, by = gx + 22, gy_bot + gh - 24
            ix = gx + gw - (38 if node["icon"] in ("database", "team") else 32)
            _draw_pipeline_icon(ax, node["icon"], ix, gy_bot + gh - 28)
            right_clear = {"database": 32, "team": 30, "schema": 24}.get(node["icon"], 19)
            title_cx = ((gx + 38) + (ix - right_clear - 4)) / 2
            is_multiline = "\n" in node["title"]
            ax.text(
                title_cx,
                gy_bot + gh - (27 if is_multiline else 24),
                node["title"],
                ha="center",
                va="center",
                fontsize=9.6 if is_multiline else 10.0,
                fontweight="bold",
                fontfamily=font_fam,
                color="#0F172A",
                linespacing=1.15,
                zorder=6,
            )
            desc_y = gy_bot + (25 if gh >= 100 else 20)
            ax.text(cx, desc_y, node["desc"], ha="center", va="center",
                    fontsize=8.9, fontfamily=font_fam, color="#334155", linespacing=1.18, zorder=6)

        ax.add_patch(Circle((bx, by), 13.5, facecolor="#FFE599", edgecolor="#333333", linewidth=1.4, zorder=9))
        ax.text(bx, by, node["num"], ha="center", va="center",
                fontsize=10.2, fontweight="bold", fontfamily=font_fam, color="#000000", zorder=10)

    for arr in diagram["arrows"]:
        pts = [(px, ph - py) for (px, py) in arr["pts"]]
        ls = arr.get("ls", "-")
        rad = arr.get("rad", 0.0)
        color = arr["color"]
        if len(pts) > 2:
            ax.plot([p[0] for p in pts[:-1]], [p[1] for p in pts[:-1]],
                    color=color, linewidth=2.2, linestyle=ls, zorder=7)
        kw = dict(arrowstyle="-|>,head_length=6.5,head_width=4.0", color=color, linewidth=2.2, linestyle=ls, zorder=7)
        if abs(rad) > 1e-6:
            kw["connectionstyle"] = f"arc3,rad={rad}"
        ax.add_patch(FancyArrowPatch(pts[-2], pts[-1], **kw))

    out_path = os.path.join(OUTPUT_DIR, diagram["png_file"])
    plt.savefig(out_path, bbox_inches="tight", pad_inches=0.14)
    plt.close()
    print(f"Đã tạo ảnh PNG: {out_path}")
    return out_path


def render_png(diagram):
    """Vẽ ảnh PNG độ phân giải 300 DPI khớp từng tọa độ với file .drawio"""
    if diagram.get("layout") == "pipeline":
        return _render_pipeline_png(diagram)
    pw = diagram["pool_w"]
    ph = diagram["pool_h"]
    lanes = _lane_map(diagram)

    fig_w = 13.8
    fig_h = fig_w * (ph + 65) / pw
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_xlim(-8, pw + 8)
    ax.set_ylim(-8, ph + 52)
    ax.axis("off")
    fig.patch.set_facecolor("#FFFFFF")
    ax.set_facecolor("#FFFFFF")

    # Tiêu đề hình chuẩn báo cáo
    ax.text(
        pw / 2,
        ph + 28,
        diagram["png_title"],
        ha="center",
        va="center",
        fontsize=13.2,
        fontweight="bold",
        color="#1E3A8A",
    )

    # Khung Pool ngoài
    pool_bg = FancyBboxPatch(
        (0, 0),
        pw,
        ph,
        boxstyle="square,pad=0",
        facecolor="#FFFFFF",
        edgecolor="#6c8ebf",
        linewidth=1.8,
        zorder=1,
    )
    ax.add_patch(pool_bg)

    # Thanh tiêu đề Pool bên trái (rộng 30px)
    pool_hdr = FancyBboxPatch(
        (0, 0),
        30,
        ph,
        boxstyle="square,pad=0",
        facecolor="#dae8fc",
        edgecolor="#6c8ebf",
        linewidth=1.6,
        zorder=2,
    )
    ax.add_patch(pool_hdr)
    ax.text(
        15,
        ph / 2,
        diagram["pool_title"],
        ha="center",
        va="center",
        rotation=90,
        fontsize=9.8,
        fontweight="bold",
        color="#0F172A",
        zorder=3,
    )

    # 3 Làn dọc (Swimlanes)
    for lane in diagram["lanes"]:
        lx, lw = lane["x"], lane["w"]
        # Thân làn trắng viền màu làn
        lane_body = FancyBboxPatch(
            (lx, 0),
            lw,
            ph,
            boxstyle="square,pad=0",
            facecolor="#FFFFFF",
            edgecolor=lane["stroke"],
            linewidth=1.5,
            zorder=1,
        )
        ax.add_patch(lane_body)
        # Tiêu đề làn (30px trên cùng)
        lane_hdr = FancyBboxPatch(
            (lx, ph - 30),
            lw,
            30,
            boxstyle="square,pad=0",
            facecolor=lane["fill"],
            edgecolor=lane["stroke"],
            linewidth=1.5,
            zorder=2,
        )
        ax.add_patch(lane_hdr)
        ax.text(
            lx + lw / 2,
            ph - 15,
            lane["title"],
            ha="center",
            va="center",
            fontsize=9.6,
            fontweight="bold",
            color="#0F172A",
            zorder=3,
        )

    # Vẽ các nút
    for node in diagram["nodes"]:
        lane_x = lanes[node["parent"]]["x"]
        gx = lane_x + node["x"]
        gy_top = node["y"]
        w, h = node["w"], node["h"]
        gy_bot = ph - (gy_top + h)
        cx, cy = gx + w / 2, gy_bot + h / 2
        fw = "bold" if node.get("bold") else "normal"

        kind = node["kind"]
        if kind in ("start", "end"):
            lw = 2.4 if kind == "end" else 1.5
            patch = Ellipse(
                (cx, cy),
                width=w,
                height=h,
                facecolor=node["fill"],
                edgecolor=node["stroke"],
                linewidth=lw,
                zorder=3,
            )
            ax.add_patch(patch)
            ax.text(cx, cy, node["label"], ha="center", va="center", fontsize=8.8, fontweight="bold", color="#0F172A", zorder=4)
        elif kind == "gateway":
            pts = [(cx, cy + h / 2), (cx + w / 2, cy), (cx, cy - h / 2), (cx - w / 2, cy)]
            poly = Polygon(pts, closed=True, facecolor=node["fill"], edgecolor=node["stroke"], linewidth=1.6, zorder=3)
            ax.add_patch(poly)
            ax.text(cx, cy, node["label"], ha="center", va="center", fontsize=8.3, fontweight="bold", color="#0F172A", linespacing=1.2, zorder=4)
        elif kind == "db":
            _draw_cylinder(ax, gx, gy_top, w, h, node["fill"], node["stroke"], node["label"], ph)
        else:
            patch = FancyBboxPatch(
                (gx, gy_bot),
                w,
                h,
                boxstyle="round,pad=0,rounding_size=7",
                facecolor=node["fill"],
                edgecolor=node["stroke"],
                linewidth=1.5,
                zorder=3,
            )
            ax.add_patch(patch)
            ax.text(cx, cy, node["label"], ha="center", va="center", fontsize=8.4, fontweight=fw, color="#0F172A", linespacing=1.25, zorder=4)

    # Vẽ các mũi tên nối vuông góc
    for edge in diagram["edges"]:
        pts = [(px, ph - py) for (px, py) in edge["pts"]]
        if len(pts) > 2:
            xs = [p[0] for p in pts[:-1]]
            ys = [p[1] for p in pts[:-1]]
            ax.plot(xs, ys, color="#334155", linewidth=1.4, zorder=2)
        arr = FancyArrowPatch(
            pts[-2],
            pts[-1],
            arrowstyle="-|>,head_length=5.5,head_width=3.5",
            color="#334155",
            linewidth=1.4,
            zorder=2,
        )
        ax.add_patch(arr)
        if edge.get("label"):
            lx, ly = edge["lpos"]
            ax.text(
                lx,
                ph - ly,
                edge["label"],
                ha="center",
                va="center",
                fontsize=8.2,
                fontweight="bold",
                color=edge.get("lcolor", "#1E293B"),
                bbox=dict(boxstyle="round,pad=0.22", facecolor="#FFFFFF", edgecolor="#CBD5E1", lw=0.7),
                zorder=6,
            )

    out_path = os.path.join(OUTPUT_DIR, diagram["png_file"])
    plt.savefig(out_path, bbox_inches="tight", pad_inches=0.18)
    plt.close()
    print(f"Đã tạo ảnh PNG: {out_path}")
    return out_path


def main():
    print("=" * 68)
    print("   TẠO BỘ SƠ ĐỒ BPMN SWIMLANE (.DRAWIO + .PNG 300 DPI) CHƯƠNG 1, 4 & 5")
    print("=" * 68)
    diagrams = build_diagrams()
    export_drawio_files(diagrams)
    for d in diagrams:
        render_png(d)
    print("=" * 68)
    print(f"--> HOÀN THÀNH! Đã lưu {len(diagrams) + 1} file .drawio và {len(diagrams)} ảnh .png tại: {OUTPUT_DIR}")
    print("=" * 68)


if __name__ == "__main__":
    main()
