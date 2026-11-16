"""
Excel Report Generator: generate_excel_report.py
Tạo file Excel tổng hợp toàn bộ đồ án (6 Sheets chuẩn báo cáo, đồng bộ 100% với MLBA.docx):
  1. 1_Dataset            — Toàn bộ 73.100 dòng x 15 cột + Bảng 1.2 (15 thuộc tính) + Bảng 4.2 (28 đặc trưng) + Thống kê công thức
  2. 2_PCCV               — Bảng Phân công công việc 5 thành viên (Khang, Hiếu, Minh, Ân, Lợi) theo Chương & Chi tiết
  3. 3_Gantt_Chart        — Biểu đồ Gantt tiến độ thực tế (20/10/2026 – 15/11/2026, 10 chặng mốc M1–M10, 18 đầu việc / 6 giai đoạn)
  4. 4_Flowchart_Ch1_Ch4  — Lưu đồ quy trình Chương 1 (8 bước EDA) vs Chương 4 (10 bước ML & Tồn kho) + Danh mục toàn bộ Sơ đồ/Hình ảnh
  5. 5_Contribution_Work  — Ma trận tỷ lệ đóng góp 5 thành viên + Danh mục Chức năng Web Studio (Chương 5)
  6. 6_Bang_Chi_So_Mo_Hinh — Bảng so sánh 5 mô hình (RMSE/MAE/MAPE), Kiểm định ADF/KPSS, Feature Importance + Biểu đồ Excel & ảnh PNG

Xuất ra đồng thời tại:
  - reports/Bao_Cao_Tong_Hop_Du_An_Retail_ML_Hub.xlsx
  - ../Bao_Cao_Tong_Hop_Du_An_Retail_ML_Hub.xlsx (thư mục gốc doan/ để mở nhanh)
"""

import datetime
import json
import os
import shutil
import sys
import time

import joblib
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")


# ---------------------------------------------------------------------------
# ĐƯỜNG DẪN DỰ ÁN & CẤU HÌNH STYLE DÙNG CHUNG (SINGLETONS CHO TỐC ĐỘ CAO)
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_CSV = os.path.join(BASE_DIR, "Dataset", "retail_store_inventory.csv")
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
SAVED_MODELS_DIR = os.path.join(BASE_DIR, "saved_models")
OUT_XLSX_REPORTS = os.path.join(REPORTS_DIR, "Bao_Cao_Tong_Hop_Du_An_Retail_ML_Hub.xlsx")
OUT_XLSX_PARENT = os.path.join(os.path.dirname(BASE_DIR), "Bao_Cao_Tong_Hop_Du_An_Retail_ML_Hub.xlsx")

FONT_NAME = "Segoe UI"
THIN_SIDE = Side(style="thin", color="CBD5E1")
MED_NAVY_SIDE = Side(style="medium", color="1E3A8A")
BORDER_CELL = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
BORDER_HEADER = Border(left=THIN_SIDE, right=THIN_SIDE, top=MED_NAVY_SIDE, bottom=MED_NAVY_SIDE)

FILL_NAVY = PatternFill("solid", fgColor="1E3A8A")
FILL_SUBHEADER = PatternFill("solid", fgColor="EFF6FF")
FILL_SEC_HEADER = PatternFill("solid", fgColor="DBEAFE")
FILL_TBL_HEADER = PatternFill("solid", fgColor="1D4ED8")
FILL_ZEBRA = PatternFill("solid", fgColor="F8FAFC")
FILL_WHITE = PatternFill("solid", fgColor="FFFFFF")
FILL_TOTAL = PatternFill("solid", fgColor="E2E8F0")
FILL_BEST = PatternFill("solid", fgColor="DCFCE7")
FILL_WARN = PatternFill("solid", fgColor="FEF3C7")

FONT_BANNER = Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF")
FONT_SUB = Font(name=FONT_NAME, size=10, italic=True, color="1E3A8A")
FONT_SEC = Font(name=FONT_NAME, size=11, bold=True, color="1E3A8A")
FONT_TH = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
FONT_TD = Font(name=FONT_NAME, size=10, bold=False, color="0F172A")
FONT_TD_BOLD = Font(name=FONT_NAME, size=10, bold=True, color="0F172A")
FONT_BEST = Font(name=FONT_NAME, size=10, bold=True, color="166534")
FONT_NAVY_BOLD = Font(name=FONT_NAME, size=10, bold=True, color="1E3A8A")

AL_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
AL_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)
AL_BANNER = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ---------------------------------------------------------------------------
# HÀM HỖ TRỢ ĐỊNH DẠNG DÙNG CHUNG
# ---------------------------------------------------------------------------
def style_banner(ws, start_col, end_col, title, subtitle):
    """Vẽ banner tiêu đề 2 dòng chuẩn cho mỗi Sheet"""
    s_let = get_column_letter(start_col)
    e_let = get_column_letter(end_col)
    ws.merge_cells(f"{s_let}1:{e_let}1")
    ws.merge_cells(f"{s_let}2:{e_let}2")

    c1 = ws.cell(1, start_col, title)
    c1.font = FONT_BANNER
    c1.fill = FILL_NAVY
    c1.alignment = AL_BANNER
    ws.row_dimensions[1].height = 30

    c2 = ws.cell(2, start_col, subtitle)
    c2.font = FONT_SUB
    c2.fill = FILL_SUBHEADER
    c2.alignment = AL_BANNER
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 10


def style_section_title(ws, row, start_col, end_col, text):
    """Vẽ thanh tiêu đề mục màu xanh nhạt"""
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.fill = FILL_SEC_HEADER
        cell.border = BORDER_CELL
    c = ws.cell(row, start_col, text)
    c.font = FONT_SEC
    c.alignment = AL_LEFT
    ws.row_dimensions[row].height = 24


def write_table_header(ws, row, headers, start_col=1, fill=None):
    """Ghi dòng tiêu đề bảng với định dạng chuẩn"""
    f = fill or FILL_TBL_HEADER
    for idx, text in enumerate(headers):
        c = ws.cell(row, start_col + idx, text)
        c.font = FONT_TH
        c.fill = f
        c.border = BORDER_HEADER
        c.alignment = AL_CENTER
    ws.row_dimensions[row].height = 26


def apply_row_style(ws, row, start_col, end_col, alignments=None, num_formats=None,
                    fill=None, font=None, bold_cols=None, height=22):
    """Áp dụng định dạng chuẩn cho 1 dòng dữ liệu bảng"""
    bold_set = set(bold_cols or [])
    for idx, col in enumerate(range(start_col, end_col + 1)):
        c = ws.cell(row, col)
        c.border = BORDER_CELL
        c.fill = fill if fill is not None else FILL_WHITE
        if font is not None:
            c.font = font
        else:
            c.font = FONT_TD_BOLD if idx in bold_set else FONT_TD
        if alignments and idx < len(alignments):
            c.alignment = alignments[idx]
        else:
            c.alignment = AL_LEFT
        if num_formats and idx in num_formats:
            c.number_format = num_formats[idx]
    if height:
        ws.row_dimensions[row].height = height


def set_col_widths(ws, widths_dict):
    """Thiết lập chiều rộng từng cột theo dictionary {col_letter_or_idx: width}"""
    for k, w in widths_dict.items():
        col_letter = get_column_letter(k) if isinstance(k, int) else k
        ws.column_dimensions[col_letter].width = w


def _estimate_wrapped_lines(text, col_width, char_factor=0.82):
    """Ước lượng số dòng sau khi wrap chữ tiếng Việt/Anh trong Excel"""
    if text is None:
        return 1
    s = str(text)
    if not s or s.startswith("="):
        return 1
    eff_chars = max(6, int((col_width - 1.5) * char_factor))
    total_lines = 0
    for para in s.split("\n"):
        if not para:
            total_lines += 1
            continue
        lines = 1
        cur_len = 0
        for word in para.split(" "):
            w_len = len(word)
            if cur_len == 0:
                if w_len > eff_chars:
                    lines += (w_len - 1) // eff_chars
                    cur_len = ((w_len - 1) % eff_chars) + 1
                else:
                    cur_len = w_len
            elif cur_len + 1 + w_len <= eff_chars:
                cur_len += 1 + w_len
            else:
                lines += 1
                if w_len > eff_chars:
                    lines += (w_len - 1) // eff_chars
                    cur_len = ((w_len - 1) % eff_chars) + 1
                else:
                    cur_len = w_len
        total_lines += lines
    return max(1, total_lines)


def auto_fit_rows(ws, start_row, end_row, col_widths, start_col=1, end_col=None,
                  char_factor=0.82, line_height=15.5, padding=10, min_height=24):
    """Tự động tính chiều cao từng hàng dựa trên nội dung chữ & bề rộng cột để không bị cắt chữ"""
    max_col = end_col or ws.max_column
    merged_widths = {}
    for mr in ws.merged_cells.ranges:
        w_sum = sum(col_widths.get(c, 12) for c in range(mr.min_col, mr.max_col + 1))
        for r in range(mr.min_row, mr.max_row + 1):
            merged_widths[(r, mr.min_col)] = w_sum

    for r in range(start_row, end_row + 1):
        max_lines = 1
        for c in range(start_col, max_col + 1):
            cell_val = ws.cell(r, c).value
            w = merged_widths.get((r, c), col_widths.get(c, 12))
            lines = _estimate_wrapped_lines(cell_val, w, char_factor=char_factor)
            if lines > max_lines:
                max_lines = lines
        calc_h = max(min_height, max_lines * line_height + padding)
        cur_h = ws.row_dimensions[r].height or 0
        ws.row_dimensions[r].height = max(cur_h, calc_h)


def add_embedded_image(ws, img_path, anchor_cell, max_width_px=640):
    """Nhúng ảnh PNG vào Excel giữ nguyên tỷ lệ khung hình"""
    if not os.path.exists(img_path):
        print(f"  [Cảnh báo] Không tìm thấy ảnh để nhúng: {img_path}")
        return 0
    with PILImage.open(img_path) as im:
        w, h = im.size
    scale = min(1.0, max_width_px / float(w))
    xl_img = XLImage(img_path)
    xl_img.width = int(round(w * scale))
    xl_img.height = int(round(h * scale))
    ws.add_image(xl_img, anchor_cell)
    return xl_img.height


# ---------------------------------------------------------------------------
# SHEET 1: 1_Dataset (73.100 DÒNG x 15 CỘT + BẢNG 1.1 + BẢNG 4.1 + KPI CÔNG THỨC)
# ---------------------------------------------------------------------------
def build_sheet_dataset(wb):
    print("[1/6] Đang tạo Sheet 1_Dataset (nạp 73.100 bản ghi gốc & đặc tả)...")
    ws = wb.active
    ws.title = "1_Dataset"
    ws.views.sheetView[0].showGridLines = True

    df = pd.read_csv(DATA_CSV)
    n_rows = len(df)
    headers = list(df.columns)

    # Banner & Header bảng gốc (cột A..O)
    style_banner(
        ws, 1, 15,
        "DỮ LIỆU GỐC BÁN LẺ & TỒN KHO — RETAIL ML HUB (73.100 DÒNG × 15 CỘT)",
        "Nguồn: Dataset/retail_store_inventory.csv | Giai đoạn: 01/01/2022 – 01/01/2024 (731 ngày) | 5 Cửa hàng (S001–S005) × 20 SKU (P0001–P0020) | 0 Giá trị Null"
    )
    write_table_header(ws, 4, headers, start_col=1)

    # Ghi nhanh 73.100 dòng bằng itertuples (tối ưu hiệu năng < 5s)
    num_fmt_cols = {
        6: "#,##0",      # Inventory Level
        7: "#,##0",      # Units Sold
        8: "#,##0",      # Units Ordered
        9: "#,##0.00",   # Demand Forecast
        10: "#,##0.00",  # Price
        11: "0",         # Discount (%)
        13: "0",         # Holiday/Promotion
        14: "#,##0.00",  # Competitor Pricing
    }
    for row_idx, rec in enumerate(df.itertuples(index=False, name=None), start=5):
        for col_idx, val in enumerate(rec, start=1):
            c = ws.cell(row_idx, col_idx, val)
            if col_idx in num_fmt_cols:
                c.number_format = num_fmt_cols[col_idx]

    ws.auto_filter.ref = f"A4:O{4 + n_rows}"
    ws.freeze_panes = "A5"

    # Cột P làm khoảng đệm; Cột Q..U đặt Bảng 1.2, Bảng 4.2 & Khối Công thức Excel
    style_banner(
        ws, 17, 21,
        "ĐẶC TẢ DỮ LIỆU GỐC (BẢNG 1.2 WORD) & 28 BIẾN ĐẶC TRƯNG (BẢNG 4.2 WORD)",
        "Đối chiếu trực tiếp Mục 1.2 (Chương 1) & Mục 4.2 (Chương 4) trong Báo cáo MLBA.docx"
    )

    # Khối A: Bảng 1.2 trong Word (15 thuộc tính)
    style_section_title(ws, 4, 17, 21, "BẢNG 1.2 (WORD): MÔ TẢ CHI TIẾT 15 THUỘC TÍNH DỮ LIỆU GỐC (MỤC 1.2)")
    t11 = pd.read_csv(os.path.join(REPORTS_DIR, "table_1_1_mo_ta_15_cot_du_lieu.csv"))
    write_table_header(ws, 5, list(t11.columns), start_col=17)
    for idx, rec in enumerate(t11.itertuples(index=False, name=None), start=6):
        for c_i, val in enumerate(rec):
            ws.cell(idx, 17 + c_i, val)
        apply_row_style(
            ws, idx, 17, 21,
            alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_LEFT, AL_LEFT],
            fill=FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE,
            bold_cols=[1],
            height=28,
        )

    # Khối B: Thống kê động bằng Công thức Excel tham chiếu A5:O73104
    last_data_row = 4 + n_rows
    stat_start = 22
    style_section_title(ws, stat_start, 17, 21, "THỐNG KÊ NHANH DỮ LIỆU GỐC (TÍNH TRỰC TIẾP BẰNG CÔNG THỨC EXCEL)")
    write_table_header(ws, stat_start + 1, ["STT", "Chỉ tiêu thống kê", "Công thức áp dụng", "Giá trị tính toán", "Đơn vị / Ghi chú"], start_col=17)
    formula_rows = [
        (1, "Tổng số bản ghi giao dịch (Rows)", f"=COUNTA(A5:A{last_data_row})", "#,##0", "Đủ 5 cửa hàng × 20 SKU × 731 ngày"),
        (2, "Số ô bị thiếu ở biến mục tiêu Units Sold", f"=COUNTBLANK(G5:G{last_data_row})", "#,##0", "Dữ liệu hoàn toàn sạch (0 Null)"),
        (3, "Tổng sản lượng tiêu thụ toàn chuỗi (Units Sold)", f"=SUM(G5:G{last_data_row})", "#,##0", "Đơn vị sản phẩm (2022–2024)"),
        (4, "Sản lượng bán trung bình / ngày / SKU", f"=AVERAGE(G5:G{last_data_row})", "#,##0.00", "Đơn vị / ngày"),
        (5, "Mức tồn kho trung bình (Inventory Level)", f"=AVERAGE(F5:F{last_data_row})", "#,##0.00", "Đơn vị / ngày"),
        (6, "Giá bán niêm yết trung bình (Price)", f"=AVERAGE(J5:J{last_data_row})", "$#,##0.00", "USD / đơn vị"),
        (7, "Tỷ lệ chiết khấu trung bình (Discount)", f"=AVERAGE(K5:K{last_data_row})", "0.00\"%\"", "Phần trăm giảm giá"),
        (8, "Giá đối thủ trung bình (Competitor Pricing)", f"=AVERAGE(N5:N{last_data_row})", "$#,##0.00", "USD / đơn vị"),
    ]
    for offset, (stt, label, formula, num_fmt, note) in enumerate(formula_rows):
        r = stat_start + 2 + offset
        ws.cell(r, 17, stt)
        ws.cell(r, 18, label)
        ws.cell(r, 19, formula.lstrip("="))
        ws.cell(r, 20, formula)
        ws.cell(r, 21, note)
        apply_row_style(
            ws, r, 17, 21,
            alignments=[AL_CENTER, AL_LEFT, AL_CENTER, AL_RIGHT, AL_LEFT],
            num_formats={3: num_fmt},
            fill=FILL_SUBHEADER if offset == 2 else (FILL_ZEBRA if r % 2 == 1 else FILL_WHITE),
            bold_cols=[1, 3],
            height=22,
        )

    # Khối C: Bảng 4.2 trong Word (28 biến đặc trưng đầu vào)
    t41_start = stat_start + len(formula_rows) + 3
    style_section_title(ws, t41_start, 17, 21, "BẢNG 4.2 (WORD): DANH SÁCH 28 BIẾN ĐẶC TRƯNG ĐẦU VÀO (FEATURE ENGINEERING)")
    t41 = pd.read_csv(os.path.join(REPORTS_DIR, "table_4_1_danh_sach_28_bien_dac_trung.csv"))
    write_table_header(ws, t41_start + 1, list(t41.columns), start_col=17)
    for idx, rec in enumerate(t41.itertuples(index=False, name=None), start=t41_start + 2):
        for c_i, val in enumerate(rec):
            ws.cell(idx, 17 + c_i, val)
        apply_row_style(
            ws, idx, 17, 21,
            alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_LEFT, AL_LEFT],
            fill=FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE,
            bold_cols=[1],
            height=26,
        )

    widths_1 = {
        1: 12, 2: 10, 3: 11, 4: 13, 5: 10, 6: 15, 7: 12, 8: 14,
        9: 16, 10: 10, 11: 10, 12: 18, 13: 17, 14: 18, 15: 13,
        16: 4, 17: 7, 18: 28, 19: 26, 20: 36, 21: 48,
    }
    set_col_widths(ws, widths_1)
    auto_fit_rows(ws, 5, t41_start + 1 + len(t41), widths_1, start_col=17, end_col=21)


# ---------------------------------------------------------------------------
# SHEET 2: 2_PCCV (BẢNG PHÂN CÔNG CÔNG VIỆC 5 THÀNH VIÊN)
# ---------------------------------------------------------------------------
def build_sheet_pccv(wb):
    print("[2/6] Đang tạo Sheet 2_PCCV (Bảng Phân công Công việc 5 thành viên)...")
    ws = wb.create_sheet("2_PCCV")
    ws.views.sheetView[0].showGridLines = True

    style_banner(
        ws, 1, 8,
        "BẢNG PHÂN CÔNG CÔNG VIỆC NHÓM (PCCV) — ĐỒ ÁN RETAIL ML HUB",
        "Học phần: Học máy trong Phân tích Kinh doanh (MLBA) | Nhóm thực hiện: Khang, Hiếu, Minh, Ân, Lợi | Cập nhật: 15/11/2026"
    )

    # Bảng 1: Tóm tắt vai trò & mức độ hoàn thành 5 thành viên (Khớp mục 2 ảnh chụp)
    style_section_title(ws, 4, 1, 8, "1. TÓM TẮT VAI TRÒ, NHIỆM VỤ CHÍNH THEO CHƯƠNG & MỨC ĐỘ HOÀN THÀNH (5 THÀNH VIÊN)")
    write_table_header(ws, 5, [
        "STT", "Thành viên", "Vai trò chính", "Chương / Module phụ trách",
        "Nhiệm vụ trọng tâm & Sản phẩm bàn giao", "Tỷ trọng CV", "Mức độ hoàn thành", "Trạng thái"
    ])

    members = [
        (1, "Khang", "Nhóm trưởng (Team Lead) & Kiến trúc Dữ liệu", "Chương 1 (Dữ liệu & EDA) & Điều phối chung",
         "Thu thập & kiểm tra dữ liệu retail_store_inventory.csv (73.100 dòng); chuẩn hóa Bảng 1.2 (15 thuộc tính); xây dựng 7 biểu đồ EDA (Hình 1.1–1.7); lập kế hoạch Gantt (20/10–15/11/2026) & rà soát toàn bộ báo cáo Word.",
         0.20, 1.00, "Hoàn thành 100%"),
        (2, "Hiếu", "Trưởng nhóm Học máy (ML Lead) & Backend FastAPI", "Chương 4 (ML) & Backend RESTful API",
         "Thiết kế pipeline kỹ thuật 28 đặc trưng (Bảng 4.2); chạy kiểm định dừng ADF & KPSS (Bảng 4.1, Mục 4.1.2); huấn luyện & lưu 5 mô hình ARIMA, SARIMA, Holt-Winters, Random Forest, XGBoost (.pkl); xây dựng FastAPI backend (backend/api.py, forecast_service.py).",
         0.20, 1.00, "Hoàn thành 100%"),
        (3, "Minh", "Chuyên viên Phương pháp & Tối ưu Chuỗi cung ứng", "Chương 2 (Phương pháp) & Tối ưu Tồn kho",
         "Soạn cơ sở lý thuyết 5 mô hình, thước đo RMSE/MAE/MAPE & lý thuyết tồn kho (Mục 2.2–2.4); cài đặt công thức tính Tồn kho an toàn (SS = 311), Điểm đặt hàng lại (ROP = 777), đề xuất đặt hàng và Ma trận ABC–XYZ (inventory_optimizer.py).",
         0.20, 1.00, "Hoàn thành 100%"),
        (4, "Ân", "Chuyên viên Nghiên cứu Tài liệu & Tổng hợp Báo cáo", "Chương 3 (Literature Review) & Chương 6",
         "Tổng hợp tài liệu nghiên cứu khoa học trong & ngoài nước về dự báo nhu cầu bán lẻ và quản trị tồn kho (Chương 3); viết Chương 6 (Kết luận thực nghiệm, Ý nghĩa, Hướng phát triển & Kiến nghị), chuẩn hóa 14 Tài liệu tham khảo.",
         0.20, 1.00, "Hoàn thành 100%"),
        (5, "Lợi", "Kỹ sư Giao diện Web Studio & Mô hình hóa BPMN", "Chương 5 (Web Studio) & Bộ Flowchart BPMN",
         "Phát triển giao diện Web Dashboard 6 màn hình (frontend/index.html, app.js, styles) tích hợp ApexCharts offline; thiết kế bộ sơ đồ Flowchart/BPMN (.drawio + PNG 300 DPI: Sơ đồ 1.1, 4.1, 5.1–5.3); chụp ảnh & thuyết minh 6 màn hình (Hình 5.1–5.12).",
         0.20, 1.00, "Hoàn thành 100%"),
    ]

    for i, row_data in enumerate(members):
        r = 6 + i
        for c_i, val in enumerate(row_data, start=1):
            ws.cell(r, c_i, val)
        apply_row_style(
            ws, r, 1, 8,
            alignments=[AL_CENTER, AL_CENTER, AL_LEFT, AL_LEFT, AL_LEFT, AL_RIGHT, AL_RIGHT, AL_CENTER],
            num_formats={5: "0.0%", 6: "0.0%"},
            fill=FILL_ZEBRA if i % 2 == 1 else FILL_WHITE,
            bold_cols=[1, 7],
            height=48,
        )
        ws.cell(r, 8).fill = FILL_BEST
        ws.cell(r, 8).font = FONT_BEST

    # Dòng tổng kết Bảng 1
    tot_r = 11
    ws.cell(tot_r, 1, "TỔNG")
    ws.cell(tot_r, 2, "5 thành viên")
    ws.cell(tot_r, 3, "Toàn bộ dự án Retail ML Hub")
    ws.cell(tot_r, 4, "Chương 1 → Chương 6 & Web App")
    ws.cell(tot_r, 5, "Đã bàn giao đầy đủ Mã nguồn, Mô hình .pkl, Flowcharts .drawio/.png, Báo cáo Word & File Excel")
    ws.cell(tot_r, 6, "=SUM(F6:F10)")
    ws.cell(tot_r, 7, "=AVERAGE(G6:G10)")
    ws.cell(tot_r, 8, "Đạt tiến độ")
    apply_row_style(
        ws, tot_r, 1, 8,
        alignments=[AL_CENTER, AL_CENTER, AL_LEFT, AL_LEFT, AL_LEFT, AL_RIGHT, AL_RIGHT, AL_CENTER],
        num_formats={5: "0.0%", 6: "0.0%"},
        fill=FILL_TOTAL,
        font=FONT_NAVY_BOLD,
        height=28,
    )

    # Bảng 2: Phân công công việc chi tiết theo từng Chương & Module
    sec2_r = 13
    style_section_title(ws, sec2_r, 1, 8, "2. BẢNG PHÂN CÔNG CÔNG VIỆC CHI TIẾT THEO CHƯƠNG MỤC & ĐẦU MỤC KỸ THUẬT")
    write_table_header(ws, sec2_r + 1, [
        "STT", "Giai đoạn / Chương", "Đầu mục công việc chi tiết", "Chủ trì",
        "Phối hợp", "Sản phẩm / Tệp mã nguồn bàn giao", "Tiến độ", "Ghi chú kiểm định"
    ])

    detail_tasks = [
        (1, "Mở đầu & Chương 1", "Xác định đề tài, mục tiêu kinh doanh, phạm vi dữ liệu & hoàn thiện Bảng 1.1", "Khang", "Ân, Minh", "Mục 1.1 & Bảng 1.1 (MLBA.docx)", 1.0, "Khớp 5 Store × 20 SKU × 731 ngày"),
        (2, "Chương 1 — Dữ liệu", "Kiểm tra kiểu dữ liệu, phát hiện Null (0), tổng hợp Bảng 1.2 mô tả 15 thuộc tính", "Khang", "Hiếu", "Mục 1.2 + table_1_1_mo_ta_15_cot_du_lieu.csv", 1.0, "Loại cột Demand Forecast tránh rò rỉ"),
        (3, "Chương 1 — EDA", "Trực quan hóa 7 biểu đồ EDA (Hình 1.1–1.7) & hoàn thiện Sơ đồ 1.1 (8 bước)", "Khang", "Lợi", "src/eda_analysis.py + Hình 1.1–1.7 (.png)", 1.0, "Xuất 300 DPI + Tổng hợp Bảng 1.3"),
        (4, "Chương 2 — Phương pháp", "Hệ thống hóa lý thuyết 5 mô hình & các chỉ số đánh giá RMSE/MAE/MAPE (Mục 2.1–2.3)", "Minh", "Hiếu", "Chương 2 (Mục 2.1–2.3 MLBA.docx) + config.py", 1.0, "Bổ sung FastAPI & công thức sai số"),
        (5, "Chương 2 — Công thức Kho", "Xây dựng lý thuyết Tồn kho an toàn (SS), Điểm đặt hàng lại (ROP) & Ma trận ABC–XYZ (Mục 2.4)", "Minh", "Khang", "Mục 2.4 (MLBA.docx) + src/inventory_optimizer.py", 1.0, "Z=1,645 (SL=95%), L=3 ngày"),
        (6, "Chương 3 — Tài liệu", "Tổng quan 7+ công trình nghiên cứu khoa học trong & ngoài nước về dự báo bán lẻ (Bảng 3.1)", "Ân", "Khang", "Chương 3 + Mục Tài liệu tham khảo (14 nguồn)", 1.0, "Trích dẫn thống nhất cuối báo cáo"),
        (7, "Chương 4 — Chẩn đoán", "Phân rã mùa vụ, vẽ ACF/PACF, Rolling Mean/Std & kiểm định ADF (-27,2824), KPSS (0,0506)", "Hiếu", "Minh", "Mục 4.1 + Bảng 4.1 + Hình 4.1–4.3 + Sơ đồ 4.1", 1.0, "Đã lưu CSV stationarity_tests"),
        (8, "Chương 4 — Đặc trưng", "Xây dựng pipeline chuẩn hóa ngày ('D'), tạo 28 biến đặc trưng (Bảng 4.2) & chia Train/Test", "Hiếu", "Minh", "Mục 4.2 + Bảng 4.2–4.3 (560 Train / 141 Test ML)", 1.0, "Chuỗi đơn biến: 584/147 ngày"),
        (9, "Chương 4 — 5 Mô hình", "Huấn luyện 5 mô hình, trích xuất Feature Importance & đối sánh (Bảng 4.5–4.6, Hình 4.4–4.10)", "Hiếu", "Khang", "src/forecasting_models.py + saved_models/*.pkl", 1.0, "RF #1 (RMSE 92,42 | MAE 73,71)"),
        (10, "Chương 4 — Tối ưu Kho", "Tính SS=311, ROP=777 (ML) & phân loại ABC–XYZ cho 20 SKU tại S001 (Mục 4.5, Bảng 4.7)", "Minh", "Hiếu", "Mục 4.5 + Hình 4.11–4.12 + abc_xyz_analysis_s001.csv", 1.0, "S001: AZ(15), BZ(3), CZ(2)"),
        (11, "Chương 5 — Backend API", "Phát triển 10 REST API endpoint trên FastAPI, RAM Cache & dự báo trực tuyến on-the-fly", "Hiếu", "Lợi", "backend/api.py + src/forecast_service.py", 1.0, "Hỗ trợ chọn 5 Store × 20 SKU"),
        (12, "Chương 5 — Web Studio UI", "Xây dựng giao diện 6 màn hình (Hình 5.1–5.12), Light/Dark & xuất CSV offline (Bảng 5.1)", "Lợi", "Khang, Minh", "frontend/index.html, js/, css/, vendor/", 1.0, "S001: 20/20 SKU mức Nguy cấp"),
        (13, "Chương 5 — BPMN & Flow", "Thiết kế bộ 6 file Draw.io BPMN Swimlane + PNG 300 DPI kèm icon (Sơ đồ 1.1, 4.1, 5.1–5.3)", "Lợi", "Hiếu", "reports/figures/flowcharts/*.drawio & *.png", 1.0, "XML verified + đồng bộ Word/Excel"),
        (14, "Chương 6 — Kết luận", "Hoàn thiện Chương 6 (6.1–6.4), đồng bộ SS=311/ROP=777 & 14 Tài liệu tham khảo", "Ân", "Khang, Hiếu", "Chương 6 + Tài liệu tham khảo (MLBA.docx)", 1.0, "Khớp 100% giữa Word & Excel"),
        (15, "Kiểm thử & Đóng gói", "Đồng bộ Báo cáo Word MLBA.docx + Excel 6 Sheet (20/10–15/11/2026) & chuẩn bị Slide bảo vệ", "Khang", "Cả nhóm", "MLBA.docx + Bao_Cao_Tong_Hop_Du_An_Retail_ML_Hub.xlsx", 1.0, "Sẵn sàng nộp & thuyết trình"),
    ]

    for idx, t in enumerate(detail_tasks):
        r = sec2_r + 2 + idx
        for c_i, val in enumerate(t, start=1):
            ws.cell(r, c_i, val)
        apply_row_style(
            ws, r, 1, 8,
            alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_CENTER, AL_CENTER, AL_LEFT, AL_RIGHT, AL_LEFT],
            num_formats={6: "0%"},
            fill=FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE,
            bold_cols=[1, 3],
            height=28,
        )

    ws.freeze_panes = "A6"
    widths_2 = {1: 8, 2: 20, 3: 32, 4: 26, 5: 48, 6: 30, 7: 16, 8: 24}
    set_col_widths(ws, widths_2)
    auto_fit_rows(ws, 5, sec2_r + 1 + len(detail_tasks), widths_2, start_col=1, end_col=8)


# ---------------------------------------------------------------------------
# SHEET 3: 3_Gantt_Chart (BIỂU ĐỒ GANTT TIẾN ĐỘ TỪ CUỐI THÁNG 10 ĐẾN 15/11/2026)
# ---------------------------------------------------------------------------
def build_sheet_gantt(wb):
    print("[3/6] Đang tạo Sheet 3_Gantt_Chart (Biểu đồ Gantt 20/10 – 15/11/2026)...")
    ws = wb.create_sheet("3_Gantt_Chart")
    ws.views.sheetView[0].showGridLines = True

    style_banner(
        ws, 1, 20,
        "BIỂU ĐỒ GANTT TIẾN ĐỘ THỰC HIỆN ĐỒ ÁN (20/10/2026 – 15/11/2026) — RETAIL ML HUB",
        "Khởi động: Cuối tháng 10 (20/10/2026) → Hoàn tất: 15/11/2026 (27 ngày, chia 10 mốc M1–M10) | 6 Giai đoạn (18 hạng mục) | Nhóm: Khang, Hiếu, Minh, Ân, Lợi"
    )

    # Chú giải màu Giai đoạn tại dòng 4
    phase_colors = {
        "GĐ1: Khởi tạo & EDA (Ch.1)": ("2563EB", "DBEAFE"),
        "GĐ2: Lý thuyết & Tài liệu (Ch.2–3)": ("0891B2", "CFFAFE"),
        "GĐ3: Đặc trưng & 5 Mô hình (Ch.4)": ("7C3AED", "EDE9FE"),
        "GĐ4: Tối ưu Tồn kho (Ch.4)": ("D97706", "FEF3C7"),
        "GĐ5: Web Studio & BPMN (Ch.5)": ("059669", "D1FAE5"),
        "GĐ6: Kiểm thử & Báo cáo (Ch.6)": ("DC2626", "FEE2E2"),
    }
    ws.cell(4, 1, "Chú giải màu:").font = FONT_NAVY_BOLD
    ws.cell(4, 1).alignment = AL_LEFT
    legend_items = list(phase_colors.items())
    legend_spans = [(2, 4), (5, 7), (8, 10), (11, 13), (14, 16), (17, 20)]
    for (p_name, (bar_hex, bg_hex)), (c_s, c_e) in zip(legend_items, legend_spans):
        ws.merge_cells(start_row=4, start_column=c_s, end_row=4, end_column=c_e)
        for col in range(c_s, c_e + 1):
            cell = ws.cell(4, col)
            cell.fill = PatternFill("solid", fgColor=bg_hex)
            cell.border = BORDER_CELL
        cell0 = ws.cell(4, c_s, f"■ {p_name}")
        cell0.font = Font(name=FONT_NAME, size=9, bold=True, color=bar_hex)
        cell0.alignment = AL_CENTER
    ws.row_dimensions[4].height = 34
    ws.row_dimensions[5].height = 8

    # 10 chặng mốc thực tế từ cuối tháng 10 (20/10) đến giữa tháng 11 (15/11/2026)
    milestone_cols = [
        "M1\n20–22/10", "M2\n23–25/10", "M3\n26–28/10", "M4\n29–31/10",
        "M5\n01–03/11", "M6\n04–06/11", "M7\n07–09/11", "M8\n10–11/11",
        "M9\n12–13/11", "M10\n14–15/11",
    ]
    headers = [
        "WBS", "Giai đoạn", "Hạng mục công việc", "Chương", "Chủ trì",
        "Phối hợp", "Ngày BĐ", "Ngày KT", "Số ngày", "Tiến độ",
        *milestone_cols
    ]
    write_table_header(ws, 6, headers, start_col=1)
    ws.row_dimensions[6].height = 34

    d = datetime.date
    # (WBS, Giai đoạn, Hạng mục, Chương, Chủ trì, Phối hợp, Ngày BĐ, Ngày KT, Mốc BĐ, Mốc KT, Tiến độ)
    gantt_tasks = [
        ("1.1", "GĐ1: Khởi tạo & EDA (Ch.1)", "Xác định đề tài, phạm vi bài toán & thu thập dữ liệu (73.100 dòng)", "Mở đầu", "Khang", "Cả nhóm", d(2026, 10, 20), d(2026, 10, 21), 1, 1, 1.0),
        ("1.2", "GĐ1: Khởi tạo & EDA (Ch.1)", "Kiểm định cấu trúc, làm sạch dữ liệu & tạo Bảng 1.2 (15 thuộc tính)", "Chương 1", "Khang", "Hiếu", d(2026, 10, 21), d(2026, 10, 24), 1, 2, 1.0),
        ("1.3", "GĐ1: Khởi tạo & EDA (Ch.1)", "Phân tích khám phá EDA, vẽ Hình 1.1–1.7 & Sơ đồ 1.1 (8 bước)", "Chương 1", "Khang", "Lợi", d(2026, 10, 23), d(2026, 10, 25), 2, 2, 1.0),
        ("2.1", "GĐ2: Lý thuyết & Tài liệu (Ch.2–3)", "Nghiên cứu phương pháp dự báo (ARIMA, SARIMA, ES, RF, XGB)", "Chương 2", "Minh", "Hiếu", d(2026, 10, 23), d(2026, 10, 28), 2, 3, 1.0),
        ("2.2", "GĐ2: Lý thuyết & Tài liệu (Ch.2–3)", "Lược khảo công trình nghiên cứu khoa học trong & ngoài nước (Bảng 3.1)", "Chương 3", "Ân", "Khang", d(2026, 10, 24), d(2026, 10, 31), 2, 4, 1.0),
        ("2.3", "GĐ2: Lý thuyết & Tài liệu (Ch.2–3)", "Chuẩn hóa công thức RMSE/MAE/MAPE, SS, ROP & ma trận ABC–XYZ (Mục 2.3–2.4)", "Chương 2", "Minh", "Khang", d(2026, 10, 26), d(2026, 10, 31), 3, 4, 1.0),
        ("3.1", "GĐ3: Đặc trưng & 5 Mô hình (Ch.4)", "Phân rã mùa vụ, ACF/PACF, Rolling & kiểm định dừng ADF/KPSS (Hình 4.1–4.3)", "Chương 4", "Hiếu", "Minh", d(2026, 10, 28), d(2026, 11, 2), 3, 5, 1.0),
        ("3.2", "GĐ3: Đặc trưng & 5 Mô hình (Ch.4)", "Thiết kế 28 biến đặc trưng (Lag, Rolling, Giá, Ngoại cảnh — Bảng 4.2)", "Chương 4", "Hiếu", "Minh", d(2026, 10, 29), d(2026, 11, 3), 4, 5, 1.0),
        ("3.3", "GĐ3: Đặc trưng & 5 Mô hình (Ch.4)", "Huấn luyện 5 mô hình dự báo (Train 80% / Test 20%) & lưu .pkl", "Chương 4", "Hiếu", "Khang", d(2026, 11, 1), d(2026, 11, 6), 5, 6, 1.0),
        ("3.4", "GĐ3: Đặc trưng & 5 Mô hình (Ch.4)", "So sánh RMSE/MAE/MAPE (Bảng 4.5) & phân tích Feature Importance (Hình 4.10)", "Chương 4", "Hiếu", "Ân", d(2026, 11, 4), d(2026, 11, 7), 6, 7, 1.0),
        ("4.1", "GĐ4: Tối ưu Tồn kho (Ch.4)", "Xây dựng động cơ tính SS=311, ROP=777 & cảnh báo đặt hàng 3 mức", "Chương 4", "Minh", "Hiếu", d(2026, 11, 5), d(2026, 11, 8), 6, 7, 1.0),
        ("4.2", "GĐ4: Tối ưu Tồn kho (Ch.4)", "Phân tích Ma trận ABC–XYZ 20 SKU (S001: AZ 15, BZ 3, CZ 2) & What-If", "Chương 4", "Minh", "Lợi", d(2026, 11, 7), d(2026, 11, 9), 7, 7, 1.0),
        ("5.1", "GĐ5: Web Studio & BPMN (Ch.5)", "Xây dựng Backend FastAPI (10 API endpoints + RAM Cache)", "Chương 5", "Hiếu", "Lợi", d(2026, 11, 6), d(2026, 11, 10), 6, 8, 1.0),
        ("5.2", "GĐ5: Web Studio & BPMN (Ch.5)", "Phát triển Frontend Web Studio 6 màn hình (Hình 5.1–5.12) + xuất CSV offline", "Chương 5", "Lợi", "Hiếu, Minh", d(2026, 11, 7), d(2026, 11, 12), 7, 9, 1.0),
        ("5.3", "GĐ5: Web Studio & BPMN (Ch.5)", "Thiết kế bộ Flowchart/BPMN kèm icon (.drawio + PNG) cho Ch.1, 4, 5", "Ch.1,4,5", "Lợi", "Khang", d(2026, 11, 10), d(2026, 11, 13), 8, 9, 1.0),
        ("6.1", "GĐ6: Kiểm thử & Báo cáo (Ch.6)", "Kiểm thử tích hợp toàn hệ thống & đối soát số liệu Word/Excel/Web", "Kiểm thử", "Khang", "Hiếu, Lợi", d(2026, 11, 11), d(2026, 11, 13), 8, 9, 1.0),
        ("6.2", "GĐ6: Kiểm thử & Báo cáo (Ch.6)", "Hoàn thiện toàn bộ Báo cáo Word (MLBA.docx) & File Excel 6 Sheet", "Chương 6", "Ân", "Khang, Lợi", d(2026, 11, 12), d(2026, 11, 15), 9, 10, 1.0),
        ("6.3", "GĐ6: Kiểm thử & Báo cáo (Ch.6)", "Đóng gói sản phẩm, kiểm tra lần cuối & chuẩn bị Slide bảo vệ", "Hoàn tất", "Khang", "Cả nhóm", d(2026, 11, 14), d(2026, 11, 15), 10, 10, 1.0),
    ]

    first_r = 7
    for idx, (wbs, phase, task, ch, lead, sup, s_dt, e_dt, s_m, e_m, prog) in enumerate(gantt_tasks):
        r = first_r + idx
        ws.cell(r, 1, wbs)
        ws.cell(r, 2, phase)
        ws.cell(r, 3, task)
        ws.cell(r, 4, ch)
        ws.cell(r, 5, lead)
        ws.cell(r, 6, sup)
        ws.cell(r, 7, s_dt)
        ws.cell(r, 8, e_dt)
        ws.cell(r, 9, f"=H{r}-G{r}+1")
        ws.cell(r, 10, prog)

        apply_row_style(
            ws, r, 1, 10,
            alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_CENTER, AL_CENTER, AL_CENTER, AL_CENTER, AL_CENTER, AL_CENTER, AL_RIGHT],
            num_formats={6: "dd/mm", 7: "dd/mm", 8: "0\" ng\"", 9: "0%"},
            fill=FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE,
            bold_cols=[0, 4],
            height=24,
        )

        bar_hex, bg_hex = phase_colors[phase]
        bar_fill = PatternFill("solid", fgColor=bar_hex)
        inactive_fill = PatternFill("solid", fgColor="F8FAFC")
        white_side = Side(style="thin", color="FFFFFF")
        bar_border = Border(left=white_side, right=white_side, top=THIN_SIDE, bottom=THIN_SIDE)

        for m_idx in range(1, 11):
            c = ws.cell(r, 10 + m_idx)
            c.alignment = AL_CENTER
            if s_m <= m_idx <= e_m:
                c.value = "■"
                c.font = Font(name=FONT_NAME, size=10, color=bar_hex)
                c.fill = bar_fill
                c.border = bar_border
            else:
                c.fill = inactive_fill
                c.border = BORDER_CELL

    last_r = first_r + len(gantt_tasks) - 1
    sum_r = last_r + 1
    ws.cell(sum_r, 1, "TỔNG")
    ws.cell(sum_r, 2, "6 Giai đoạn")
    ws.cell(sum_r, 3, f"{len(gantt_tasks)} hạng mục (20/10/2026 – 15/11/2026 | Tổng cộng 27 ngày lịch)")
    ws.cell(sum_r, 4, "Ch.1–6")
    ws.cell(sum_r, 5, "5 TV")
    ws.cell(sum_r, 6, "100%")
    ws.cell(sum_r, 7, f"=MIN(G{first_r}:G{last_r})")
    ws.cell(sum_r, 8, f"=MAX(H{first_r}:H{last_r})")
    ws.cell(sum_r, 9, f"=H{sum_r}-G{sum_r}+1")
    ws.cell(sum_r, 10, f"=AVERAGE(J{first_r}:J{last_r})")
    apply_row_style(
        ws, sum_r, 1, 10,
        alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_CENTER, AL_CENTER, AL_CENTER, AL_CENTER, AL_CENTER, AL_CENTER, AL_RIGHT],
        num_formats={6: "dd/mm", 7: "dd/mm", 8: "0\" ng\"", 9: "0.0%"},
        fill=FILL_TOTAL,
        font=FONT_NAVY_BOLD,
        height=24,
    )
    for m_idx in range(1, 11):
        c_let = get_column_letter(10 + m_idx)
        c = ws.cell(sum_r, 10 + m_idx, f"=COUNTIF({c_let}{first_r}:{c_let}{last_r},\"■\")")
        c.font = FONT_NAVY_BOLD
        c.fill = FILL_TOTAL
        c.border = BORDER_CELL
        c.alignment = AL_CENTER

    # Bảng tóm tắt 6 Cột mốc Bàn giao (Milestones) bên dưới Gantt
    ms_sec_r = sum_r + 2
    style_section_title(ws, ms_sec_r, 1, 10, "TỔNG HỢP 6 CỘT MỐC BÀN GIAO THEO GIAI ĐOẠN (20/10/2026 – 15/11/2026)")
    write_table_header(ws, ms_sec_r + 1, [
        "Mốc", "Giai đoạn", "Sản phẩm bàn giao trọng tâm", "Chương", "Chủ trì",
        "Hỗ trợ", "Từ ngày", "Đến ngày", "Số ngày", "Trạng thái"
    ])
    milestones = [
        ("MS1", "GĐ1: Khởi tạo & EDA (Ch.1)", "Chốt đề tài, làm sạch 73.100 dòng, Bảng 1.1–1.3 & 7 biểu đồ EDA + Sơ đồ 1.1", "Ch.1", "Khang", "Hiếu, Lợi", d(2026, 10, 20), d(2026, 10, 25), "Đạt 100%"),
        ("MS2", "GĐ2: Lý thuyết & Tài liệu (Ch.2–3)", "Hoàn thành Chương 2 (5 mô hình, RMSE/MAE/MAPE, SS/ROP, ABC-XYZ) & Chương 3", "Ch.2–3", "Minh, Ân", "Khang", d(2026, 10, 23), d(2026, 10, 31), "Đạt 100%"),
        ("MS3", "GĐ3: Đặc trưng & 5 Mô hình (Ch.4)", "Kiểm định ADF/KPSS, 28 features, huấn luyện & so sánh 5 mô hình (RF RMSE 92,42)", "Ch.4", "Hiếu", "Minh, Khang", d(2026, 10, 28), d(2026, 11, 7), "Đạt 100%"),
        ("MS4", "GĐ4: Tối ưu Tồn kho (Ch.4)", "Hoàn tất động cơ SS=311, ROP=777, Ma trận ABC–XYZ (S001) & mô phỏng What-If", "Ch.4", "Minh", "Hiếu, Lợi", d(2026, 11, 5), d(2026, 11, 9), "Đạt 100%"),
        ("MS5", "GĐ5: Web Studio & BPMN (Ch.5)", "Chạy ổn định FastAPI + 6 màn hình Web Studio offline + Bộ 6 file Draw.io BPMN", "Ch.5", "Lợi, Hiếu", "Cả nhóm", d(2026, 11, 6), d(2026, 11, 13), "Đạt 100%"),
        ("MS6", "GĐ6: Kiểm thử & Báo cáo (Ch.6)", "Đồng bộ hoàn chỉnh Báo cáo Word MLBA.docx + File Excel 6 Sheet + Slide bảo vệ", "Ch.1–6", "Khang, Ân", "Cả nhóm", d(2026, 11, 11), d(2026, 11, 15), "Hoàn tất"),
    ]
    for idx, (ms_id, ph, deliv, ch, lead, sup, dt1, dt2, st) in enumerate(milestones):
        r = ms_sec_r + 2 + idx
        vals = (ms_id, ph, deliv, ch, lead, sup, dt1, dt2, f"=H{r}-G{r}+1", st)
        for c_i, val in enumerate(vals, start=1):
            ws.cell(r, c_i, val)
        apply_row_style(
            ws, r, 1, 10,
            alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_CENTER, AL_CENTER, AL_CENTER, AL_CENTER, AL_CENTER, AL_CENTER, AL_CENTER],
            num_formats={6: "dd/mm", 7: "dd/mm", 8: "0\" ng\""},
            fill=FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE,
            bold_cols=[0, 4],
            height=24,
        )
        ws.cell(r, 10).fill = FILL_BEST
        ws.cell(r, 10).font = FONT_BEST

    ws.freeze_panes = "D7"
    widths = {1: 8, 2: 32, 3: 52, 4: 11, 5: 11, 6: 13, 7: 10, 8: 10, 9: 10, 10: 11}
    for m_idx in range(1, 11):
        widths[10 + m_idx] = 8.2
    set_col_widths(ws, widths)
    auto_fit_rows(ws, 7, ms_sec_r + 1 + len(milestones), widths, start_col=1, end_col=20)


# ---------------------------------------------------------------------------
# SHEET 4: 4_Flowchart_Ch1_Ch4 (LƯU ĐỒ CHƯƠNG 1 & 4 + DANH MỤC BIỂU ĐỒ/SƠ ĐỒ)
# ---------------------------------------------------------------------------
def build_sheet_flowchart(wb):
    print("[4/6] Đang tạo Sheet 4_Flowchart_Ch1_Ch4 (Lưu đồ Chương 1 vs 4 & Danh mục Sơ đồ)...")
    ws = wb.create_sheet("4_Flowchart_Ch1_Ch4")
    ws.views.sheetView[0].showGridLines = True

    style_banner(
        ws, 1, 7,
        "LƯU ĐỒ QUY TRÌNH (FLOWCHART CHƯƠNG 1 VS CHƯƠNG 4) & DANH MỤC BIỂU ĐỒ — SƠ ĐỒ",
        "Đối chiếu trực tiếp Sơ đồ 1.1 (8 bước EDA đa nhánh), Sơ đồ 4.1 (10 bước đối chứng 5 Mô hình & Tồn kho) và Thống kê Toàn bộ Tài sản Hình ảnh/Draw.io"
    )

    # Phần 1: Bảng thuyết minh so sánh từng bước Flowchart Chương 1 vs Chương 4
    style_section_title(ws, 4, 1, 7, "1. BẢNG ĐỐI CHIẾU TỪNG BƯỚC LƯU ĐỒ QUY TRÌNH: SƠ ĐỒ 1.1 (8 BƯỚC EDA) VS SƠ ĐỒ 4.1 (10 BƯỚC ML & TỒN KHO)")
    write_table_header(ws, 5, [
        "Bước", "Giai đoạn trong Lưu đồ", "Chương 1 — Sơ đồ 1.1 (8 bước Khám phá Dữ liệu & EDA)",
        "Đầu ra Chương 1", "Chương 4 — Sơ đồ 4.1 (10 bước 5 Mô hình ML & Tối ưu Tồn kho)",
        "Đầu ra Chương 4", "File nguồn Draw.io"
    ])

    flow_steps = [
        ("B1", "Khởi tạo & Nạp dữ liệu", "Nạp tập dữ liệu gốc retail_store_inventory.csv (73.100 dòng × 15 cột, 2022–2024).", "DataFrame 73.100 dòng", "Trích xuất chuỗi đại diện SKU S001–P0001 (get_sku_data) và chuẩn hóa tần suất ngày 'D'.", "Chuỗi 731 ngày sạch", "BPMN_0a / BPMN_0b"),
        ("B2", "Kiểm định & Chẩn đoán", "Kiểm tra khuyết Null/NaN (0 thiếu), ép kiểu Date, xác định biến mục tiêu Units Sold.", "Bảng 1.2 (15 thuộc tính)", "Phân rã chuỗi thời gian, ACF/PACF, Rolling Mean/Std & kiểm định dừng ADF / KPSS.", "ADF=-27,2824 | KPSS=0,0506", "BPMN_0a / BPMN_0b"),
        ("B3", "Xử lý & Đặc trưng hóa", "Loại cột rò rỉ Demand Forecast / Units Ordered; tính doanh thu thuần & tổng hợp KPI.", "9,97M đv | $494,97M", "Kỹ thuật 28 đặc trưng (Gốc, Lịch, 6 Lag, 4 Rolling, 2 Giá, 6 One-Hot); dropna().", "701 mẫu × 28 đặc trưng", "BPMN_0a / BPMN_0b"),
        ("B4", "Nhánh EDA 1 / Chia Train–Test", "Nhánh 1: Phân tích xu hướng bán hàng theo ngày kèm trung bình trượt 7 & 30 ngày.", "Hình 1.1 (Xu hướng)", "Chia theo thời gian 80/20: 584/147 ngày (Thống kê) và 560/141 ngày (Học máy sau lag_30).", "Bảng 4.3 (Train / Test)", "BPMN_0a / BPMN_0b"),
        ("B5", "Nhánh EDA 2 / Nhánh TS cổ điển", "Nhánh 2: Cơ cấu sản lượng tiêu thụ theo 5 Ngành hàng, 5 Cửa hàng & 4 Khu vực.", "Hình 1.2 – Hình 1.3", "Nhánh A (Thống kê đơn biến): Huấn luyện ARIMA(2,1,1), SARIMA(1,1,1)(1,0,1,7), Holt-Winters.", "Hình 4.4 – Hình 4.6", "BPMN_0a / BPMN_0b"),
        ("B6", "Nhánh EDA 3 / Nhánh Học máy", "Nhánh 3: Phân tích tác động ngoại cảnh (Thời tiết, Mùa vụ & Chương trình Khuyến mãi).", "Hình 1.4 – Hình 1.5", "Nhánh B (Học máy đa biến): Huấn luyện Random Forest (100 cây) & XGBoost (150 cây).", "Hình 4.7 – Hình 4.8", "BPMN_0a / BPMN_0b"),
        ("B7", "Nhánh EDA 4 / Đối sánh chỉ số", "Nhánh 4: Độ co giãn cầu theo tỷ số giá cạnh tranh & Ma trận tương quan Pearson.", "Hình 1.6 – Hình 1.7", "Đối sánh ngoại mẫu theo RMSE, MAE, MAPE & trích xuất Feature Importance.", "Bảng 4.5–4.6 | Hình 4.9–4.10", "BPMN_0a / BPMN_0b"),
        ("B8", "Kết luận EDA / Xuất Best Model", "Tổng hợp phát hiện EDA: xác định các biến chi phối nhu cầu làm tiền đề Chương 4.", "Bảng 1.3 (Insight EDA)", "Chọn mô hình tối ưu #1 Random Forest (RMSE 92,42; MAE 73,71) và lưu ra file .pkl.", "saved_models/*.pkl", "BPMN_0a / BPMN_0b"),
        ("B9", "Nạp dữ liệu Kho / Tối ưu SS–ROP", "(Chuyển tiếp dữ liệu sạch & KPI toàn chuỗi sang phân hệ Tổng quan Web Studio).", "Tầng Dữ liệu sạch", "Nạp cầu dự báo (155,1 đv/ngày) tính SS=311, ROP=777 & phân nhóm ABC–XYZ (AZ:15, BZ:3, CZ:2).", "Hình 4.11–4.12 | Bảng 4.7", "BPMN_0b"),
        ("B10", "Tích hợp Web Studio (Ch.5)", "Hiển thị trực tiếp trên Màn hình 1 (#/overview) với 5 thẻ KPI & 5 biểu đồ ApexCharts.", "Sơ đồ 5.1 | Hình 5.1–5.2", "Phục vụ dự báo đệ quy, cảnh báo tồn kho 3 mức & mô phỏng What-If trên FastAPI Web App.", "Màn hình #/forecast→#/whatif", "BPMN_Retail_ML_Hub_All_In_One.drawio"),
    ]

    for idx, row_data in enumerate(flow_steps):
        r = 6 + idx
        for c_i, val in enumerate(row_data, start=1):
            ws.cell(r, c_i, val)
        apply_row_style(
            ws, r, 1, 7,
            alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_LEFT, AL_LEFT, AL_LEFT, AL_LEFT],
            fill=FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE,
            bold_cols=[0, 1],
            height=38,
        )

    # Phần 2: Bảng Danh mục Toàn bộ Biểu đồ & Sơ đồ toàn đồ án (Khớp số hiệu Hình/Sơ đồ trong MLBA.docx)
    cat_r = 6 + len(flow_steps) + 1
    style_section_title(ws, cat_r, 1, 7, "2. DANH MỤC THỐNG KÊ TOÀN BỘ SƠ ĐỒ (FLOWCHART), BIỂU ĐỒ EDA, MÔ HÌNH & ẢNH CHỤP UI (ĐỒNG BỘ MLBA.DOCX)")
    write_table_header(ws, cat_r + 1, [
        "STT", "Chương", "Phân loại", "Tên Hình / Sơ đồ trong Báo cáo Word (MLBA.docx)",
        "Đường dẫn tệp (trong thư mục reports/figures/)", "Nội dung & Ý nghĩa đối chiếu", "Trạng thái tệp"
    ])

    catalog_entries = [
        # Flowcharts (5 PNG + 6 Drawio)
        ("Chương 1", "Flowchart EDA", "Sơ đồ 1.1: Quy trình Khám phá Dữ liệu & EDA (8 bước)", "flowcharts/chuong_1_flowchart_quy_trinh_kham_pha_du_lieu_eda.png", "Lưu đồ 8 bước đa nhánh từ CSV 73.100 dòng đến tổng hợp Insight EDA (kèm icon).",),
        ("Chương 4", "Flowchart ML", "Sơ đồ 4.1: Quy trình Huấn luyện 5 Mô hình & Tối ưu Tồn kho (10 bước)", "flowcharts/chuong_4_flowchart_quy_trinh_huan_luyen_5_mo_hinh.png", "Lưu đồ 10 bước 2 nhánh (Thống kê vs Học máy) đến xuất .pkl & SS/ROP/ABC-XYZ.",),
        ("Chương 5", "Flowchart BPMN", "Sơ đồ 5.1: BPMN Quy trình Khởi động & Tổng quan EDA", "flowcharts/chuong_5_hinh_5_1_flowchart_tong_quan_he_thong.png", "Luồng tương tác FastAPI RAM Cache, bộ lọc Store/SKU/Chu kỳ & 5 biểu đồ Web.",),
        ("Chương 5", "Flowchart BPMN", "Sơ đồ 5.2: BPMN Quy trình Dự báo SKU & So sánh 5 Mô hình", "flowcharts/chuong_5_hinh_5_2_flowchart_du_bao_va_so_sanh_mo_hinh.png", "Phân nhánh nạp .pkl (S001–P0001) vs huấn luyện trực tuyến và dự báo đệ quy.",),
        ("Chương 5", "Flowchart BPMN", "Sơ đồ 5.3: BPMN Quy trình Tối ưu Tồn kho SS/ROP & What-If", "flowcharts/chuong_5_hinh_5_3_flowchart_ton_kho_rop_abcxyz_whatif.png", "Tính SS, ROP, cảnh báo 3 mức, ABC–XYZ và mô phỏng độ nhạy L', SL'.",),
        ("Ch.1, 4, 5", "Nguồn Draw.io", "File nguồn Draw.io gộp 5 Tab chuẩn (kèm icon nhúng Base64)", "flowcharts/BPMN_Retail_ML_Hub_All_In_One.drawio", "Tệp XML đa tab chỉnh sửa trực tiếp toàn bộ 5 flowchart trên diagrams.net / VS Code.",),
        # Chương 1 - EDA
        ("Chương 1", "Biểu đồ EDA", "Hình 1.1: Xu hướng và Tổng Sản lượng Bán Hàng ngày", "eda/chuong_1_hinh_1_1_xu_huong_ban_hang_ngay.png", "Chuỗi bán hàng ngày 2022–2024 kèm đường trung bình trượt 7 ngày & 30 ngày.",),
        ("Chương 1", "Biểu đồ EDA", "Hình 1.2: Tổng Sản lượng Bán theo 5 Ngành hàng", "eda/chuong_1_hinh_1_2_san_luong_theo_nganh_hang.png", "So sánh sản lượng giữa Clothing, Electronics, Furniture, Groceries, Toys.",),
        ("Chương 1", "Biểu đồ EDA", "Hình 1.3: So sánh Sản lượng theo Cửa hàng & Khu vực", "eda/chuong_1_hinh_1_3_san_luong_cua_hang_khu_vuc.png", "Hiệu quả tiêu thụ giữa 5 cửa hàng (S001–S005) và 4 vùng địa lý.",),
        ("Chương 1", "Biểu đồ EDA", "Hình 1.4: Tác động của Điều kiện Thời tiết", "eda/chuong_1_hinh_1_4_tac_dong_thoi_tiet.png", "Biểu đồ Boxplot phân phối sản lượng bán theo Nắng, Mưa, Nhiều mây, Tuyết.",),
        ("Chương 1", "Biểu đồ EDA", "Hình 1.5: Mùa vụ và Chương trình Khuyến mãi", "eda/chuong_1_hinh_1_5_mua_vu_va_khuyen_mai.png", "So sánh sản lượng trung bình ngày thường vs ngày lễ/khuyến mãi qua 4 mùa.",),
        ("Chương 1", "Biểu đồ EDA", "Hình 1.6: Độ co giãn Giá (Cửa hàng / Đối thủ)", "eda/chuong_1_hinh_1_6_do_co_gian_gia.png", "Phân tán hồi quy giữa Price_Ratio và sản lượng tiêu thụ Units Sold.",),
        ("Chương 1", "Biểu đồ EDA", "Hình 1.7: Ma trận Tương quan Biến Định lượng", "eda/chuong_1_hinh_1_7_ma_tran_tuong_quan.png", "Heatmap hệ số tương quan Pearson giữa 7 biến số định lượng.",),
        # Chương 4 - Chẩn đoán chuỗi thời gian
        ("Chương 4", "Chẩn đoán TS", "Hình 4.1: Phân rã Chuỗi thời gian (Seasonal Decomposition)", "diagnostics/chuong_4_hinh_4_1_phan_ra_chuoi_thoi_gian.png", "Tách thành phần Observed, Trend, Seasonal (chu kỳ 7 ngày) và Residual.",),
        ("Chương 4", "Chẩn đoán TS", "Hình 4.2: Biểu đồ Tự tương quan ACF và PACF", "diagnostics/chuong_4_hinh_4_2_tu_tuong_quan_acf_pacf.png", "Xác định bậc tự hồi quy (p) và trung bình trượt (q) cho ARIMA/SARIMA.",),
        ("Chương 4", "Chẩn đoán TS", "Hình 4.3: Trung bình trượt & Độ lệch chuẩn trượt", "diagnostics/chuong_4_hinh_4_3_trung_binh_do_lech_truot.png", "Kiểm chứng trực quan tính dừng kết hợp cùng phép thử ADF/KPSS.",),
        # Chương 4 - 5 Mô hình & So sánh
        ("Chương 4", "Mô hình ML/TS", "Hình 4.4: Thực tế vs Dự báo — Mô hình ARIMA(2,1,1)", "models/chuong_4_hinh_4_4_arima.png", "Đối chiếu tập Train 80% / Test 20% và đường dự báo ARIMA (RMSE 107,07).",),
        ("Chương 4", "Mô hình ML/TS", "Hình 4.5: Thực tế vs Dự báo — Exponential Smoothing", "models/chuong_4_hinh_4_5_exponential_smoothing.png", "Kết quả mô hình Holt-Winters mùa vụ 7 ngày trên Test 20% (RMSE 108,21).",),
        ("Chương 4", "Mô hình ML/TS", "Hình 4.6: Thực tế vs Dự báo — SARIMA(1,1,1)(1,0,1,7)", "models/chuong_4_hinh_4_6_sarima.png", "Mô hình thống kê tốt nhất trong nhóm chuỗi thời gian (RMSE 107,03).",),
        ("Chương 4", "Mô hình ML/TS", "Hình 4.7: Thực tế vs Dự báo — Random Forest Regressor", "models/chuong_4_hinh_4_7_random_forest.png", "Mô hình tốt nhất toàn đồ án trên 28 đặc trưng (RMSE 92,42 | MAE 73,71).",),
        ("Chương 4", "Mô hình ML/TS", "Hình 4.8: Thực tế vs Dự báo — XGBoost Regressor", "models/chuong_4_hinh_4_8_xgboost.png", "Mô hình xếp hạng #2 trên tập kiểm thử (RMSE 95,73 | MAE 76,24).",),
        ("Chương 4", "So sánh & FI", "Hình 4.9: Biểu đồ So sánh Hiệu năng 5 Mô hình", "models/chuong_4_hinh_4_9_so_sanh_hieu_nang_mo_hinh.png", "Biểu đồ đối chiếu trực quan RMSE, MAE và MAPE giữa 5 thuật toán.",),
        ("Chương 4", "So sánh & FI", "Hình 4.10: Top 10 Biến Đặc trưng Quan trọng nhất", "models/chuong_4_hinh_4_10_tam_quan_trong_dac_trung.png", "Thứ tự đóng góp thông tin dự báo từ mô hình cây quyết định (Feature Importance).",),
        # Chương 4 - Tồn kho
        ("Chương 4", "Tồn kho & ROP", "Hình 4.11: Mô phỏng Diễn biến Tồn kho & Ngưỡng SS/ROP", "inventory/chuong_4_hinh_4_11_mo_phong_ton_kho_rop.png", "Mô phỏng 60 ngày tồn kho SKU S001–P0001 kèm ngưỡng ROP=707/777 & SS=312/311.",),
        ("Chương 4", "Tồn kho & ROP", "Hình 4.12: Ma trận Phân loại Tồn kho ABC – XYZ", "inventory/chuong_4_hinh_4_12_ma_tran_abc_xyz.png", "Phân bổ 20 SKU cửa hàng S001 theo Doanh thu (ABC) & Biến động CV (XYZ).",),
        # Chương 5 - 6 Giao diện UI Web Studio (Hình 5.1, 5.3, 5.5, 5.7, 5.9, 5.11 + Flowchart màn hình 5.2..5.12)
        ("Chương 5", "Ảnh chụp UI", "Hình 5.1 (UI) & Hình 5.2 (Flow): Trang Tổng quan chuỗi & EDA (#/overview)", "ui_screenshots/chuong_5_hinh_5_4_giao_dien_tong_quan_eda.png", "Hiển thị 5 thẻ KPI tổng quan, bộ lọc Store/SKU/Chu kỳ & 5 biểu đồ EDA.",),
        ("Chương 5", "Ảnh chụp UI", "Hình 5.3 (UI) & Hình 5.4 (Flow): Dự báo Nhu cầu SKU (#/forecast)", "ui_screenshots/chuong_5_hinh_5_5_giao_dien_du_bao_nhu_cau.png", "Dự báo 7/14/30 ngày (1.139 đv), dải tin cậy 95%, bảng chi tiết & xuất CSV.",),
        ("Chương 5", "Ảnh chụp UI", "Hình 5.5 (UI) & Hình 5.6 (Flow): So sánh 5 Mô hình & FI (#/models)", "ui_screenshots/chuong_5_hinh_5_6_giao_dien_so_sanh_5_mo_hinh.png", "Bảng xếp hạng trực tiếp, biểu đồ RMSE/MAE, đối chiếu kiểm thử & Top 10 Features.",),
        ("Chương 5", "Ảnh chụp UI", "Hình 5.7 (UI) & Hình 5.8 (Flow): Tồn kho An toàn & Cảnh báo ROP (#/inventory)", "ui_screenshots/chuong_5_hinh_5_7_giao_dien_ton_kho_rop.png", "Thẻ cảnh báo 20/20 SKU Nguy cấp tại S001, biểu đồ 90 ngày (SS=311, ROP=777) & xuất CSV.",),
        ("Chương 5", "Ảnh chụp UI", "Hình 5.9 (UI) & Hình 5.10 (Flow): Ma trận Chiến lược ABC – XYZ (#/abcxyz)", "ui_screenshots/chuong_5_hinh_5_8_giao_dien_ma_tran_abc_xyz.png", "Ma trận 3×3 lọc ô AZ (15 SKU), BZ (3), CZ (2), khuyến nghị & biểu đồ phân tán.",),
        ("Chương 5", "Ảnh chụp UI", "Hình 5.11 (UI) & Hình 5.12 (Flow): Mô phỏng Kịch bản What-If (#/whatif)", "ui_screenshots/chuong_5_hinh_5_9_giao_dien_mo_phong_what_if.png", "Kịch bản L'=5 ngày, SL=95%: SS=402 (+91), ROP=1.204 (+427), tổng nhập 26.784 đv.",),
    ]

    for idx, (ch, cat, title, rel_path, desc) in enumerate(catalog_entries):
        r = cat_r + 2 + idx
        abs_p = os.path.join(REPORTS_DIR, "figures", rel_path.replace("/", os.sep))
        exists_txt = "Sẵn sàng (OK)" if os.path.exists(abs_p) else "Thiếu tệp"
        vals = (idx + 1, ch, cat, title, f"reports/figures/{rel_path}", desc, exists_txt)
        for c_i, v in enumerate(vals, start=1):
            ws.cell(r, c_i, v)
        apply_row_style(
            ws, r, 1, 7,
            alignments=[AL_CENTER, AL_CENTER, AL_CENTER, AL_LEFT, AL_LEFT, AL_LEFT, AL_CENTER],
            fill=FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE,
            bold_cols=[0, 3],
            height=26,
        )
        ws.cell(r, 7).fill = FILL_BEST if os.path.exists(abs_p) else FILL_WARN
        ws.cell(r, 7).font = FONT_BEST if os.path.exists(abs_p) else FONT_TD_BOLD

    # Phần 3: Nhúng trực tiếp ảnh Flowchart Chương 1 & Chương 4 bên dưới Bảng Danh mục
    img_sec_r = cat_r + len(catalog_entries) + 3
    style_section_title(ws, img_sec_r, 1, 7, "3. HÌNH ẢNH LƯU ĐỒ FLOWCHART CHƯƠNG 1 VS CHƯƠNG 4 (KÈM ICON NHÚNG & PHONG CÁCH TRỰC QUAN)")
    ws.cell(img_sec_r + 1, 1, "Sơ đồ 1.1 — Quy trình Khám phá & Phân tích Dữ liệu (EDA — 8 bước)").font = FONT_NAVY_BOLD
    ws.cell(img_sec_r + 1, 5, "Sơ đồ 4.1 — Quy trình Huấn luyện 5 Mô hình ML & Tối ưu Tồn kho (10 bước)").font = FONT_NAVY_BOLD
    ws.row_dimensions[img_sec_r + 1].height = 22

    img_ch1 = os.path.join(REPORTS_DIR, "figures", "flowcharts", "chuong_1_flowchart_quy_trinh_kham_pha_du_lieu_eda.png")
    img_ch4 = os.path.join(REPORTS_DIR, "figures", "flowcharts", "chuong_4_flowchart_quy_trinh_huan_luyen_5_mo_hinh.png")
    add_embedded_image(ws, img_ch1, f"A{img_sec_r + 2}", max_width_px=650)
    add_embedded_image(ws, img_ch4, f"E{img_sec_r + 2}", max_width_px=650)

    ws.freeze_panes = "A6"
    widths_4 = {1: 8, 2: 20, 3: 34, 4: 28, 5: 44, 6: 42, 7: 22}
    set_col_widths(ws, widths_4)
    auto_fit_rows(ws, 5, cat_r + 1 + len(catalog_entries), widths_4, start_col=1, end_col=7)


# ---------------------------------------------------------------------------
# SHEET 5: 5_Contribution_Work (ĐÓNG GÓP CÔNG VIỆC & CHỨC NĂNG WEB STUDIO)
# ---------------------------------------------------------------------------
def build_sheet_contribution(wb):
    print("[5/6] Đang tạo Sheet 5_Contribution_Work (Contribution Evaluation & Danh mục Web Studio)...")
    ws = wb.create_sheet("5_Contribution_Work")
    ws.views.sheetView[0].showGridLines = True

    # Style chuẩn theo mẫu Contribution Evaluation (Rose / Dark Red / Times New Roman / Viền đen)
    rose_title_fill = PatternFill("solid", fgColor="E6B8AF")
    rose_header_fill = PatternFill("solid", fgColor="F4CCCC")
    rose_total_fill = PatternFill("solid", fgColor="F9EAE6")
    black_thin = Side(style="thin", color="000000")
    black_border = Border(left=black_thin, right=black_thin, top=black_thin, bottom=black_thin)

    font_ce_title = Font(name="Times New Roman", size=20, bold=True, color="980000")
    font_ce_th = Font(name="Times New Roman", size=12, bold=True, color="980000")
    font_ce_name = Font(name="Times New Roman", size=12, bold=True, color="000000")
    font_ce_body = Font(name="Times New Roman", size=12, bold=False, color="000000")

    al_ce_top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)
    al_ce_top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    al_ce_mid_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_ce_mid_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def write_contribution_block(start_row, title_text, headers, rows_data, total_row_data):
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=5)
        for col_i in range(1, 6):
            cell = ws.cell(start_row, col_i)
            cell.fill = rose_title_fill
            cell.border = black_border
        t_cell = ws.cell(start_row, 1, title_text)
        t_cell.font = font_ce_title
        t_cell.alignment = al_ce_mid_center
        ws.row_dimensions[start_row].height = 36

        hdr_row = start_row + 1
        for col_i, h_text in enumerate(headers, start=1):
            c = ws.cell(hdr_row, col_i, h_text)
            c.font = font_ce_th
            c.fill = rose_header_fill
            c.border = black_border
            c.alignment = al_ce_mid_center
        ws.row_dimensions[hdr_row].height = 34

        for idx, rec in enumerate(rows_data):
            r = hdr_row + 1 + idx
            for col_i, val in enumerate(rec, start=1):
                c = ws.cell(r, col_i, val)
                c.fill = FILL_WHITE
                c.border = black_border
                if col_i in (1, 2):
                    c.font = font_ce_name
                    c.alignment = al_ce_top_center
                elif col_i in (3, 4):
                    c.font = font_ce_body
                    c.alignment = al_ce_top_left
                else:
                    c.font = font_ce_body
                    c.alignment = al_ce_top_center
                    c.number_format = "0%"
            ws.row_dimensions[r].height = 76

        tot_r = hdr_row + 1 + len(rows_data)
        for col_i, val in enumerate(total_row_data, start=1):
            c = ws.cell(tot_r, col_i, val)
            c.fill = rose_total_fill
            c.border = black_border
            c.font = font_ce_name
            if col_i in (1, 2, 5):
                c.alignment = al_ce_mid_center
            else:
                c.alignment = al_ce_mid_left
            if col_i == 5:
                c.number_format = "0%"
        ws.row_dimensions[tot_r].height = 46

    # Khối 1 (A1:E8): Bảng Contribution Evaluation chuẩn mẫu (Tiếng Anh)
    ce_en_rows = [
        ("Khang", "ID1",
         "Wrote Chapter 1 in report (data collection, Table 1.2 & 7 EDA charts), created project Gantt chart (Oct 20 – Nov 15, 2026), assisted in frontend UI review, report editing, and helped with demo.",
         "Positive, cooperative, and always willing to support the team in data analysis, report editing, and demonstration tasks.",
         0.20),
        ("Hiếu", "ID2",
         "Wrote Chapter 4 in report (28 engineered features, ADF/KPSS stationarity tests, trained 5 ML/TS models: ARIMA, SARIMA, Holt-Winters, Random Forest, XGBoost), and built FastAPI backend.",
         "Active, supportive, and showed strong responsibility in machine learning modeling, backend API development, and report writing.",
         0.20),
        ("Minh", "ID3",
         "Wrote Chapter 2 in report (theoretical methodology, accuracy metrics & inventory theory), implemented Safety Stock (SS=311), Reorder Point (ROP=777), ABC-XYZ matrix, and What-If logic.",
         "Responsible, supportive, and completed assigned work well while actively contributing to inventory optimization and slide presentation.",
         0.20),
        ("Ân", "ID4",
         "Wrote Chapter 3 (Literature Review) and Chapter 6 (Conclusion, significance, future directions & recommendations), standardized 14 IEEE references, and contributed to slide preparation.",
         "Worked responsibly, maintained good teamwork, and completed assigned tasks with care and consistency.",
         0.20),
        ("Lợi", "ID5",
         "Developed Chapter 5 Web Studio frontend (6 interactive screens with offline ApexCharts), created and drew 5 flowchart diagrams with embedded icons (.drawio & 300 DPI PNG) for Chapters 1, 4, and 5.",
         "Hardworking, collaborative, and contributed effectively to frontend UI development, flowchart/BPMN system diagrams, and integration tasks.",
         0.20),
    ]
    assert abs(sum(r[4] for r in ce_en_rows) - 1.0) < 1e-9

    write_contribution_block(
        start_row=1,
        title_text="Contribution Evaluation",
        headers=["Full Name", "Student ID", "Assigned task", "Attitude", "Contribution"],
        rows_data=ce_en_rows,
        total_row_data=(
            "Total",
            "5 Members",
            "Completed all 6 report chapters, 5 trained models (.pkl), 5 BPMN flowcharts (.drawio/.png), and Web Studio app.",
            "100% team consensus and equal commitment across all project deliverables.",
            "=SUM(E3:E7)",
        ),
    )

    ws.row_dimensions[9].height = 16

    # Khối 2 (A10:E17): Bản Tiếng Việt cùng cấu trúc & màu sắc chuẩn mẫu
    ce_vi_rows = [
        ("Khang", "ID1",
         "Viết Chương 1 (thu thập 73.100 dòng dữ liệu, Bảng 1.1–1.3 & 7 biểu đồ EDA), lập biểu đồ Gantt tiến độ (20/10–15/11/2026), hỗ trợ rà soát giao diện Web UI, biên tập báo cáo Word và trình bày demo.",
         "Tích cực, hợp tác tốt, luôn sẵn sàng hỗ trợ các thành viên trong khâu xử lý dữ liệu, hoàn thiện báo cáo và thuyết trình.",
         0.20),
        ("Hiếu", "ID2",
         "Viết Chương 4 (kiểm định ADF/KPSS, 28 biến đặc trưng, huấn luyện 5 mô hình ARIMA, SARIMA, Holt-Winters, Random Forest, XGBoost) và xây dựng hệ thống Backend FastAPI.",
         "Chủ động, nhiệt tình hỗ trợ đồng đội và thể hiện tinh thần trách nhiệm cao trong lập trình học máy, API và viết báo cáo.",
         0.20),
        ("Minh", "ID3",
         "Viết Chương 2 (cơ sở lý thuyết, chỉ số đánh giá & tối ưu tồn kho Mục 2.2–2.4), cài đặt công thức Tồn kho an toàn (SS=311), Điểm đặt hàng lại (ROP=777), ma trận ABC–XYZ và mô phỏng What-If.",
         "Có trách nhiệm, phối hợp nhịp nhàng, hoàn thành tốt phần việc được giao và đóng góp tích cực vào tối ưu tồn kho & slide.",
         0.20),
        ("Ân", "ID4",
         "Viết Chương 3 (Tổng quan nghiên cứu khoa học) và Chương 6 (Kết luận, Ý nghĩa, Hướng phát triển & Kiến nghị), chuẩn hóa 14 tài liệu tham khảo và hỗ trợ chuẩn bị slide.",
         "Làm việc nghiêm túc, giữ tinh thần đồng đội tốt, hoàn thành các hạng mục được phân công cẩn thận và nhất quán.",
         0.20),
        ("Lợi", "ID5",
         "Phát triển giao diện Web Studio Chương 5 (6 màn hình Hình 5.1–5.12 với ApexCharts offline), thiết kế bộ sơ đồ Flowchart/BPMN kèm icon (.drawio & PNG 300 DPI) cho Chương 1, 4 và 5.",
         "Chăm chỉ, phối hợp chặt chẽ, đóng góp hiệu quả cho giao diện Frontend Web, sơ đồ hệ thống Flowchart/BPMN và tích hợp toàn hệ thống.",
         0.20),
    ]
    assert abs(sum(r[4] for r in ce_vi_rows) - 1.0) < 1e-9

    write_contribution_block(
        start_row=10,
        title_text="Contribution Evaluation — Bảng Đánh Giá Đóng Góp Thành Viên (Bản Tiếng Việt)",
        headers=[
            "Họ và tên (Full Name)",
            "MSSV (Student ID)",
            "Nhiệm vụ được giao (Assigned task)",
            "Thái độ làm việc (Attitude)",
            "Đóng góp (Contribution)",
        ],
        rows_data=ce_vi_rows,
        total_row_data=(
            "Tổng cộng (Total)",
            "5 Thành viên",
            "Hoàn thiện đủ 6 chương báo cáo, 5 mô hình (.pkl), 5 sơ đồ BPMN (.drawio/.png) và ứng dụng Web Studio.",
            "Đồng thuận 100% và cam kết đóng góp đồng đều giữa các thành viên trong suốt đồ án.",
            "=SUM(E12:E16)",
        ),
    )

    ws.row_dimensions[18].height = 16

    # Khối 3 (A19:I28): Danh mục Chức năng Web Studio (Chương 5)
    mod_r = 19
    style_section_title(ws, mod_r, 1, 9, "DANH MỤC CHỨC NĂNG HỆ THỐNG WEB STUDIO (RETAIL ML HUB — CHƯƠNG 5)")
    write_table_header(ws, mod_r + 1, [
        "STT", "Module / Nhóm tính năng", "Trang Web (Route)", "API Endpoint / Dịch vụ",
        "Đầu vào (Người dùng)", "Xử lý hệ thống (Backend / ML)", "Đầu ra & Giá trị ra quyết định",
        "Phụ trách", "Trạng thái"
    ])

    web_modules = [
        (1, "Hệ thống Bộ lọc đa chiều toàn cục & theo trang", "Tất cả 6 trang (#/overview → #/whatif)", "GET /api/options",
         "Cửa hàng (Toàn chuỗi / S001–S005), Sản phẩm (Tất cả / P0001–P0020), Chu kỳ (Ngày/Tuần/Tháng), Ô Ma trận ABC–XYZ.",
         "Truy vấn RAM Cache tức thì theo Store × SKU × Timeframe và đồng bộ trạng thái giữa các tab.",
         "Cập nhật toàn bộ KPI, biểu đồ và bảng dữ liệu trong thời gian thực dưới 1 giây.", "Lợi, Hiếu", "Đã hoàn thiện"),
        (2, "Hệ thống Thẻ KPI Tổng quan & Chỉ số thực nghiệm", "#/overview, #/forecast, #/inventory, #/whatif", "GET /api/kpi",
         "Tự động khi mở Web App hoặc khi người dùng đổi cửa hàng/sản phẩm/kịch bản.",
         "Tổng hợp 73.100 dòng (9.975.582 đv bán; $494,971,375 doanh thu; 274,5 tồn kho TB) & các delta.",
         "Hiển thị trực quan các thẻ KPI tổng quan (5 thẻ), Forecast (4 thẻ), ROP (4 thẻ) & What-If (4 thẻ Delta).", "Lợi, Khang", "Đã hoàn thiện"),
        (3, "Dự báo Nhu cầu SKU & Khoảng tin cậy 95%", "#/forecast", "POST /api/forecast",
         "Chọn Store ID, Product ID, Mô hình (RF/XGB/SARIMA/ARIMA/ES) & Kỳ hạn (7/14/30 ngày).",
         "Nạp .pkl (nếu S001–P0001) hoặc huấn luyện trực tuyến; dự báo đệ quy đa bước kèm cận 95%.",
         "Biểu đồ 60 ngày thực tế + đường dự báo đứt nét, 4 thẻ lỗi & bảng chi tiết theo ngày (Xuất CSV).", "Hiếu, Lợi", "Đã hoàn thiện"),
        (4, "So sánh 5 Mô hình & Feature Importance", "#/models", "GET /api/model-comparison, /api/feature-importance",
         "Chọn Store ID & Product ID bất kỳ cần đánh giá chéo mô hình.",
         "Huấn luyện/đánh giá cùng lúc 5 thuật toán trên tập Test 20%, xếp hạng RMSE & trích Top 10 đặc trưng.",
         "Bảng xếp hạng #1–#5, biểu đồ cột RMSE/MAE, biểu đồ bật/tắt 5 đường dự báo & thanh Top 10 Features.", "Hiếu", "Đã hoàn thiện"),
        (5, "Cảnh báo Tồn kho An toàn (SS) & Điểm đặt hàng lại (ROP)", "#/inventory", "GET /api/inventory-status, /api/inventory-history",
         "Chọn Cửa hàng, Sản phẩm, Lead Time L (ngày), Service Level (90/95/99%) & Nguồn cầu (ML / TB 90 ngày).",
         "Tính SS = ceil(Z×σ_D×√L) & ROP = ceil(D_hat×L + SS), phân loại 3 mức: Nguy cấp / Cảnh báo / An toàn.",
         "Đếm trạng thái 20 SKU (S001: 20/20 Nguy cấp), vẽ biểu đồ tồn kho 90 ngày kèm ngưỡng SS & ROP.", "Minh, Hiếu", "Đã hoàn thiện"),
        (6, "Chức năng Tự động sinh & Xuất Đơn đặt hàng (CSV Export)", "#/inventory, #/forecast, #/abcxyz", "Client CSV Exporter + InventoryOptimizer",
         "Nhấn nút \"Xuất danh sách đặt hàng (CSV)\" sau khi kiểm tra ngưỡng ROP.",
         "Tự động tính số lượng cần nhập bù lên mức mục tiêu 1,5×ROP: Q = max(0, ceil(1.5×ROP - I_curr)).",
         "Sinh ngay file CSV danh sách SKU cần đặt gấp kèm số lượng Q để chuyển cho bộ phận thu mua.", "Minh, Lợi", "Đã hoàn thiện"),
        (7, "Ma trận Phân loại Tồn kho Chiến lược ABC – XYZ", "#/abcxyz", "GET /api/abc-xyz",
         "Chọn Cửa hàng (S001–S005) và nhấp tương tác trực tiếp vào từng ô trong ma trận 9 nhóm.",
         "Phân hạng Doanh thu tích lũy ABC (80%/15%/5%) & Hệ số biến thiên CV=σ_D/Mean cho XYZ (0,40 / 0,70).",
         "Ma trận 3×3 (tại S001: AZ 15, BZ 3, CZ 2), hộp khuyến nghị chiến lược động & biểu đồ phân tán.", "Minh, Lợi", "Đã hoàn thiện"),
        (8, "Mô phỏng Kịch bản Giả định Chuỗi cung ứng (What-If)", "#/whatif", "GET /api/inventory-status (Dual Policy)",
         "Kéo thanh trượt Thời gian giao hàng mới (L' = 1–14 ngày), Service Level mới (SL') & SKU trọng tâm.",
         "Tính song song 2 chính sách Mặc định vs Kịch bản mới cho từng SKU và toàn bộ 20 SKU của cửa hàng.",
         "4 thẻ Delta (ΔSS, ΔROP, ΔQ SKU, Tổng Q toàn cửa hàng), biểu đồ cột đối chiếu & bảng chênh lệch.", "Minh, Hiếu", "Đã hoàn thiện"),
    ]

    for idx, mod in enumerate(web_modules):
        r = mod_r + 2 + idx
        for c_i, v in enumerate(mod, start=1):
            ws.cell(r, c_i, v)
        apply_row_style(
            ws, r, 1, 9,
            alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_LEFT, AL_LEFT, AL_LEFT, AL_LEFT, AL_CENTER, AL_CENTER],
            fill=FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE,
            bold_cols=[0, 1],
            height=42,
        )
        ws.cell(r, 9).fill = FILL_BEST
        ws.cell(r, 9).font = FONT_BEST

    ws.freeze_panes = "A3"
    widths_5 = {1: 20, 2: 26, 3: 52, 4: 46, 5: 22, 6: 36, 7: 36, 8: 14, 9: 16}
    set_col_widths(ws, widths_5)
    auto_fit_rows(ws, 2, 8, widths_5, start_col=1, end_col=5, char_factor=0.72, line_height=17.0, padding=14, min_height=28)
    auto_fit_rows(ws, 11, 17, widths_5, start_col=1, end_col=5, char_factor=0.72, line_height=17.0, padding=14, min_height=28)
    auto_fit_rows(ws, mod_r + 1, mod_r + 1 + len(web_modules), widths_5, start_col=1, end_col=9)


# ---------------------------------------------------------------------------
# SHEET 6: 6_Bang_Chi_So_Mo_Hinh (CHỈ SỐ 5 MÔ HÌNH + ADF/KPSS + FEATURE IMPORTANCE + CHART)
# ---------------------------------------------------------------------------
def build_sheet_models(wb):
    print("[6/6] Đang tạo Sheet 6_Bang_Chi_So_Mo_Hinh (Chỉ số 5 Mô hình, ADF/KPSS, Feature Importance & Biểu đồ)...")
    ws = wb.create_sheet("6_Bang_Chi_So_Mo_Hinh")
    ws.views.sheetView[0].showGridLines = True

    with open(os.path.join(SAVED_MODELS_DIR, "model_metadata.json"), encoding="utf-8") as f:
        meta = json.load(f)

    style_banner(
        ws, 1, 9,
        "BẢNG CHỈ SỐ ĐÁNH GIÁ HIỆU NĂNG 5 MÔ HÌNH DỰ BÁO & KIỂM ĐỊNH CHUỖI THỜI GIAN",
        f"Đánh giá định lượng SKU {meta['store_id']} – {meta['product_id']} (Test 20%: 141 ngày ML [2023-08-14 → 2024-01-01] / 147 ngày gốc) | Tối ưu nhất: {meta['best_model']} (RMSE = {meta['metrics']['RMSE']:.2f})"
    )

    # Phần 1: Bảng 4.5 / 4.6 Word — So sánh 5 mô hình dự báo
    style_section_title(ws, 4, 1, 9, "1. BẢNG SO SÁNH ĐỊNH LƯỢNG 5 MÔ HÌNH DỰ BÁO ĐÃ HUẤN LUYỆN (ĐỐI CHIẾU BẢNG 4.5–4.6 TRONG MLBA.DOCX)")
    write_table_header(ws, 5, [
        "Hạng", "Mô hình (Model)", "Nhóm phương pháp", "Cấu hình siêu tham số chính",
        "RMSE", "MAE", "MAPE (%)", "% Cải thiện RMSE so với ARIMA", "Kết luận & Nhận xét nổi bật"
    ])

    model_extra = {
        "Random Forest": ("Học máy (Ensemble Tree)", "n_estimators=100, max_depth=12, random_state=42",
                          "TỐI ƯU NHẤT (#1) — RMSE thấp nhất 92,42; vượt trội hơn rõ rệt nhóm thống kê cổ điển nhờ 28 biến đặc trưng."),
        "XGBoost": ("Học máy (Gradient Boosting)", "n_estimators=150, lr=0.05, max_depth=6, subsample=0.8",
                    "Hạng #2 — Bám sát Random Forest (RMSE 95,73), phản ánh tốt các nhịp tăng giảm ngắn hạn."),
        "SARIMA": ("Thống kê Chuỗi thời gian", "order=(1, 1, 1), seasonal_order=(1, 0, 1, 7)",
                   "Tốt nhất trong nhóm Thống kê cổ điển nhờ nắm bắt chu kỳ lặp 7 ngày trong tuần."),
        "ARIMA": ("Thống kê Chuỗi thời gian", "order=(2, 1, 1)",
                  "Mô hình tham chiếu cổ điển đơn biến (Benchmark cơ sở để tính % cải thiện)."),
        "Exponential Smoothing": ("Thống kê Làm mượt hàm mũ", "Holt-Winters (trend='add', seasonal='add', period=7)",
                                  "Thích ứng mượt nhưng khó bám các biến động đột biến theo ngày của bán lẻ."),
    }

    sorted_models = sorted(meta["all_results"], key=lambda x: x["RMSE"])
    arima_row_num = 6 + [m["Model"] for m in sorted_models].index("ARIMA")

    for idx, m in enumerate(sorted_models):
        r = 6 + idx
        name = m["Model"]
        grp, params, note = model_extra[name]
        ws.cell(r, 1, f"#{idx + 1}")
        ws.cell(r, 2, name)
        ws.cell(r, 3, grp)
        ws.cell(r, 4, params)
        ws.cell(r, 5, float(m["RMSE"]))
        ws.cell(r, 6, float(m["MAE"]))
        ws.cell(r, 7, float(m["MAPE (%)"]) / 100.0)
        ws.cell(r, 8, f"=($E${arima_row_num}-E{r})/$E${arima_row_num}")
        ws.cell(r, 9, note)

        is_best = idx == 0
        apply_row_style(
            ws, r, 1, 9,
            alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_LEFT, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_LEFT],
            num_formats={4: "#,##0.00", 5: "#,##0.00", 6: "0.00%", 7: "+0.00%;-0.00%;0.00%"},
            fill=FILL_BEST if is_best else (FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE),
            font=FONT_BEST if is_best else FONT_TD,
            bold_cols=[0, 1] if not is_best else [0, 1, 4, 5, 6, 7],
            height=30 if is_best else 24,
        )

    # Dòng Min & Trung bình cho Bảng 1
    min_r = 11
    avg_r = 12
    ws.cell(min_r, 1, "MIN")
    ws.cell(min_r, 2, "Giá trị tốt nhất (Thấp nhất)")
    ws.cell(min_r, 3, "Tối ưu toàn cục")
    ws.cell(min_r, 4, "=INDEX(B6:B10, MATCH(MIN(E6:E10), E6:E10, 0))")
    ws.cell(min_r, 5, "=MIN(E6:E10)")
    ws.cell(min_r, 6, "=MIN(F6:F10)")
    ws.cell(min_r, 7, "=MIN(G6:G10)")
    ws.cell(min_r, 8, "=MAX(H6:H10)")
    ws.cell(min_r, 9, "Random Forest đạt giá trị tối ưu đồng thời ở cả 3 chỉ số RMSE, MAE và MAPE.")
    apply_row_style(
        ws, min_r, 1, 9,
        alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_CENTER, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_LEFT],
        num_formats={4: "#,##0.00", 5: "#,##0.00", 6: "0.00%", 7: "+0.00%;-0.00%;0.00%"},
        fill=FILL_SUBHEADER,
        font=FONT_NAVY_BOLD,
        height=24,
    )

    ws.cell(avg_r, 1, "AVG")
    ws.cell(avg_r, 2, "Trung bình cả 5 mô hình")
    ws.cell(avg_r, 3, "Benchmark chung")
    ws.cell(avg_r, 4, "5 mô hình (3 TS + 2 ML)")
    ws.cell(avg_r, 5, "=AVERAGE(E6:E10)")
    ws.cell(avg_r, 6, "=AVERAGE(F6:F10)")
    ws.cell(avg_r, 7, "=AVERAGE(G6:G10)")
    ws.cell(avg_r, 8, "=AVERAGE(H6:H10)")
    ws.cell(avg_r, 9, "Nhóm Học máy kéo trung bình sai số giảm đáng kể so với chuỗi đơn biến.")
    apply_row_style(
        ws, avg_r, 1, 9,
        alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_CENTER, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_LEFT],
        num_formats={4: "#,##0.00", 5: "#,##0.00", 6: "0.00%", 7: "+0.00%;-0.00%;0.00%"},
        fill=FILL_TOTAL,
        font=FONT_TD_BOLD,
        height=24,
    )

    # Phần 2: Bảng Kiểm định Tính dừng Chuỗi thời gian (ADF & KPSS)
    stat_r = 14
    style_section_title(ws, stat_r, 1, 9, "2. KẾT QUẢ KIỂM ĐỊNH TÍNH DỪNG CHUỖI THỜI GIAN — ADF & KPSS (S001 – P0001)")
    st_df = pd.read_csv(os.path.join(REPORTS_DIR, "stationarity_tests_s001_p0001.csv"))
    write_table_header(ws, stat_r + 1, [
        "Phép kiểm định", "Statistic", "p-value", "Lags Used", "Observations",
        "Ngưỡng 1%", "Ngưỡng 5%", "Ngưỡng 10%", "Kết luận ở mức α = 0,05"
    ])
    for idx, rec in enumerate(st_df.itertuples(index=False, name=None)):
        r = stat_r + 2 + idx
        for c_i, val in enumerate(rec, start=1):
            ws.cell(r, c_i, val)
        apply_row_style(
            ws, r, 1, 9,
            alignments=[AL_LEFT, AL_RIGHT, AL_RIGHT, AL_CENTER, AL_CENTER, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_LEFT],
            num_formats={1: "0.0000", 2: "0.0000", 3: "0", 4: "#,##0", 5: "0.0000", 6: "0.0000", 7: "0.0000"},
            fill=FILL_BEST if idx == 0 else FILL_SUBHEADER,
            bold_cols=[0, 1, 8],
            height=26,
        )

    # Phần 3: Top 15 Biến Đặc trưng Quan trọng nhất (trích trực tiếp từ .pkl)
    fi_r = 19
    style_section_title(ws, fi_r, 1, 9, "3. XẾP HẠNG TOP 15 BIẾN ĐẶC TRƯNG QUAN TRỌNG NHẤT (TRÍCH XUẤT TỪ RANDOM FOREST & XGBOOST .PKL)")
    write_table_header(ws, fi_r + 1, [
        "Hạng", "Tên đặc trưng (Feature)", "Nhóm đặc trưng", "Random Forest (%)",
        "XGBoost (%)", "Trung bình 2 mô hình", "Chênh lệch (RF - XGB)", "Mức ảnh hưởng", "Diễn giải vai trò kinh doanh"
    ])

    feats = meta["features"]
    rf_model = joblib.load(os.path.join(SAVED_MODELS_DIR, "random_forest_model.pkl"))
    xgb_model = joblib.load(os.path.join(SAVED_MODELS_DIR, "xgboost_model.pkl"))
    rf_map = dict(zip(feats, [float(v) for v in rf_model.feature_importances_]))
    xgb_map = dict(zip(feats, [float(v) for v in xgb_model.feature_importances_]))

    t41_df = pd.read_csv(os.path.join(REPORTS_DIR, "table_4_1_danh_sach_28_bien_dac_trung.csv"))
    feat_meta = {
        row["Tên biến đặc trưng (Feature)"]: (row["Nhóm biến"], row["Vai trò đối với mô hình dự báo"])
        for _, row in t41_df.iterrows()
    }

    top_feats = sorted(feats, key=lambda k: rf_map[k], reverse=True)[:15]
    for idx, f_name in enumerate(top_feats):
        r = fi_r + 2 + idx
        grp, desc = feat_meta.get(f_name, ("Đặc trưng dẫn xuất", "Hỗ trợ dự báo nhu cầu SKU"))
        rf_val = rf_map[f_name]
        xgb_val = xgb_map[f_name]
        tier = "Rất cao (Chủ chốt)" if rf_val >= 0.04 else ("Trung bình" if rf_val >= 0.025 else "Bổ trợ")

        ws.cell(r, 1, f"#{idx + 1}")
        ws.cell(r, 2, f_name)
        ws.cell(r, 3, grp)
        ws.cell(r, 4, rf_val)
        ws.cell(r, 5, xgb_val)
        ws.cell(r, 6, f"=AVERAGE(D{r}:E{r})")
        ws.cell(r, 7, f"=D{r}-E{r}")
        ws.cell(r, 8, tier)
        ws.cell(r, 9, desc)
        apply_row_style(
            ws, r, 1, 9,
            alignments=[AL_CENTER, AL_LEFT, AL_LEFT, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_CENTER, AL_LEFT],
            num_formats={3: "0.00%", 4: "0.00%", 5: "0.00%", 6: "+0.00%;-0.00%;0.00%"},
            fill=FILL_BEST if idx == 0 else (FILL_ZEBRA if idx % 2 == 1 else FILL_WHITE),
            font=FONT_BEST if idx == 0 else FONT_TD,
            bold_cols=[0, 1] if idx > 0 else [0, 1, 3, 7],
            height=24,
        )

    fi_sum_r = fi_r + 2 + len(top_feats)
    ws.cell(fi_sum_r, 1, "TOP 15")
    ws.cell(fi_sum_r, 2, "Tổng tỷ trọng Top 15 đặc trưng")
    ws.cell(fi_sum_r, 3, "15 / 28 biến")
    ws.cell(fi_sum_r, 4, f"=SUM(D{fi_r + 2}:D{fi_sum_r - 1})")
    ws.cell(fi_sum_r, 5, f"=SUM(E{fi_r + 2}:E{fi_sum_r - 1})")
    ws.cell(fi_sum_r, 6, f"=AVERAGE(D{fi_sum_r}:E{fi_sum_r})")
    ws.cell(fi_sum_r, 7, f"=D{fi_sum_r}-E{fi_sum_r}")
    ws.cell(fi_sum_r, 8, "Tập trung cao")
    ws.cell(fi_sum_r, 9, "Top 15 đặc trưng chiếm phần lớn sức giải thích của cây quyết định.")
    apply_row_style(
        ws, fi_sum_r, 1, 9,
        alignments=[AL_CENTER, AL_LEFT, AL_CENTER, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_RIGHT, AL_CENTER, AL_LEFT],
        num_formats={3: "0.00%", 4: "0.00%", 5: "0.00%", 6: "+0.00%;-0.00%;0.00%"},
        fill=FILL_TOTAL,
        font=FONT_NAVY_BOLD,
        height=24,
    )

    # Phần 4: Biểu đồ Cột Excel原生 (Native BarChart) + Nhúng ảnh 300 DPI bên dưới bảng
    chart_sec_r = fi_sum_r + 2
    style_section_title(ws, chart_sec_r, 1, 9, "4. TRỰC QUAN HÓA SO SÁNH 5 MÔ HÌNH (BIỂU ĐỒ EXCEL TƯƠNG TÁC & ẢNH BÁO CÁO 300 DPI)")

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.overlap = -15
    chart.title = "So sánh sai số RMSE & MAE của 5 Mô hình Dự báo (Giá trị thấp hơn là tốt hơn)"
    chart.y_axis.title = "Giá trị sai số (Đơn vị)"
    chart.x_axis.title = "Mô hình Dự báo"
    chart.style = 10
    chart.width = 20
    chart.height = 11

    data_ref = Reference(ws, min_col=5, max_col=6, min_row=5, max_row=10)
    cats_ref = Reference(ws, min_col=2, min_row=6, max_row=10)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.legend.position = "t"
    ws.add_chart(chart, f"A{chart_sec_r + 2}")

    # Nhúng kèm ảnh thực nghiệm Hình 4.9 ở phía dưới biểu đồ Excel
    img_49 = os.path.join(REPORTS_DIR, "figures", "models", "chuong_4_hinh_4_9_so_sanh_hieu_nang_mo_hinh.png")
    img_410 = os.path.join(REPORTS_DIR, "figures", "models", "chuong_4_hinh_4_10_tam_quan_trong_dac_trung.png")
    h_img49 = add_embedded_image(ws, img_49, f"F{chart_sec_r + 2}", max_width_px=540)
    row_offset = max(23, int(h_img49 // 20) + 3)
    add_embedded_image(ws, img_410, f"A{chart_sec_r + 2 + row_offset}", max_width_px=640)

    ws.freeze_panes = "A6"
    widths_6 = {1: 16, 2: 26, 3: 26, 4: 38, 5: 13, 6: 16, 7: 16, 8: 22, 9: 56}
    set_col_widths(ws, widths_6)
    auto_fit_rows(ws, 5, fi_sum_r, widths_6, start_col=1, end_col=9)


def main():
    t0 = time.time()
    print("=" * 72)
    print("   TẠO FILE EXCEL TỔNG HỢP ĐỒ ÁN (6 SHEETS) — RETAIL ML HUB")
    print("=" * 72)

    wb = Workbook()
    build_sheet_dataset(wb)
    build_sheet_pccv(wb)
    build_sheet_gantt(wb)
    build_sheet_flowchart(wb)
    build_sheet_contribution(wb)
    build_sheet_models(wb)

    print(f"--> Đang lưu workbook ra file: {OUT_XLSX_REPORTS}...")
    wb.save(OUT_XLSX_REPORTS)
    shutil.copy2(OUT_XLSX_REPORTS, OUT_XLSX_PARENT)

    dt = time.time() - t0
    sz_mb = os.path.getsize(OUT_XLSX_REPORTS) / (1024 * 1024)
    print(f"Đã tạo thành công ({sz_mb:.2f} MB trong {dt:.1f}s):")
    print(f"  1) {OUT_XLSX_REPORTS}")
    print(f"  2) {OUT_XLSX_PARENT}")
    print("=" * 72)


if __name__ == "__main__":
    main()
