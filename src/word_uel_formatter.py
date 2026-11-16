r"""
word_uel_formatter.py
Chuẩn hóa định dạng báo cáo Word (MLBA.docx) theo chuẩn học thuật Khoa Hệ thống Thông tin — UEL
(đối chiếu trực tiếp từ D:/HK7/Phát triển web kinh doanh/Final_Project_Web.docx):
  - Đồng bộ Header/Footer (đánh số trang tự động từ trang Mục lục), Ảnh bìa UEL & word/styles.xml.
  - Tạo Trang bìa Section 1 (P0..P18 với <w:sectPr><w:titlePg/></w:sectPr>).
  - Tạo Mục lục (TOC \o "1-3"), Danh sách hình ảnh (36 mục TIeuDeanh), Danh sách bảng (13 mục Caption).
  - Bố cục phần mở đầu chuẩn: Thành viên nhóm, Phân công nhiệm vụ + Gantt PNG (20/10–15/11/2026), Từ viết tắt, Lời cảm ơn, Lời cam kết.
  - Chuẩn hóa toàn bộ Chương 1–6 & Tài liệu tham khảo (sửa 4.3.3, 4.4.1, thêm Bảng 6.1, căn giữa ảnh/toán, TableGrid).
"""

import copy
import io
import math
import os
import re
import zipfile
import xml.etree.ElementTree as ET
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from PIL import Image

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"

REF_DOCX_PATH = r"D:\HK7\Phát triển web kinh doanh\Final_Project_Web.docx"

TEAM_MEMBERS = [
    ("1", "Phùng Trịnh Trường Khang", "K244060721", "Nhóm trưởng", "23%", "100%"),
    ("2", "Huỳnh Mai Hiếu", "K244060716", "Thành viên", "22%", "100%"),
    ("3", "Minh", "—", "Thành viên", "20%", "100%"),
    ("4", "Ân", "—", "Thành viên", "17%", "100%"),
    ("5", "Cai Đặng Phúc Lợi", "K244060723", "Thành viên", "18%", "100%"),
]

TASK_ASSIGNMENTS = [
    (
        "1",
        "Phùng Trịnh Trường Khang\n(K244060721)",
        "Nhóm trưởng",
        "Chủ trì định hướng kiến trúc đề tài; trực tiếp thực nghiệm & đối chiếu 5 mô hình (Chương 4); kiểm soát đồng bộ số liệu Word – Excel – Web Studio; biên tập Slide & thuyết trình.",
        "Kiến trúc tổng thể, Bảng 4.5–4.6, 23% đóng góp (100% đạt)",
    ),
    (
        "2",
        "Huỳnh Mai Hiếu\n(K244060716)",
        "Thành viên",
        "Xây dựng pipeline Feature Engineering (28 biến), kiểm định ADF/KPSS, tính toán tồn kho SS/ROP/ABC–XYZ; phát triển Backend FastAPI & Web Studio 6 màn hình (Chương 5).",
        "Code & Web Studio, Bảng 4.1–4.2, Chương 5, 22% đóng góp (100% đạt)",
    ),
    (
        "3",
        "Minh",
        "Thành viên",
        "Viết Chương 2 (Cơ sở lý thuyết & thuật toán) phần toán học; tổng hợp phần Tối ưu tồn kho & Chương 6 (Kết luận, Bảng 6.1); hỗ trợ kiểm thử toàn hệ thống.",
        "Chương 2 & Chương 6, Bảng 6.1, Kịch bản kiểm thử, 20% đóng góp (100% đạt)",
    ),
    (
        "4",
        "Ân",
        "Thành viên",
        "Thực hiện Chương 3 (Lược khảo 11 công trình quốc tế & trong nước); chuẩn hóa trích dẫn IEEE & Danh mục tài liệu tham khảo [1]–[14]; hỗ trợ giao diện Web & Slide.",
        "Chương 3, Bảng 3.1, Tài liệu tham khảo IEEE, 17% đóng góp (100% đạt)",
    ),
    (
        "5",
        "Cai Đặng Phúc Lợi\n(K244060723)",
        "Thành viên",
        "Phụ trách Chương 1 (Tổng quan, 7 biểu đồ EDA, Bảng 1.1–1.3); thiết kế bộ sơ đồ BPMN Draw.io (Sơ đồ 1.1, 4.1, 5.1–5.3 & 6 lưu đồ màn hình); tổng hợp Workbook Excel.",
        "Chương 1, 11 sơ đồ BPMN & Excel 6 sheet, 18% đóng góp (100% đạt)",
    ),
]

ACRONYMS = [
    ("1", "AI", "Artificial Intelligence — Trí tuệ nhân tạo"),
    ("2", "ML", "Machine Learning — Học máy"),
    ("3", "EDA", "Exploratory Data Analysis — Phân tích khám phá dữ liệu"),
    ("4", "SKU", "Stock Keeping Unit — Đơn vị lưu kho chuẩn hóa (Store × Product)"),
    ("5", "ARIMA", "AutoRegressive Integrated Moving Average — Mô hình tự hồi quy tích hợp trung bình trượt"),
    ("6", "SARIMA", "Seasonal ARIMA — Mô hình ARIMA kết hợp thành phần mùa vụ"),
    ("7", "HW", "Holt-Winters Exponential Smoothing — San bằng mũ ba thành phần"),
    ("8", "RF", "Random Forest Regressor — Mô hình rừng ngẫu nhiên hồi quy"),
    ("9", "XGBoost", "eXtreme Gradient Boosting — Thuật toán tăng cường độ dốc cực hạn"),
    ("10", "RMSE", "Root Mean Squared Error — Căn bậc hai trung bình bình phương sai số"),
    ("11", "MAE", "Mean Absolute Error — Sai số tuyệt đối trung bình"),
    ("12", "MAPE", "Mean Absolute Percentage Error — Phần trăm sai số tuyệt đối trung bình"),
    ("13", "ADF", "Augmented Dickey-Fuller Test — Kiểm định nghiệm đơn vị / tính dừng"),
    ("14", "KPSS", "Kwiatkowski–Phillips–Schmidt–Shin Test — Kiểm định giả thuyết chuỗi dừng"),
    ("15", "ACF", "Autocorrelation Function — Hàm tự tương quan"),
    ("16", "PACF", "Partial Autocorrelation Function — Hàm tự tương quan riêng phần"),
    ("17", "SS", "Safety Stock — Tồn kho an toàn"),
    ("18", "ROP", "Reorder Point — Điểm đặt hàng lại"),
    ("19", "ABC–XYZ", "Phân loại tồn kho hai chiều theo giá trị doanh thu (ABC) và độ ổn định nhu cầu (XYZ)"),
    ("20", "CV", "Coefficient of Variation — Hệ số biến thiên nhu cầu"),
    ("21", "API", "Application Programming Interface — Giao diện lập trình ứng dụng"),
    ("22", "UI/UX", "User Interface / User Experience — Giao diện và trải nghiệm người dùng"),
    ("23", "BPMN", "Business Process Model and Notation — Ký hiệu mô hình hóa quy trình nghiệp vụ"),
    ("24", "UEL", "University of Economics and Law — Trường Đại học Kinh tế – Luật, ĐHQG-HCM"),
]


def _w(local):
    return f"{{{W_NS}}}{local}"


def _r(local):
    return f"{{{R_NS}}}{local}"


def _wp(local):
    return f"{{{WP_NS}}}{local}"


def _a(local):
    return f"{{{A_NS}}}{local}"


def _pic(local):
    return f"{{{PIC_NS}}}{local}"


def _ptxt(el):
    return "".join(t.text or "" for t in el.findall(f".//{_w('t')}"))


def _ttxt(el):
    return "".join(t.text or "" for t in el.findall(f".//{_w('t')}"))


def _make_rpr(font_name="Times New Roman", sz=26, bold=False, italic=False, color="000000",
              underline=None, rstyle=None, no_proof=False, web_hidden=False):
    rpr = ET.Element(_w("rPr"))
    if rstyle:
        ET.SubElement(rpr, _w("rStyle")).set(_w("val"), rstyle)
    rf = ET.SubElement(rpr, _w("rFonts"))
    for k in ("ascii", "eastAsia", "hAnsi", "cs"):
        rf.set(_w(k), font_name)
    if bold:
        ET.SubElement(rpr, _w("b"))
        ET.SubElement(rpr, _w("bCs"))
    if italic:
        ET.SubElement(rpr, _w("i"))
        ET.SubElement(rpr, _w("iCs"))
    if no_proof:
        ET.SubElement(rpr, _w("noProof"))
    if web_hidden:
        ET.SubElement(rpr, _w("webHidden"))
    if color:
        ET.SubElement(rpr, _w("color")).set(_w("val"), color)
    if sz is not None:
        ET.SubElement(rpr, _w("sz")).set(_w("val"), str(sz))
        ET.SubElement(rpr, _w("szCs")).set(_w("val"), str(sz))
    if underline is not None:
        ET.SubElement(rpr, _w("u")).set(_w("val"), underline)
    return rpr


def _make_p(runs_spec, align="both", style=None, before=120, after=120, line=312,
            first_line=None, left=None, hanging=None, keep_next=False, keep_lines=False,
            page_break_before=False, tabs=None, rpr_sz=26, rpr_bold=False, rpr_color="000000"):
    p = ET.Element(_w("p"))
    ppr = ET.SubElement(p, _w("pPr"))
    if style:
        ET.SubElement(ppr, _w("pStyle")).set(_w("val"), style)
    if keep_next:
        ET.SubElement(ppr, _w("keepNext"))
    if keep_lines:
        ET.SubElement(ppr, _w("keepLines"))
    if page_break_before:
        ET.SubElement(ppr, _w("pageBreakBefore"))
    if tabs:
        tel = ET.SubElement(ppr, _w("tabs"))
        for tval, pos, leader in tabs:
            t = ET.SubElement(tel, _w("tab"))
            t.set(_w("val"), tval)
            if leader:
                t.set(_w("leader"), leader)
            t.set(_w("pos"), str(pos))
    sp = ET.SubElement(ppr, _w("spacing"))
    if before is not None:
        sp.set(_w("before"), str(before))
    if after is not None:
        sp.set(_w("after"), str(after))
    if line is not None:
        sp.set(_w("line"), str(line))
        sp.set(_w("lineRule"), "auto")
    if first_line or left or hanging:
        ind = ET.SubElement(ppr, _w("ind"))
        if left:
            ind.set(_w("left"), str(left))
        if first_line:
            ind.set(_w("firstLine"), str(first_line))
        if hanging:
            ind.set(_w("hanging"), str(hanging))
    if align:
        ET.SubElement(ppr, _w("jc")).set(_w("val"), align)
    ppr.append(_make_rpr(sz=rpr_sz, bold=rpr_bold, color=rpr_color))

    if isinstance(runs_spec, str):
        runs_spec = [(runs_spec, rpr_bold, False, rpr_color, rpr_sz)]
    for item in runs_spec:
        if item is None:
            continue
        text, b, it, col, sz = (item + (rpr_bold, False, rpr_color, rpr_sz))[:5] if len(item) < 5 else item[:5]
        r = ET.SubElement(p, _w("r"))
        r.append(_make_rpr(sz=sz, bold=b, italic=it, color=col))
        if text == "\t":
            ET.SubElement(r, _w("tab"))
        else:
            t = ET.SubElement(r, _w("t"))
            if text.startswith(" ") or text.endswith(" ") or "\n" in text:
                t.set(f"{{{XML_NS}}}space", "preserve")
            t.text = text
    return p


def _make_marker(tag_text):
    p = ET.Element(_w("p"))
    r = ET.SubElement(p, _w("r"))
    t = ET.SubElement(r, _w("t"))
    t.text = tag_text
    return p


def _make_inline_image_p(rid, w_px, h_px, docpr_id, name, max_w_emu=5730000, max_h_emu=4200000, keep_next=True):
    scale = min(max_w_emu / max(w_px * 9525, 1), max_h_emu / max(h_px * 9525, 1), 1.0)
    cx = int(w_px * 9525 * scale)
    cy = int(h_px * 9525 * scale)

    p = ET.Element(_w("p"))
    ppr = ET.SubElement(p, _w("pPr"))
    if keep_next:
        ET.SubElement(ppr, _w("keepNext"))
    sp = ET.SubElement(ppr, _w("spacing"))
    sp.set(_w("before"), "120")
    sp.set(_w("after"), "60")
    sp.set(_w("line"), "240")
    sp.set(_w("lineRule"), "auto")
    ET.SubElement(ppr, _w("jc")).set(_w("val"), "center")

    r = ET.SubElement(p, _w("r"))
    rpr = ET.SubElement(r, _w("rPr"))
    ET.SubElement(rpr, _w("noProof"))
    dr = ET.SubElement(r, _w("drawing"))
    inl = ET.SubElement(dr, _wp("inline"))
    for k, v in (("distT", "0"), ("distB", "0"), ("distL", "0"), ("distR", "0")):
        inl.set(k, v)
    ext = ET.SubElement(inl, _wp("extent"))
    ext.set("cx", str(cx))
    ext.set("cy", str(cy))
    eff = ET.SubElement(inl, _wp("effectExtent"))
    for k in ("l", "t", "r", "b"):
        eff.set(k, "0")
    dp = ET.SubElement(inl, _wp("docPr"))
    dp.set("id", str(docpr_id))
    dp.set("name", name)
    cnv = ET.SubElement(inl, _wp("cNvGraphicFramePr"))
    gfl = ET.SubElement(cnv, _a("graphicFrameLocks"))
    gfl.set("noChangeAspect", "1")
    gr = ET.SubElement(inl, _a("graphic"))
    gd = ET.SubElement(gr, _a("graphicData"))
    gd.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")
    pic = ET.SubElement(gd, _pic("pic"))
    nv = ET.SubElement(pic, _pic("nvPicPr"))
    cnvp = ET.SubElement(nv, _pic("cNvPr"))
    cnvp.set("id", "0")
    cnvp.set("name", name)
    ET.SubElement(nv, _pic("cNvPicPr"))
    bf = ET.SubElement(pic, _pic("blipFill"))
    bl = ET.SubElement(bf, _a("blip"))
    bl.set(_r("embed"), rid)
    st = ET.SubElement(bf, _a("stretch"))
    ET.SubElement(st, _a("fillRect"))
    spp = ET.SubElement(pic, _pic("spPr"))
    xf = ET.SubElement(spp, _a("xfrm"))
    off = ET.SubElement(xf, _a("off"))
    off.set("x", "0")
    off.set("y", "0")
    xext = ET.SubElement(xf, _a("ext"))
    xext.set("cx", str(cx))
    xext.set("cy", str(cy))
    pg = ET.SubElement(spp, _a("prstGeom"))
    pg.set("prst", "rect")
    ET.SubElement(pg, _a("avLst"))
    return p


def _make_grid_table(headers, rows, col_widths_dxa, col_aligns=None, font_sz=23):
    tbl = ET.Element(_w("tbl"))
    tbl_pr = ET.SubElement(tbl, _w("tblPr"))
    ET.SubElement(tbl_pr, _w("tblStyle")).set(_w("val"), "TableGrid")
    tw = ET.SubElement(tbl_pr, _w("tblW"))
    tw.set(_w("w"), str(sum(col_widths_dxa)))
    tw.set(_w("type"), "dxa")
    ET.SubElement(tbl_pr, _w("jc")).set(_w("val"), "center")
    tbd = ET.SubElement(tbl_pr, _w("tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = ET.SubElement(tbd, _w(side))
        b.set(_w("val"), "single")
        b.set(_w("sz"), "4")
        b.set(_w("space"), "0")
        b.set(_w("color"), "BFBFBF")
    ET.SubElement(tbl_pr, _w("tblLayout")).set(_w("type"), "fixed")
    cm = ET.SubElement(tbl_pr, _w("tblCellMar"))
    for k, val in (("top", "80"), ("left", "110"), ("bottom", "80"), ("right", "110")):
        m = ET.SubElement(cm, _w(k))
        m.set(_w("w"), val)
        m.set(_w("type"), "dxa")
    lk = ET.SubElement(tbl_pr, _w("tblLook"))
    lk.set(_w("val"), "04A0")

    tgrid = ET.SubElement(tbl, _w("tblGrid"))
    for cw in col_widths_dxa:
        ET.SubElement(tgrid, _w("gridCol")).set(_w("w"), str(cw))

    all_rows = [headers] + rows
    for r_idx, row in enumerate(all_rows):
        is_hdr = r_idx == 0
        tr = ET.SubElement(tbl, _w("tr"))
        tr_pr = ET.SubElement(tr, _w("trPr"))
        ET.SubElement(tr_pr, _w("cantSplit"))
        if is_hdr:
            ET.SubElement(tr_pr, _w("tblHeader"))
        for c_idx, cell_text in enumerate(row):
            tc = ET.SubElement(tr, _w("tc"))
            tc_pr = ET.SubElement(tc, _w("tcPr"))
            tcw = ET.SubElement(tc_pr, _w("tcW"))
            tcw.set(_w("w"), str(col_widths_dxa[c_idx]))
            tcw.set(_w("type"), "dxa")
            shd = ET.SubElement(tc_pr, _w("shd"))
            shd.set(_w("val"), "clear")
            shd.set(_w("color"), "auto")
            shd.set(_w("fill"), "D9E2F3" if is_hdr else ("F8FAFC" if r_idx % 2 == 0 else "FFFFFF"))
            ET.SubElement(tc_pr, _w("vAlign")).set(_w("val"), "center")

            al = "center" if is_hdr else (col_aligns[c_idx] if col_aligns else "left")
            lines = str(cell_text).split("\n")
            for li, line_str in enumerate(lines):
                p = _make_p(
                    line_str,
                    align=al,
                    before=30 if li == 0 else 10,
                    after=30 if li == len(lines) - 1 else 10,
                    line=264,
                    rpr_sz=font_sz,
                    rpr_bold=is_hdr or (c_idx == 0 and len(line_str) <= 5),
                    rpr_color="0F172A" if is_hdr else "000000",
                )
                tc.append(p)
    return tbl


# ---------------------------------------------------------------------------
# 1. TẠO HÌNH BIỂU ĐỒ GANTT CHUẨN TIẾN ĐỘ THỰC TẾ (20/10/2026 – 15/11/2026)
# ---------------------------------------------------------------------------
def generate_gantt_png(out_path):
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    milestones = [
        ("M1 (GĐ1): Khởi động, chốt đề tài & EDA 73.100 dòng", 0, 2, "Khang, Lợi", "20–21/10", "#1E3A8A"),
        ("M2 (GĐ1): Vẽ Sơ đồ 1.1 EDA & Hoàn thiện Chương 1", 2, 4, "Lợi, Khang", "22–25/10", "#1E3A8A"),
        ("M3 (GĐ2): Cơ sở lý thuyết toán học ML & Tồn kho (Ch. 2)", 6, 3, "Minh, Ân", "26–28/10", "#0284C7"),
        ("M4 (GĐ2): Lược khảo 11 công trình & Chốt Chương 3", 9, 4, "Ân, Minh", "29/10–01/11", "#0284C7"),
        ("M5 (GĐ3): Kiểm định ADF/KPSS & 28 đặc trưng (lag_30)", 13, 3, "Hiếu", "02–04/11", "#0F766E"),
        ("M6 (GĐ3): Train & so sánh 5 mô hình (RF RMSE 92,42)", 16, 4, "Khang, Hiếu", "05–08/11", "#0F766E"),
        ("M7 (GĐ4): Mô phỏng SS=311, ROP=777 & Ma trận ABC–XYZ", 20, 3, "Hiếu, Minh", "09–11/11", "#D97706"),
        ("M8 (GĐ4): FastAPI & 6 màn hình Web Studio + 11 BPMN", 23, 2, "Hiếu, Lợi, Khang", "12–13/11", "#7C3AED"),
        ("M9 (GĐ5): Đồng bộ Chương 5, Chương 6 & Workbook Excel", 25, 1, "Minh, Lợi", "14/11", "#BE123C"),
        ("M10 (GĐ6): Kiểm thử tổng thể, chuẩn hóa Word UEL & Nộp", 26, 1, "Cả nhóm (100%)", "15/11", "#15803D"),
    ]

    fig, ax = plt.subplots(figsize=(10.6, 4.7), dpi=300)
    fig.patch.set_facecolor("#FFFFFF")
    ax.set_facecolor("#FAFBFC")

    week_bands = [
        (0, 6, "Tuần 1 (20/10–25/10)", "#EFF6FF"),
        (6, 13, "Tuần 2 (26/10–01/11)", "#F0FDF4"),
        (13, 20, "Tuần 3 (02/11–08/11)", "#FEFCE8"),
        (20, 27, "Tuần 4 (09/11–15/11)", "#FDF2F8"),
    ]
    for s, e, lbl, bg in week_bands:
        ax.axvspan(s, e, color=bg, alpha=0.65, zorder=0)
        ax.text((s + e) / 2, len(milestones) - 0.12, lbl, ha="center", va="bottom",
                fontsize=8.2, fontweight="bold", color="#334155", family="DejaVu Sans")

    y_pos = list(range(len(milestones) - 1, -1, -1))
    for idx, (label, start, dur, owner, d_str, col) in enumerate(milestones):
        y = y_pos[idx]
        bar_w = max(dur, 0.85)
        ax.barh(y, bar_w, left=start, height=0.58, color=col, edgecolor="#FFFFFF", linewidth=1.2, zorder=3, alpha=0.92)
        tag = f" {d_str} | {owner}"
        if start + bar_w < 19.5:
            ax.text(start + bar_w + 0.25, y, tag, va="center", ha="left", fontsize=7.6, color="#0F172A", fontweight="semibold")
        else:
            ax.text(start - 0.25, y, f"{d_str} | {owner} ", va="center", ha="right", fontsize=7.6, color="#0F172A", fontweight="semibold")

    ax.set_yticks(y_pos)
    ax.set_yticklabels([m[0] for m in milestones], fontsize=8.1, color="#0F172A", fontweight="medium")
    xticks = [0, 3, 6, 9, 13, 16, 20, 23, 26]
    xlabels = ["20/10", "23/10", "26/10", "29/10", "02/11", "05/11", "09/11", "12/11", "15/11"]
    ax.set_xticks(xticks)
    ax.set_xticklabels(xlabels, fontsize=8.2, color="#334155")
    ax.set_xlim(-0.3, 27.5)
    ax.set_ylim(-0.7, len(milestones) + 0.35)
    ax.grid(axis="x", linestyle="--", linewidth=0.6, color="#CBD5E1", zorder=1)
    ax.grid(axis="y", linestyle=":", linewidth=0.5, color="#E2E8F0", zorder=1)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#94A3B8")
    ax.spines["bottom"].set_color("#94A3B8")

    ax.set_title(
        "LỊCH TRÌNH GANTT THỰC HIỆN ĐỒ ÁN RETAIL ML HUB (20/10/2026 – 15/11/2026)",
        fontsize=10.5, fontweight="bold", color="#1E3A8A", pad=16
    )

    legends = [
        mpatches.Patch(color="#1E3A8A", label="GĐ1: Khởi động & EDA"),
        mpatches.Patch(color="#0284C7", label="GĐ2: Lý thuyết & Lược khảo"),
        mpatches.Patch(color="#0F766E", label="GĐ3: Đặc trưng & 5 Mô hình ML"),
        mpatches.Patch(color="#D97706", label="GĐ4a: Tồn kho SS/ROP & ABC–XYZ"),
        mpatches.Patch(color="#7C3AED", label="GĐ4b: Web Studio & BPMN"),
        mpatches.Patch(color="#BE123C", label="GĐ5: Đồng bộ Excel & Báo cáo"),
        mpatches.Patch(color="#15803D", label="GĐ6: Kiểm thử & Nộp bài"),
    ]
    ax.legend(handles=legends, loc="upper center", bbox_to_anchor=(0.48, -0.09), ncol=4, frameon=False, fontsize=7.5)
    plt.tight_layout()
    fig.savefig(out_path, dpi=300, bbox_inches="tight")
    plt.close(fig)
    return out_path


# ---------------------------------------------------------------------------
# 2. ĐỒNG BỘ ẢNH BÌA UEL, HEADER/FOOTER, CONTENT_TYPES VÀ WORD/STYLES.XML
# ---------------------------------------------------------------------------
def sync_uel_styles_and_resources(zip_files, ctx, ref_docx_path=REF_DOCX_PATH):
    if not os.path.exists(ref_docx_path):
        raise FileNotFoundError(f"Không tìm thấy tài liệu mẫu UEL: {ref_docx_path}")

    with zipfile.ZipFile(ref_docx_path, "r") as refz:
        ref_cover_frame = refz.read("word/media/image1.jpeg")
        ref_cover_bg = refz.read("word/media/image2.jpg")
        ref_header1 = refz.read("word/header1.xml")
        ref_footer1 = refz.read("word/footer1.xml")
        ref_styles_root = ET.fromstring(refz.read("word/styles.xml"))

    # Lưu ảnh nền bìa & khung bìa
    zip_files["word/media/uel_cover_frame.jpeg"] = ref_cover_frame
    zip_files["word/media/uel_cover_bg.jpg"] = ref_cover_bg
    zip_files["word/header1.xml"] = ref_header1
    zip_files["word/footer1.xml"] = ref_footer1
    zip_files["word/_rels/header1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/uel_cover_bg.jpg"/>'
        '</Relationships>'
    ).encode("utf-8")

    # Đăng ký relationships trong word/_rels/document.xml.rels
    def _add_rel(rel_type, target):
        for rel in ctx.rels_root.findall(f"{{{REL_NS}}}Relationship"):
            if rel.attrib.get("Target") == target and rel.attrib.get("Type") == rel_type:
                return rel.attrib["Id"]
        rid = f"rId{ctx.next_rid}"
        ctx.next_rid += 1
        rel = ET.SubElement(ctx.rels_root, f"{{{REL_NS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", rel_type)
        rel.set("Target", target)
        ctx.rel_map[rid] = target
        return rid

    cover_frame_rid = _add_rel(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        "media/uel_cover_frame.jpeg",
    )
    header_rid = _add_rel(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
        "header1.xml",
    )
    footer_rid = _add_rel(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        "footer1.xml",
    )

    # Cập nhật [Content_Types].xml
    ct_root = ET.fromstring(zip_files["[Content_Types].xml"])
    existing_defaults = {el.attrib.get("Extension", "").lower() for el in ct_root.findall(f"{{{CT_NS}}}Default")}
    for ext, mime in (("jpeg", "image/jpeg"), ("jpg", "image/jpeg"), ("png", "image/png")):
        if ext not in existing_defaults:
            d = ET.SubElement(ct_root, f"{{{CT_NS}}}Default")
            d.set("Extension", ext)
            d.set("ContentType", mime)
    existing_overrides = {el.attrib.get("PartName", "") for el in ct_root.findall(f"{{{CT_NS}}}Override")}
    for part, ctype in (
        ("/word/header1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"),
        ("/word/footer1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"),
    ):
        if part not in existing_overrides:
            ov = ET.SubElement(ct_root, f"{{{CT_NS}}}Override")
            ov.set("PartName", part)
            ov.set("ContentType", ctype)
    ET.register_namespace("", CT_NS)
    zip_files["[Content_Types].xml"] = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)

    # Đồng bộ word/styles.xml (giữ Table1..Table8 của MLBA, merge đầy đủ styles từ mẫu UEL)
    orig_styles_root = ET.fromstring(zip_files["word/styles.xml"])
    doc_def = orig_styles_root.find(_w("docDefaults"))
    if doc_def is not None:
        rpr_def = doc_def.find(_w("rPrDefault"))
        if rpr_def is None:
            rpr_def = ET.SubElement(doc_def, _w("rPrDefault"))
        old_rpr = rpr_def.find(_w("rPr"))
        if old_rpr is not None:
            rpr_def.remove(old_rpr)
        new_rpr = _make_rpr(font_name="Times New Roman", sz=26, color="000000")
        lang = ET.SubElement(new_rpr, _w("lang"))
        lang.set(_w("val"), "vi-VN")
        lang.set(_w("eastAsia"), "en-US")
        rpr_def.append(new_rpr)

        ppr_def = doc_def.find(_w("pPrDefault"))
        if ppr_def is None:
            ppr_def = ET.SubElement(doc_def, _w("pPrDefault"))
        ppr = ppr_def.find(_w("pPr"))
        if ppr is None:
            ppr = ET.SubElement(ppr_def, _w("pPr"))
        old_sp = ppr.find(_w("spacing"))
        if old_sp is not None:
            ppr.remove(old_sp)
        sp = ET.SubElement(ppr, _w("spacing"))
        sp.set(_w("after"), "120")
        sp.set(_w("line"), "312")
        sp.set(_w("lineRule"), "auto")

    import_ids = {
        "Normal", "Heading1", "Heading2", "Heading3", "Heading2Char", "Heading3Char",
        "TOC1", "TOC2", "TOC3", "Hyperlink", "TableofFigures", "Caption", "CaptionChar",
        "TIeuDeanh", "TIeuDeanhChar", "Header", "HeaderChar", "Footer", "FooterChar",
        "TableGrid", "DefaultParagraphFont", "TableNormal",
    }
    ref_style_map = {}
    for st in ref_styles_root.findall(_w("style")):
        sid = st.attrib.get(_w("styleId"))
        if sid in import_ids:
            ref_style_map[sid] = copy.deepcopy(st)

    # Xóa bản cũ nếu trùng ID rồi chèn bản từ mẫu
    for st in list(orig_styles_root.findall(_w("style"))):
        if st.attrib.get(_w("styleId")) in ref_style_map:
            orig_styles_root.remove(st)

    def _set_style_props(st_el, sz=None, bold=None, color="000000", jc=None,
                         before=None, after=None, line=None, ind_left=None, strip_italic=False):
        if jc is not None or before is not None or after is not None or line is not None or ind_left is not None:
            ppr = st_el.find(_w("pPr"))
            if ppr is None:
                ppr = ET.SubElement(st_el, _w("pPr"))
            if before is not None or after is not None or line is not None:
                sp = ppr.find(_w("spacing"))
                if sp is None:
                    sp = ET.SubElement(ppr, _w("spacing"))
                if before is not None:
                    sp.set(_w("before"), str(before))
                if after is not None:
                    sp.set(_w("after"), str(after))
                if line is not None:
                    sp.set(_w("line"), str(line))
                    sp.set(_w("lineRule"), "auto")
            if ind_left is not None:
                ind = ppr.find(_w("ind"))
                if ind is None:
                    ind = ET.SubElement(ppr, _w("ind"))
                ind.set(_w("left"), str(ind_left))
            if jc is not None:
                j = ppr.find(_w("jc"))
                if j is None:
                    j = ET.SubElement(ppr, _w("jc"))
                j.set(_w("val"), jc)

        old_rpr = st_el.find(_w("rPr"))
        if st_el.attrib.get(_w("type")) in ("paragraph", "character"):
            if old_rpr is not None and strip_italic:
                for itag in (_w("i"), _w("iCs")):
                    for x in list(old_rpr.findall(itag)):
                        old_rpr.remove(x)
            if old_rpr is not None:
                st_el.remove(old_rpr)
            if sz is not None or bold is not None or color is not None:
                st_el.append(_make_rpr(sz=sz or 26, bold=bool(bold), color=color))

    if "Normal" in ref_style_map:
        _set_style_props(ref_style_map["Normal"], sz=26, bold=False, color="000000", after=120, line=312)
    if "Heading1" in ref_style_map:
        _set_style_props(ref_style_map["Heading1"], sz=28, bold=True, color="000000", jc="center", before=0, after=240)
    if "Heading2" in ref_style_map:
        _set_style_props(ref_style_map["Heading2"], sz=26, bold=True, color="000000", jc="left", before=240, after=120)
    if "Heading3" in ref_style_map:
        _set_style_props(ref_style_map["Heading3"], sz=26, bold=True, color="000000", jc="left", before=160, after=80)
    if "Heading2Char" in ref_style_map:
        _set_style_props(ref_style_map["Heading2Char"], sz=26, bold=True, color="000000")
    if "Heading3Char" in ref_style_map:
        _set_style_props(ref_style_map["Heading3Char"], sz=26, bold=True, color="000000")
    if "Caption" in ref_style_map:
        _set_style_props(ref_style_map["Caption"], sz=26, bold=True, color="000000", jc="center", before=160, after=80, strip_italic=True)
    if "CaptionChar" in ref_style_map:
        _set_style_props(ref_style_map["CaptionChar"], sz=26, bold=True, color="000000", strip_italic=True)
    if "TIeuDeanh" in ref_style_map:
        _set_style_props(ref_style_map["TIeuDeanh"], sz=26, bold=True, color="000000", jc="center", before=80, after=160, strip_italic=True)
    if "TIeuDeanhChar" in ref_style_map:
        _set_style_props(ref_style_map["TIeuDeanhChar"], sz=26, bold=True, color="000000", strip_italic=True)
    if "TOC1" in ref_style_map:
        _set_style_props(ref_style_map["TOC1"], sz=26, bold=True, color="000000", before=40, after=80, line=288)
    if "TOC2" in ref_style_map:
        _set_style_props(ref_style_map["TOC2"], sz=26, bold=False, color="000000", before=20, after=60, line=288, ind_left=240)
    if "TOC3" in ref_style_map:
        _set_style_props(ref_style_map["TOC3"], sz=26, bold=False, color="000000", before=10, after=50, line=288, ind_left=480)
    if "TableofFigures" in ref_style_map:
        _set_style_props(ref_style_map["TableofFigures"], sz=26, bold=False, color="000000", before=20, after=60, line=288)
    if "Footer" in ref_style_map:
        _set_style_props(ref_style_map["Footer"], sz=22, bold=False, color="000000")
    if "Header" in ref_style_map:
        _set_style_props(ref_style_map["Header"], sz=22, bold=False, color="000000")

    for sid in import_ids:
        if sid in ref_style_map:
            orig_styles_root.append(ref_style_map[sid])

    zip_files["word/styles.xml"] = ET.tostring(orig_styles_root, encoding="utf-8", xml_declaration=True)
    return {
        "cover_frame_rid": cover_frame_rid,
        "header_rid": header_rid,
        "footer_rid": footer_rid,
    }


# ---------------------------------------------------------------------------
# 3. TẠO TRANG BÌA (SECTION 1: P0..P18) & PHẦN MỞ ĐẦU (SECTION 2)
# ---------------------------------------------------------------------------
def _make_cover_anchor_p(cover_frame_rid, docpr_id):
    p = ET.Element(_w("p"))
    ppr = ET.SubElement(p, _w("pPr"))
    sp = ET.SubElement(ppr, _w("spacing"))
    sp.set(_w("after"), "360")
    sp.set(_w("line"), "240")
    sp.set(_w("lineRule"), "auto")
    ET.SubElement(ppr, _w("jc")).set(_w("val"), "center")
    ppr.append(_make_rpr(sz=36, bold=True))

    r1 = ET.SubElement(p, _w("r"))
    r1.append(_make_rpr(sz=36, no_proof=True))
    dr = ET.SubElement(r1, _w("drawing"))
    anc = ET.SubElement(dr, _wp("anchor"))
    for k, v in (
        ("distT", "0"), ("distB", "0"), ("distL", "114300"), ("distR", "114300"),
        ("simplePos", "0"), ("relativeHeight", "251659264"), ("behindDoc", "1"),
        ("locked", "0"), ("layoutInCell", "1"), ("allowOverlap", "1"),
    ):
        anc.set(k, v)
    spos = ET.SubElement(anc, _wp("simplePos"))
    spos.set("x", "0")
    spos.set("y", "0")
    ph = ET.SubElement(anc, _wp("positionH"))
    ph.set("relativeFrom", "page")
    ET.SubElement(ph, _wp("align")).text = "right"
    pv = ET.SubElement(anc, _wp("positionV"))
    pv.set("relativeFrom", "paragraph")
    ET.SubElement(pv, _wp("posOffset")).text = "-903514"
    ext = ET.SubElement(anc, _wp("extent"))
    ext.set("cx", "7550785")
    ext.set("cy", "10666095")
    eff = ET.SubElement(anc, _wp("effectExtent"))
    eff.set("l", "0")
    eff.set("t", "0")
    eff.set("r", "0")
    eff.set("b", "1905")
    ET.SubElement(anc, _wp("wrapNone"))
    dp = ET.SubElement(anc, _wp("docPr"))
    dp.set("id", str(docpr_id))
    dp.set("name", "UEL Cover Frame")
    cnv = ET.SubElement(anc, _wp("cNvGraphicFramePr"))
    gfl = ET.SubElement(cnv, _a("graphicFrameLocks"))
    gfl.set("noChangeAspect", "1")
    gr = ET.SubElement(anc, _a("graphic"))
    gd = ET.SubElement(gr, _a("graphicData"))
    gd.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")
    pic = ET.SubElement(gd, _pic("pic"))
    nv = ET.SubElement(pic, _pic("nvPicPr"))
    cnvp = ET.SubElement(nv, _pic("cNvPr"))
    cnvp.set("id", "0")
    cnvp.set("name", "uel_cover_frame.jpeg")
    ET.SubElement(nv, _pic("cNvPicPr"))
    bf = ET.SubElement(pic, _pic("blipFill"))
    bl = ET.SubElement(bf, _a("blip"))
    bl.set(_r("embed"), cover_frame_rid)
    bl.set("cstate", "print")
    st = ET.SubElement(bf, _a("stretch"))
    ET.SubElement(st, _a("fillRect"))
    spp = ET.SubElement(pic, _pic("spPr"))
    xf = ET.SubElement(spp, _a("xfrm"))
    off = ET.SubElement(xf, _a("off"))
    off.set("x", "0")
    off.set("y", "0")
    xext = ET.SubElement(xf, _a("ext"))
    xext.set("cx", "7550785")
    xext.set("cy", "10666095")
    pg = ET.SubElement(spp, _a("prstGeom"))
    pg.set("prst", "rect")
    ET.SubElement(pg, _a("avLst"))

    r2 = ET.SubElement(p, _w("r"))
    r2.append(_make_rpr(sz=36, bold=True))
    t = ET.SubElement(r2, _w("t"))
    t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = " "
    return p


def build_uel_cover_and_front_matter(ctx, rids, gantt_png_path):
    elems = []

    # --- SECTION 1: TRANG BÌA (P0..P18 giống chính xác Final_Project_Web.docx) ---
    elems.append(_make_cover_anchor_p(rids["cover_frame_rid"], ctx.next_shape_id()))  # P0
    elems.append(_make_p([], align="center", before=None, after=360, line=240, rpr_sz=36, rpr_bold=True))  # P1
    elems.append(_make_p([("KHOA HỆ THỐNG THÔNG TIN", True, False, "000000", 36)],
                         align="center", before=None, after=360, line=240, rpr_sz=36, rpr_bold=True))  # P2
    for _ in range(3):  # P3..P5 (vượt qua vòng tròn logo UEL)
        elems.append(_make_p([], align="center", before=None, after=360, line=240, rpr_sz=48, rpr_bold=True))
    elems.append(_make_p([("BÁO CÁO CUỐI KỲ", True, False, "000000", 48)],
                         align="center", before=None, after=360, line=240, rpr_sz=48, rpr_bold=True))  # P6
    elems.append(_make_p([("HỌC MÁY TRONG PHÂN TÍCH KINH DOANH", True, False, "000000", 36)],
                         align="center", before=None, after=360, line=240, rpr_sz=36, rpr_bold=True))  # P7
    elems.append(_make_p([("ĐỀ TÀI: XÂY DỰNG HỆ THỐNG DỰ BÁO NHU CẦU", True, False, "365F91", 40)],
                         align="center", before=200, after=120, line=240, rpr_sz=40, rpr_color="365F91"))  # P8
    elems.append(_make_p([("VÀ TỐI ƯU HÓA TỒN KHO BÁN LẺ", True, False, "365F91", 40)],
                         align="center", before=None, after=40, line=240, rpr_sz=40, rpr_color="365F91"))  # P9
    elems.append(_make_p([("TÍCH HỢP HỌC MÁY (RETAIL ML HUB)", True, False, "365F91", 40)],
                         align="center", before=None, after=480, line=240, rpr_sz=40, rpr_color="365F91"))  # P10

    elems.append(_make_p([
        ("Giảng viên hướng dẫn: ", True, False, "000000", 36),
        ("ThS. [Tên Giảng viên hướng dẫn]", False, False, "000000", 36),
    ], align="center", before=360, after=80, line=240, rpr_sz=36))  # P11
    elems.append(_make_p([
        ("Lớp học phần: ", True, False, "000000", 36),
        ("Học máy trong Phân tích Kinh doanh (HK7)", False, False, "000000", 36),
    ], align="center", before=360, after=80, line=240, rpr_sz=36))  # P12
    elems.append(_make_p([
        ("Nhóm thực hiện: ", True, False, "000000", 36),
        ("Nhóm Đồ án MLBA", False, False, "000000", 36),
    ], align="center", before=360, after=80, line=240, rpr_sz=36))  # P13

    elems.append(_make_p([], align="center", before=500, after=0, line=240, rpr_sz=32, rpr_color="FFFFFF"))  # P14
    elems.append(_make_p([], align="center", before=360, after=80, line=240, rpr_sz=40))  # P15
    cover_tabs = [("center", 4514, None), ("left", 7251, None)]
    elems.append(_make_p([], align="left", before=0, after=0, line=240, tabs=cover_tabs, rpr_sz=26))  # P16
    elems.append(_make_p([], align="left", before=0, after=0, line=240, tabs=cover_tabs, rpr_sz=26))  # P17

    p18 = _make_p([
        ("TP. HỒ CHÍ MINH, THÁNG 11 NĂM 2026", True, False, "FFFFFF", 32),
        ("\t", False, False, "000000", 26),
    ], align="right", before=240, after=0, line=240, tabs=cover_tabs, rpr_sz=26)  # P18
    p18_ppr = p18.find(_w("pPr"))
    s1 = ET.SubElement(p18_ppr, _w("sectPr"))
    fref = ET.SubElement(s1, _w("footerReference"))
    fref.set(_w("type"), "default")
    fref.set(_r("id"), rids["footer_rid"])
    href = ET.SubElement(s1, _w("headerReference"))
    href.set(_w("type"), "first")
    href.set(_r("id"), rids["header_rid"])
    pgsz = ET.SubElement(s1, _w("pgSz"))
    pgsz.set(_w("w"), "11909")
    pgsz.set(_w("h"), "16834")
    pgmar = ET.SubElement(s1, _w("pgMar"))
    for k, v in (("top", "1440"), ("right", "1440"), ("bottom", "1440"), ("left", "1440"),
                 ("header", "720"), ("footer", "720"), ("gutter", "0")):
        pgmar.set(_w(k), v)
    ET.SubElement(s1, _w("cols")).set(_w("space"), "720")
    ET.SubElement(s1, _w("titlePg"))
    elems.append(p18)

    # --- SECTION 2: PHẦN MỞ ĐẦU CHUẨN MẪU UEL ---
    # 1. MỤC LỤC (Normal bold centered 16pt, KHÔNG đưa vào Heading1)
    elems.append(_make_p([("MỤC LỤC", True, False, "000000", 32)],
                         align="center", before=0, after=240, line=288, keep_next=True, rpr_sz=32, rpr_bold=True))
    elems.append(_make_marker("[[TOC_PLACEHOLDER]]"))

    # 2. DANH SÁCH HÌNH ẢNH
    elems.append(_make_p([("DANH SÁCH HÌNH ẢNH", True, False, "000000", 28)],
                         align="center", style="Heading1", before=0, after=240, line=288,
                         keep_next=True, keep_lines=True, page_break_before=True, rpr_sz=28, rpr_bold=True))
    elems.append(_make_marker("[[LOF_PLACEHOLDER]]"))

    # 3. DANH SÁCH BẢNG
    elems.append(_make_p([("DANH SÁCH BẢNG", True, False, "000000", 28)],
                         align="center", style="Heading1", before=0, after=240, line=288,
                         keep_next=True, keep_lines=True, page_break_before=True, rpr_sz=28, rpr_bold=True))
    elems.append(_make_marker("[[LOT_PLACEHOLDER]]"))

    # 4. DANH SÁCH THÀNH VIÊN NHÓM (Table #1)
    elems.append(_make_p([("DANH SÁCH THÀNH VIÊN NHÓM", True, False, "000000", 28)],
                         align="center", style="Heading1", before=0, after=240, line=288,
                         keep_next=True, keep_lines=True, page_break_before=True, rpr_sz=28, rpr_bold=True))
    elems.append(_make_p(
        "Đồ án môn học được thực hiện bởi nhóm 05 sinh viên dưới sự hướng dẫn chuyên môn của Giảng viên bộ môn Học máy trong Phân tích Kinh doanh. Tỷ lệ đóng góp được đánh giá thống nhất giữa các thành viên (tổng cộng 100%, mức độ hoàn thành 100%):",
        first_line=720, before=80, after=120,
    ))
    elems.append(_make_grid_table(
        headers=["STT", "Họ và tên", "MSSV", "Vai trò", "Đóng góp", "Hoàn thành", "Ký tên"],
        rows=[list(m) + [""] for m in TEAM_MEMBERS],
        col_widths_dxa=[620, 2766, 1380, 1280, 960, 1020, 1000],
        col_aligns=["center", "left", "center", "center", "center", "center", "center"],
        font_sz=23,
    ))

    # 5. PHÂN CÔNG NHIỆM VỤ (Table #2 + Gantt PNG)
    elems.append(_make_p([("PHÂN CÔNG NHIỆM VỤ", True, False, "000000", 28)],
                         align="center", style="Heading1", before=0, after=240, line=288,
                         keep_next=True, keep_lines=True, page_break_before=True, rpr_sz=28, rpr_bold=True))
    elems.append(_make_p(
        "Bảng phân công dưới đây tóm tắt vai trò, đầu việc và sản phẩm bàn giao của từng thành viên xuyên suốt 6 giai đoạn triển khai từ ngày 20/10/2026 đến ngày 15/11/2026:",
        first_line=720, before=80, after=120,
    ))
    elems.append(_make_grid_table(
        headers=["STT", "Thành viên", "Vai trò", "Phạm vi phụ trách & Nhiệm vụ chính", "Sản phẩm & Đánh giá"],
        rows=[list(row) for row in TASK_ASSIGNMENTS],
        col_widths_dxa=[600, 1850, 1080, 3446, 2050],
        col_aligns=["center", "left", "center", "left", "left"],
        font_sz=22,
    ))
    elems.append(_make_p(
        [("Biểu đồ Gantt kế hoạch và tiến độ thực hiện đồ án (20/10/2026 – 15/11/2026)", True, False, "000000", 26)],
        align="center", before=200, after=80, line=264, keep_next=True, rpr_sz=26, rpr_bold=True,
    ))
    gantt_rid, _, gw, gh = ctx.register_new_png(gantt_png_path)
    elems.append(_make_inline_image_p(
        gantt_rid, gw, gh, ctx.next_shape_id(), "gantt_tien_do_du_an.png",
        max_w_emu=5700000, max_h_emu=2800000, keep_next=False,
    ))

    # 6. DANH MỤC TỪ VIẾT TẮT (Table #3)
    elems.append(_make_p([("DANH MỤC TỪ VIẾT TẮT", True, False, "000000", 28)],
                         align="center", style="Heading1", before=0, after=240, line=288,
                         keep_next=True, keep_lines=True, page_break_before=True, rpr_sz=28, rpr_bold=True))
    elems.append(_make_grid_table(
        headers=["STT", "Từ viết tắt", "Diễn giải thuật ngữ đầy đủ"],
        rows=[list(r) for r in ACRONYMS],
        col_widths_dxa=[800, 1800, 6426],
        col_aligns=["center", "center", "left"],
        font_sz=23,
    ))

    # 7. LỜI CẢM ƠN
    elems.append(_make_p([("LỜI CẢM ƠN", True, False, "000000", 28)],
                         align="center", style="Heading1", before=0, after=240, line=288,
                         keep_next=True, keep_lines=True, page_break_before=True, rpr_sz=28, rpr_bold=True))
    for para in (
        "Nhóm xin chân thành cảm ơn quý thầy cô Khoa Hệ thống Thông tin, Trường Đại học Kinh tế – Luật, Đại học Quốc gia Thành phố Hồ Chí Minh đã tận tâm truyền đạt nền tảng kiến thức về phân tích dữ liệu, học máy ứng dụng và phân tích kinh doanh trong suốt quá trình học tập.",
        "Đặc biệt, nhóm xin gửi lời tri ân sâu sắc đến Giảng viên hướng dẫn bộ môn Học máy trong Phân tích Kinh doanh đã nhiệt tình định hướng, góp ý chuyên môn và hỗ trợ nhóm hoàn thiện hệ thống dự báo nhu cầu kết hợp tối ưu hóa tồn kho bán lẻ (Retail ML Hub).",
        "Nhóm cũng xin trân trọng cảm ơn gia đình, bạn bè và các bạn cùng lớp đã luôn động viên, chia sẻ tài liệu và hỗ trợ nhóm hoàn thành đồ án đúng tiến độ đề ra.",
    ):
        elems.append(_make_p(para, first_line=720, before=120, after=120, line=312))
    elems.append(_make_p([("Nhóm thực hiện", True, False, "000000", 26)],
                         align="right", before=240, after=120, line=288, rpr_sz=26, rpr_bold=True))

    # 8. LỜI CAM KẾT
    elems.append(_make_p([("LỜI CAM KẾT", True, False, "000000", 28)],
                         align="center", style="Heading1", before=0, after=240, line=288,
                         keep_next=True, keep_lines=True, page_break_before=True, rpr_sz=28, rpr_bold=True))
    for para in (
        "Nhóm xin cam kết toàn bộ nội dung báo cáo đồ án “Xây dựng hệ thống dự báo nhu cầu và tối ưu hóa tồn kho bán lẻ tích hợp Học máy (Retail ML Hub)”, bao gồm mã nguồn huấn luyện mô hình Python, ứng dụng Web Studio, bảng tính tổng hợp Excel và bộ sơ đồ BPMN là kết quả nghiên cứu, thiết kế và triển khai thực tế của tập thể nhóm.",
        "Tất cả dữ liệu thực nghiệm (73.100 bản ghi), kết quả huấn luyện (5 mô hình dự báo với Random Forest RMSE = 92,42), tham số tồn kho (SS = 311, ROP = 777) và phân tích ma trận ABC–XYZ đều trung thực, nhất quán và được kiểm chứng trực tiếp từ mã nguồn thực thi.",
        "Nhóm không sao chép bất hợp pháp công trình của bất kỳ cá nhân hay tổ chức nào; các tài liệu, công trình tham khảo sử dụng trong báo cáo đều được trích dẫn đầy đủ và đúng chuẩn mực tại mục Tài liệu tham khảo.",
    ):
        elems.append(_make_p(para, first_line=720, before=120, after=120, line=312))
    elems.append(_make_p([("Nhóm thực hiện", True, False, "000000", 26)],
                         align="right", before=240, after=120, line=288, rpr_sz=26, rpr_bold=True))

    return elems


# ---------------------------------------------------------------------------
# 4. CHUẨN HÓA TOÀN BỘ THÂN BÀI, FIX 4.3.3 / 4.4.1 / BẢNG 6.1 & DỰNG TOC/LOF/LOT
# ---------------------------------------------------------------------------
def _set_para_text_preserve_first(p, new_text):
    runs = p.findall(_w("r"))
    if not runs:
        r = ET.SubElement(p, _w("r"))
        r.append(_make_rpr(sz=26, bold=True))
        ET.SubElement(r, _w("t")).text = new_text
        return
    kept_rpr = copy.deepcopy(runs[0].find(_w("rPr")))
    for r in runs:
        if r.find(_w("drawing")) is None:
            p.remove(r)
    r = ET.SubElement(p, _w("r"))
    r.append(kept_rpr if kept_rpr is not None else _make_rpr(sz=26, bold=True))
    t = ET.SubElement(r, _w("t"))
    t.text = new_text


def _format_para_ppr_and_runs(p, align="both", style=None, before=120, after=120, line=312,
                              first_line=None, left=None, hanging=None, keep_next=False, keep_lines=False,
                              page_break_before=False, bold=None, sz=26, color="000000", strip_italic=False,
                              preserve_numpr=False, preserve_run_bold=False):
    old_ppr = p.find(_w("pPr"))
    saved_numpr = copy.deepcopy(old_ppr.find(_w("numPr"))) if (preserve_numpr and old_ppr is not None and old_ppr.find(_w("numPr")) is not None) else None
    if old_ppr is not None:
        p.remove(old_ppr)

    new_ppr = ET.Element(_w("pPr"))
    if style:
        ET.SubElement(new_ppr, _w("pStyle")).set(_w("val"), style)
    if keep_next:
        ET.SubElement(new_ppr, _w("keepNext"))
    if keep_lines:
        ET.SubElement(new_ppr, _w("keepLines"))
    if page_break_before:
        ET.SubElement(new_ppr, _w("pageBreakBefore"))
    if saved_numpr is not None:
        new_ppr.append(saved_numpr)
    sp = ET.SubElement(new_ppr, _w("spacing"))
    sp.set(_w("before"), str(before))
    sp.set(_w("after"), str(after))
    sp.set(_w("line"), str(line))
    sp.set(_w("lineRule"), "auto")
    if first_line or left or hanging:
        ind = ET.SubElement(new_ppr, _w("ind"))
        if left:
            ind.set(_w("left"), str(left))
        if first_line:
            ind.set(_w("firstLine"), str(first_line))
        if hanging:
            ind.set(_w("hanging"), str(hanging))
    if align:
        ET.SubElement(new_ppr, _w("jc")).set(_w("val"), align)
    new_ppr.append(_make_rpr(sz=sz, bold=bool(bold), color=color))
    p.insert(0, new_ppr)

    # Chuẩn hóa runs trực tiếp và bên trong hyperlink (bảo toàn <w:drawing>, <m:oMath>)
    def _norm_run(r, force_hyperlink=False):
        if r.find(_w("drawing")) is not None:
            return
        old_rpr = r.find(_w("rPr"))
        was_b = False
        was_i = False
        if old_rpr is not None:
            b_el = old_rpr.find(_w("b"))
            was_b = b_el is not None and b_el.attrib.get(_w("val"), "1") != "0"
            i_el = old_rpr.find(_w("i"))
            was_i = i_el is not None and i_el.attrib.get(_w("val"), "1") != "0"
            r.remove(old_rpr)
        final_b = was_b if (bold is None and preserve_run_bold) else bool(bold)
        final_i = False if strip_italic else was_i
        new_rpr = _make_rpr(
            sz=sz,
            bold=final_b,
            italic=final_i,
            color="0563C1" if force_hyperlink else color,
            underline="single" if force_hyperlink else None,
        )
        r.insert(0, new_rpr)

    for child in list(p):
        if child.tag == _w("r"):
            _norm_run(child, force_hyperlink=False)
        elif child.tag == _w("hyperlink"):
            for hr in child.findall(_w("r")):
                _norm_run(hr, force_hyperlink=True)


def _normalize_existing_table(tbl):
    tbl_pr = tbl.find(_w("tblPr"))
    if tbl_pr is not None:
        tbl.remove(tbl_pr)
    new_pr = ET.Element(_w("tblPr"))
    ET.SubElement(new_pr, _w("tblStyle")).set(_w("val"), "TableGrid")
    tw = ET.SubElement(new_pr, _w("tblW"))
    tw.set(_w("w"), "9026")
    tw.set(_w("type"), "dxa")
    ET.SubElement(new_pr, _w("jc")).set(_w("val"), "center")
    tbd = ET.SubElement(new_pr, _w("tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = ET.SubElement(tbd, _w(side))
        b.set(_w("val"), "single")
        b.set(_w("sz"), "4")
        b.set(_w("space"), "0")
        b.set(_w("color"), "BFBFBF")
    ET.SubElement(new_pr, _w("tblLayout")).set(_w("type"), "fixed")
    cm = ET.SubElement(new_pr, _w("tblCellMar"))
    for k, val in (("top", "80"), ("left", "105"), ("bottom", "80"), ("right", "105")):
        m = ET.SubElement(cm, _w(k))
        m.set(_w("w"), val)
        m.set(_w("type"), "dxa")
    ET.SubElement(new_pr, _w("tblLook")).set(_w("val"), "04A0")
    tbl.insert(0, new_pr)

    # Cân đối chiều rộng cột về đúng 9026 dxa
    grid = tbl.find(_w("tblGrid"))
    col_widths = []
    if grid is not None:
        cols = grid.findall(_w("gridCol"))
        raw_w = [max(450, round(float(c.attrib.get(_w("w"), "1200")))) for c in cols]
        tot = sum(raw_w) or 1
        col_widths = [max(450, round(w * 9026 / tot)) for w in raw_w]
        diff = 9026 - sum(col_widths)
        if col_widths:
            col_widths[-1] += diff
        for c, nw in zip(cols, col_widths):
            c.set(_w("w"), str(nw))

    trs = tbl.findall(_w("tr"))
    num_cols = len(col_widths) if col_widths else (len(trs[0].findall(_w("tc"))) if trs else 1)
    cell_sz = 22 if num_cols >= 6 else 23
    for r_idx, tr in enumerate(trs):
        is_hdr = r_idx == 0
        tr_pr = tr.find(_w("trPr"))
        if tr_pr is None:
            tr_pr = ET.Element(_w("trPr"))
            tr.insert(0, tr_pr)
        if tr_pr.find(_w("cantSplit")) is None:
            ET.SubElement(tr_pr, _w("cantSplit"))
        if is_hdr and tr_pr.find(_w("tblHeader")) is None:
            ET.SubElement(tr_pr, _w("tblHeader"))

        c_pos = 0
        for tc in tr.findall(_w("tc")):
            tc_pr = tc.find(_w("tcPr"))
            if tc_pr is None:
                tc_pr = ET.Element(_w("tcPr"))
                tc.insert(0, tc_pr)
            gs_el = tc_pr.find(_w("gridSpan"))
            span = int(gs_el.attrib.get(_w("val"), "1")) if gs_el is not None else 1
            if col_widths and c_pos < len(col_widths):
                w_sum = sum(col_widths[c_pos:c_pos + span])
                tcw = tc_pr.find(_w("tcW"))
                if tcw is None:
                    tcw = ET.SubElement(tc_pr, _w("tcW"))
                tcw.set(_w("w"), str(w_sum))
                tcw.set(_w("type"), "dxa")
            c_pos += span

            for shd_old in list(tc_pr.findall(_w("shd"))):
                tc_pr.remove(shd_old)
            shd = ET.SubElement(tc_pr, _w("shd"))
            shd.set(_w("val"), "clear")
            shd.set(_w("color"), "auto")
            shd.set(_w("fill"), "D9E2F3" if is_hdr else ("F8FAFC" if r_idx % 2 == 0 else "FFFFFF"))
            va = tc_pr.find(_w("vAlign"))
            if va is None:
                va = ET.SubElement(tc_pr, _w("vAlign"))
            va.set(_w("val"), "center")

            for p in tc.findall(_w("p")):
                al = "center" if is_hdr else None
                if not is_hdr:
                    old_ppr = p.find(_w("pPr"))
                    old_jc = old_ppr.find(_w("jc")) if old_ppr is not None else None
                    al = old_jc.attrib.get(_w("val"), "left") if old_jc is not None else "left"
                    if al == "both":
                        al = "left"
                _format_para_ppr_and_runs(
                    p, align=al, before=30, after=30, line=260,
                    bold=True if is_hdr else None, sz=cell_sz,
                    color="0F172A" if is_hdr else "000000",
                    preserve_run_bold=not is_hdr, preserve_numpr=True,
                )


def _clamp_drawings_and_renumber_ids(el, ctx):
    for dp in el.findall(f".//{_wp('docPr')}"):
        dp.set("id", ctx.next_shape_id())
    for cnv in el.findall(f".//{_pic('cNvPr')}"):
        cnv.set("id", "0")
    max_w_emu = 5730000
    for tag_parent in (_wp("inline"), _wp("anchor")):
        for box in el.findall(f".//{tag_parent}"):
            if tag_parent == _wp("anchor") and box.attrib.get("behindDoc") == "1":
                continue
            ext = box.find(_wp("extent"))
            if ext is not None:
                cx = int(ext.attrib.get("cx", "0"))
                cy = int(ext.attrib.get("cy", "0"))
                if cx > max_w_emu and cx > 0:
                    ratio = max_w_emu / cx
                    new_cx = max_w_emu
                    new_cy = int(cy * ratio)
                    ext.set("cx", str(new_cx))
                    ext.set("cy", str(new_cy))
                    for aext in box.findall(f".//{_a('ext')}"):
                        aext.set("cx", str(new_cx))
                        aext.set("cy", str(new_cy))


def _wrap_bookmark(p, bm_id, bm_name):
    children = list(p)
    ppr = p.find(_w("pPr"))
    insert_idx = list(p).index(ppr) + 1 if ppr is not None else 0
    bs = ET.Element(_w("bookmarkStart"))
    bs.set(_w("id"), str(bm_id))
    bs.set(_w("name"), bm_name)
    be = ET.Element(_w("bookmarkEnd"))
    be.set(_w("id"), str(bm_id))
    p.insert(insert_idx, bs)
    p.append(be)


def _build_field_list_paragraphs(field_instr, entries, default_style="TableofFigures", is_main_toc=False):
    """Tạo cụm field Word hoàn chỉnh (begin/separate + các hyperlink PAGEREF + fldChar end)."""
    if not entries:
        return []
    out = []
    for idx, (text, bm_name, page_num, level) in enumerate(entries):
        is_first = idx == 0
        is_last = idx == len(entries) - 1
        if is_main_toc:
            pstyle = f"TOC{level}"
            is_bold = level == 1
        else:
            pstyle = default_style
            is_bold = False

        p = ET.Element(_w("p"))
        ppr = ET.SubElement(p, _w("pPr"))
        ET.SubElement(ppr, _w("pStyle")).set(_w("val"), pstyle)
        tabs = ET.SubElement(ppr, _w("tabs"))
        tab = ET.SubElement(tabs, _w("tab"))
        tab.set(_w("val"), "right")
        tab.set(_w("leader"), "dot")
        tab.set(_w("pos"), "9019")
        ppr.append(_make_rpr(sz=26, bold=is_bold, no_proof=True))

        if is_first:
            rb = ET.SubElement(p, _w("r"))
            rb.append(_make_rpr(sz=26, bold=is_bold))
            ET.SubElement(rb, _w("fldChar")).set(_w("fldCharType"), "begin")

            ri = ET.SubElement(p, _w("r"))
            ri.append(_make_rpr(sz=26, bold=is_bold))
            it = ET.SubElement(ri, _w("instrText"))
            it.set(f"{{{XML_NS}}}space", "preserve")
            it.text = f" {field_instr} "

            rs = ET.SubElement(p, _w("r"))
            rs.append(_make_rpr(sz=26, bold=is_bold))
            ET.SubElement(rs, _w("fldChar")).set(_w("fldCharType"), "separate")

        hl = ET.SubElement(p, _w("hyperlink"))
        hl.set(_w("anchor"), bm_name)
        hl.set(_w("history"), "1")

        rt = ET.SubElement(hl, _w("r"))
        rt.append(_make_rpr(sz=26, bold=is_bold, color="000000", underline="none", rstyle="Hyperlink", no_proof=True))
        tt = ET.SubElement(rt, _w("t"))
        if text.startswith(" ") or text.endswith(" "):
            tt.set(f"{{{XML_NS}}}space", "preserve")
        tt.text = text

        rtab = ET.SubElement(hl, _w("r"))
        rtab.append(_make_rpr(sz=26, bold=is_bold, no_proof=True, web_hidden=True))
        ET.SubElement(rtab, _w("tab"))

        rpb = ET.SubElement(hl, _w("r"))
        rpb.append(_make_rpr(sz=26, bold=is_bold, no_proof=True, web_hidden=True))
        ET.SubElement(rpb, _w("fldChar")).set(_w("fldCharType"), "begin")

        rpi = ET.SubElement(hl, _w("r"))
        rpi.append(_make_rpr(sz=26, bold=is_bold, no_proof=True, web_hidden=True))
        pit = ET.SubElement(rpi, _w("instrText"))
        pit.set(f"{{{XML_NS}}}space", "preserve")
        pit.text = f" PAGEREF {bm_name} \\h "

        rps = ET.SubElement(hl, _w("r"))
        rps.append(_make_rpr(sz=26, bold=is_bold, no_proof=True, web_hidden=True))
        ET.SubElement(rps, _w("fldChar")).set(_w("fldCharType"), "separate")

        rpn = ET.SubElement(hl, _w("r"))
        rpn.append(_make_rpr(sz=26, bold=is_bold, no_proof=True, web_hidden=True))
        ET.SubElement(rpn, _w("t")).text = str(page_num)

        rpe = ET.SubElement(hl, _w("r"))
        rpe.append(_make_rpr(sz=26, bold=is_bold, no_proof=True, web_hidden=True))
        ET.SubElement(rpe, _w("fldChar")).set(_w("fldCharType"), "end")

        if is_last:
            rend = ET.SubElement(p, _w("r"))
            rend.append(_make_rpr(sz=26, bold=is_bold))
            ET.SubElement(rend, _w("fldChar")).set(_w("fldCharType"), "end")

        out.append(p)
    return out


def _estimate_element_pt(el, style_id=""):
    if el.tag == _w("tbl"):
        trs = el.findall(_w("tr"))
        tot_h = 0.0
        for tr in trs:
            max_lines = 1
            for tc in tr.findall(_w("tc")):
                txt = _ptxt(tc)
                max_lines = max(max_lines, math.ceil(max(len(txt), 1) / 28))
            tot_h += max_lines * 16.0 + 10.0
        return tot_h + 14.0

    exts = el.findall(f".//{_wp('extent')}")
    if exts:
        cy = max(int(e.attrib.get("cy", "0")) for e in exts)
        return (cy / 12700.0) + 18.0

    txt = _ptxt(el)
    if style_id == "Heading1":
        return 36.0
    if style_id in ("Heading2", "Heading3"):
        return 32.0
    if style_id in ("TIeuDeanh", "Caption"):
        return 24.0
    if txt in ("[[TOC_PLACEHOLDER]]", "[[LOF_PLACEHOLDER]]", "[[LOT_PLACEHOLDER]]"):
        return 0.0
    lines = max(1, math.ceil(max(len(txt), 1) / 68))
    return lines * 17.5 + 10.0


def _compute_page_map(elems, style_by_idx, toc_n=60, lof_n=36, lot_n=13):
    page = 1  # Cover P0..P18
    used = 0.0
    usable = 695.0
    page_of_idx = {}

    for i, el in enumerate(elems):
        if i <= 18:
            page_of_idx[i] = 1
            continue
        if i == 19:
            page = 2
            used = 0.0

        txt = _ptxt(el).strip() if el.tag == _w("p") else ""
        st = style_by_idx.get(i, "")
        ppr = el.find(_w("pPr")) if el.tag == _w("p") else None
        has_pb = ppr is not None and ppr.find(_w("pageBreakBefore")) is not None

        if txt == "[[TOC_PLACEHOLDER]]":
            for _ in range(toc_n):
                if used + 22.0 > usable:
                    page += 1
                    used = 0.0
                used += 22.0
            continue
        if txt == "[[LOF_PLACEHOLDER]]":
            for _ in range(lof_n):
                if used + 22.0 > usable:
                    page += 1
                    used = 0.0
                used += 22.0
            continue
        if txt == "[[LOT_PLACEHOLDER]]":
            for _ in range(lot_n):
                if used + 22.0 > usable:
                    page += 1
                    used = 0.0
                used += 22.0
            continue

        if has_pb and used > 0.0:
            page += 1
            used = 0.0

        h = _estimate_element_pt(el, st)
        if used + h > usable and used > 0.0:
            page += 1
            used = 0.0
        page_of_idx[i] = page
        used += h
        while used >= usable:
            page += 1
            used -= usable
    return page_of_idx


def finalize_uel_document(body_elements, sect_pr, footer_rid, ctx):
    # Tách phần Section 1 (Cover P0..P18) và Section 2 (front matter + body)
    cover_els = list(body_elements[:19])
    sec2_raw = list(body_elements[19:])

    # BƯỚC A: Tiền xử lý cấu trúc (tách 4.3.3 ARIMA kết hợp ảnh, bỏ dòng trống thừa, sửa Bảng 4.6, thêm Bảng 6.1)
    cleaned = []
    for el in sec2_raw:
        if el.tag == _w("tbl"):
            cleaned.append(el)
            continue
        txt = _ptxt(el).strip()
        drawings = el.findall(f".//{_w('drawing')}")
        maths = el.findall(f".//{_m_tag('oMath')}") if False else el.findall(f".//{{{M_NS}}}oMath")
        has_br = bool(el.findall(f".//{_w('br')}"))
        has_sect = el.find(_w("pPr")) is not None and el.find(_w("pPr")).find(_w("sectPr")) is not None

        # Fix 4.3.3: tách đoạn 'Hình 4.4: Mô hình ARIMA(2, 1, 1)' + drawing thành tiêu đề phụ + đoạn ảnh riêng
        if txt.startswith("Hình 4.4: Mô hình ARIMA") and drawings:
            cleaned.append(_make_p([("a) Mô hình ARIMA(2, 1, 1)", True, False, "000000", 26)],
                                   align="left", before=160, after=80, line=288, keep_next=True))
            img_p = ET.Element(_w("p"))
            ppr = ET.SubElement(img_p, _w("pPr"))
            ET.SubElement(ppr, _w("keepNext"))
            sp = ET.SubElement(ppr, _w("spacing"))
            sp.set(_w("before"), "120")
            sp.set(_w("after"), "60")
            ET.SubElement(ppr, _w("jc")).set(_w("val"), "center")
            for r in el.findall(_w("r")):
                if r.find(_w("drawing")) is not None:
                    img_p.append(copy.deepcopy(r))
            cleaned.append(img_p)
            continue

        # Bỏ các đoạn trống dư từ bản Google Docs gốc (giữ marker & đoạn có hình/toán/break)
        if not txt and not drawings and not maths and not has_br and not has_sect:
            continue

        # Chuẩn hóa tên tiêu đề & caption trước khi phân loại
        if txt.startswith("Hình 4.4. Kết quả huấn luyện"):
            _set_para_text_preserve_first(el, "Hình 4.4: Kết quả huấn luyện và dự báo mô hình ARIMA(2, 1, 1)")
        elif txt.startswith("Hình 4.5: Mô hình Holt-Winters"):
            _set_para_text_preserve_first(el, "b) Mô hình Holt-Winters Exponential Smoothing")
        elif txt.startswith("Hình 4.5 Kết quả huấn luyện"):
            _set_para_text_preserve_first(el, "Hình 4.5: Kết quả huấn luyện và dự báo mô hình Exponential Smoothing")
        elif txt.startswith("Hình 4.6: Mô hình SARIMA"):
            _set_para_text_preserve_first(el, "c) Mô hình SARIMA(1, 1, 1)(1, 0, 1, 7)")
        elif txt.startswith("Hình 4.6 Kết quả huấn luyện"):
            _set_para_text_preserve_first(el, "Hình 4.6: Kết quả huấn luyện và dự báo mô hình SARIMA")
        elif txt.startswith("Hình 4.7: Mô hình Random Forest"):
            _set_para_text_preserve_first(el, "d) Mô hình Random Forest Regressor")
        elif txt.startswith("Hình 4.7 Kết quả huấn luyện"):
            _set_para_text_preserve_first(el, "Hình 4.7: Kết quả huấn luyện và dự báo mô hình Random Forest")
        elif txt.startswith("Hình 4.8: Mô hình XGBoost"):
            _set_para_text_preserve_first(el, "e) Mô hình XGBoost Regressor")
        elif txt.startswith("Hình 4.8 Kết quả huấn luyện"):
            _set_para_text_preserve_first(el, "Hình 4.8: Kết quả huấn luyện và dự báo mô hình XGBoost")
        elif re.match(r"^Hình 1\.[1-7]\s+[—–-]\s+", txt):
            _set_para_text_preserve_first(el, re.sub(r"^(Hình 1\.[1-7])\s+[—–-]\s+", r"\1: ", txt))
        elif txt.startswith("Bảng 3.1. "):
            _set_para_text_preserve_first(el, txt.replace("Bảng 3.1. ", "Bảng 3.1: ", 1))
        elif txt.startswith("5.8 Bảng tổng hợp các chức năng"):
            _set_para_text_preserve_first(el, "5.8. Bảng tổng hợp các chức năng hệ thống")

        cleaned.append(el)

    # Đảo caption Bảng 4.6 lên ngay phía trên bảng 4.6 (nếu nằm dưới bảng)
    i = 0
    while i < len(cleaned) - 1:
        if cleaned[i].tag == _w("tbl") and cleaned[i + 1].tag == _w("p") and _ptxt(cleaned[i + 1]).strip().startswith("Bảng 4.6:"):
            cleaned[i], cleaned[i + 1] = cleaned[i + 1], cleaned[i]
        i += 1

    # Chèn Bảng 6.1 ở cuối Mục 6.1 (ngay trước 6.2)
    b61_caption = _make_p([("Bảng 6.1: Đối chiếu mục tiêu đề ra và kết quả đạt được của đồ án", True, False, "000000", 26)],
                          align="center", style="Caption", before=160, after=80, line=264, keep_next=True, rpr_sz=26, rpr_bold=True)
    b61_tbl = _make_grid_table(
        headers=["STT", "Mục tiêu nghiên cứu (Chương 1)", "Kết quả đạt được thực tế", "Mức đạt"],
        rows=[
            ["1", "Khám phá dữ liệu bán hàng & tác nhân nhu cầu", "Phân tích trọn vẹn 73.100 giao dịch, 7 biểu đồ EDA, Sơ đồ 1.1 & Bảng 1.1–1.3", "100%"],
            ["2", "Phát triển 5 mô hình dự báo (Thống kê & ML)", "Huấn luyện đồng bộ trên 701 mẫu (560 Train / 141 Test) với 28 đặc trưng", "100%"],
            ["3", "Đánh giá định lượng bằng RMSE, MAE, MAPE", "Random Forest đạt kết quả tối ưu nhất (RMSE 92,42; MAE 73,71; giảm 10,83%)", "100%"],
            ["4", "Đề xuất tối ưu hóa tồn kho theo dự báo ML", "Xác định SS = 311, ROP = 777 (95% SL, L=3) và Ma trận ABC–XYZ (20 SKU)", "100%"],
            ["5", "Triển khai Web Studio & Mô phỏng kinh doanh", "Hoàn chỉnh Retail ML Hub FastAPI (6 màn hình, 23 chức năng, What-If & ROP)", "100%"],
        ],
        col_widths_dxa=[620, 2900, 4456, 1050],
        col_aligns=["center", "left", "left", "center"],
        font_sz=23,
    )
    inserted_b61 = False
    rebuilt_sec2 = []
    for el in cleaned:
        if not inserted_b61 and el.tag == _w("p") and _ptxt(el).strip().startswith("6.2."):
            rebuilt_sec2.append(b61_caption)
            rebuilt_sec2.append(b61_tbl)
            inserted_b61 = True
        rebuilt_sec2.append(el)

    all_els = cover_els + rebuilt_sec2

    # Xóa bookmark cũ ở Section 2 & đánh lại toàn bộ docPr ID + giới hạn chiều rộng ảnh <= 5730000 EMU
    for idx, el in enumerate(all_els):
        if idx >= 19:
            for btag in (_w("bookmarkStart"), _w("bookmarkEnd")):
                for bm in list(el.findall(btag)):
                    el.remove(bm)
        _clamp_drawings_and_renumber_ids(el, ctx)

    # BƯỚC B: Chuẩn hóa kiểu dáng, cỡ chữ, căn lề, khoảng cách cho toàn bộ Section 2
    front_h1_titles = {
        "DANH SÁCH HÌNH ẢNH", "DANH SÁCH BẢNG", "DANH SÁCH THÀNH VIÊN NHÓM",
        "PHÂN CÔNG NHIỆM VỤ", "DANH MỤC TỪ VIẾT TẮT", "LỜI CẢM ƠN", "LỜI CAM KẾT",
    }
    in_chapters = False
    in_refs = False
    style_by_idx = {}
    toc_items = []   # (idx, level, txt, bm)
    lof_items = []   # (idx, txt, bm)
    lot_items = []   # (idx, txt, bm)
    bm_seq = 0
    toc_counter = 245000001

    for idx in range(19, len(all_els)):
        el = all_els[idx]
        if el.tag == _w("tbl"):
            _normalize_existing_table(el)
            continue

        txt = _ptxt(el).strip()
        if txt in ("[[TOC_PLACEHOLDER]]", "[[LOF_PLACEHOLDER]]", "[[LOT_PLACEHOLDER]]"):
            continue
        if txt == "MỤC LỤC":
            _format_para_ppr_and_runs(el, align="center", before=0, after=240, line=288,
                                      keep_next=True, bold=True, sz=32)
            continue

        if txt.startswith("CHƯƠNG 1"):
            in_chapters = True
        if txt == "TÀI LIỆU THAM KHẢO":
            in_refs = True

        # 1) Heading 1
        if txt in front_h1_titles or re.match(r"^CHƯƠNG\s+[1-6]\b", txt) or txt == "TÀI LIỆU THAM KHẢO":
            _format_para_ppr_and_runs(
                el, align="center", style="Heading1", before=0, after=240, line=288,
                keep_next=True, keep_lines=True, page_break_before=True, bold=True, sz=28,
            )
            bm = f"_Toc{toc_counter}"
            toc_counter += 1
            _wrap_bookmark(el, bm_seq, bm)
            bm_seq += 1
            style_by_idx[idx] = "Heading1"
            toc_items.append((idx, 1, txt, bm))
            continue

        # 2) Heading 3 (kiểm tra 3 số trước 2 số!)
        if in_chapters and not in_refs and re.match(r"^\d+\.\d+\.\d+\.?\s+", txt) and len(txt) < 180 and not re.match(r"^(Hình|Sơ đồ|Bảng)\s+", txt):
            _format_para_ppr_and_runs(
                el, align="left", style="Heading3", before=160, after=80, line=288,
                keep_next=True, keep_lines=True, bold=True, sz=26,
            )
            bm = f"_Toc{toc_counter}"
            toc_counter += 1
            _wrap_bookmark(el, bm_seq, bm)
            bm_seq += 1
            style_by_idx[idx] = "Heading3"
            toc_items.append((idx, 3, txt, bm))
            continue

        # 3) Heading 2
        if in_chapters and not in_refs and re.match(r"^\d+\.\d+\.?\s+", txt) and len(txt) < 180 and not re.match(r"^(Hình|Sơ đồ|Bảng)\s+", txt):
            _format_para_ppr_and_runs(
                el, align="left", style="Heading2", before=240, after=120, line=288,
                keep_next=True, keep_lines=True, bold=True, sz=26,
            )
            bm = f"_Toc{toc_counter}"
            toc_counter += 1
            _wrap_bookmark(el, bm_seq, bm)
            bm_seq += 1
            style_by_idx[idx] = "Heading2"
            toc_items.append((idx, 2, txt, bm))
            continue

        # 4) Chú thích hình / sơ đồ -> TIeuDeanh (36 mục)
        if in_chapters and re.match(r"^(Hình|Sơ đồ)\s+\d+\.\d+:", txt) and len(txt) < 240:
            _format_para_ppr_and_runs(
                el, align="center", style="TIeuDeanh", before=80, after=160, line=264,
                bold=True, sz=26, strip_italic=True,
            )
            bm = f"_Toc{toc_counter}"
            toc_counter += 1
            _wrap_bookmark(el, bm_seq, bm)
            bm_seq += 1
            style_by_idx[idx] = "TIeuDeanh"
            lof_items.append((idx, txt, bm))
            continue

        # 5) Chú thích bảng -> Caption (13 mục)
        if in_chapters and re.match(r"^Bảng\s+\d+\.\d+:", txt) and len(txt) < 240:
            _format_para_ppr_and_runs(
                el, align="center", style="Caption", before=160, after=80, line=264,
                keep_next=True, bold=True, sz=26, strip_italic=True,
            )
            bm = f"_Toc{toc_counter}"
            toc_counter += 1
            _wrap_bookmark(el, bm_seq, bm)
            bm_seq += 1
            style_by_idx[idx] = "Caption"
            lot_items.append((idx, txt, bm))
            continue

        # 6) Đoạn chứa hình vẽ
        if el.findall(f".//{_w('drawing')}"):
            next_is_fig_cap = (
                idx + 1 < len(all_els)
                and all_els[idx + 1].tag == _w("p")
                and bool(re.match(r"^(Hình|Sơ đồ)\s+\d+\.\d+:", _ptxt(all_els[idx + 1]).strip()))
            )
            _format_para_ppr_and_runs(
                el, align="center", before=120, after=60, line=240,
                keep_next=next_is_fig_cap, sz=26,
            )
            continue

        # 7) Đoạn công thức độc lập hoặc chứa toán
        maths = el.findall(f".//{{{M_NS}}}oMath")
        if maths:
            pure_txt = "".join(t.text or "" for t in el.findall(f"./{_w('r')}/{_w('t')}")).strip()
            if len(pure_txt) <= 2:
                _format_para_ppr_and_runs(
                    el, align="center", before=120, after=120, line=288,
                    sz=26, preserve_run_bold=True,
                )
                continue

        # 8) Chữ ký phải (Nhóm thực hiện)
        if txt == "Nhóm thực hiện":
            _format_para_ppr_and_runs(el, align="right", before=240, after=120, line=288, bold=True, sz=26)
            continue

        # 9) Tiêu đề Gantt (không đưa vào TIeuDeanh)
        if txt.startswith("Biểu đồ Gantt kế hoạch"):
            _format_para_ppr_and_runs(el, align="center", before=180, after=60, line=264, keep_next=True, bold=True, sz=26)
            continue

        # 10) Tiểu mục hạ cấp a., b., c. (Chương 3) & a)..e) (Chương 4.3.3)
        if re.match(r"^[a-e][\.\)]\s+", txt) and len(txt) < 140:
            _format_para_ppr_and_runs(
                el, align="left", style="Normal", before=140, after=80, line=288,
                keep_next=True, bold=True, sz=26,
            )
            continue

        # 11) Tài liệu tham khảo [1]..[14]
        if in_refs and re.match(r"^\[\d+\]\s+", txt):
            _format_para_ppr_and_runs(
                el, align="both", style="Normal", before=60, after=80, line=288,
                left=480, hanging=480, bold=None, sz=26, preserve_run_bold=True,
            )
            continue

        # 12) Danh sách bullet hoặc có numPr gốc
        old_ppr = el.find(_w("pPr"))
        has_numpr = old_ppr is not None and old_ppr.find(_w("numPr")) is not None
        if has_numpr:
            _format_para_ppr_and_runs(
                el, align="both", style="Normal", before=80, after=80, line=312,
                bold=None, sz=26, preserve_run_bold=True, preserve_numpr=True,
            )
            continue
        if txt.startswith("•") or re.match(r"^\d+\.\s+", txt):
            _format_para_ppr_and_runs(
                el, align="both", style="Normal", before=80, after=80, line=312,
                left=360, hanging=220, bold=None, sz=26, preserve_run_bold=True,
            )
            continue

        # 13) Văn bản học thuật tiêu chuẩn (Times New Roman 13pt, both, firstLine=720, line=1.3)
        _format_para_ppr_and_runs(
            el, align="both", style="Normal", before=120, after=120, line=312,
            first_line=720, bold=None, sz=26, preserve_run_bold=True,
        )

    # BƯỚC C: Ước lượng số trang 2 lượt và mở rộng 3 field Mục lục / Hình ảnh / Bảng
    page_map = _compute_page_map(all_els, style_by_idx, toc_n=len(toc_items), lof_n=len(lof_items), lot_n=len(lot_items))
    toc_entries = [(txt, bm, page_map.get(idx, 2), lvl) for idx, lvl, txt, bm in toc_items]
    lof_entries = [(txt, bm, page_map.get(idx, 10), 1) for idx, txt, bm in lof_items]
    lot_entries = [(txt, bm, page_map.get(idx, 12), 1) for idx, txt, bm in lot_items]

    final_elements = []
    for el in all_els:
        if el.tag == _w("p"):
            t = _ptxt(el).strip()
            if t == "[[TOC_PLACEHOLDER]]":
                final_elements.extend(_build_field_list_paragraphs('TOC \\o "1-3" \\h \\z \\u', toc_entries, is_main_toc=True))
                continue
            if t == "[[LOF_PLACEHOLDER]]":
                final_elements.extend(_build_field_list_paragraphs('TOC \\h \\z \\t "TIeuDeanh" \\c', lof_entries, default_style="TableofFigures", is_main_toc=False))
                continue
            if t == "[[LOT_PLACEHOLDER]]":
                final_elements.extend(_build_field_list_paragraphs('TOC \\h \\z \\t "Caption" \\c', lot_entries, default_style="TableofFigures", is_main_toc=False))
                continue
        final_elements.append(el)

    # BƯỚC D: Chuẩn hóa <w:sectPr> cuối Section 2 (footer đánh số trang tự động, A4, margins 1440)
    if sect_pr is not None:
        for ch in list(sect_pr):
            sect_pr.remove(ch)
        fref = ET.SubElement(sect_pr, _w("footerReference"))
        fref.set(_w("type"), "default")
        fref.set(_r("id"), footer_rid)
        pgsz = ET.SubElement(sect_pr, _w("pgSz"))
        pgsz.set(_w("w"), "11909")
        pgsz.set(_w("h"), "16834")
        pgsz.set(_w("orient"), "portrait")
        pgmar = ET.SubElement(sect_pr, _w("pgMar"))
        for k, v in (("top", "1440"), ("right", "1440"), ("bottom", "1440"), ("left", "1440"),
                     ("header", "720"), ("footer", "720"), ("gutter", "0")):
            pgmar.set(_w(k), v)
        ET.SubElement(sect_pr, _w("cols")).set(_w("space"), "720")

    return final_elements, {
        "toc_count": len(toc_entries),
        "lof_count": len(lof_entries),
        "lot_count": len(lot_entries),
        "total_tables": sum(1 for el in final_elements if el.tag == _w("tbl")),
    }
