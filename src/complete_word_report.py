"""
Complete Word Report Generator: complete_word_report.py
Hoàn thiện toàn bộ các phần còn thiếu, sửa các lỗi/mâu thuẫn số liệu và đồng bộ 100%
file Word D:/HK7/Machine_Learning/doan/MLBA.docx với file Excel & Mã nguồn/Web Studio:
  - Luôn sao lưu bản gốc tại D:/HK7/Machine_Learning/doan/MLBA_backup_origin.docx (idempotent).
  - Bổ sung Trang bìa chuẩn học thuật UEL, Lời cảm ơn & Lời cam đoan, Bảng PCCV + Đánh giá
    đóng góp + Lịch trình Gantt (20/10/2026 – 15/11/2026), Mục lục & Danh mục Bảng/Hình.
  - Bổ sung Tiêu đề cấp Chương (Heading1): CHƯƠNG 1 → CHƯƠNG 6 & TÀI LIỆU THAM KHẢO.
  - Hoàn thiện Chương 1 (Sửa T00, chèn Mục 1.2 + Sơ đồ 1.1 EDA + Bảng 1.2, đồng bộ Bảng 1.3).
  - Hoàn thiện Chương 2 (Bổ sung FastAPI ở 2.1, chèn Mục 2.3 RMSE/MAE/MAPE & Mục 2.4 SS/ROP/ABC-XYZ).
  - Hoàn thiện Chương 3 (Đánh số trích dẫn thống nhất [4]–[11], bổ sung SARIMA, gom URL về Tài liệu tham khảo).
  - Hoàn thiện Chương 4 (Chèn Mục 4.1 + Sơ đồ 4.1 + Hình 4.1–4.3 + Bảng 4.1 ADF/KPSS,
    chèn Mục 4.2 + Bảng 4.2 28 đặc trưng, đồng bộ chia tập 560 Train / 141 Test ML vs 731 gốc,
    thay ảnh 300 DPI cho Hình 4.9–4.10, thêm chú thích Hình 4.10, chèn Mục 4.5 + Hình 4.11–4.12 + Bảng 4.7).
  - Hoàn thiện Chương 5 (Chèn 3 Sơ đồ BPMN 5.1–5.3 ở Mục 5.1.2, thêm đủ 12 chú thích Hình 5.1–5.12,
    sửa 4 chỗ lệch số liệu so với ảnh chụp Web Studio, thêm tiêu đề Bảng 5.1).
  - Hoàn thiện Chương 6 (Đồng bộ SS = 311, ROP = 777 & gom 14 Tài liệu tham khảo chuẩn IEEE).
"""

import csv
import io
import json
import os
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET
from PIL import Image

try:
    from .word_uel_formatter import (
        build_uel_cover_and_front_matter,
        finalize_uel_document,
        generate_gantt_png,
        sync_uel_styles_and_resources,
    )
except ImportError:
    from word_uel_formatter import (
        build_uel_cover_and_front_matter,
        finalize_uel_document,
        generate_gantt_png,
        sync_uel_styles_and_resources,
    )

if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")

# ---------------------------------------------------------------------------
# ĐƯỜNG DẪN TÀI NGUYÊN DỰ ÁN
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PARENT_DOAN_DIR = os.path.dirname(BASE_DIR)
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
FIGURES_DIR = os.path.join(REPORTS_DIR, "figures")
SAVED_MODELS_DIR = os.path.join(BASE_DIR, "saved_models")

TARGET_DOCX = os.path.join(PARENT_DOAN_DIR, "MLBA.docx")
BACKUP_DOCX = os.path.join(PARENT_DOAN_DIR, "MLBA_backup_origin.docx")
COPY_REPORTS_DOCX = os.path.join(REPORTS_DIR, "MLBA.docx")

# OOXML Namespaces
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

PREFIX_MAP = {
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "r": R_NS,
    "m": M_NS,
    "v": "urn:schemas-microsoft-com:vml",
    "wp": WP_NS,
    "w10": "urn:schemas-microsoft-com:office:word",
    "w": W_NS,
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "a": A_NS,
    "pic": PIC_NS,
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "lc": "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "w14": W14_NS,
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "cr": "http://schemas.microsoft.com/office/comments/2020/reactions",
}
for pfx, uri in PREFIX_MAP.items():
    ET.register_namespace(pfx, uri)


def wtag(local):
    return f"{{{W_NS}}}{local}"


def rtag(local):
    return f"{{{R_NS}}}{local}"


def wptag(local):
    return f"{{{WP_NS}}}{local}"


def atag(local):
    return f"{{{A_NS}}}{local}"


def pictag(local):
    return f"{{{PIC_NS}}}{local}"


def p_text(el):
    return "".join(t.text or "" for t in el.findall(f".//{wtag('t')}"))


def tbl_text(el):
    return "".join(t.text or "" for t in el.findall(f".//{wtag('t')}"))


# ---------------------------------------------------------------------------
# NGỮ CẢNH QUẢN LÝ MEDIA & RELATIONSHIPS TRONG DOCX
# ---------------------------------------------------------------------------
class DocxMediaContext:
    def __init__(self, zip_files, rels_root):
        self.zip_files = zip_files
        self.rels_root = rels_root
        self.rel_map = {}
        max_rid = 0
        for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
            rid = rel.attrib["Id"]
            self.rel_map[rid] = rel.attrib["Target"]
            m = re.match(r"rId(\d+)$", rid)
            if m:
                max_rid = max(max_rid, int(m.group(1)))
        self.next_rid = max_rid + 1

        max_img = 0
        for name in zip_files:
            m = re.match(r"word/media/image(\d+)\.png$", name)
            if m:
                max_img = max(max_img, int(m.group(1)))
        self.next_img_idx = max_img + 1
        self.next_docpr_id = 1000

    def register_new_png(self, png_path):
        if not os.path.exists(png_path):
            raise FileNotFoundError(f"Không tìm thấy tệp hình ảnh: {png_path}")
        with open(png_path, "rb") as f:
            data = f.read()
        with Image.open(io.BytesIO(data)) as im:
            w_px, h_px = im.size

        img_name = f"image{self.next_img_idx}.png"
        self.next_img_idx += 1
        zip_member = f"word/media/{img_name}"
        self.zip_files[zip_member] = data

        rid = f"rId{self.next_rid}"
        self.next_rid += 1
        rel_el = ET.SubElement(self.rels_root, f"{{{REL_NS}}}Relationship")
        rel_el.set("Id", rid)
        rel_el.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        rel_el.set("Target", f"media/{img_name}")
        self.rel_map[rid] = f"media/{img_name}"
        return rid, img_name, w_px, h_px

    def replace_existing_png(self, rid, new_png_path):
        target = self.rel_map.get(rid)
        if not target:
            raise KeyError(f"Không tìm thấy relationship {rid}")
        if not os.path.exists(new_png_path):
            raise FileNotFoundError(f"Không tìm thấy ảnh thay thế: {new_png_path}")
        with open(new_png_path, "rb") as f:
            data = f.read()
        with Image.open(io.BytesIO(data)) as im:
            w_px, h_px = im.size
        zip_member = f"word/{target}" if not target.startswith("word/") else target
        self.zip_files[zip_member] = data
        return w_px, h_px

    def next_shape_id(self):
        self.next_docpr_id += 1
        return str(self.next_docpr_id)


# ---------------------------------------------------------------------------
# HÀM TẠO PHẦN TỬ OOXML CHUẨN HỌC THUẬT (TIMES NEW ROMAN)
# ---------------------------------------------------------------------------
def _make_rpr(font_name="Times New Roman", sz=26, bold=False, italic=False, color=None, underline=False):
    rpr = ET.Element(wtag("rPr"))
    rf = ET.SubElement(rpr, wtag("rFonts"))
    for k in ("ascii", "hAnsi", "cs", "eastAsia"):
        rf.set(wtag(k), font_name)
    if bold:
        ET.SubElement(rpr, wtag("b")).set(wtag("val"), "1")
        ET.SubElement(rpr, wtag("bCs")).set(wtag("val"), "1")
    if italic:
        ET.SubElement(rpr, wtag("i")).set(wtag("val"), "1")
        ET.SubElement(rpr, wtag("iCs")).set(wtag("val"), "1")
    if underline:
        ET.SubElement(rpr, wtag("u")).set(wtag("val"), "single")
    if color:
        ET.SubElement(rpr, wtag("color")).set(wtag("val"), color)
    ET.SubElement(rpr, wtag("sz")).set(wtag("val"), str(sz))
    ET.SubElement(rpr, wtag("szCs")).set(wtag("val"), str(sz))
    return rpr


def make_p(
    runs,
    align="both",
    space_before=0,
    space_after=120,
    line=288,
    first_line=720,
    left_indent=0,
    p_style=None,
    keep_next=False,
    page_break_before=False,
    border_bottom=None,
    border_box=None,
    shading_fill=None,
    default_sz=26,
    default_color=None,
):
    """Tạo đoạn văn <w:p>. `runs` có thể là str hoặc list các tuple `(text, bold, italic, color, sz)`."""
    p = ET.Element(wtag("p"))
    ppr = ET.SubElement(p, wtag("pPr"))
    if p_style:
        ET.SubElement(ppr, wtag("pStyle")).set(wtag("val"), p_style)
    if keep_next:
        ET.SubElement(ppr, wtag("keepNext")).set(wtag("val"), "1")
        ET.SubElement(ppr, wtag("keepLines")).set(wtag("val"), "1")
    if page_break_before:
        ET.SubElement(ppr, wtag("pageBreakBefore")).set(wtag("val"), "1")
    if border_bottom or border_box:
        pbdr = ET.SubElement(ppr, wtag("pBdr"))
        if border_box:
            for side in ("top", "left", "bottom", "right"):
                b = ET.SubElement(pbdr, wtag(side))
                b.set(wtag("val"), "single")
                b.set(wtag("sz"), str(border_box.get("sz", 12)))
                b.set(wtag("space"), str(border_box.get("space", 12)))
                b.set(wtag("color"), border_box.get("color", "1E3A8A"))
        elif border_bottom:
            b = ET.SubElement(pbdr, wtag("bottom"))
            b.set(wtag("val"), "single")
            b.set(wtag("sz"), str(border_bottom.get("sz", 10)))
            b.set(wtag("space"), str(border_bottom.get("space", 4)))
            b.set(wtag("color"), border_bottom.get("color", "1E3A8A"))
    if shading_fill:
        shd = ET.SubElement(ppr, wtag("shd"))
        shd.set(wtag("val"), "clear")
        shd.set(wtag("color"), "auto")
        shd.set(wtag("fill"), shading_fill)

    sp = ET.SubElement(ppr, wtag("spacing"))
    sp.set(wtag("before"), str(space_before))
    sp.set(wtag("after"), str(space_after))
    sp.set(wtag("line"), str(line))
    sp.set(wtag("lineRule"), "auto")

    if first_line or left_indent:
        ind = ET.SubElement(ppr, wtag("ind"))
        if left_indent:
            ind.set(wtag("left"), str(left_indent))
        if first_line:
            ind.set(wtag("firstLine"), str(first_line))
    if align:
        ET.SubElement(ppr, wtag("jc")).set(wtag("val"), align)

    if isinstance(runs, str):
        runs = [(runs, False, False, default_color, default_sz)]

    for item in runs:
        if isinstance(item, str):
            text, bold, italic, color, sz = item, False, False, default_color, default_sz
        else:
            text = item[0]
            bold = item[1] if len(item) > 1 else False
            italic = item[2] if len(item) > 2 else False
            color = item[3] if len(item) > 3 and item[3] is not None else default_color
            sz = item[4] if len(item) > 4 and item[4] is not None else default_sz
        r = ET.SubElement(p, wtag("r"))
        r.append(_make_rpr(sz=sz, bold=bold, italic=italic, color=color))
        t = ET.SubElement(r, wtag("t"))
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
    return p


def make_page_break():
    p = ET.Element(wtag("p"))
    r = ET.SubElement(p, wtag("r"))
    br = ET.SubElement(r, wtag("br"))
    br.set(wtag("type"), "page")
    return p


def make_chapter_heading(title_text, page_break=True):
    return make_p(
        [(title_text, True, False, "1E3A8A", 32)],
        align="center",
        space_before=120,
        space_after=240,
        line=288,
        first_line=0,
        p_style="Heading1",
        keep_next=True,
        page_break_before=page_break,
        border_bottom={"sz": 12, "space": 6, "color": "1E3A8A"},
    )


def make_h2(title_text):
    return make_p(
        [(title_text, True, False, "1D4ED8", 28)],
        align="left",
        space_before=240,
        space_after=100,
        line=276,
        first_line=0,
        p_style="Heading2",
        keep_next=True,
    )


def make_h3(title_text):
    return make_p(
        [(title_text, True, False, "0F172A", 26)],
        align="left",
        space_before=160,
        space_after=80,
        line=276,
        first_line=0,
        p_style="Heading3",
        keep_next=True,
    )


def make_caption(text, is_table=False):
    return make_p(
        [(text, True, not is_table, "1E293B", 24)],
        align="center",
        space_before=120 if is_table else 60,
        space_after=60 if is_table else 180,
        line=264,
        first_line=0,
        keep_next=is_table,
    )


def make_bullet(bold_prefix, body_text):
    runs = [("•  ", True, False, "1E3A8A", 26)]
    if bold_prefix:
        runs.append((bold_prefix, True, False, "0F172A", 26))
    runs.append((body_text, False, False, "0F172A", 26))
    return make_p(
        runs,
        align="both",
        space_before=40,
        space_after=80,
        line=288,
        first_line=0,
        left_indent=540,
    )


def make_image_block(png_path, caption_text, ctx, max_w_emu=5600000, max_h_emu=6000000):
    rid, img_name, w_px, h_px = ctx.register_new_png(png_path)
    scale = min(max_w_emu / float(w_px), max_h_emu / float(h_px))
    cx = str(int(round(w_px * scale)))
    cy = str(int(round(h_px * scale)))
    docpr_id = ctx.next_shape_id()

    p = ET.Element(wtag("p"))
    ppr = ET.SubElement(p, wtag("pPr"))
    ET.SubElement(ppr, wtag("keepNext")).set(wtag("val"), "1")
    sp = ET.SubElement(ppr, wtag("spacing"))
    sp.set(wtag("before"), "120")
    sp.set(wtag("after"), "60")
    sp.set(wtag("line"), "240")
    sp.set(wtag("lineRule"), "auto")
    ET.SubElement(ppr, wtag("jc")).set(wtag("val"), "center")

    r = ET.SubElement(p, wtag("r"))
    drawing = ET.SubElement(r, wtag("drawing"))
    inline = ET.SubElement(drawing, wptag("inline"))
    for k, v in (("distT", "0"), ("distB", "0"), ("distL", "0"), ("distR", "0")):
        inline.set(k, v)
    ext = ET.SubElement(inline, wptag("extent"))
    ext.set("cx", cx)
    ext.set("cy", cy)
    eff = ET.SubElement(inline, wptag("effectExtent"))
    for k in ("l", "t", "r", "b"):
        eff.set(k, "0")
    docpr = ET.SubElement(inline, wptag("docPr"))
    docpr.set("id", docpr_id)
    docpr.set("name", img_name)

    graphic = ET.SubElement(inline, atag("graphic"))
    gdata = ET.SubElement(graphic, atag("graphicData"))
    gdata.set("uri", PIC_NS)
    pic = ET.SubElement(gdata, pictag("pic"))

    nv = ET.SubElement(pic, pictag("nvPicPr"))
    cnv = ET.SubElement(nv, pictag("cNvPr"))
    cnv.set("id", "0")
    cnv.set("name", img_name)
    cnv_pic = ET.SubElement(nv, pictag("cNvPicPr"))
    cnv_pic.set("preferRelativeResize", "0")

    bf = ET.SubElement(pic, pictag("blipFill"))
    blip = ET.SubElement(bf, atag("blip"))
    blip.set(rtag("embed"), rid)
    sr = ET.SubElement(bf, atag("srcRect"))
    for k in ("b", "l", "r", "t"):
        sr.set(k, "0")
    stretch = ET.SubElement(bf, atag("stretch"))
    ET.SubElement(stretch, atag("fillRect"))

    sppr = ET.SubElement(pic, pictag("spPr"))
    xfrm = ET.SubElement(sppr, atag("xfrm"))
    off = ET.SubElement(xfrm, atag("off"))
    off.set("x", "0")
    off.set("y", "0")
    aext = ET.SubElement(xfrm, atag("ext"))
    aext.set("cx", cx)
    aext.set("cy", cy)
    prst = ET.SubElement(sppr, atag("prstGeom"))
    prst.set("prst", "rect")
    ET.SubElement(prst, atag("avLst"))

    return [p, make_caption(caption_text, is_table=False)]


def make_table(headers, rows, col_widths_dxa, aligns=None, bold_cols=None, font_sz=22, header_fill="1E3A8A"):
    """Tạo bảng chuẩn học thuật với tiêu đề lặp qua trang, viền gọn & căn chỉnh rõ ràng."""
    bold_set = set(bold_cols or [])
    total_w = sum(col_widths_dxa)

    tbl = ET.Element(wtag("tbl"))
    tblPr = ET.SubElement(tbl, wtag("tblPr"))
    tblW = ET.SubElement(tblPr, wtag("tblW"))
    tblW.set(wtag("w"), str(total_w))
    tblW.set(wtag("type"), "dxa")
    ET.SubElement(tblPr, wtag("jc")).set(wtag("val"), "center")

    bdrs = ET.SubElement(tblPr, wtag("tblBorders"))
    for side, sz, col in (
        ("top", "8", header_fill),
        ("bottom", "8", header_fill),
        ("left", "4", "CBD5E1"),
        ("right", "4", "CBD5E1"),
        ("insideH", "4", "CBD5E1"),
        ("insideV", "4", "CBD5E1"),
    ):
        b = ET.SubElement(bdrs, wtag(side))
        b.set(wtag("val"), "single")
        b.set(wtag("sz"), sz)
        b.set(wtag("space"), "0")
        b.set(wtag("color"), col)

    ET.SubElement(tblPr, wtag("tblLayout")).set(wtag("type"), "fixed")

    grid = ET.SubElement(tbl, wtag("tblGrid"))
    for cw in col_widths_dxa:
        gc = ET.SubElement(grid, wtag("gridCol"))
        gc.set(wtag("w"), str(cw))

    def _add_row(cells, is_header=False, fill_hex="FFFFFF"):
        tr = ET.SubElement(tbl, wtag("tr"))
        trPr = ET.SubElement(tr, wtag("trPr"))
        ET.SubElement(trPr, wtag("cantSplit"))
        if is_header:
            ET.SubElement(trPr, wtag("tblHeader"))

        for c_i, val in enumerate(cells):
            tc = ET.SubElement(tr, wtag("tc"))
            tcPr = ET.SubElement(tc, wtag("tcPr"))
            tcW = ET.SubElement(tcPr, wtag("tcW"))
            tcW.set(wtag("w"), str(col_widths_dxa[c_i]))
            tcW.set(wtag("type"), "dxa")
            shd = ET.SubElement(tcPr, wtag("shd"))
            shd.set(wtag("val"), "clear")
            shd.set(wtag("color"), "auto")
            shd.set(wtag("fill"), header_fill if is_header else fill_hex)

            mar = ET.SubElement(tcPr, wtag("tcMar"))
            for mside, mw in (("top", "70"), ("bottom", "70"), ("left", "100"), ("right", "100")):
                m_el = ET.SubElement(mar, wtag(mside))
                m_el.set(wtag("w"), mw)
                m_el.set(wtag("type"), "dxa")
            ET.SubElement(tcPr, wtag("vAlign")).set(wtag("val"), "center")

            jc_val = "center" if is_header else (aligns[c_i] if aligns and c_i < len(aligns) else "left")
            is_b = True if is_header else (c_i in bold_set)
            color = "FFFFFF" if is_header else "0F172A"
            p = make_p(
                [(str(val), is_b, False, color, font_sz)],
                align=jc_val,
                space_before=0,
                space_after=0,
                line=252,
                first_line=0,
            )
            tc.append(p)

    _add_row(headers, is_header=True)
    for r_i, row_vals in enumerate(rows):
        row_fill = "EFF6FF" if (len(rows) > 1 and r_i == len(rows) - 1 and str(row_vals[0]).startswith("TỔNG")) else ("F8FAFC" if r_i % 2 == 1 else "FFFFFF")
        _add_row(row_vals, is_header=False, fill_hex=row_fill)

    return tbl


# ---------------------------------------------------------------------------
# HÀM CHỈNH SỬA IN-PLACE PHẦN TỬ GỐC (BẢO TOÀN STYLE & CÔNG THỨC OMML)
# ---------------------------------------------------------------------------
def replace_cell_text(tbl_el, row_idx, col_idx, new_text):
    rows = tbl_el.findall(wtag("tr"))
    cells = rows[row_idx].findall(wtag("tc"))
    tc = cells[col_idx]
    ps = tc.findall(wtag("p"))
    if not ps:
        return
    p0 = ps[0]
    ts = p0.findall(f".//{wtag('t')}")
    if ts:
        ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        ts[0].text = new_text
        for extra_t in ts[1:]:
            extra_t.text = ""
    else:
        p0.append(make_p(new_text, first_line=0, space_after=0).find(wtag("r")))


def replace_para_text(p_el, new_text, preserve_first_rpr=True):
    """Thay toàn bộ nội dung văn bản của đoạn văn gốc nhưng giữ nguyên pPr và rPr đầu tiên."""
    ppr = p_el.find(wtag("pPr"))
    first_rpr = None
    first_r = p_el.find(wtag("r"))
    if preserve_first_rpr and first_r is not None:
        rpr = first_r.find(wtag("rPr"))
        if rpr is not None:
            first_rpr = ET.fromstring(ET.tostring(rpr, encoding="unicode"))

    for child in list(p_el):
        if child is not ppr:
            p_el.remove(child)

    r = ET.SubElement(p_el, wtag("r"))
    if first_rpr is not None:
        r.append(first_rpr)
    t = ET.SubElement(r, wtag("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = new_text


def replace_para_substrings(p_el, replacements):
    """Thay thế từng chuỗi con trong các thẻ <w:t> của đoạn văn gốc (giữ nguyên toàn bộ OMML <m:oMath>)."""
    for t in p_el.findall(f".//{wtag('t')}"):
        if t.text:
            txt = t.text
            for old, new in replacements:
                txt = txt.replace(old, new)
            t.text = txt
    # Xử lý trường hợp từ khóa bị tách qua nhiều <w:t>
    full = p_text(p_el)
    for old, new in replacements:
        if old in full:
            combined = full
            for o, n in replacements:
                combined = combined.replace(o, n)
            if not p_el.findall(f".//{{{M_NS}}}oMath"):
                replace_para_text(p_el, combined)
            break


def replace_image_in_para(p_el, new_png_path, ctx, max_w_emu=5731200, max_h_emu=4600000):
    blip = p_el.find(f".//{atag('blip')}")
    if blip is None:
        return
    rid = blip.attrib.get(rtag("embed"))
    w_px, h_px = ctx.replace_existing_png(rid, new_png_path)
    scale = min(max_w_emu / float(w_px), max_h_emu / float(h_px))
    cx = str(int(round(w_px * scale)))
    cy = str(int(round(h_px * scale)))
    for ext in p_el.findall(f".//{wptag('extent')}") + p_el.findall(f".//{atag('ext')}"):
        ext.set("cx", cx)
        ext.set("cy", cy)


# ---------------------------------------------------------------------------
# DỰNG NỘI DUNG BỔ SUNG TỪNG PHẦN (FRONT MATTER & CHƯƠNG 1..6 & TÀI LIỆU TK)
# ---------------------------------------------------------------------------
def build_front_matter(ctx):
    elems = []

    # 1. TRANG BÌA CHUẨN HỌC THUẬT UEL
    elems.append(make_p([("ĐẠI HỌC QUỐC GIA THÀNH PHỐ HỒ CHÍ MINH", True, False, "0F172A", 28)], align="center", space_before=60, space_after=40, first_line=0))
    elems.append(make_p([("TRƯỜNG ĐẠI HỌC KINH TẾ - LUẬT", True, False, "0F172A", 28)], align="center", space_before=0, space_after=40, first_line=0))
    elems.append(make_p([("KHOA HỆ THỐNG THÔNG TIN", True, False, "1E3A8A", 28)], align="center", space_before=0, space_after=560, first_line=0, border_bottom={"sz": 8, "space": 6, "color": "1E3A8A"}))

    elems.append(make_p([("BÁO CÁO ĐỒ ÁN MÔN HỌC", True, False, "1E3A8A", 36)], align="center", space_before=480, space_after=120, first_line=0))
    elems.append(make_p([("HỌC MÁY TRONG PHÂN TÍCH KINH DOANH", True, False, "0F172A", 30)], align="center", space_before=0, space_after=60, first_line=0))
    elems.append(make_p([("(MACHINE LEARNING IN BUSINESS ANALYTICS)", False, True, "334155", 26)], align="center", space_before=0, space_after=480, first_line=0))

    elems.append(
        make_p(
            [
                ("ĐỀ TÀI:\n", True, False, "1E3A8A", 28),
                ("ỨNG DỤNG HỌC MÁY TRONG DỰ BÁO NHU CẦU BÁN HÀNG\nVÀ TỐI ƯU HÓA TỒN KHO CHO CHUỖI CỬA HÀNG BÁN LẺ\n(RETAIL ML HUB STUDIO)", True, False, "0F172A", 30),
            ],
            align="center",
            space_before=120,
            space_after=540,
            line=320,
            first_line=0,
            border_box={"sz": 14, "space": 14, "color": "1E3A8A"},
            shading_fill="F8FAFC",
        )
    )

    info_rows = [
        ("Giảng viên hướng dẫn:", "TS. Trần Duy Thanh"),
        ("Môn học:", "Học máy trong Phân tích Kinh doanh (MLBA)"),
        ("Thời gian thực hiện:", "20/10/2026 – 15/11/2026 (27 ngày cao điểm)"),
        ("Nhóm thực hiện:", "Khang (Nhóm trưởng), Hiếu, Minh, Ân, Lợi"),
    ]
    elems.append(make_table(["Thông tin đồ án", "Chi tiết"], info_rows, [3200, 5800], aligns=["left", "left"], bold_cols=[0], font_sz=24))

    elems.append(make_p([("Thành phố Hồ Chí Minh, Tháng 11 năm 2026", True, False, "1E3A8A", 26)], align="center", space_before=840, space_after=60, first_line=0))
    elems.append(make_page_break())

    # 2. LỜI CẢM ƠN & LỜI CAM ĐOAN
    elems.append(make_chapter_heading("LỜI CẢM ƠN VÀ LỜI CAM ĐOAN", page_break=False))
    elems.append(make_h2("Lời cảm ơn"))
    elems.append(make_p(
        "Trong suốt quá trình học tập học phần Học máy trong Phân tích Kinh doanh và thực hiện đề tài “Ứng dụng Học máy trong dự báo nhu cầu bán hàng và tối ưu hóa tồn kho cho chuỗi cửa hàng bán lẻ (Retail ML Hub Studio)”, nhóm nghiên cứu đã nhận được sự hướng dẫn khoa học, tận tâm và những góp ý chuyên môn sâu sắc từ TS. Trần Duy Thanh."
    ))
    elems.append(make_p(
        "Thầy không chỉ truyền đạt nền tảng kiến thức bài bản về các thuật toán phân tích chuỗi thời gian, học máy ứng dụng và quy trình xây dựng hệ thống hỗ trợ ra quyết định, mà còn định hướng cho nhóm cách tiếp cận bài toán thực tiễn một cách chặt chẽ từ khâu khám phá dữ liệu (EDA), kiểm định thống kê, huấn luyện đối chứng mô hình cho đến chuyển hóa kết quả dự báo thành quyết định quản trị chuỗi cung ứng cụ thể."
    ))
    elems.append(make_p(
        "Nhóm xin trân trọng gửi lời cảm ơn chân thành đến TS. Trần Duy Thanh cùng quý Thầy, Cô Khoa Hệ thống Thông tin — Trường Đại học Kinh tế - Luật (ĐHQG TP.HCM) đã tạo điều kiện học thuật thuận lợi để nhóm hoàn thành trọn vẹn đồ án này."
    ))

    elems.append(make_h2("Lời cam đoan"))
    elems.append(make_p(
        "Nhóm tác giả xin cam đoan toàn bộ nội dung báo cáo, mã nguồn xây dựng mô hình học máy, hệ thống FastAPI & giao diện Retail ML Hub Studio, bộ sơ đồ lưu đồ nghiệp vụ và các kết quả thực nghiệm trình bày trong đồ án là công trình nghiên cứu trung thực do tập thể 5 thành viên trực tiếp triển khai và kiểm chứng. Các số liệu, tài liệu khoa học trích dẫn đều được dẫn nguồn đầy đủ, rõ ràng; mọi sai số thực nghiệm (kể cả các hạn chế như chỉ số MAPE cao trên chuỗi nhiễu mạnh) đều được nhóm trình bày khách quan đúng theo kết quả chạy thực tế."
    ))
    elems.append(make_page_break())

    # 3. BẢNG PHÂN CÔNG CÔNG VIỆC, ĐÁNH GIÁ ĐÓNG GÓP & TIẾN ĐỘ GANTT (ĐỒNG BỘ EXCEL)
    elems.append(make_chapter_heading("PHÂN CÔNG CÔNG VIỆC, ĐÓNG GÓP & TIẾN ĐỘ ĐỒ ÁN", page_break=False))
    elems.append(make_h2("1. Bảng phân công công việc và Đánh giá mức độ đóng góp (Contribution Evaluation)"))
    elems.append(make_p(
        "Nhóm gồm 5 thành viên phối hợp triển khai đồng bộ theo quy trình khép kín từ thu thập – khám phá dữ liệu, xây dựng mô hình dự báo học máy đến lập trình ứng dụng Web Studio và biên soạn báo cáo. Tỷ lệ đóng góp được toàn nhóm thống nhất đánh giá minh bạch như trong Bảng P.1 (khớp 100% với Sheet 2_PCCV và Sheet 5_Contribution_Work trong file Excel tổng hợp)."
    ))

    pccv_rows = [
        ("1", "Khang", "Nhóm trưởng & Kiến trúc dữ liệu", "Chương 1 (Dữ liệu, Bảng 1.1–1.3, Hình 1.1–1.7), lập Gantt (20/10–15/11), biên tập toàn báo cáo.", "Tích cực, điều phối sát sao tiến độ và rà soát chéo số liệu.", "20% (100%)"),
        ("2", "Hiếu", "Trưởng nhóm ML & Backend API", "Chương 4 (ADF/KPSS, 28 đặc trưng, huấn luyện & lưu 5 mô hình .pkl) và Backend FastAPI.", "Chủ động kỹ thuật, xử lý tốt pipeline ML và RAM Cache API.", "20% (100%)"),
        ("3", "Minh", "Phương pháp & Tối ưu Tồn kho", "Chương 2 (Mục 2.1–2.4), thuật toán SS/ROP, cảnh báo 3 mức, ABC–XYZ & mô phỏng What-If.", "Trách nhiệm, chuẩn hóa chặt chẽ công thức chuỗi cung ứng.", "20% (100%)"),
        ("4", "Ân", "Nghiên cứu tài liệu & Tổng kết", "Chương 3 (Tổng quan 11+ nghiên cứu), Chương 6 (6.1–6.4) và 14 Tài liệu tham khảo chuẩn IEEE.", "Cẩn thận, nhất quán trong tổng hợp học thuật và slide.", "20% (100%)"),
        ("5", "Lợi", "Giao diện Web & Sơ đồ Flowchart", "Chương 5 (6 màn hình Web Studio offline ApexCharts) & bộ 6 tệp Flowchart/BPMN kèm icon (.drawio).", "Sáng tạo, tỉ mỉ trong thiết kế UI/UX và sơ đồ nghiệp vụ.", "20% (100%)"),
        ("TỔNG", "5 TV", "Toàn bộ đồ án Retail ML Hub", "6 Chương báo cáo, 5 mô hình .pkl, 5 sơ đồ Flowchart/BPMN & Web App Studio + File Excel 6 Sheet.", "Đồng thuận 100% toàn nhóm.", "100%"),
    ]
    elems.append(make_caption("Bảng P.1: Bảng phân công nhiệm vụ và đánh giá mức độ đóng góp của 5 thành viên", is_table=True))
    elems.append(make_table(
        ["STT", "Thành viên", "Vai trò chính", "Nhiệm vụ & Sản phẩm bàn giao", "Thái độ hợp tác", "Đóng góp"],
        pccv_rows,
        [600, 1000, 1500, 2900, 1900, 1100],
        aligns=["center", "center", "left", "left", "left", "center"],
        bold_cols=[0, 1, 5],
        font_sz=20,
    ))

    elems.append(make_h2("2. Bảng tiến độ thực hiện đồ án (20/10/2026 – 15/11/2026)"))
    elems.append(make_p(
        "Dự án được nhóm khởi động từ cuối tháng 10 (ngày 20/10/2026) và hoàn tất đóng gói vào ngày 15/11/2026 (tổng thời gian 27 ngày lịch liên tục, chia thành 6 giai đoạn và 10 chặng mốc M1–M10 như trình bày chi tiết tại Sheet 3_Gantt_Chart trong file Excel)."
    ))
    gantt_rows = [
        ("GĐ1", "Khởi tạo đề tài & Khám phá EDA (Ch.1)", "20/10 – 25/10", "M1 – M2", "Khang", "Chốt 73.100 dòng dữ liệu sạch, Bảng 1.1–1.3, 7 biểu đồ EDA & Sơ đồ 1.1.", "Hoàn thành"),
        ("GĐ2", "Cơ sở lý thuyết & Lược khảo (Ch.2–3)", "23/10 – 31/10", "M2 – M4", "Minh, Ân", "Hoàn thành Chương 2 (5 mô hình, metrics, SS/ROP, ABC–XYZ) & Chương 3.", "Hoàn thành"),
        ("GĐ3", "Chẩn đoán, 28 Đặc trưng & 5 Mô hình (Ch.4)", "28/10 – 07/11", "M3 – M7", "Hiếu", "Kiểm định ADF/KPSS, 28 features, huấn luyện 5 mô hình (RF RMSE 92,42).", "Hoàn thành"),
        ("GĐ4", "Tối ưu hóa Tồn kho & ABC–XYZ (Ch.4)", "05/11 – 09/11", "M6 – M7", "Minh", "Tính SS = 311, ROP = 777, ma trận ABC–XYZ (S001) & chính sách What-If.", "Hoàn thành"),
        ("GĐ5", "Web Studio & Bộ Flowchart/BPMN (Ch.5)", "06/11 – 13/11", "M6 – M9", "Lợi, Hiếu", "FastAPI + 6 màn hình Web offline + Bộ sơ đồ Draw.io/PNG kèm icon.", "Hoàn thành"),
        ("GĐ6", "Kiểm thử tích hợp, Báo cáo & Slide (Ch.6)", "11/11 – 15/11", "M8 – M10", "Cả nhóm", "Đồng bộ hoàn chỉnh Báo cáo Word MLBA.docx + File Excel 6 Sheet + Slide.", "Hoàn thành"),
        ("TỔNG", "Toàn dự án (18 đầu việc WBS)", "20/10 – 15/11", "27 ngày", "5 Thành viên", "Đạt 100% tất cả các mốc bàn giao đúng tiến độ đề ra.", "100%"),
    ]
    elems.append(make_caption("Bảng P.2: Tóm tắt tiến độ 6 giai đoạn thực hiện đồ án từ 20/10/2026 đến 15/11/2026", is_table=True))
    elems.append(make_table(
        ["GĐ", "Giai đoạn thực hiện", "Thời gian", "Mốc Gantt", "Chủ trì", "Sản phẩm bàn giao trọng tâm", "Trạng thái"],
        gantt_rows,
        [600, 2100, 1150, 950, 1000, 2300, 900],
        aligns=["center", "left", "center", "center", "center", "left", "center"],
        bold_cols=[0, 4, 6],
        font_sz=20,
    ))
    elems.append(make_page_break())

    # 4. MỤC LỤC TỔNG QUÁT & DANH MỤC BẢNG / HÌNH
    elems.append(make_chapter_heading("MỤC LỤC VÀ DANH MỤC BẢNG BIỂU — HÌNH ẢNH", page_break=False))
    elems.append(make_h2("Mục lục tổng quát"))
    toc_lines = [
        ("PHẦN MỞ ĐẦU", "Lời cảm ơn, Lời cam đoan, Phân công công việc & Tiến độ thực hiện (20/10 – 15/11/2026)"),
        ("CHƯƠNG 1: ", "TỔNG QUAN ĐỀ TÀI VÀ KHÁM PHÁ DỮ LIỆU (1.1 Đặt vấn đề — 1.2 Mô tả 15 thuộc tính & Quy trình EDA — 1.3 Khám phá 7 biểu đồ EDA)"),
        ("CHƯƠNG 2: ", "CƠ SỞ LÝ THUYẾT VÀ PHƯƠNG PHÁP NGHIÊN CỨU (2.1 Công nghệ phát triển — 2.2 Thuật toán dự báo — 2.3 Chỉ số RMSE/MAE/MAPE — 2.4 Lý thuyết SS, ROP & ABC–XYZ)"),
        ("CHƯƠNG 3: ", "LƯỢC KHẢO CÁC CÔNG TRÌNH NGHIÊN CỨU LIÊN QUAN (3.1 Tổng quan nghiên cứu quốc tế, trong nước & khoảng trống nghiên cứu)"),
        ("CHƯƠNG 4: ", "TRIỂN KHAI THỰC NGHIỆM, ĐÁNH GIÁ MÔ HÌNH VÀ TỐI ƯU HÓA TỒN KHO (4.1 Chẩn đoán chuỗi thời gian — 4.2 Kỹ thuật 28 đặc trưng — 4.3 Huấn luyện 5 mô hình — 4.4 Thảo luận định lượng — 4.5 Tối ưu hóa Tồn kho SS/ROP & ABC–XYZ)"),
        ("CHƯƠNG 5: ", "XÂY DỰNG HỆ THỐNG HỖ TRỢ RA QUYẾT ĐỊNH RETAIL ML HUB STUDIO (5.1 Kiến trúc & 3 Sơ đồ BPMN — 5.2 đến 5.7 Thuyết minh 6 màn hình Web — 5.8 Bảng tổng hợp 23 chức năng)"),
        ("CHƯƠNG 6: ", "KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN (6.1 Tổng kết — 6.2 Ý nghĩa — 6.3 Hướng phát triển — 6.4 Kiến nghị)"),
        ("TÀI LIỆU THAM KHẢO", "Danh mục 14 công trình nghiên cứu và nguồn dữ liệu trích dẫn chuẩn IEEE"),
    ]
    for prefix, desc in toc_lines:
        elems.append(make_p([(prefix, True, False, "1E3A8A", 24), (desc, False, False, "0F172A", 24)], align="left", space_before=30, space_after=60, first_line=0, left_indent=360))

    elems.append(make_h2("Danh mục Bảng biểu chính"))
    tbl_list = [
        "Bảng P.1 – P.2: Bảng phân công công việc, đánh giá đóng góp và lịch trình Gantt (20/10 – 15/11/2026)",
        "Bảng 1.1: Đối tượng và phạm vi nghiên cứu (5 cửa hàng × 20 SKU × 731 ngày = 73.100 bản ghi)",
        "Bảng 1.2: Mô tả chi tiết 15 thuộc tính trong tập dữ liệu retail_store_inventory.csv",
        "Bảng 1.3: Tổng hợp các phát hiện chính từ 7 biểu đồ khám phá dữ liệu (EDA)",
        "Bảng 3.1: Tổng hợp các nghiên cứu tiêu biểu về dự báo nhu cầu bán lẻ",
        "Bảng 4.1: Kết quả kiểm định tính dừng ADF và KPSS trên chuỗi nhu cầu S001–P0001",
        "Bảng 4.2: Danh sách 28 biến đặc trưng đầu vào cho mô hình học máy",
        "Bảng 4.3: Chi tiết phân bổ tập huấn luyện và kiểm thử (731 ngày gốc & 701 mẫu ML sau lag_30)",
        "Bảng 4.4: Các bước triển khai thực nghiệm huấn luyện 5 mô hình",
        "Bảng 4.5: Đối sánh định lượng hiệu năng dự báo của 5 mô hình trên tập kiểm thử",
        "Bảng 4.6: Mức độ cải thiện RMSE theo từng nhóm phương pháp so với SARIMA",
        "Bảng 4.7: Tổng hợp phân loại chiến lược tồn kho ABC–XYZ cho 20 sản phẩm tại cửa hàng S001",
        "Bảng 5.1: Bảng tổng hợp 23 chức năng của hệ thống Retail ML Hub Studio",
    ]
    for item in tbl_list:
        elems.append(make_p([("•  " + item, False, False, "1E293B", 22)], align="left", space_before=20, space_after=40, first_line=0, left_indent=360))

    elems.append(make_h2("Danh mục Sơ đồ và Hình ảnh chính"))
    fig_list = [
        "Sơ đồ 1.1 & Hình 1.1 – Hình 1.7 (Chương 1): Quy trình EDA 8 bước đa nhánh và 7 biểu đồ khám phá dữ liệu bán lẻ",
        "Sơ đồ 4.1 & Hình 4.1 – Hình 4.3 (Chương 4): Quy trình huấn luyện 10 bước, phân rã chuỗi, ACF/PACF và thống kê trượt",
        "Hình 4.4 – Hình 4.10 (Chương 4): Đồ thị dự báo 5 mô hình, biểu đồ so sánh RMSE/MAE/MAPE và Feature Importance",
        "Hình 4.11 – Hình 4.12 (Chương 4): Mô phỏng tồn kho theo ngưỡng SS/ROP và Ma trận nhiệt phân loại ABC–XYZ",
        "Sơ đồ 5.1 – Sơ đồ 5.3 & Hình 5.1 – Hình 5.12 (Chương 5): 3 Sơ đồ BPMN kiến trúc hệ thống, 6 giao diện Web & 6 lưu đồ màn hình",
    ]
    for item in fig_list:
        elems.append(make_p([("•  " + item, False, False, "1E293B", 22)], align="left", space_before=20, space_after=40, first_line=0, left_indent=360))

    elems.append(make_page_break())
    return elems


def build_section_1_2(ctx):
    elems = []
    elems.append(make_h2("1.2. Mô tả bộ dữ liệu và Quy trình khám phá dữ liệu (EDA)"))
    elems.append(make_p(
        "Bộ dữ liệu thực nghiệm được sử dụng trong đồ án là tập giao dịch bán lẻ và quản trị tồn kho theo ngày (Dataset/retail_store_inventory.csv), gồm tổng cộng 73.100 bản ghi hoàn chỉnh kéo dài trong 731 ngày liên tục từ ngày 01/01/2022 đến ngày 01/01/2024 (tương ứng 24 tháng). Dữ liệu bao phủ 5 cửa hàng (S001 – S005) thuộc 4 khu vực địa lý (North, South, East, West) với danh mục 20 mã sản phẩm (P0001 – P0020) trải rộng trên 5 ngành hàng (Clothing, Electronics, Furniture, Groceries, Toys). Mỗi ngày có đúng 100 cặp Cửa hàng × Sản phẩm được ghi nhận đầy đủ; kết quả kiểm định chất lượng cho thấy 0 giá trị bị khuyết (Null/NaN) trên toàn bộ 15 cột thuộc tính."
    ))
    elems.append(make_p(
        "Quy trình chuẩn hóa dữ liệu gốc và phân tích khám phá đa chiều (Exploratory Data Analysis – EDA) được chuẩn hóa thành 8 bước đa nhánh như minh họa tại Sơ đồ 1.1: từ nạp tệp CSV và kiểm tra khuyết/ép kiểu ngày (B1–B2), xác định biến mục tiêu Units Sold và loại trừ các trường gây rò rỉ dữ liệu như Demand Forecast hay Units Ordered (B3), rẽ thành 4 nhánh phân tích song song gồm xu hướng chuỗi thời gian – cơ cấu ngành hàng/khu vực – tác động ngoại cảnh thời tiết/khuyến mãi – độ co giãn giá & ma trận tương quan (B4–B7), trước khi tổng hợp hệ thống giả thuyết làm đầu vào cho bước thiết kế 28 đặc trưng ở Chương 4 (B8)."
    ))

    eda_flow = os.path.join(FIGURES_DIR, "flowcharts", "chuong_1_flowchart_quy_trinh_kham_pha_du_lieu_eda.png")
    elems.extend(make_image_block(
        eda_flow,
        "Sơ đồ 1.1: Quy trình thu thập, làm sạch dữ liệu và khám phá đa chiều (EDA — 8 bước)",
        ctx,
        max_w_emu=5500000,
        max_h_emu=5000000,
    ))

    elems.append(make_p(
        "Bảng 1.2 trình bày đặc tả kỹ thuật chi tiết của tất cả 15 trường thuộc tính trong bộ dữ liệu gốc, kèm miền giá trị thực tế và vai trò xử lý trước khi đưa vào mô hình dự báo."
    ))
    csv_11 = os.path.join(REPORTS_DIR, "table_1_1_mo_ta_15_cot_du_lieu.csv")
    rows_12 = []
    with open(csv_11, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            rows_12.append([
                row["STT"],
                row["Tên cột (Column)"],
                row["Kiểu dữ liệu"],
                row["Miền giá trị / Đơn vị"],
                row["Ý nghĩa nghiệp vụ"],
            ])
    elems.append(make_caption("Bảng 1.2: Mô tả chi tiết 15 thuộc tính trong tập dữ liệu retail_store_inventory.csv", is_table=True))
    elems.append(make_table(
        ["STT", "Tên thuộc tính", "Kiểu dữ liệu", "Miền giá trị / Đơn vị", "Ý nghĩa nghiệp vụ & Vai trò mô hình"],
        rows_12,
        [550, 1550, 1650, 2100, 3150],
        aligns=["center", "left", "left", "left", "left"],
        bold_cols=[0, 1],
        font_sz=20,
    ))
    return elems


def build_chapter2_additions():
    elems = []
    # Mục 2.3
    elems.append(make_h2("2.3. Các chỉ số đánh giá độ chính xác của mô hình dự báo"))
    elems.append(make_p(
        "Để đối chiếu định lượng năng lực dự báo giữa nhóm phương pháp chuỗi thời gian thống kê cổ điển và nhóm học máy trên cùng tập kiểm thử ngoại mẫu gồm n quan sát (y_i là sản lượng bán thực tế và ŷ_i là giá trị mô hình dự đoán tại ngày i), nghiên cứu áp dụng đồng thời bộ ba thước đo chuẩn:"
    ))
    elems.append(make_bullet(
        "Root Mean Squared Error (RMSE): ",
        "RMSE = √[(1/n) Σ (y_i − ŷ_i)²]. Thước đo căn bậc hai của trung bình bình phương sai số, có cùng đơn vị đo với biến mục tiêu (sản phẩm/ngày). Nhờ phép bình phương độ lệch, RMSE phạt nặng các sai số dự báo cực lớn (đặc biệt tại các ngày cao điểm đột biến), giúp bảo đảm an toàn tồn kho trước rủi ro đứt gãy nguồn cung. Do đó, RMSE được chọn làm tiêu chí xếp hạng chính."
    ))
    elems.append(make_bullet(
        "Mean Absolute Error (MAE): ",
        "MAE = (1/n) Σ |y_i − ŷ_i|. Phản ánh độ lệch tuyệt đối trung bình giữa dự báo và thực tế mỗi ngày, ít nhạy cảm với giá trị ngoại lai hơn RMSE và mang lại diễn giải trực quan về quy mô sai số đơn vị hàng hóa trong lập kế hoạch kho vận."
    ))
    elems.append(make_bullet(
        "Mean Absolute Percentage Error (MAPE %): ",
        "MAPE = (100% / n) Σ |(y_i − ŷ_i) / y_i| (tính trên tập các ngày có y_i ≠ 0). Cung cấp góc nhìn sai số tương đối theo tỷ lệ phần trăm. Tuy nhiên, đối với chuỗi bán lẻ có những ngày sức mua chạm đáy gần bằng 0, mẫu số y_i rất nhỏ sẽ khuếch đại MAPE lên rất cao; vì vậy MAPE được dùng như chỉ số tham khảo bổ sung bên cạnh cặp RMSE và MAE."
    ))

    # Mục 2.4
    elems.append(make_h2("2.4. Cơ sở lý thuyết về Quản trị và Tối ưu hóa Tồn kho"))
    elems.append(make_p(
        "Giá trị thực tiễn của dự báo nhu cầu chỉ phát huy đầy đủ khi kết quả đầu ra được chuyển hóa tự động thành định mức tồn kho và quyết định đặt hàng hợp lý nhằm cân bằng giữa mục tiêu đáp ứng khách hàng (Service Level) và chi phí vốn lưu động. Hệ thống Retail ML Hub triển khai hai công cụ quản trị tồn kho chuẩn mực sau:"
    ))
    elems.append(make_bullet(
        "Mức tồn kho an toàn (Safety Stock — SS): ",
        "Đóng vai trò lượng dự trữ đệm chống đỡ sự biến động bất thường của sức mua trong khoảng thời gian chờ hàng từ nhà cung cấp (Lead Time L). Công thức chuẩn giả định phân phối nhu cầu: SS = ⌈ Z_α × σ_D × √L ⌉, trong đó Z_α là hệ số phân vị chuẩn ứng với mức độ phục vụ mục tiêu α (ví dụ α = 95% tương ứng Z ≈ 1,645), σ_D là độ lệch chuẩn nhu cầu hàng ngày trong 90 ngày gần nhất và L là thời gian giao hàng (mặc định L = 3 ngày)."
    ))
    elems.append(make_bullet(
        "Điểm đặt hàng lại (Reorder Point — ROP) và Đề xuất đặt hàng (Q): ",
        "Xác định ngưỡng tồn kho kích hoạt phát lệnh nhập bổ sung hàng mới: ROP = ⌈ D_hat × L + SS ⌉, với D_hat là nhu cầu tiêu thụ dự kiến mỗi ngày đầu ra từ mô hình học máy tối ưu (hoặc trung bình lịch sử). Dựa trên mức tồn kho thực tế I_curr, hệ thống phân chia 3 ngưỡng cảnh báo: Nguy cấp nếu I_curr ≤ ROP; Cảnh báo nếu ROP < I_curr ≤ 1,2 × ROP; An toàn nếu I_curr > 1,2 × ROP. Khi ở trạng thái Nguy cấp, lượng đặt hàng bổ sung được tính theo mức tồn mục tiêu (1,5 × ROP): Q = max(0, ⌈ 1,5 × ROP − I_curr ⌉)."
    ))
    elems.append(make_bullet(
        "Ma trận phân loại chiến lược ABC – XYZ: ",
        "Kết hợp phân tích Pareto theo đóng góp doanh thu thuần (Nhóm A chiếm 80% doanh thu tích lũy đầu tiên, Nhóm B chiếm 15% tiếp theo, Nhóm C chiếm 5% còn lại) với hệ số biến thiên nhu cầu CV = σ_D / μ_D (Nhóm X ổn định: CV ≤ 0,40; Nhóm Y dao động vừa: 0,40 < CV ≤ 0,70; Nhóm Z biến động mạnh: CV > 0,70). Ma trận 9 ô (AX → CZ) giúp nhà quản trị phân bổ chính sách kiểm soát tồn kho và tần suất rà soát đơn hàng phù hợp cho từng mặt hàng."
    ))
    return elems


def build_chapter4_prefix(ctx):
    """Mục 4.1 (Sơ đồ 4.1 + Hình 4.1–4.3 + Bảng 4.1 ADF/KPSS) & Mục 4.2 (Bảng 4.2 28 Features)."""
    elems = []
    elems.append(make_h2("4.1. Quy trình thực nghiệm và Phân tích chẩn đoán chuỗi thời gian"))
    elems.append(make_h3("4.1.1. Tổng quan quy trình huấn luyện 5 mô hình và tối ưu hóa tồn kho"))
    elems.append(make_p(
        "Toàn bộ thực nghiệm định lượng được thiết kế chuẩn hóa theo quy trình khép kín gồm 10 bước như thể hiện trong Sơ đồ 4.1. Từ tập dữ liệu sạch của chuỗi đại diện Store S001 – Product P0001 (tần suất ngày liên tục 731 ngày), hệ thống tiến hành kiểm định chẩn đoán thống kê và phân rã thành phần chuỗi thời gian, trích xuất ma trận 28 đặc trưng đa chiều và chia theo đúng trật tự thời gian 80:20. Tiếp theo, luồng xử lý rẽ làm hai nhánh đối chứng song song: Nhánh A gồm 3 mô hình chuỗi thời gian cổ điển (ARIMA, Holt-Winters, SARIMA) và Nhánh B gồm 2 mô hình học máy đa biến (Random Forest, XGBoost). Mô hình đạt sai số RMSE thấp nhất được lưu định dạng .pkl để cấp đầu vào nhu cầu dự báo cho mô-đun tối ưu hóa tồn kho (SS, ROP, ma trận ABC–XYZ) và tích hợp lên Web Studio."
    ))

    ch4_flow = os.path.join(FIGURES_DIR, "flowcharts", "chuong_4_flowchart_quy_trinh_huan_luyen_5_mo_hinh.png")
    elems.extend(make_image_block(
        ch4_flow,
        "Sơ đồ 4.1: Quy trình huấn luyện đối chứng 5 mô hình dự báo và chuyển giao tối ưu tồn kho (10 bước)",
        ctx,
        max_w_emu=5500000,
        max_h_emu=5000000,
    ))

    elems.append(make_h3("4.1.2. Phân rã chuỗi thời gian, Tự tương quan và Kiểm định tính dừng (ADF & KPSS)"))
    elems.append(make_p(
        "Trước khi thiết lập bậc tham số cho các mô hình chuỗi thời gian và xác định cấu trúc độ trễ cho mô hình học máy, chuỗi sản lượng tiêu thụ hàng ngày Units Sold của SKU S001 – P0001 (731 quan sát từ 01/01/2022 đến 01/01/2024) được phân rã theo mô hình cộng tính (Seasonal Decomposition) với chu kỳ tuần (period = 7), khảo sát hàm tự tương quan ACF/PACF đến độ trễ 40 ngày và kiểm chứng bằng hai phép thử ADF – KPSS (triển khai trong src/time_series_diagnostics.py)."
    ))

    h41 = os.path.join(FIGURES_DIR, "diagnostics", "chuong_4_hinh_4_1_phan_ra_chuoi_thoi_gian.png")
    elems.extend(make_image_block(
        h41,
        "Hình 4.1: Phân rã chuỗi thời gian doanh số SKU S001–P0001 theo mô hình cộng tính (Observed – Trend – Seasonal – Residual)",
        ctx,
        max_w_emu=5600000,
        max_h_emu=4800000,
    ))

    h42 = os.path.join(FIGURES_DIR, "diagnostics", "chuong_4_hinh_4_2_tu_tuong_quan_acf_pacf.png")
    elems.extend(make_image_block(
        h42,
        "Hình 4.2: Đồ thị hàm tự tương quan (ACF) và tự tương quan riêng phần (PACF) với độ trễ 40 ngày",
        ctx,
        max_w_emu=5600000,
        max_h_emu=3800000,
    ))

    h43 = os.path.join(FIGURES_DIR, "diagnostics", "chuong_4_hinh_4_3_trung_binh_do_lech_truot.png")
    elems.extend(make_image_block(
        h43,
        "Hình 4.3: Biến động trung bình trượt và độ lệch chuẩn trượt (cửa sổ 7 ngày) của chuỗi S001–P0001",
        ctx,
        max_w_emu=5600000,
        max_h_emu=3800000,
    ))

    # Bảng 4.1 ADF & KPSS đọc từ CSV
    st_csv = os.path.join(REPORTS_DIR, "stationarity_tests_s001_p0001.csv")
    st_rows = []
    with open(st_csv, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            p_str = "< 0,0001 (0,0000)" if float(r["p_value"]) == 0.0 else f"≥ {float(r['p_value']):.4f}".replace(".", ",")
            st_rows.append([
                r["Test"],
                f"{float(r['Statistic']):.4f}".replace(".", ","),
                p_str,
                r["Lags_Used"],
                r["Observations"],
                f"{r['Critical_1pct']} / {r['Critical_5pct']} / {r['Critical_10pct']}".replace(".", ","),
                r["Conclusion"],
            ])
    elems.append(make_caption("Bảng 4.1: Kết quả kiểm định tính dừng ADF và KPSS cho chuỗi nhu cầu S001 – P0001", is_table=True))
    elems.append(make_table(
        ["Phép kiểm định", "Statistic", "p-value", "Lags", "Quan sát", "Ngưỡng 1% / 5% / 10%", "Kết luận (α = 0,05)"],
        st_rows,
        [1550, 850, 1050, 550, 750, 1750, 2500],
        aligns=["left", "right", "center", "center", "center", "center", "left"],
        bold_cols=[0, 1, 2],
        font_sz=19,
    ))

    elems.append(make_p(
        "Kết quả phân tích Hình 4.1 – Hình 4.3 và Bảng 4.1 mang lại hai căn cứ thống kê quan trọng: (1) Cả hai phép thử ADF (Statistic = -27,2824 < -3,4393 tại mức 1%, p < 0,0001, bác bỏ H0 có nghiệm đơn vị) và KPSS (Statistic = 0,0506 < 0,4630 tại mức 5%, p ≥ 0,1000, không bác bỏ H0 chuỗi dừng) đồng thuận khẳng định chuỗi nhu cầu gốc có tính dừng vững chắc quanh giá trị kỳ vọng ổn định; (2) Mặc dù dừng về trung bình, độ lệch chuẩn dao động hàng ngày vẫn rất cao (quan sát phần dư Residual ở Hình 4.1 và biên độ độ lệch chuẩn trượt ở Hình 4.3). Hiện tượng này chứng tỏ chỉ dựa vào các độ trễ tuyến tính đơn biến trên ACF/PACF (Hình 4.2) là chưa đủ, mà đòi hỏi kết hợp các biến ngoại sinh về tồn kho, giá bán, thời tiết và mùa vụ thông qua các thuật toán học máy."
    ))

    # Mục 4.2 Feature Engineering
    elems.append(make_h2("4.2. Kỹ thuật tạo đặc trưng đa chiều (Feature Engineering — 28 biến đầu vào)"))
    elems.append(make_p(
        "Để giúp hai mô hình học máy có giám sát (Random Forest Regressor và XGBoost Regressor) khai thác đồng thời động lực tự hồi quy theo thời gian và tác động ngoại cảnh tại điểm bán, hệ thống xây dựng bộ biến đổi đặc trưng trong mô-đun src/data_loader.py (hàm engineer_features) tạo ra 28 biến đầu vào thuộc 5 nhóm: (i) 5 biến gốc tại điểm bán; (ii) 5 biến lịch thời gian; (iii) 6 biến độ trễ lag_1, lag_2, lag_7, lag_14, lag_21, lag_30; (iv) 4 biến thống kê trượt (rolling_mean_7/14, rolling_std_7/14 được dịch lùi 1 ngày bằng shift(1) nhằm chống rò rỉ dữ liệu tương lai); và (v) 8 biến tương tác giá & mã hóa One-Hot thời tiết, mùa vụ (Bảng 4.2). Hai trường Units Ordered và Demand Forecast bị loại hoàn toàn khỏi ma trận X."
    ))

    feat_csv = os.path.join(REPORTS_DIR, "table_4_1_danh_sach_28_bien_dac_trung.csv")
    feat_rows = []
    with open(feat_csv, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            feat_rows.append([
                r["STT"],
                r["Tên biến đặc trưng (Feature)"],
                r["Nhóm biến"],
                r["Công thức / Cách tạo"],
                r["Vai trò đối với mô hình dự báo"],
            ])
    elems.append(make_caption("Bảng 4.2: Danh sách 28 biến đặc trưng đầu vào (Features) cho mô hình học máy", is_table=True))
    elems.append(make_table(
        ["STT", "Tên biến (Feature)", "Nhóm biến", "Công thức / Cách tạo", "Vai trò đối với dự báo nhu cầu"],
        feat_rows,
        [520, 1880, 1450, 2150, 3000],
        aligns=["center", "left", "left", "left", "left"],
        bold_cols=[0, 1],
        font_sz=19,
    ))
    return elems


def build_section_4_5(ctx):
    """Mục 4.5: Ứng dụng kết quả dự báo vào Tối ưu hóa Tồn kho (SS, ROP & ABC-XYZ)."""
    elems = []
    elems.append(make_h2("4.5. Ứng dụng kết quả dự báo vào Tối ưu hóa Tồn kho (SS, ROP & Ma trận ABC–XYZ)"))
    elems.append(make_p(
        "Sau khi lựa chọn được mô hình tối ưu nhất là Random Forest Regressor (RMSE = 92,42, MAE = 73,71, lưu tại saved_models/best_forecasting_model.pkl), đầu ra nhu cầu dự báo theo ngày D_hat được truyền trực tiếp sang động cơ tối ưu tồn kho (src/inventory_optimizer.py) để tính toán hai tham số vận hành thiết yếu với thời gian giao hàng mặc định L = 3 ngày và mức độ phục vụ mục tiêu Service Level = 95%:"
    ))
    elems.append(make_bullet(
        "Chế độ tích hợp dự báo ML Random Forest trên Web Studio (Màn hình 4 & 6): ",
        "Với chuỗi SKU S001 – P0001 có nhu cầu dự báo trung bình 7 ngày tới D_hat = 155,1 đơn vị/ngày, độ lệch chuẩn 90 ngày gần nhất σ_D = 109,15 đơn vị và phân vị Z = 1,645 (ứng với Service Level 95%), hệ thống tính ra Mức tồn kho an toàn SS = ⌈ 1,645 × 109,15 × √3 ⌉ = 311 đơn vị và Điểm đặt hàng lại ROP = ⌈ 155,1 × 3 + 311 ⌉ = 777 đơn vị. So với tồn kho cuối kỳ thực tế đạt 223 đơn vị (thấp hơn rất nhiều so với ngưỡng 777), sản phẩm rơi vào trạng thái Nguy cấp và cần phát lệnh đặt hàng bổ sung đề xuất Q = ⌈ 1,5 × 777 − 223 ⌉ = 942 đơn vị."
    ))
    elems.append(make_bullet(
        "Chế độ tham chiếu cơ sở theo trung bình lịch sử 90 ngày (Hình 4.11): ",
        "Khi dùng tham số cầu bình quân lịch sử và hệ số làm tròn Z = 1,65 trong cấu hình tĩnh (src/config.py), ngưỡng an toàn đạt SS = 312 đơn vị và điểm đặt hàng lại đạt ROP = 707 đơn vị. Đồ thị mô phỏng 60 ngày gần nhất ở Hình 4.11 cho thấy mức tồn kho thực tế của S001 – P0001 dao động quanh 55 – 480 đơn vị, nằm trọn trong vùng rủi ro thiếu hàng (dưới ngưỡng ROP), khẳng định tính cấp thiết của việc nâng tần suất và quy mô đặt hàng."
    ))

    h411 = os.path.join(FIGURES_DIR, "inventory", "chuong_4_hinh_4_11_mo_phong_ton_kho_rop.png")
    elems.extend(make_image_block(
        h411,
        "Hình 4.11: Mô phỏng biến động tồn kho thực tế 60 ngày so với ngưỡng Đặt hàng lại (ROP = 707 lịch sử / 777 ML) và Tồn an toàn (SS = 312 / 311) của SKU S001–P0001",
        ctx,
        max_w_emu=5600000,
        max_h_emu=3800000,
    ))

    # Đọc trực tiếp kết quả phân nhóm ABC-XYZ từ reports/abc_xyz_analysis_s001.csv
    abc_csv = os.path.join(REPORTS_DIR, "abc_xyz_analysis_s001.csv")
    groups = {"AZ": [], "BZ": [], "CZ": []}
    all_rev = 0.0
    all_cvs = []
    with open(abc_csv, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            cls = r["ABC_XYZ"]
            rev = float(r["Total_Revenue"])
            cv = float(r["CV"])
            all_rev += rev
            all_cvs.append(cv)
            groups.setdefault(cls, []).append((r["Product ID"], r["Category"], rev, cv))

    elems.append(make_p(
        f"Đồng thời, hệ thống quét toàn bộ 20 mã sản phẩm tại cửa hàng S001 trong 731 ngày (tổng doanh thu đạt ${all_rev:,.0f}) để lập Ma trận phân loại chiến lược ABC – XYZ (Hình 4.12 và Bảng 4.7). Do đặc thù nhu cầu tiêu dùng biến động mạnh giữa các ngày thường và ngày lễ/khuyến mãi, hệ số biến thiên CV của cả 20 mã đều nằm trong khoảng {min(all_cvs):.3f} – {max(all_cvs):.3f} (> 0,70), khiến toàn bộ danh mục thuộc cột Z (Biến động mạnh) và phân bổ vào 3 ô chiến lược AZ (15 SKU), BZ (3 SKU) và CZ (2 SKU)."
    ))

    h412 = os.path.join(FIGURES_DIR, "inventory", "chuong_4_hinh_4_12_ma_tran_abc_xyz.png")
    elems.extend(make_image_block(
        h412,
        "Hình 4.12: Ma trận phân loại tồn kho chiến lược ABC – XYZ cho 20 sản phẩm tại cửa hàng S001",
        ctx,
        max_w_emu=4600000,
        max_h_emu=4200000,
    ))

    policy_map = {
        "AZ": "Giá trị doanh thu rất cao nhưng nhu cầu thất thường: ưu tiên dự báo ML hàng ngày, duy trì SS = 95–99% và đặt hàng linh hoạt theo ROP.",
        "BZ": "Giá trị trung bình, biến động cao: kết hợp dự báo Machine Learning theo tuần và đặt hàng định kỳ có kiểm soát trần tồn kho.",
        "CZ": "Tỷ trọng doanh thu thấp, biến động lớn: tránh ôm hàng dài hạn; đặt lô nhỏ khi chạm ngưỡng ROP hoặc gom đơn từ nhà cung cấp.",
    }
    abc_rows = []
    for cls in ("AZ", "BZ", "CZ"):
        items = groups.get(cls, [])
        share = sum(x[2] for x in items) / all_rev * 100.0
        cvs = [x[3] for x in items]
        skus = ", ".join(x[0] for x in items[:6]) + (f", ... (+{len(items)-6})" if len(items) > 6 else "")
        abc_rows.append([
            cls,
            f"{len(items)} SKU ({len(items)*5}%)",
            f"{share:.2f}%".replace(".", ","),
            f"{min(cvs):.3f} – {max(cvs):.3f}".replace(".", ","),
            skus,
            policy_map[cls],
        ])
    abc_rows.append(["TỔNG", "20 SKU (100%)", "100,00%", f"{min(all_cvs):.3f} – {max(all_cvs):.3f}".replace(".", ","), "P0001 – P0020", "100% mã tại S001 tập trung cột Z đòi hỏi theo dõi ROP sát sao."])

    elems.append(make_caption("Bảng 4.7: Tổng hợp kết quả phân loại chiến lược tồn kho ABC–XYZ tại cửa hàng S001", is_table=True))
    elems.append(make_table(
        ["Nhóm", "Số lượng SKU", "Tỷ trọng DT", "Khoảng CV", "Mã sản phẩm tiêu biểu", "Khuyến nghị chính sách quản trị tồn kho"],
        abc_rows,
        [700, 1100, 1000, 1150, 1950, 3100],
        aligns=["center", "center", "right", "center", "left", "left"],
        bold_cols=[0, 2],
        font_sz=19,
    ))
    return elems


def build_references_section():
    elems = []
    elems.append(make_chapter_heading("TÀI LIỆU THAM KHẢO", page_break=True))
    elems.append(make_p(
        "Danh mục các công trình nghiên cứu khoa học, báo cáo thống kê chính thức và tài liệu chuẩn mực học thuật được sử dụng làm cơ sở tham chiếu và đối chứng xuyên suốt từ Chương 1 đến Chương 6:",
        space_after=160,
        first_line=0,
    ))
    refs = [
        "[1] Tổng cục Thống kê Việt Nam (NSO), “Thông cáo báo chí tình hình kinh tế – xã hội quý IV và năm 2025,” Hà Nội, tháng 01/2026. [Trực tuyến]. Địa chỉ: https://www.nso.gov.vn/tin-tuc-thong-ke/2026/01/thong-cao-bao-chi-tinh-hinh-kinh-te-xa-hoi-quy-iv-va-nam-2025/.",
        "[2] Tổng cục Thống kê Việt Nam (NSO), “Thông cáo báo chí tình hình kinh tế – xã hội quý IV và năm 2024,” Hà Nội, tháng 01/2025. [Trực tuyến]. Địa chỉ: https://www.nso.gov.vn/du-lieu-va-so-lieu-thong-ke/2025/01/thong-cao-bao-chi-tinh-hinh-kinh-te-xa-hoi-quy-iv-va-nam-2024/.",
        "[3] T. U. Aye, A. Assad Neto, và E. Ribeiro da Silva, “Knowing your surroundings: Tactical retail demand forecasting through exogeneity-aware machine learning,” IEEE Engineering Management Review, 2025. [Trực tuyến]. Địa chỉ: https://ieeexplore.ieee.org/abstract/document/11303048.",
        "[4] S. Makridakis, E. Spiliotis, và V. Assimakopoulos, “M5 accuracy competition: Results, findings, and conclusions,” International Journal of Forecasting, tập 38, số 4, tr. 1346–1364, 2022. [Trực tuyến]. Địa chỉ: https://www.sciencedirect.com/science/article/pii/S0169207021001187.",
        "[5] K. Bandara, C. Bergmeir, và S. Smyl (hoặc nhóm tác giả M5 Benchmark), “Machine learning and statistical time series forecasting in retail M5 datasets,” International Journal of Forecasting, 2022. [Trực tuyến]. Địa chỉ: https://www.sciencedirect.com/science/article/pii/S0169207021001874.",
        "[6] M. Ulrich và cộng sự, “Comparative evaluation of SARIMA, SARIMAX and LSTM for retail demand forecasting with promotional effects,” Procedia Computer Science, 2022. [Trực tuyến]. Địa chỉ: https://www.sciencedirect.com/science/article/pii/S1877050922003076.",
        "[7] Nghiên cứu dự báo doanh số bán lẻ đa kênh bằng học máy tổ hợp (RF, XGBoost, Gradient Boosting, AdaBoost, ANN và RF-XGBoost-LR), PubMed Central (PMC9514716), 2022. [Trực tuyến]. Địa chỉ: https://pmc.ncbi.nlm.nih.gov/articles/PMC9514716/.",
        "[8] K. Shah và V. Dimitrov, “Univariate retail sales forecasting: A comparative study of Holt-Winters, SARIMA, LSTM and Prophet,” Inderscience International Journal, 2022. [Trực tuyến]. Địa chỉ: https://www.inderscience.com/info/inarticle.php?artid=126806.",
        "[9] Nguyễn Thị Liễu Điền, “Dự đoán nhu cầu hàng hóa thương mại điện tử bằng các phương pháp chuỗi thời gian và học máy,” Thư viện số Trường Đại học Kinh tế TP. Hồ Chí Minh (UEH), 2026. [Trực tuyến]. Địa chỉ: https://digital.lib.ueh.edu.vn/handle/UEH/79342?mode=full.",
        "[10] Nghiên cứu đối sánh các mô hình cây quyết định (Random Forest, XGBoost, LightGBM) và học sâu (LSTM, TCN) với biến ngoại sinh trong dự báo bán lẻ, IEEE Xplore, 2025. [Trực tuyến]. Địa chỉ: https://ieeexplore.ieee.org/document/11232957.",
        "[11] Nghiên cứu ứng dụng ARIMA, Gradient Boosting, LSTM và BiLSTM trong dự báo nhu cầu bán lẻ hiện đại, ScienceDirect, 2026. [Trực tuyến]. Địa chỉ: https://www.sciencedirect.com/science/article/pii/S2352146526002851.",
        "[12] R. J. Hyndman và G. Athanasopoulos, Forecasting: Principles and Practice, tái bản lần 3, OTexts, Melbourne, Australia, 2021 (Cơ sở lý thuyết ARIMA, SARIMA, Holt-Winters Exponential Smoothing và kiểm định ADF/KPSS).",
        "[13] L. Breiman, “Random Forests,” Machine Learning, tập 45, số 1, tr. 5–32, 2001; và T. Chen & C. Guestrin, “XGBoost: A Scalable Tree Boosting System,” trong Proc. 22nd ACM SIGKDD, tr. 785–794, 2016.",
        "[14] E. A. Silver, D. F. Pyke, và D. J. Thomas, Inventory and Production Management in Supply Chains, tái bản lần 4, CRC Press, 2017 (Cơ sở lý thuyết Safety Stock, Reorder Point và Ma trận phân loại ABC–XYZ).",
    ]
    for ref in refs:
        elems.append(make_p(
            [(ref, False, False, "0F172A", 24)],
            align="both",
            space_before=40,
            space_after=80,
            line=276,
            first_line=0,
            left_indent=420,
        ))
    return elems


# ---------------------------------------------------------------------------
# HÀM CHÍNH XÂY DỰNG LẠI TOÀN BỘ DOCUMENT.XML & ĐÓNG GÓI DOCX
# ---------------------------------------------------------------------------
def main():
    print("=" * 72)
    print("   HOÀN THIỆN BÁO CÁO ĐỒ ÁN WORD (MLBA.DOCX) & ĐỒNG BỘ HỆ THỐNG")
    print("=" * 72)

    if not os.path.exists(BACKUP_DOCX):
        print(f"--> Sao lưu file gốc lần đầu ra: {BACKUP_DOCX}")
        shutil.copy2(TARGET_DOCX, BACKUP_DOCX)
    else:
        print(f"--> Đọc từ bản gốc đã sao lưu: {BACKUP_DOCX}")

    with zipfile.ZipFile(BACKUP_DOCX, "r") as zin:
        zip_files = {name: zin.read(name) for name in zin.namelist()}

    doc_raw = zip_files["word/document.xml"].decode("utf-8")
    rels_raw = zip_files["word/_rels/document.xml.rels"].decode("utf-8")

    # Trích xuất chính xác thẻ mở <w:document ...> để giữ toàn bộ xmlns gốc
    m_hdr = re.match(r"^(<\?xml[^>]+\?>\s*<w:document[^>]+>)", doc_raw)
    orig_doc_header = m_hdr.group(1) if m_hdr else None

    doc_root = ET.fromstring(doc_raw)
    rels_root = ET.fromstring(rels_raw)
    ctx = DocxMediaContext(zip_files, rels_root)

    body = doc_root.find(wtag("body"))
    sect_pr = body.find(wtag("sectPr"))
    orig_elements = [el for el in list(body) if el is not sect_pr]

    # Khởi tạo danh sách phần tử mới theo chuẩn UEL (Trang bìa, Mục lục, Danh mục, Lời cảm ơn/cam kết, PCCV & Gantt)
    rids = sync_uel_styles_and_resources(zip_files, ctx)
    gantt_png = generate_gantt_png(os.path.join(FIGURES_DIR, "flowcharts", "gantt_tien_do_du_an.png"))
    new_body = []
    new_body.extend(build_uel_cover_and_front_matter(ctx, rids, gantt_png))

    # Các cờ kiểm soát trạng thái duyệt qua body gốc
    skip_until_marker = None
    ch5_pair_idx = 0
    ch5_screen_labels = [
        ("Hình 5.1: Giao diện Màn hình 1 — Tổng quan chuỗi bán lẻ và Khám phá dữ liệu EDA (#/overview)",
         "Hình 5.2: Lưu đồ xử lý nghiệp vụ Màn hình 1 — Tổng quan chuỗi và EDA"),
        ("Hình 5.3: Giao diện Màn hình 2 — Dự báo nhu cầu sản phẩm theo SKU (#/forecast)",
         "Hình 5.4: Lưu đồ xử lý nghiệp vụ Màn hình 2 — Dự báo nhu cầu sản phẩm theo SKU"),
        ("Hình 5.5: Giao diện Màn hình 3 — So sánh 5 mô hình dự báo và Tầm quan trọng đặc trưng (#/models)",
         "Hình 5.6: Lưu đồ xử lý nghiệp vụ Màn hình 3 — So sánh 5 mô hình và Feature Importance"),
        ("Hình 5.7: Giao diện Màn hình 4 — Kiểm tra tồn kho an toàn (SS) và Cảnh báo ROP (#/inventory)",
         "Hình 5.8: Lưu đồ xử lý nghiệp vụ Màn hình 4 — Quản trị tồn kho và Cảnh báo điểm đặt hàng lại"),
        ("Hình 5.9: Giao diện Màn hình 5 — Ma trận phân loại tồn kho chiến lược ABC – XYZ (#/abcxyz)",
         "Hình 5.10: Lưu đồ xử lý nghiệp vụ Màn hình 5 — Ma trận phân loại chiến lược ABC – XYZ"),
        ("Hình 5.11: Giao diện Màn hình 6 — Mô phỏng kịch bản What-If chuỗi cung ứng (#/whatif)",
         "Hình 5.12: Lưu đồ xử lý nghiệp vụ Màn hình 6 — Mô phỏng kịch bản What-If chuỗi cung ứng"),
    ]
    ch5_img_sub_idx = 0
    in_ch5 = False

    # Kiểm tra trước một số bảng đặc biệt bằng nội dung
    for idx, el in enumerate(orig_elements):
        is_tbl = el.tag == wtag("tbl")
        txt = (tbl_text(el) if is_tbl else p_text(el)).strip()

        # Bỏ qua 2 khối tài liệu tham khảo rời rạc cũ ở cuối 1.1 và cuối Chương 3 (đã gom về cuối báo cáo)
        if txt.startswith("[1] Tổng cục Thống kê") or txt.startswith("[2] https://www.nso.gov.vn") or txt.startswith("[3] Aye, T. U."):
            continue
        if any(txt.startswith(f"[{k}] https://") for k in range(1, 9)):
            continue

        # Bổ sung tiêu đề CHƯƠNG 1 + 1.1 trước đoạn mở đầu đầu tiên
        if not is_tbl and txt == "Bối cảnh và lý do thực hiện đề tài":
            new_body.append(make_chapter_heading("CHƯƠNG 1: TỔNG QUAN ĐỀ TÀI VÀ KHÁM PHÁ DỮ LIỆU", page_break=False))
            new_body.append(make_h2("1.1. Đặt vấn đề và Tổng quan đề tài"))
            replace_para_text(el, "1.1.1. Bối cảnh và lý do thực hiện đề tài")
            new_body.append(el)
            continue

        if not is_tbl and txt == "Tính cấp thiết của đề tài":
            replace_para_text(el, "1.1.2. Tính cấp thiết của đề tài")
            new_body.append(el)
            continue
        if not is_tbl and txt == "Mục tiêu nghiên cứu (Objectives)":
            replace_para_text(el, "1.1.3. Mục tiêu nghiên cứu")
            new_body.append(el)
            continue
        if not is_tbl and txt == "Đối tượng và phạm vi nghiên cứu":
            replace_para_text(el, "1.1.4. Đối tượng và phạm vi nghiên cứu")
            new_body.append(el)
            continue
        if not is_tbl and txt.startswith("Nghiên cứu tập trung vào hoạt động dự báo"):
            new_body.append(el)
            new_body.append(make_caption("Bảng 1.1: Đối tượng và phạm vi nghiên cứu của đề tài", is_table=True))
            continue

        # Sửa lỗi chính tả và dữ liệu trong Bảng T00 (Bảng 1.1)
        if is_tbl and "5 cửa hàng báun lẻ" in txt:
            replace_cell_text(el, 1, 1, "5 cửa hàng bán lẻ (S001 – S005) thuộc hệ thống, phân bố tại 4 khu vực địa lý (North, South, East, West)")
            replace_cell_text(el, 2, 1, "20 mặt hàng (P0001 – P0020) thuộc 5 ngành hàng (Clothing, Electronics, Furniture, Groceries, Toys)")
            replace_cell_text(el, 3, 1, "Từ ngày 01/01/2022 đến ngày 01/01/2024 (731 ngày liên tục, tổng cộng 73.100 bản ghi giao dịch)")
            replace_cell_text(el, 4, 1, "Theo ngày (Daily — 100 cặp Store × SKU mỗi ngày)")
            new_body.append(el)
            continue

        if not is_tbl and txt == "Thách thức nghiệp vụ":
            replace_para_text(el, "1.1.5. Thách thức nghiệp vụ")
            new_body.append(el)
            continue

        # Trước 1.3.1: chèn trọn vẹn Mục 1.2 (Sơ đồ 1.1 + Bảng 1.2) và tiêu đề 1.3
        if not is_tbl and txt.startswith("1.3.1. Tổng quan về dữ liệu"):
            new_body.extend(build_section_1_2(ctx))
            new_body.append(make_h2("1.3. Phân tích khám phá dữ liệu đa chiều (EDA)"))
            new_body.append(el)
            continue

        # Chèn chú thích Bảng 1.3 ngay trên bảng tổng kết EDA T01
        if is_tbl and txt.startswith("Biểu đồ (Phát hiện từ EDA)"):
            new_body.append(make_caption("Bảng 1.3: Tổng hợp các phát hiện chính từ 7 biểu đồ khám phá dữ liệu (EDA)", is_table=True))
            new_body.append(el)
            continue

        # CHƯƠNG 2
        if not is_tbl and txt.startswith("2.1. Ngôn ngữ lập trình"):
            new_body.append(make_chapter_heading("CHƯƠNG 2: CƠ SỞ LÝ THUYẾT VÀ PHƯƠNG PHÁP NGHIÊN CỨU", page_break=True))
            new_body.append(el)
            continue

        if not is_tbl and txt.startswith("ApexCharts:"):
            new_body.append(el)
            new_body.append(make_p([
                ("FastAPI & Uvicorn (Backend RESTful API): ", True, False, "0F172A", 26),
                ("Framework Python hiệu năng cao được sử dụng ở tầng máy chủ để nạp bộ nhớ đệm dữ liệu (RAM Cache), phục vụ các endpoint suy luận mô hình học máy theo thời gian thực (/api/forecast, /api/model-comparison, /api/inventory-status, /api/abc-xyz) và kết nối liền mạch với giao diện HTML5/ApexCharts đóng gói offline.", False, False, "0F172A", 26),
            ], space_before=60, space_after=120, first_line=720))
            continue

        if not is_tbl and txt.startswith("Thuật toán XGBoost Regressor cũng là một kỹ thuật"):
            new_body.append(el)
            new_body.extend(build_chapter2_additions())
            continue

        # CHƯƠNG 3
        if not is_tbl and txt.startswith("3.1. Tổng quan nghiên cứu"):
            new_body.append(make_chapter_heading("CHƯƠNG 3: LƯỢC KHẢO CÁC CÔNG TRÌNH NGHIÊN CỨU LIÊN QUAN", page_break=True))
            new_body.append(el)
            continue

        if not is_tbl and any(tag in txt for tag in ("[1, 2]", "[3]", "[4]", "[5]", "[6]", "[7]", "[8]")):
            if "42.840 chuỗi thời gian" in txt:
                replace_para_substrings(el, [("[1, 2]", "[4, 5]")])
            elif "tại Áo [3]" in txt:
                replace_para_substrings(el, [("[3]", "[6]")])
            elif "nghiên cứu năm 2022 [4]" in txt:
                replace_para_substrings(el, [("[4]", "[7]")])
            elif "Shah và Dimitrov" in txt:
                replace_para_substrings(el, [("[5]", "[8]")])
            elif "Nguyễn Thị Liễu Điền" in txt:
                replace_para_substrings(el, [("[6]", "[9]")])
            elif "năm 2025 [7]" in txt or "năm 2026 [8]" in txt:
                replace_para_substrings(el, [("[7]", "[10]"), ("[8]", "[11]")])

        if is_tbl and "Makridakis và cộng sự" in txt:
            replace_cell_text(el, 6, 1, "Nghiên cứu dự báo bán lẻ bằng mô hình học máy [10]")
            replace_cell_text(el, 7, 1, "Nghiên cứu dự báo nhu cầu bán lẻ [11]")
            new_body.append(el)
            continue

        if not is_tbl and txt.startswith("Trên cơ sở đó, đề tài tập trung so sánh nhóm thống kê"):
            replace_para_substrings(el, [("ARIMA, Holt-Winters", "ARIMA, SARIMA, Holt-Winters")])
            new_body.append(el)
            continue

        # CHƯƠNG 4
        if not is_tbl and txt.startswith("4.3. TRIỂN KHAI THỰC NGHIỆM"):
            new_body.append(make_chapter_heading("CHƯƠNG 4: TRIỂN KHAI THỰC NGHIỆM, ĐÁNH GIÁ MÔ HÌNH VÀ TỐI ƯU HÓA TỒN KHO", page_break=True))
            new_body.extend(build_chapter4_prefix(ctx))
            new_body.append(el)
            continue

        # Đồng bộ chia Train/Test trong 4.3.1 – 4.4 (560 Train / 141 Test theo main_train_and_evaluate.py)
        if not is_tbl and txt.startswith("Dữ liệu thực nghiệm được trích xuất cho cặp"):
            replace_para_text(
                el,
                "Dữ liệu thực nghiệm được trích xuất cho cặp đơn vị lưu kho chuẩn hóa SKU mục tiêu (Store ID = S001, Product ID = P0001), gồm 731 ngày liên tục từ 01/01/2022 đến 01/01/2024 (tương ứng tỷ lệ 80:20 thô là 584/147 ngày). Sau khi tạo hệ thống đặc trưng với độ trễ tối đa 30 ngày (lag_30) và loại bỏ 30 ngày khởi tạo đầu chuỗi (dropna()), ma trận dữ liệu sạch đưa vào huấn luyện và đánh giá đồng bộ cả 5 mô hình gồm 701 mẫu hợp lệ (từ 31/01/2022 đến 01/01/2024), chia theo tỷ lệ 80:20 như trong Bảng 4.3."
            )
            new_body.append(el)
            continue

        if not is_tbl and txt.startswith("Tập huấn luyện gồm 584 ngày"):
            replace_para_text(
                el,
                "Cụ thể, tập huấn luyện (Train Set) gồm 560 ngày liên tục (80%), từ 31/01/2022 đến 13/08/2023, được dùng để ước lượng tham số và học quy luật xu hướng, chu kỳ mùa vụ cũng như mối quan hệ phi tuyến với các đặc trưng ngoại cảnh. Tập kiểm thử (Test Set) gồm 141 ngày liên tục (20%), từ 14/08/2023 đến 01/01/2024, hoàn toàn tách biệt với quá trình huấn luyện và được dùng để đo sai số ngoại mẫu của cả 5 mô hình trên cùng một khoảng thời gian."
            )
            new_body.append(el)
            continue

        if not is_tbl and txt.startswith("Bảng 4.1: Chi tiết phân bổ tập huấn luyện"):
            replace_para_text(el, "Bảng 4.3: Chi tiết phân bổ tập huấn luyện và kiểm thử cho SKU S001–P0001")
            new_body.append(el)
            continue

        if is_tbl and "584 ngày" in txt and "Phân vùng" in txt:
            replace_cell_text(el, 1, 2, "560 ngày")
            replace_cell_text(el, 1, 3, "31/01/2022 – 13/08/2023")
            replace_cell_text(el, 2, 2, "141 ngày")
            replace_cell_text(el, 2, 3, "14/08/2023 – 01/01/2024")
            replace_cell_text(el, 3, 1, "100% (sau lag_30)")
            replace_cell_text(el, 3, 2, "701 ngày (từ 731 gốc)")
            replace_cell_text(el, 3, 3, "31/01/2022 – 01/01/2024")
            new_body.append(el)
            continue

        if not is_tbl and txt.startswith("Trục thời gian phân chia được bố trí liền kề"):
            replace_para_text(
                el,
                "Trục thời gian phân chia được bố trí liền kề và không chồng lấn: giai đoạn huấn luyện kết thúc vào ngày 13/08/2023 và giai đoạn kiểm thử bắt đầu ngay từ ngày 14/08/2023 đến 01/01/2024 (tổng cộng 141 ngày kiểm thử đồng bộ cho cả nhóm Thống kê lẫn Học máy)."
            )
            new_body.append(el)
            continue

        if not is_tbl and "như trong Bảng 4.2." in txt:
            replace_para_substrings(el, [("Bảng 4.2", "Bảng 4.4")])
            new_body.append(el)
            continue

        if not is_tbl and txt.startswith("Bảng 4.2: Các bước triển khai thực nghiệm"):
            replace_para_text(el, "Bảng 4.4: Các bước triển khai thực nghiệm huấn luyện 5 mô hình")
            new_body.append(el)
            continue

        if is_tbl and "Nạp và tiền xử lý dữ liệu" in txt:
            replace_cell_text(el, 2, 2, "Tạo 28 đặc trưng (lag_30), dropna() còn 701 ngày; chia theo thời gian: 560 ngày huấn luyện (31/01/2022 – 13/08/2023) và 141 ngày kiểm thử (14/08/2023 – 01/01/2024).")
            replace_cell_text(el, 5, 2, "Dự báo 141 ngày trên tập kiểm thử, tính RMSE, MAE và MAPE (%), xuất đồ thị đối sánh (Hình 4.4 – 4.8).")
            new_body.append(el)
            continue

        if not is_tbl and any(s in txt for s in ("trong suốt 147 ngày", "dự báo liên tục 147 ngày", "Năm mô hình được đánh giá trên cùng tập kiểm thử 147 ngày", "một khoảng kiểm thử 147 ngày")):
            replace_para_substrings(el, [
                ("147 ngày (08/08/2023 – 01/01/2024)", "141 ngày (14/08/2023 – 01/01/2024)"),
                ("147 ngày", "141 ngày"),
                ("Bảng 4.3", "Bảng 4.5"),
            ])
            new_body.append(el)
            continue

        if not is_tbl and txt.startswith("Bảng 4.3: Đối sánh định lượng"):
            replace_para_text(el, "Bảng 4.5: Đối sánh định lượng hiệu năng dự báo của 5 mô hình trên tập kiểm thử (141 ngày)")
            new_body.append(el)
            continue

        if not is_tbl and txt.startswith("Bảng 4.2: Bảng so sánh định lượng"):
            replace_para_text(el, "Bảng 4.6: Tổng hợp mức độ cải thiện RMSE theo nhóm phương pháp so với SARIMA")
            new_body.append(el)
            continue

        # Thay ảnh độ phân giải thấp P228 (Hình 4.9) bằng ảnh 300 DPI
        if not is_tbl and idx > 220 and len(el.findall(f".//{atag('blip')}")) > 0 and idx + 1 < len(orig_elements) and "Hình 4.9" in p_text(orig_elements[idx + 1]):
            replace_image_in_para(el, os.path.join(FIGURES_DIR, "models", "chuong_4_hinh_4_9_so_sanh_hieu_nang_mo_hinh.png"), ctx, max_w_emu=5700000, max_h_emu=3200000)
            new_body.append(el)
            continue

        # Thay ảnh độ phân giải thấp P237 (Hình 4.10) bằng ảnh 300 DPI + chèn dòng chú thích bị thiếu
        if not is_tbl and idx > 235 and len(el.findall(f".//{atag('blip')}")) > 0 and idx + 1 < len(orig_elements) and "ngoài yếu tố nguồn cung sẵn có" in p_text(orig_elements[idx + 1]):
            replace_image_in_para(el, os.path.join(FIGURES_DIR, "models", "chuong_4_hinh_4_10_tam_quan_trong_dac_trung.png"), ctx, max_w_emu=5600000, max_h_emu=3600000)
            new_body.append(el)
            new_body.append(make_caption("Hình 4.10: Top 10 đặc trưng quan trọng nhất trong dự báo nhu cầu (XGBoost Feature Importance)", is_table=False))
            continue

        if not is_tbl and "Cụ thể, ngoài yếu tố nguồn cung sẵn có là mức tồn kho" in txt:
            new_body.append(el)
            new_body.append(make_p(
                "Lưu ý đối chiếu giữa hai mô hình cây quyết định: Hình 4.10 minh họa mức độ quan trọng Gain từ XGBoost (trong đó Inventory Level dẫn đầu với ~0,082 hay 8,18%, theo sau là price_ratio 0,0516 và rolling_std_14 0,0485); tương tự trên mô hình tối ưu Random Forest Regressor (tệp random_forest_model.pkl và Sheet 6 Excel), Inventory Level tiếp tục giữ vị trí #1 áp đảo với tỷ trọng 42,33%, kế tiếp là lag_21 (4,82%), rolling_std_14 (4,45%) và price_ratio (4,42%). Kết quả thống nhất từ cả hai mô hình khẳng định độ chính xác của việc đưa các biến tồn kho, độ trễ và định giá vào dự báo nhu cầu bán lẻ."
            ))
            continue

        if not is_tbl and txt.startswith("Thiếu cơ chế dự báo đệ quy cho bài toán dài hạn"):
            new_body.append(el)
            new_body.extend(build_section_4_5(ctx))
            continue

        # CHƯƠNG 5
        if not is_tbl and txt.startswith("5.1. THIẾT KẾ TỔNG THỂ GIAO DIỆN"):
            in_ch5 = True
            new_body.append(make_chapter_heading("CHƯƠNG 5: XÂY DỰNG HỆ THỐNG HỖ TRỢ RA QUYẾT ĐỊNH RETAIL ML HUB STUDIO", page_break=True))
            new_body.append(el)
            continue

        # Thêm 3 Sơ đồ BPMN tổng thể vào cuối mục 5.1.2
        if not is_tbl and txt.startswith("Máy chủ xử lý: Đóng vai trò là trung tâm tính toán"):
            new_body.append(el)
            new_body.append(make_p(
                "Nhằm chuẩn hóa luồng tương tác giữa 3 thành phần (Người quản lý vận hành — Giao diện Web & Máy chủ FastAPI — Động cơ ML & Kho dữ liệu lịch sử), kiến trúc nghiệp vụ của hệ thống được mô hình hóa bằng bộ ba sơ đồ BPMN 3 làn dọc (tệp nguồn .drawio và PNG 300 DPI) thể hiện qua Sơ đồ 5.1, Sơ đồ 5.2 và Sơ đồ 5.3 dưới đây."
            ))
            new_body.extend(make_image_block(
                os.path.join(FIGURES_DIR, "flowcharts", "chuong_5_hinh_5_1_flowchart_tong_quan_he_thong.png"),
                "Sơ đồ 5.1: Sơ đồ BPMN Quy trình Khởi động Hệ thống và Phân tích Tổng quan / EDA",
                ctx,
                max_w_emu=5500000,
                max_h_emu=4600000,
            ))
            new_body.extend(make_image_block(
                os.path.join(FIGURES_DIR, "flowcharts", "chuong_5_hinh_5_2_flowchart_du_bao_va_so_sanh_mo_hinh.png"),
                "Sơ đồ 5.2: Sơ đồ BPMN Quy trình Dự báo Nhu cầu SKU (đệ quy) và So sánh 5 Mô hình",
                ctx,
                max_w_emu=5500000,
                max_h_emu=4600000,
            ))
            new_body.extend(make_image_block(
                os.path.join(FIGURES_DIR, "flowcharts", "chuong_5_hinh_5_3_flowchart_ton_kho_rop_abcxyz_whatif.png"),
                "Sơ đồ 5.3: Sơ đồ BPMN Quy trình Tối ưu Tồn kho (SS/ROP), Ma trận ABC–XYZ và Mô phỏng What-If",
                ctx,
                max_w_emu=5500000,
                max_h_emu=4600000,
            ))
            continue

        # Đếm cặp màn hình 5.2 -> 5.7 để gắn đúng 12 chú thích Hình 5.1 .. Hình 5.12 & sửa sai lệch
        if in_ch5 and not is_tbl:
            for s_idx, prefix in enumerate(["5.2. MÀN HÌNH 1", "5.3. MÀN HÌNH 2", "5.4. MÀN HÌNH 3", "5.5. MÀN HÌNH 4", "5.6. MÀN HÌNH 5", "5.7. MÀN HÌNH 6"]):
                if txt.startswith(prefix):
                    ch5_pair_idx = s_idx
                    ch5_img_sub_idx = 0

            if txt.startswith("Màn hình Tổng quan chuỗi cung cấp"):
                replace_para_substrings(el, [("đường trung bình 7 ngày", "đường trung bình trượt (như 4 tuần và 12 tuần)")])
            elif txt.startswith("Phân hệ Dự báo nhu cầu hỗ trợ"):
                replace_para_substrings(el, [("162,7 đơn vị/ngày", "162,71 đơn vị/ngày")])
            elif txt.startswith("Phân hệ Đánh giá mô hình cung cấp"):
                replace_para_substrings(el, [("biểu đồ kiểm định 147 ngày", "biểu đồ kiểm định 141 ngày (tập kiểm thử cuối chuỗi)")])
            elif txt.startswith("Phân hệ Quản trị tồn kho chuyển hóa"):
                replace_para_substrings(el, [
                    ("gồm 20 mã sản phẩm được phân loại thành nhóm Nguy cấp (15 mã), Cảnh báo (2 mã) và An toàn (3 mã) có thể nhấn vào để lọc bảng.",
                     "gồm Tổng số mã SKU (20 mã), Nguy cấp (20 mã cần phát lệnh đặt hàng ngay do tồn cuối kỳ tại S001 thấp hơn ROP), Cảnh báo (0 mã) và An toàn (0 mã), cho phép nhấp từng thẻ để lọc nhanh danh sách."),
                ])
            elif txt.startswith("Người quản lý bắt đầu bằng việc chọn chi nhánh cửa hàng cần phân tích để hệ thống"):
                replace_para_substrings(el, [("28 tháng", "24 tháng (731 ngày từ 01/01/2022 đến 01/01/2024)")])
            elif txt.startswith("Phân hệ Mô phỏng What-If cung cấp công cụ"):
                replace_para_substrings(el, [
                    ("và nâng mức độ dịch vụ lên 99%", "trong khi giữ mức độ dịch vụ cân bằng 95% (kịch bản mặc định L = 3 ngày, SL = 95%)"),
                ])
            elif txt.startswith("5.8 Bảng tổng hợp các chức năng"):
                in_ch5 = False
                new_body.append(el)
                new_body.append(make_caption("Bảng 5.1: Bảng tổng hợp 23 chức năng của hệ thống Retail ML Hub Studio", is_table=True))
                continue

            if len(el.findall(f".//{atag('blip')}")) > 0:
                new_body.append(el)
                cap = ch5_screen_labels[ch5_pair_idx][min(ch5_img_sub_idx, 1)]
                new_body.append(make_caption(cap, is_table=False))
                ch5_img_sub_idx += 1
                continue

        # CHƯƠNG 6
        if not is_tbl and txt.startswith("6.1. Tổng kết các nội dung chính"):
            new_body.append(make_chapter_heading("CHƯƠNG 6: KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", page_break=True))
            new_body.append(el)
            continue

        if not is_tbl and "SS = 313 sản phẩm" in txt:
            replace_para_substrings(el, [
                ("Mức tồn kho an toàn (SS = 313 sản phẩm) và Điểm đặt hàng lại (ROP = 707 sản phẩm)",
                 "Mức tồn kho an toàn (SS = 311 sản phẩm) và Điểm đặt hàng lại (ROP = 777 sản phẩm) khi tích hợp nhu cầu dự báo Random Forest với Service Level 95% (Z = 1,645, L = 3 ngày; tương ứng mức tham chiếu lịch sử 90 ngày là SS = 312 và ROP = 707)"),
            ])
            new_body.append(el)
            continue

        new_body.append(el)

    # Thêm Mục Tài liệu tham khảo chuẩn IEEE vào cuối tài liệu
    new_body.extend(build_references_section())

    # Chuẩn hóa cấu trúc UEL, cập nhật TOC/LOF/LOT động, đánh lại docPr và Section 2 sectPr
    new_body, uel_stats = finalize_uel_document(new_body, sect_pr, rids["footer_rid"], ctx)
    print(f"[INFO] UEL format stats: {uel_stats}")

    # Ghi đè lại các phần tử con trong <w:body> và giữ nguyên <w:sectPr> ở cuối
    for child in list(body):
        body.remove(child)
    for el in new_body:
        body.append(el)
    if sect_pr is not None:
        body.append(sect_pr)

    # Xuất XML và khôi phục nguyên bản thẻ gốc <w:document ...> nếu có namespace default
    doc_xml_str = ET.tostring(doc_root, encoding="utf-8", xml_declaration=True).decode("utf-8")
    if orig_doc_header:
        doc_xml_str = re.sub(r"^<\?xml[^>]*\?>\s*<w:document[^>]*>", orig_doc_header, doc_xml_str, count=1)
    zip_files["word/document.xml"] = doc_xml_str.encode("utf-8")

    # Rel XML
    ET.register_namespace("", REL_NS)
    rels_xml_str = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True).decode("utf-8")
    zip_files["word/_rels/document.xml.rels"] = rels_xml_str.encode("utf-8")

    # Ghi file nguyên tử qua tệp tạm rồi thay thế
    tmp_out = TARGET_DOCX + ".tmp"
    with zipfile.ZipFile(tmp_out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in zip_files.items():
            zout.writestr(name, data)

    try:
        os.replace(tmp_out, TARGET_DOCX)
    except PermissionError as exc:
        print(f"[LỖI KHÓA FILE] Vui lòng đóng Microsoft Word đang mở '{TARGET_DOCX}' rồi chạy lại script!")
        raise exc

    shutil.copy2(TARGET_DOCX, COPY_REPORTS_DOCX)

    sz_mb = os.path.getsize(TARGET_DOCX) / (1024 * 1024)
    print(f"--> Hoàn tất xuất báo cáo Word đồng bộ ({sz_mb:.2f} MB):")
    print(f"    1) {TARGET_DOCX}")
    print(f"    2) {COPY_REPORTS_DOCX}")
    print(f"    3) Bản gốc dự phòng: {BACKUP_DOCX}")
    print("=" * 72)


if __name__ == "__main__":
    main()
