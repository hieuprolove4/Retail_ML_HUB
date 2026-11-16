r"""
patch_mlba3_review.py
Chỉnh sửa toàn diện tệp D:\HK7\Machine_Learning\doan\MLBA (3).docx dựa trên bản sao lưu
D:\HK7\Machine_Learning\doan\MLBA (3)_backup.docx theo đánh giá chuyên môn và chuẩn trích dẫn APA 7th:
  - Bảo toàn nguyên vẹn số hiệu đề mục (1.1., 4.1., 6.1...), tiêu đề Hình/Bảng, ngày tháng, URL và các công thức <m:oMath> chính.
  - Chuẩn hóa dấu thập phân ',' tiếng Việt cho toàn bộ số liệu trong văn bản và bảng.
  - Thống nhất học vị giảng viên (ThS. Trần Duy Thanh) và tỷ lệ đóng góp thực tế của 5 thành viên (23%, 22%, 20%, 18%, 17%).
  - Bổ sung Baseline Naive/Seasonal Naive, đối chứng loại bỏ Inventory Level (RF/XGB 27 đặc trưng),
    đối chứng đệ quy thuần 141 bước, kiểm định Ljung-Box và Diebold-Mariano vào Chương 4.
  - Sửa sai số học Q = 943, giải thích quy tắc 1,5 × ROP, tái tạo ROP = 707 (D = 131,50) vs 721 (D90 = 136,50),
    giải thích Hình 4.11 ([50, 500]) và tính suy biến của ma trận ABC-XYZ (CV 0,758–0,834 > 0,70).
  - Giảm các diễn đạt quá mức ở Chương 4 & Chương 6, viết lại Bảng 6.1 trung thực theo đúng kết quả đo được.
  - Chuyển toàn bộ trích dẫn trong bài & Danh mục Tài liệu tham khảo sang chuẩn APA (7th ed.).
"""

import copy
import io
import os
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET

if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
XML_NS = "http://www.w3.org/XML/1998/namespace"

DOCX_PATH = r"D:\HK7\Machine_Learning\doan\MLBA (3).docx"
BACKUP_PATH = r"D:\HK7\Machine_Learning\doan\MLBA (3)_backup.docx"

_PARA_ID_COUNTER = 0x7B000001


def _next_para_id():
    global _PARA_ID_COUNTER
    val = f"{_PARA_ID_COUNTER:08X}"
    _PARA_ID_COUNTER += 1
    return val


def _w(tag):
    return f"{{{W_NS}}}{tag}"


def _p_style(p_el):
    ppr = p_el.find(_w("pPr"))
    if ppr is not None:
        ps = ppr.find(_w("pStyle"))
        if ps is not None:
            return ps.attrib.get(_w("val"), "")
    return ""


def _p_text(el):
    return "".join(
        n.text or ""
        for n in el.iter()
        if n.tag in (_w("t"), f"{{{M_NS}}}t")
    )


def _set_plain_text_preserve_format(p_el, new_text, strip_omath=False, italic=False):
    """
    Thay toàn bộ văn bản của <w:p> giữ nguyên <w:pPr>, <w:rPr> mẫu và bookmarks.
    """
    first_r = p_el.find(_w("r"))
    rpr_template = (
        copy.deepcopy(first_r.find(_w("rPr")))
        if first_r is not None and first_r.find(_w("rPr")) is not None
        else None
    )

    keep_after = []
    for ch in list(p_el):
        if ch.tag == _w("r") or (strip_omath and ch.tag == f"{{{M_NS}}}oMath"):
            p_el.remove(ch)
        elif ch.tag == _w("bookmarkEnd"):
            keep_after.append(ch)
            p_el.remove(ch)

    r = ET.Element(_w("r"))
    if rpr_template is not None:
        r.append(rpr_template)
    elif italic:
        rpr = ET.SubElement(r, _w("rPr"))
        ET.SubElement(rpr, _w("i"))
    t = ET.SubElement(r, _w("t"))
    if new_text.startswith(" ") or new_text.endswith(" "):
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = new_text
    p_el.append(r)

    for bk in keep_after:
        p_el.append(bk)


def _set_apa_reference_runs(p_el, parts):
    """Ghi đoạn văn tài liệu tham khảo APA với các đoạn (text, is_italic) giữ nguyên <w:pPr> & font."""
    first_r = p_el.find(_w("r"))
    base_rpr = (
        copy.deepcopy(first_r.find(_w("rPr")))
        if first_r is not None and first_r.find(_w("rPr")) is not None
        else ET.Element(_w("rPr"))
    )
    for ch in list(p_el):
        if ch.tag == _w("r"):
            p_el.remove(ch)

    for txt, is_italic in parts:
        r = ET.Element(_w("r"))
        rpr = copy.deepcopy(base_rpr)
        for old_i in list(rpr.findall(_w("i"))):
            rpr.remove(old_i)
        if is_italic:
            ET.SubElement(rpr, _w("i"))
        if len(rpr):
            r.append(rpr)
        t = ET.SubElement(r, _w("t"))
        if txt.startswith(" ") or txt.endswith(" "):
            t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = txt
        p_el.append(r)


def _replace_in_wt_runs(p_el, old_sub, new_sub):
    replaced = False
    for t in p_el.iter(_w("t")):
        if t.text and old_sub in t.text:
            t.text = t.text.replace(old_sub, new_sub)
            if t.text.startswith(" ") or t.text.endswith(" "):
                t.set(f"{{{XML_NS}}}space", "preserve")
            replaced = True
    return replaced


def _set_cell_text(tc_el, text):
    ps = tc_el.findall(_w("p"))
    if not ps:
        return
    _set_plain_text_preserve_format(ps[0], text, strip_omath=True)
    for extra in ps[1:]:
        tc_el.remove(extra)


def _clone_tr_with_values(template_tr, values):
    new_tr = copy.deepcopy(template_tr)
    w14_pid = f"{{{W14_NS}}}paraId"
    if w14_pid in new_tr.attrib:
        new_tr.set(w14_pid, _next_para_id())
    for p in new_tr.iter(_w("p")):
        if w14_pid in p.attrib:
            p.set(w14_pid, _next_para_id())
    tcs = new_tr.findall(_w("tc"))
    for tc, val in zip(tcs, values):
        _set_cell_text(tc, val)
    return new_tr


def _find_paragraph_by_prefix(body_children, prefix, start_idx=139):
    for i in range(start_idx, len(body_children)):
        el = body_children[i]
        if el.tag == _w("p"):
            txt = _p_text(el).strip()
            if txt.startswith(prefix):
                return i, el
    raise ValueError(f"Không tìm thấy đoạn văn bắt đầu bằng: {prefix!r}")


def _find_table_after_header(body_children, header_fragment, start_idx=0):
    for i in range(start_idx, len(body_children)):
        el = body_children[i]
        if el.tag == _w("tbl"):
            rows = el.findall(_w("tr"))
            if rows:
                hdr_txt = " | ".join(_p_text(tc).strip() for tc in rows[0].findall(_w("tc")))
                if header_fragment in hdr_txt:
                    return i, el
    raise ValueError(f"Không tìm thấy bảng có tiêu đề: {header_fragment!r}")


_SECTION_OR_REF_PREFIX_RE = re.compile(
    r"(?:Bảng|Hình|Sơ đồ|Chương|Mục|hình|bảng|mục|chương|Python|v)\s+$",
    re.IGNORECASE,
)


def _convert_decimal_dots_in_text(text):
    """
    Chuyển dấu thập phân '.' -> ',' cho số liệu khoa học,
    BẢO TOÀN 100%: số đề mục (1.1., 4.1., 6.1.), Hình/Bảng/Mục (Hình 4.10, Bảng 1.1-1.3),
    số hàng nghìn (73.100, 42.840, 1.139), Python 3.11, ngày tháng và URL.
    """
    if not text or "http" in text or "www." in text:
        return text

    def repl(m):
        full = m.group(0)
        start, end = m.start(), m.end()
        before = text[:start]
        after = text[end:]

        # Bảo toàn đề mục kiểu "1.1.", "4.3.3", "6.4." (theo sau là dấu chấm hoặc đầu câu kết thúc bởi ".")
        if after.startswith(".") or before.endswith("."):
            return full
        if _SECTION_OR_REF_PREFIX_RE.search(before):
            return full
        if re.search(r"[-–]\s*$", before) and re.search(r"(?:Bảng|Hình|Sơ đồ|Mục)\s+\d+\.\d+\s*[-–]\s*$", before):
            return full
        if full == "3.11" and "Python" in text:
            return full

        left, right = full.split(".")
        # Bảo toàn dấu phân tách hàng nghìn (vd 73.100, 42.840, 26.784) trừ các số 0.xxx
        if len(right) == 3 and left != "0" and (start == 0 or not before[-1].isdigit()):
            return full
        if len(right) == 3 and re.match(r"^\.\d{3}", after):
            return full

        return f"{left},{right}"

    return re.sub(r"(?<![A-Za-z0-9/])\d+\.\d+(?![A-Za-z0-9/])", repl, text)


def _apa_references():
    """Danh mục 15 tài liệu tham khảo chuẩn APA (7th ed.), xếp theo vần A–Z tác giả đầu tiên."""
    return [
        [
            ("Aye, T. U., Assad Neto, A., & Ribeiro da Silva, E. (2025). Knowing your surroundings: Tactical retail demand forecasting through exogeneity-aware machine learning. ", False),
            ("IEEE Engineering Management Review", True),
            (". https://ieeexplore.ieee.org/abstract/document/11303048", False),
        ],
        [
            ("Bertsimas, D., & Kallus, N. (2020). From predictive to prescriptive analytics. ", False),
            ("Management Science", True),
            (", ", False),
            ("66", True),
            ("(3), 1025–1044. https://doi.org/10.1287/mnsc.2018.3253", False),
        ],
        [
            ("Boylan, J. E., & Syntetos, A. A. (2021). ", False),
            ("Intermittent demand forecasting: Context, methods and applications", True),
            (". John Wiley & Sons.", False),
        ],
        [
            ("Breiman, L. (2001). Random forests. ", False),
            ("Machine Learning", True),
            (", ", False),
            ("45", True),
            ("(1), 5–32. https://doi.org/10.1023/A:1010933404324", False),
        ],
        [
            ("Carbonneau, R., Laframboise, K., & Vahidov, R. (2008). Application of machine learning techniques for supply chain demand forecasting. ", False),
            ("European Journal of Operational Research", True),
            (", ", False),
            ("184", True),
            ("(3), 1140–1154. https://doi.org/10.1016/j.ejor.2006.12.004", False),
        ],
        [
            ("Chen, T., & Guestrin, C. (2016). XGBoost: A scalable tree boosting system. In ", False),
            ("Proceedings of the 22nd ACM SIGKDD International Conference on Knowledge Discovery and Data Mining", True),
            (" (pp. 785–794). ACM. https://doi.org/10.1145/2939672.2939785", False),
        ],
        [
            ("Falatouri, T., Darbanian, F., Brandtner, P., & Udokwu, C. (2022). Predictive analytics for demand forecasting: A comparison of SARIMA and LSTM in retail SCM. ", False),
            ("Procedia Computer Science", True),
            (", ", False),
            ("200", True),
            (", 993–1003. https://www.sciencedirect.com/science/article/pii/S1877050922003076", False),
        ],
        [
            ("Hyndman, R. J., & Athanasopoulos, G. (2021). ", False),
            ("Forecasting: Principles and practice", True),
            (" (3rd ed.). OTexts. https://otexts.com/fpp3/", False),
        ],
        [
            ("Makridakis, S., Spiliotis, E., & Assimakopoulos, V. (2022). M5 accuracy competition: Results, findings, and conclusions. ", False),
            ("International Journal of Forecasting", True),
            (", ", False),
            ("38", True),
            ("(4), 1346–1364. https://www.sciencedirect.com/science/article/pii/S0169207021001187", False),
        ],
        [
            ("Mitra, A., Jain, A., Kishore, A., & Kumar, P. (2022). A comparative study of demand forecasting models for a multi-channel retail company: A novel hybrid machine learning approach. ", False),
            ("Operations Research Forum", True),
            (", ", False),
            ("3", True),
            (", Article 58. https://pmc.ncbi.nlm.nih.gov/articles/PMC9514716/", False),
        ],
        [
            ("Nguyễn, T. L. Đ. (2026). ", False),
            ("Dự đoán nhu cầu hàng hóa thương mại điện tử bằng các phương pháp chuỗi thời gian và học máy", True),
            (" [Đồ án tốt nghiệp, Trường Đại học Kinh tế TP. Hồ Chí Minh]. Thư viện số UEH. https://digital.lib.ueh.edu.vn/handle/UEH/79342?mode=full", False),
        ],
        [
            ("Shah, K., & Dimitrov, V. (2022). Univariate retail sales forecasting: A comparative study of Holt-Winters, SARIMA, LSTM and Prophet. ", False),
            ("Inderscience International Journal", True),
            (". https://www.inderscience.com/info/inarticle.php?artid=126806", False),
        ],
        [
            ("Silver, E. A., Pyke, D. F., & Thomas, D. J. (2017). ", False),
            ("Inventory and production management in supply chains", True),
            (" (4th ed.). CRC Press.", False),
        ],
        [
            ("Syntetos, A. A., Babai, Z., Boylan, J. E., Kolassa, S., & Nikolopoulos, K. (2016). Supply chain forecasting: Theory, practice, their gap and the future. ", False),
            ("European Journal of Operational Research", True),
            (", ", False),
            ("252", True),
            ("(1), 1–26. https://doi.org/10.1016/j.ejor.2015.11.010", False),
        ],
        [
            ("Tổng cục Thống kê. (2026). ", False),
            ("Thông cáo báo chí tình hình kinh tế – xã hội quý IV và năm 2025", True),
            (". https://www.nso.gov.vn/tin-tuc-thong-ke/2026/01/thong-cao-bao-chi-tinh-hinh-kinh-te-xa-hoi-quy-iv-va-nam-2025/", False),
        ],
    ]


def main():
    if not os.path.exists(BACKUP_PATH):
        shutil.copy2(DOCX_PATH, BACKUP_PATH)
        print(f"--> Đã tạo bản sao lưu gốc tại: {BACKUP_PATH}")

    with zipfile.ZipFile(BACKUP_PATH, "r") as zin:
        zip_files = {name: zin.read(name) for name in zin.namelist()}

    doc_raw = zip_files["word/document.xml"]
    namespaces = dict([node for _, node in ET.iterparse(io.BytesIO(doc_raw), events=["start-ns"])])
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)

    doc_root = ET.fromstring(doc_raw)
    body = doc_root.find(_w("body"))
    children = list(body)

    # ------------------------------------------------------------------ #
    # 1. ĐỒNG BỘ PHẦN MỞ ĐẦU: HỌC VỊ GIẢNG VIÊN & TỶ LỆ ĐÓNG GÓP NHÓM
    # ------------------------------------------------------------------ #
    _, p_camket = _find_paragraph_by_prefix(children, "Nhóm xin cam kết toàn bộ", start_idx=0)
    _set_plain_text_preserve_format(
        p_camket,
        _p_text(p_camket).replace("TS.Trần Duy Thanh", "ThS. Trần Duy Thanh"),
    )

    _, tbl_members = _find_table_after_header(children, "STT | Họ và tên | MSSV | Đóng góp (%)")
    m_rows = tbl_members.findall(_w("tr"))
    shares_m = {"K244060703": "17%", "K244060716": "22%", "K244060721": "23%", "K244060723": "20%", "K244060726": "18%"}
    for r in m_rows[1:]:
        tcs = r.findall(_w("tc"))
        mssv = _p_text(tcs[2]).strip()
        if mssv in shares_m:
            _set_cell_text(tcs[3], shares_m[mssv])

    _, tbl_pccv = _find_table_after_header(children, "STT | Thành viên | Vai trò | Phạm vi phụ trách")
    p_rows = tbl_pccv.findall(_w("tr"))
    eval_map = {
        "Phùng Trịnh Trường Khang": "23% (Hoàn thành tốt)",
        "Huỳnh Mai Hiếu": "22% (Hoàn thành tốt)",
        "Phùng Anh Minh": "18% (Hoàn thành tốt)",
        "Nguyễn Huỳnh Thiên Ân": "17% (Hoàn thành tốt)",
        "Cai Đặng Phúc Lợi": "20% (Hoàn thành tốt)",
    }
    for r in p_rows[1:]:
        tcs = r.findall(_w("tc"))
        name = _p_text(tcs[1]).strip()
        if name in eval_map:
            _set_cell_text(tcs[4], eval_map[name])

    # ------------------------------------------------------------------ #
    # 2. CHƯƠNG 1 & THỐNG NHẤT NGUỒN DỮ LIỆU + TRÍCH DẪN APA
    # ------------------------------------------------------------------ #
    _, p_c1_1 = _find_paragraph_by_prefix(children, "Thị trường bán lẻ Việt Nam")
    _set_plain_text_preserve_format(
        p_c1_1,
        "Thị trường bán lẻ Việt Nam vẫn duy trì đà tăng trưởng. Theo Tổng cục Thống kê (2026), tổng mức bán lẻ hàng hóa và doanh thu dịch vụ tiêu dùng năm 2025 ước đạt 7.008,9 nghìn tỷ đồng, tăng 9,2% so với năm 2024. Quy mô này đặt ra yêu cầu ngày càng cao đối với năng lực dự báo và điều phối nguồn cung của doanh nghiệp.",
    )

    _, p_c1_2 = _find_paragraph_by_prefix(children, "Nhu cầu tiêu dùng thường biến động")
    _set_plain_text_preserve_format(
        p_c1_2,
        "Nhu cầu tiêu dùng thường biến động theo thời gian, sản phẩm, khu vực, giá bán, khuyến mãi, thời tiết và các dịp lễ, Tết. Nghiên cứu của Aye và cộng sự (2025) trên một nhà bán lẻ trực tuyến lớn tại châu Âu cho thấy các mô hình học máy có tích hợp biến ngoại sinh phù hợp có thể giảm sai số dự báo tới 22% so với phương pháp truyền thống.",
    )

    _, p_12 = _find_paragraph_by_prefix(children, "Bộ dữ liệu thực nghiệm được sử dụng trong đồ án")
    _set_plain_text_preserve_format(
        p_12,
        "Bộ dữ liệu được sử dụng trong đồ án là bộ dữ liệu mô phỏng công khai trên Kaggle (Retail Store Inventory Forecasting Dataset), mô phỏng giao dịch bán lẻ và quản trị tồn kho theo ngày gồm tổng cộng 73.100 bản ghi hoàn chỉnh kéo dài trong 731 ngày liên tục từ ngày 01/01/2022 đến ngày 01/01/2024 (tương ứng 24 tháng). Dữ liệu bao phủ 5 cửa hàng (S001 - S005) thuộc 4 khu vực địa lý (North, South, East, West) với danh mục 20 mã sản phẩm (P0001 - P0020) trải rộng trên 5 ngành hàng (Clothing, Electronics, Furniture, Groceries, Toys). Mỗi ngày có đúng 100 cặp Cửa hàng × Sản phẩm được ghi nhận đầy đủ; kết quả kiểm định chất lượng cho thấy 0 giá trị bị khuyết (Null/NaN) trên toàn bộ 15 cột thuộc tính.",
    )

    _, p_131 = _find_paragraph_by_prefix(children, "Dữ liệu phân tích bao gồm 73.100 giao dịch mô phỏng")
    _set_plain_text_preserve_format(
        p_131,
        _p_text(p_131).replace(
            "73.100 giao dịch mô phỏng của 20 sản phẩm",
            "73.100 giao dịch mô phỏng công khai (Kaggle — Retail Store Inventory Forecasting Dataset) của 20 sản phẩm",
        ),
    )

    # ------------------------------------------------------------------ #
    # 3. CHƯƠNG 2 & CHƯƠNG 3: CHUẨN HÓA TRÍCH DẪN APA + TÁC GIẢ ĐẦY ĐỦ
    # ------------------------------------------------------------------ #
    _, p_arima_def = _find_paragraph_by_prefix(children, "Mô hình ARIMA là phương pháp thống kê")
    _set_plain_text_preserve_format(
        p_arima_def,
        _p_text(p_arima_def).replace("không dừng.", "không dừng (Hyndman & Athanasopoulos, 2021)."),
    )
    _, p_sarima_def = _find_paragraph_by_prefix(children, "Mô hình SARIMA là bản mở rộng")
    _set_plain_text_preserve_format(
        p_sarima_def,
        _p_text(p_sarima_def).replace("tính mùa vụ trong chuỗi dữ liệu.", "tính mùa vụ trong chuỗi dữ liệu (Hyndman & Athanasopoulos, 2021)."),
    )
    _, p_hw_def = _find_paragraph_by_prefix(children, "Phương pháp San bằng mũ Holt-Winters")
    _set_plain_text_preserve_format(
        p_hw_def,
        _p_text(p_hw_def).replace("san bằng mũ bộ ba.", "san bằng mũ bộ ba (Hyndman & Athanasopoulos, 2021)."),
    )
    _, p_rf_def = _find_paragraph_by_prefix(children, "Thuật toán Random Forest Regressor")
    _set_plain_text_preserve_format(
        p_rf_def,
        _p_text(p_rf_def).replace("dữ liệu phi tuyến.", "dữ liệu phi tuyến (Breiman, 2001)."),
    )
    _, p_xgb_def = _find_paragraph_by_prefix(children, "Thuật toán XGBoost Regressor")
    _set_plain_text_preserve_format(
        p_xgb_def,
        _p_text(p_xgb_def).replace("vượt trội.", "hiệu quả cho dữ liệu dạng bảng (Chen & Guestrin, 2016)."),
    )

    _, p_mape_c2 = _find_paragraph_by_prefix(children, "Mean Absolute Percentage Error (MAPE")
    _replace_in_wt_runs(
        p_mape_c2,
        "bên cạnh cặp RMSE và MAE.",
        "bên cạnh cặp RMSE và MAE (Boylan & Syntetos, 2021).",
    )

    _, p_ss_c2 = _find_paragraph_by_prefix(children, "Mức tồn kho an toàn (SS):")
    _replace_in_wt_runs(
        p_ss_c2,
        "Công thức chuẩn giả định phân phối nhu cầu:",
        "Công thức chuẩn (Silver, Pyke, & Thomas, 2017) giả định phân phối nhu cầu:",
    )

    _, p_rop_c2 = _find_paragraph_by_prefix(children, "Điểm đặt hàng lại (ROP) và Đề xuất đặt hàng:")
    _replace_in_wt_runs(
        p_rop_c2,
        "Khi ở trạng thái Nguy cấp, lượng đặt hàng bổ sung được tính theo mức tồn mục tiêu",
        "Trong nguyên mẫu hiện tại (chưa thu thập chi phí đặt hàng K và chi phí lưu kho h để giải EOQ), hệ thống áp dụng quy tắc thực nghiệm (s, S) tạo vùng đệm 50% phía trên ngưỡng đặt hàng lại với mức tồn mục tiêu",
    )

    _, p_abc_c2 = _find_paragraph_by_prefix(children, "(Nhóm X ổn định:")
    _set_plain_text_preserve_format(
        p_abc_c2,
        "(Nhóm X ổn định: CV <= 0,40; Nhóm Y dao động vừa: 0,40 < CV <= 0,70; Nhóm Z biến động mạnh: CV > 0,70). Ma trận 9 ô (AX -> CZ) giúp phân bổ chính sách kiểm soát tồn kho, song các ngưỡng cố định này cần được hiệu chỉnh theo phân phối CV thực tế của từng bộ dữ liệu, vì trong y văn phân loại nhu cầu (Syntetos và cộng sự, 2016), ngưỡng nhu cầu biến động mạnh thường đặt cao hơn (ví dụ CV > 1,0).",
    )

    # Chương 3 APA author-year (xóa [1..11] và bổ sung tác giả đầy đủ)
    _, p_m5 = _find_paragraph_by_prefix(children, "Cuộc thi dự báo M5 là một công trình")
    _set_plain_text_preserve_format(
        p_m5,
        _p_text(p_m5).replace("Cuộc thi dự báo M5 là một công trình có quy mô lớn", "Cuộc thi dự báo M5 (Makridakis và cộng sự, 2022) là một công trình có quy mô lớn"),
    )
    _, p_falatouri = _find_paragraph_by_prefix(children, "Một nghiên cứu so sánh SARIMA và LSTM")
    _set_plain_text_preserve_format(
        p_falatouri,
        _p_text(p_falatouri).replace("Một nghiên cứu so sánh SARIMA và LSTM trên dữ liệu bán hàng", "Nghiên cứu của Falatouri và cộng sự (2022) so sánh SARIMA và LSTM trên dữ liệu bán hàng"),
    )
    _, p_mitra = _find_paragraph_by_prefix(children, "Đối với doanh nghiệp bán lẻ đa kênh")
    _set_plain_text_preserve_format(
        p_mitra,
        _p_text(p_mitra).strip().replace("nghiên cứu năm 2022 so sánh", "Mitra và cộng sự (2022) so sánh"),
    )
    # Shah & Dimitrov (2022) ở [296] đã đúng APA
    _, p_exo_c3 = _find_paragraph_by_prefix(children, "Các nghiên cứu gần đây cũng đánh giá học máy")
    _set_plain_text_preserve_format(
        p_exo_c3,
        "Các nghiên cứu về dự báo trong chuỗi cung ứng cũng khẳng định vai trò của biến ngoại sinh và quan hệ phi tuyến. Carbonneau, Laframboise và Vahidov (2008) chứng minh các kỹ thuật học máy phi tuyến cải thiện khả năng nắm bắt tín hiệu nhu cầu bị khuếch đại qua các tầng chuỗi cung ứng so với bình quân trượt hay hồi quy tuyến tính. Gần đây hơn, Aye và cộng sự (2025) cho thấy việc tích hợp có chọn lọc các biến ngoại cảnh như khuyến mãi, thời tiết và lịch sự kiện giúp mô hình dựa trên cây quyết định giảm sai số tới 22% so với mô hình đơn biến truyền thống.",
    )
    _, p_dien = _find_paragraph_by_prefix(children, "Một công trình tiêu biểu là đề án của Nguyễn Thị Liễu Điền")
    _set_plain_text_preserve_format(
        p_dien,
        _p_text(p_dien).replace("Nguyễn Thị Liễu Điền (2026) [9]", "Nguyễn Thị Liễu Điền (2026)"),
    )

    _, p_316_gap = _find_paragraph_by_prefix(children, "Từ đó, có thể xác định một số khoảng trống")
    _set_plain_text_preserve_format(
        p_316_gap,
        _p_text(p_316_gap).replace("bộ dữ liệu bán lẻ thực tế", "cùng một bộ dữ liệu bán lẻ mô phỏng công khai (Kaggle)"),
    )
    _, p_316_scope = _find_paragraph_by_prefix(children, "Trên cơ sở đó, đề tài tập trung so sánh")
    _set_plain_text_preserve_format(
        p_316_scope,
        _p_text(p_316_scope).replace("một chuỗi doanh số thực tế", "một chuỗi doanh số từ dữ liệu mô phỏng công khai (Kaggle)"),
    )

    # Bảng 3.1 chuẩn APA 7th, không còn số [10]/[11] hay tác giả khuyết
    _, tbl_31 = _find_table_after_header(children, "Tác giả | Năm | Phương pháp nghiên cứu")
    t31_rows = tbl_31.findall(_w("tr"))
    _set_cell_text(t31_rows[1].findall(_w("tc"))[1], "Makridakis, Spiliotis và Assimakopoulos")
    _set_cell_text(t31_rows[2].findall(_w("tc"))[1], "Shah và Dimitrov")
    _set_cell_text(t31_rows[3].findall(_w("tc"))[1], "Falatouri, Darbanian, Brandtner và Udokwu")
    _set_cell_text(t31_rows[4].findall(_w("tc"))[1], "Mitra, Jain, Kishore và Kumar")
    _set_cell_text(t31_rows[5].findall(_w("tc"))[1], "Nguyễn Thị Liễu Điền")
    _set_cell_text(t31_rows[6].findall(_w("tc"))[1], "Aye, Assad Neto và Ribeiro da Silva")
    _set_cell_text(t31_rows[6].findall(_w("tc"))[2], "2025")
    _set_cell_text(t31_rows[6].findall(_w("tc"))[3], "Mô hình học máy tích hợp biến ngoại sinh (Exogeneity-aware ML)")
    _set_cell_text(t31_rows[6].findall(_w("tc"))[4], "Tích hợp có chọn lọc khuyến mãi, thời tiết và biến lịch giúp giảm sai số dự báo tới 22% so với phương pháp truyền thống tại nhà bán lẻ châu Âu.")
    _set_cell_text(t31_rows[7].findall(_w("tc"))[1], "Carbonneau, Laframboise và Vahidov")
    _set_cell_text(t31_rows[7].findall(_w("tc"))[2], "2008")
    _set_cell_text(t31_rows[7].findall(_w("tc"))[3], "Mạng nơ-ron, RNN, SVM, hồi quy tuyến tính và trung bình trượt")
    _set_cell_text(t31_rows[7].findall(_w("tc"))[4], "Các thuật toán học máy phi tuyến xử lý tốt tín hiệu nhu cầu biến dạng trong chuỗi cung ứng, nhưng hiệu quả phụ thuộc cấu trúc dữ liệu.")

    # Bảng 3.2 row 4 ("Syntetos và cộng sự (2021)" -> "Boylan và Syntetos (2021)")
    _, tbl_32 = _find_table_after_header(children, "Công trình / hướng tiếp cận | Đóng góp chính")
    t32_rows = tbl_32.findall(_w("tr"))
    _set_cell_text(t32_rows[4].findall(_w("tc"))[0], "Boylan và Syntetos (2021)")

    # ------------------------------------------------------------------ #
    # 4. CHƯƠNG 4: ĐẦY ĐỦ ĐỐI CHỨNG (NAIVE, NO-INVENTORY, ĐỆ QUY 141 BƯỚC, LJUNG-BOX, DM)
    # ------------------------------------------------------------------ #
    _, p_stat_conc = _find_paragraph_by_prefix(children, "Kết quả phân tích hình 4.1, hình 4.3 và bảng 4.1")
    _set_plain_text_preserve_format(
        p_stat_conc,
        _p_text(p_stat_conc)
        + " Lưu ý về bậc sai phân d = 1: mặc dù hai kiểm định ADF (p < 0,0001) và KPSS (p >= 0,1000) cùng không bác bỏ giả thuyết chuỗi dừng quanh trung bình ổn định, bậc sai phân d = 1 vẫn được chọn theo hướng thực nghiệm cho ARIMA(2,1,1) và SARIMA(1,1,1)(1,0,1)7 (Hyndman & Athanasopoulos, 2021). Do chưa áp dụng tìm kiếm tự động (auto_arima) hay so sánh chỉ số AIC giữa d = 0 và d = 1, cấu hình này tiềm ẩn rủi ro sai phân thừa (over-differencing).",
    )

    _, p_arima_cfg = _find_paragraph_by_prefix(children, "Cấu hình: bậc tự tương quan p = 2")
    _set_plain_text_preserve_format(
        p_arima_cfg,
        "Cấu hình: bậc tự tương quan p = 2, bậc sai phân d = 1 (lựa chọn thực nghiệm dù ADF/KPSS gợi ý chuỗi dừng), bậc trung bình trượt q = 1. Kết quả trên tập kiểm thử (dự báo 141 bước): RMSE = 107,07, MAE = 90,04, MAPE = 221,29%. Kiểm định Ljung-Box trên phần dư tập huấn luyện đạt Q(7) = 3,02 (p = 0,8827) và Q(14) = 7,68 (p = 0,9051); với p > 0,05, kiểm định không bác bỏ giả thuyết không về tính độc lập của phần dư ở độ trễ 7 và 14 ngày.",
    )

    _, p_sarima_cfg = _find_paragraph_by_prefix(children, "Cấu hình: thành phần phi mùa vụ (1, 1, 1)")
    _set_plain_text_preserve_format(
        p_sarima_cfg,
        "Cấu hình: thành phần phi mùa vụ (1, 1, 1) với d = 1 lựa chọn thực nghiệm, kết hợp thành phần mùa vụ (1, 0, 1) theo chu kỳ s = 7 ngày. Kết quả trên tập kiểm thử (dự báo 141 bước): RMSE = 107,03, MAE = 90,03, MAPE = 221,23%. Kiểm định Ljung-Box trên phần dư tập huấn luyện đạt Q(7) = 3,03 (p = 0,8818) và Q(14) = 7,34 (p = 0,9208), tức chưa có bằng chứng bác bỏ giả thuyết phần dư độc lập theo thời gian ở mức ý nghĩa 5%.",
    )

    _, p_rf_exp = _find_paragraph_by_prefix(children, "Đường dự báo của Random Forest bám sát")
    _set_plain_text_preserve_format(
        p_rf_exp,
        "Đường dự báo một bước của Random Forest (28 đặc trưng) bám sát các bước biến động của chuỗi thực tế nhờ tổ hợp 100 cây quyết định và tập biến ngoại cảnh. Trong điều kiện biết trước các biến đầu vào của ngày dự báo (gồm tồn kho đầu ngày Inventory Level) và biến trễ thực tới ngày t-1, RMSE giảm 13,65% so với SARIMA đa bước (từ 107,03 xuống 92,42 đơn vị) và giảm 13,68% so với ARIMA (107,07). Tuy nhiên, kết quả thực nghiệm loại bỏ đặc trưng và kiểm thử đệ quy thuần ở mục 4.3.4 cho thấy lợi thế này phụ thuộc mạnh vào biến Inventory Level và giao thức dự báo một bước.",
    )

    # 4.3.4 & Bảng 4.5 + Bảng 4.6
    _, p_434_intro = _find_paragraph_by_prefix(children, "Năm mô hình được đánh giá trên cùng tập kiểm thử 141 ngày")
    _set_plain_text_preserve_format(
        p_434_intro,
        "Năm mô hình chính được đối chiếu cùng hai mô hình cơ sở (Naive dự báo phẳng và Seasonal Naive chu kỳ 7 ngày) và hai thí nghiệm kiểm chứng giao thức cho nhóm học máy: (i) loại bỏ cột Inventory Level (còn 27 đặc trưng, dự báo 1 bước) và (ii) dự báo đệ quy thuần 141 bước từ điểm cắt 13/08/2023 (giả định biến ngoại sinh bằng trung bình 14 ngày cuối tập huấn luyện, giống cơ chế vận hành của Web Studio). Tất cả được đo trên cùng 141 ngày kiểm thử (14/08/2023 - 01/01/2024). Lưu ý: MAPE của toàn bộ mô hình bị phóng đại lên 173,29%–344,00% do mẫu số nhỏ ở các ngày có nhu cầu gần bằng 0, nên đánh giá chính dựa vào RMSE và MAE.",
    )

    _, tbl_45 = _find_table_after_header(children, "STT | Mô hình | Cấu hình tham số | RMSE | MAE | MAPE")
    t45_rows = tbl_45.findall(_w("tr"))
    tpl_45 = copy.deepcopy(t45_rows[1])
    for tr in t45_rows[1:]:
        tbl_45.remove(tr)

    rows_45_data = [
        ["1", "Random Forest (28 biến, 1 bước)", "100 cây, depth 12, dùng đủ Inventory Level & ngoại sinh ngày t", "92,42", "73,71", "173,29%", "1 (nhóm 1 bước)"],
        ["2", "XGBoost (28 biến, 1 bước)", "150 cây, lr=0,05, depth 6, dùng đủ Inventory Level & ngoại sinh ngày t", "95,73", "76,24", "181,99%", "2 (nhóm 1 bước)"],
        ["3", "SARIMA (Đa bước 141 ngày)", "(1,1,1)(1,0,1)7, đơn biến, không cập nhật thực tế", "107,03", "90,03", "221,23%", "1 (nhóm đa bước)"],
        ["4", "ARIMA (Đa bước 141 ngày)", "p=2, d=1, q=1, đơn biến, không cập nhật thực tế", "107,07", "90,04", "221,29%", "2 (nhóm đa bước)"],
        ["5", "Naive (Đa bước 141 ngày)", "Dự báo phẳng y_last = 145 (cơ sở tham chiếu)", "107,52", "91,48", "231,07%", "3 (nhóm đa bước)"],
        ["6", "Exponential Smoothing (Đa bước)", "Holt-Winters cộng tính, s=7 (cơ sở theo mùa)", "108,21", "90,23", "219,37%", "4 (nhóm đa bước)"],
        ["7", "Random Forest (27 biến, bỏ Inv)", "100 cây, depth 12, bỏ Inventory Level (ablation 1 bước)", "113,19", "93,75", "229,77%", "Đối chứng bỏ Inv"],
        ["8", "XGBoost (27 biến, bỏ Inv)", "150 cây, lr=0,05, depth 6, bỏ Inventory Level (ablation 1 bước)", "114,29", "92,47", "216,71%", "Đối chứng bỏ Inv"],
        ["9", "XGBoost (Đệ quy thuần 141 bước)", "150 cây, dự báo đa bước từ 13/08/2023 (ngoại sinh TB 14 ngày)", "116,85", "101,89", "291,11%", "5 (nhóm đa bước)"],
        ["10", "Random Forest (Đệ quy thuần 141 bước)", "100 cây, dự báo đa bước từ 13/08/2023 (ngoại sinh TB 14 ngày)", "119,09", "104,91", "304,03%", "6 (nhóm đa bước)"],
        ["11", "Seasonal Naive (Đa bước 141 ngày)", "Lặp chu kỳ 7 ngày cuối tập train (cơ sở mùa vụ)", "197,64", "154,55", "344,00%", "7 (nhóm đa bước)"],
    ]
    for rdata in rows_45_data:
        tbl_45.append(_clone_tr_with_values(tpl_45, rdata))

    _, p_cap45 = _find_paragraph_by_prefix(children, "Bảng 4.5: Đối sánh định lượng")
    _set_plain_text_preserve_format(
        p_cap45,
        "Bảng 4.5: Đối sánh hiệu năng dự báo giữa các mô hình chính, mô hình cơ sở và thí nghiệm đối chứng (141 ngày kiểm thử, S001–P0001)",
    )
    for el in children[:140]:
        if el.tag == _w("p") and _p_text(el).startswith("Bảng 4.5:"):
            _replace_in_wt_runs(
                el,
                "của 5 mô hình trên tập kiểm thử",
                "giữa các mô hình chính, cơ sở và đối chứng (S001–P0001)",
            )

    _, p_434_c1 = _find_paragraph_by_prefix(children, "Thứ nhất, trên cặp SKU và cửa sổ kiểm thử")
    _set_plain_text_preserve_format(
        p_434_c1,
        "Thứ nhất, kết quả đối sánh cho thấy sự khác biệt rõ rệt tùy thuộc giao thức đánh giá và vai trò của biến Inventory Level. Khi đánh giá ở chế độ một bước với đủ 28 đặc trưng, Random Forest đạt RMSE = 92,42 và XGBoost đạt RMSE = 95,73 (thấp hơn 13,65% và 10,56% so với SARIMA 141 bước = 107,03). Tuy nhiên, khi loại bỏ Inventory Level (còn 27 đặc trưng, vẫn ở chế độ 1 bước), RMSE của Random Forest tăng vọt lên 113,19 và XGBoost tăng lên 114,29 — đều kém hơn cả SARIMA (107,03), ARIMA (107,07) và Naive (107,52). Hơn nữa, nếu đặt nhóm học máy vào đúng giao thức dự báo đệ quy thuần 141 bước (không biết trước ngoại sinh tập test, cùng điều kiện với nhóm thống kê và cùng cơ chế mà Web Studio sử dụng để tạo D_hat = 155,1), RMSE của XGBoost và Random Forest tăng lên lần lượt là 116,85 và 119,09. Ở bài toán đa bước thuần, các mô hình đơn biến ít tham số (SARIMA 107,03; ARIMA 107,07; Naive 107,52) tỏ ra ổn định hơn do chuỗi nhu cầu có tính dừng quanh mức trung bình ~137–145 đơn vị, trong khi Seasonal Naive bị nhiễu nặng (RMSE = 197,64) do lặp lại hai đỉnh đột biến 335 và 455 đơn vị ở tuần cuối tập train.",
    )

    _, p_434_c2 = _find_paragraph_by_prefix(children, "Thứ hai, Random Forest đạt thứ hạng cao nhất")
    _set_plain_text_preserve_format(
        p_434_c2,
        "Thứ hai, tệp mô hình random_forest_model.pkl được giữ làm mô hình mặc định cho giao diện nguyên mẫu Chương 5 nhằm khai thác khả năng kết hợp biến ngoại sinh khi dự báo ngắn hạn khi người dùng biết tồn kho đầu kỳ và kế hoạch giá/khuyến mãi, nhưng không còn được khẳng định là vượt trội tuyệt đối trong mọi giao thức như nhận định ban đầu.",
    )

    # 4.3.5 Hạn chế (viết trung thực, nhấn mạnh cross-protocol & DM caveat & multi-step)
    _, p_435_1 = _find_paragraph_by_prefix(children, "Sau khi diễn giải kết quả, có một số hạn chế")
    _set_plain_text_preserve_format(
        p_435_1,
        "Từ kết quả Bảng 4.5, thực nghiệm bộc lộ năm hạn chế phương pháp luận quan trọng. Thứ nhất (nghiêm trọng nhất) là giao thức đánh giá không đồng nhất giữa hai nhóm mô hình: trong src/data_loader.py (prepare_ml_data / engineer_features), các biến trễ (lag_1..lag_30) và thống kê trượt trên 141 ngày kiểm thử của RF và XGBoost được tính từ Units Sold thực tế (dự báo một bước — one-step-ahead, kết hợp ngoại sinh ngày t đã biết), trong khi ARIMA, SARIMA, Holt-Winters, Naive và Seasonal Naive dự báo liên tục 141 bước mà không thấy dữ liệu trung gian. Vì vậy, mọi so sánh và cả kiểm định Diebold-Mariano giữa RF một bước (92,42) với SARIMA 141 bước (107,03) là phép so sánh chéo khác giao thức, chưa đủ căn cứ thống kê để khẳng định mô hình ML vượt trội trên cùng một bài toán. Đáng chú ý, giá trị D_hat = 155,1 đơn vị/ngày đưa vào tính SS/ROP ở Mục 4.5 lại là dự báo đa bước đệ quy (giao thức có RMSE = 119,09 trên kiểm thử 141 ngày).",
    )

    _, p_435_2 = _find_paragraph_by_prefix(children, "Thứ hai, tập kiểm thử được dùng để so sánh")
    _set_plain_text_preserve_format(
        p_435_2,
        "Thứ hai, phép thử loại bỏ đặc trưng xác nhận độ nhạy quá mức đối với cột Inventory Level (chiếm 42,33% tầm quan trọng trong RF): chỉ cần bỏ cột này khỏi ma trận đầu vào, RMSE của RF một bước tăng từ 92,42 lên 113,19 (chênh lệch Diebold-Mariano DM = -4,03; p < 0,0001) và thua SARIMA có ý nghĩa thống kê (DM = +2,99; p = 0,0033). Kết quả chứng tỏ phần lớn độ chính xác của RF trong bộ dữ liệu Kaggle đến từ ràng buộc dữ liệu mô phỏng giữa tồn kho và sản lượng bán hơn là quy luật thời gian tự thân.",
    )

    _, p_435_3 = _find_paragraph_by_prefix(children, "Thứ ba, phạm vi thực nghiệm mới bao gồm một cặp SKU")
    _set_plain_text_preserve_format(
        p_435_3,
        "Thứ ba, thực nghiệm mới dừng ở 1 cặp SKU (S001–P0001) với một lần chia cố định 560/141 ngày (chưa có tập validation riêng, chưa rolling-origin backtesting trên 100 cặp Store × SKU). Khi thực hiện kiểm định Diebold-Mariano (hiệu chỉnh Harvey, sai số bình phương MSE) giữa các cặp cùng giao thức, khác biệt giữa RF một bước và XGBoost một bước (92,42 vs 95,73) không đạt ý nghĩa thống kê ở mức 5% (DM = -1,78; p = 0,0769 theo MSE; p = 0,1332 theo MAE), tương tự ARIMA và SARIMA cũng không khác biệt có ý nghĩa thống kê (DM = 0,47; p = 0,6365). Riêng phép so giữa RF một bước và SARIMA đa bước (DM = -3,24; p = 0,0015) chỉ mang tính tham khảo do lệch giao thức; trên cùng giao thức đa bước thuần 141 ngày, RF (119,09) lại cao hơn SARIMA (DM = +3,42; p = 0,0008).",
    )

    _, p_435_5 = _find_paragraph_by_prefix(children, "Các hạn chế trên không làm mất giá trị")
    _set_plain_text_preserve_format(
        p_435_5,
        "Do đó, để hoàn thiện thành nghiên cứu sản phẩm thực tế, hệ thống cần được đánh giá lại bằng giao thức dự báo đa bước theo từng chân trời dự báo (1–3 ngày trùng Lead Time và 7–14 ngày) kết hợp rolling-origin trên toàn bộ 100 cặp Store × SKU, bổ sung mô phỏng đóng vòng để đo chi phí lưu kho và tỷ lệ thiếu hàng thực tế thay vì chỉ tính công thức tĩnh.",
    )

    # Bảng 4.6 (đồng bộ tất cả mô hình, cơ sở và đối chứng)
    _, tbl_46 = _find_table_after_header(children, "Nhóm mô hình | Mô hình dự báo | RMSE")
    t46_rows = tbl_46.findall(_w("tr"))
    tpl_46 = copy.deepcopy(t46_rows[1])
    for tr in t46_rows[1:]:
        tbl_46.remove(tr)
    rows_46_data = [
        ["ML (1 bước, đủ 28 biến)", "Random Forest", "92,42", "73,71", "173,29%", "+13,65% (lệch giao thức)", "1 (nhóm 1 bước)"],
        ["ML (1 bước, đủ 28 biến)", "XGBoost", "95,73", "76,24", "181,99%", "+10,56% (lệch giao thức)", "2 (nhóm 1 bước)"],
        ["Thống kê (141 bước)", "SARIMA", "107,03", "90,03", "221,23%", "0,00% (mốc chuẩn)", "1 (nhóm đa bước)"],
        ["Thống kê (141 bước)", "ARIMA", "107,07", "90,04", "221,29%", "-0,04%", "2 (nhóm đa bước)"],
        ["Cơ sở (141 bước)", "Naive (y_last = 145)", "107,52", "91,48", "231,07%", "-0,46%", "3 (nhóm đa bước)"],
        ["Thống kê (141 bước)", "Exponential Smoothing", "108,21", "90,23", "219,37%", "-1,10%", "4 (nhóm đa bước)"],
        ["ML Ablation (bỏ Inv)", "Random Forest (27 biến)", "113,19", "93,75", "229,77%", "-5,76%", "Đối chứng bỏ Inv"],
        ["ML Ablation (bỏ Inv)", "XGBoost (27 biến)", "114,29", "92,47", "216,71%", "-6,78%", "Đối chứng bỏ Inv"],
        ["ML (Đệ quy 141 bước)", "XGBoost (đa bước thuần)", "116,85", "101,89", "291,11%", "-9,17%", "5 (nhóm đa bước)"],
        ["ML (Đệ quy 141 bước)", "Random Forest (đa bước thuần)", "119,09", "104,91", "304,03%", "-11,27%", "6 (nhóm đa bước)"],
        ["Cơ sở (141 bước)", "Seasonal Naive (s=7)", "197,64", "154,55", "344,00%", "-84,66%", "7 (nhóm đa bước)"],
    ]
    for rdata in rows_46_data:
        tbl_46.append(_clone_tr_with_values(tpl_46, rdata))

    # 4.4.1 / 4.4.2 / Hình 4.10 / 4.4.3 (diễn đạt chừng mực, không khẳng định vượt quá bằng chứng)
    _, p_441_ml = _find_paragraph_by_prefix(children, "Nhóm mô hình học máy (Random Forest, XGBoost)")
    _set_plain_text_preserve_format(
        p_441_ml,
        "Nhóm mô hình học máy (Random Forest, XGBoost) đạt sai số thấp hơn nhóm thống kê cổ điển khi đánh giá ở chế độ một bước với đầy đủ biến ngoại sinh đầu ngày (Random Forest đạt RMSE = 92,42, MAE = 73,71; XGBoost đạt RMSE = 95,73, MAE = 76,24; khác biệt giữa hai mô hình cây không có ý nghĩa thống kê ở mức 5% với DM = -1,78, p = 0,0769). Tuy nhiên, khi loại bỏ Inventory Level (RMSE tăng lên 113,19 và 114,29) hoặc khi ép dự báo đệ quy 141 bước không có ngoại sinh tương lai (RMSE tăng lên 119,09 và 116,85), các mô hình cây đều có sai số cao hơn nhóm thống kê đơn biến.",
        strip_omath=True,
    )

    _, p_442_1 = _find_paragraph_by_prefix(children, "Xác lập mô hình tốt nhất với mức giảm sai số vượt trội:")
    _set_plain_text_preserve_format(
        p_442_1,
        "Kết quả thực nghiệm có điều kiện trên SKU S001–P0001: Ở giao thức dự báo một bước với đủ 28 đặc trưng, Random Forest đạt RMSE = 92,42, MAE = 73,71 và MAPE = 173,29% (thấp hơn 13,65% RMSE so với SARIMA 141 bước có RMSE = 107,03, MAE = 90,03, và thấp hơn 18,13% MAE). Vì đây là phép so giữa hai giao thức khác nhau và phụ thuộc mạnh vào biến Inventory Level, kết quả này chỉ cho thấy tiềm năng của mô hình học máy trong dự báo rất ngắn hạn khi đã quan sát tồn kho đầu kỳ và thông tin giá, không hàm ý ưu thế phổ quát ở chân trời dài hạn.",
        strip_omath=True,
    )

    _, p_442_2 = _find_paragraph_by_prefix(children, "Nắm bắt xuất sắc tác động phi tuyến")
    _set_plain_text_preserve_format(
        p_442_2,
        "Phân tích mức độ đóng góp của đặc trưng trên mô hình cây: Biểu đồ xếp hạng tầm quan trọng biến dưới đây trình bày tỷ trọng phân tách của 28 đặc trưng đầu vào đối với hai mô hình phi tuyến.",
    )

    _, p_410_1 = _find_paragraph_by_prefix(children, "Cụ thể, ngoài yếu tố nguồn cung sẵn có là mức tồn kho")
    _set_plain_text_preserve_format(
        p_410_1,
        "Trong Hình 4.10, hai mô hình cây phân bổ trọng số lớn nhất cho thuộc tính tồn kho Inventory Level (Gain ≈ 0,082 hay 8,18% ở XGBoost; 42,33% ở Random Forest), theo sau là tỷ lệ giá chiết khấu (price_ratio), độ trễ lag_21, lag_30, giá bán thực tế (effective_price) và các thống kê trượt 7–14 ngày (rolling_std_14, rolling_std_7). Khi loại bỏ Inventory Level trong thí nghiệm kiểm chứng, top 5 biến quan trọng chuyển sang lag_2 (7,61%), rolling_std_14 (7,35%), price_ratio (7,16%), rolling_std_7 (6,54%) và lag_7 (6,25%), đồng thời RMSE của Random Forest tăng từ 92,42 lên 113,19.",
        strip_omath=True,
    )

    _, p_410_2 = _find_paragraph_by_prefix(children, "Lưu ý đối chiếu giữa hai mô hình cây quyết định:")
    _set_plain_text_preserve_format(
        p_410_2,
        "Giải thích cơ chế biến Inventory Level: Thuộc tính Inventory Level phản ánh lượng tồn kho ghi nhận vào đầu ngày (trước khi bán hàng trong ngày t), nên về mặt nhân quả không phải là rò rỉ nhãn mục tiêu cùng ngày. Tuy vậy, trong bộ dữ liệu mô phỏng Kaggle, sản lượng bán trong ngày bị chặn trên trực tiếp bởi lượng tồn kho sẵn có (hệ số tương quan r ≈ 0,59), tạo nên mối liên kết rất mạnh trong mô phỏng. Phép thử loại bỏ Inventory Level (RMSE tăng +20,77 đơn vị) chứng minh mô hình phụ thuộc chủ yếu vào biến này, và khi dự báo nhiều bước về tương lai mà tồn kho đầu ngày t+k chưa biết trước, độ chính xác của mô hình học máy giảm đáng kể.",
    )

    _, p_443_mape = _find_paragraph_by_prefix(children, "Chỉ số MAPE duy trì ở mức cao")
    _replace_in_wt_runs(p_443_mape, "Trong tập dữ liệu bán lẻ thực tế,", "Trong tập dữ liệu bán lẻ mô phỏng,")
    _replace_in_wt_runs(p_443_mape, "221.23%", "221,23%")
    _replace_in_wt_runs(p_443_mape, "173.29%", "173,29%")

    _, p_443_rec = _find_paragraph_by_prefix(children, "Thiếu cơ chế dự báo đệ quy cho bài toán dài hạn:")
    _set_plain_text_preserve_format(
        p_443_rec,
        "Sự chênh lệch giữa đánh giá một bước và dự báo đệ quy đa bước: Như kết quả kiểm chứng ở Bảng 4.5, khi chuyển Random Forest từ đánh giá một bước (RMSE = 92,42) sang dự báo đệ quy thuần 141 bước không có ngoại sinh tương lai (cơ chế được src/forecast_service.py dùng cho Web Studio), RMSE tăng lên 119,09 do sai số tích lũy qua từng bước đệ quy và do thiếu biến ngoại sinh cập nhật.",
        strip_omath=True,
    )

    # ------------------------------------------------------------------ #
    # 5. CHƯƠNG 4.5: SỬA SỐ HỌC Q = 943, GIẢI THÍCH 1,5×ROP, TÁI TẠO 707/721, HÌNH 4.11 & ABC-XYZ
    # ------------------------------------------------------------------ #
    _, p_440 = _find_paragraph_by_prefix(children, "Chế độ tích hợp dự báo ML Random Forest")
    _replace_in_wt_runs(
        p_440,
        "trung bình 7 ngày tới",
        "trung bình trong 3 ngày giao hàng tới (xuất ra bởi cơ chế đệ quy đa bước của Random Forest trong src/forecast_service.py — giao thức có RMSE = 119,09 khi thử 141 bước)",
    )
    _replace_in_wt_runs(
        p_440,
        " = 942 đơn vị.",
        " = ⌈942,5⌉ = 943 đơn vị (ghi chú: mã nguồn src/inventory_optimizer.py dùng giá trị ROP chưa làm tròn là 776,27 nên tính ra ⌈1,5 × 776,27 − 223⌉ = 942 đơn vị). Trong nguyên mẫu khi chưa có chi phí đặt hàng K và chi phí lưu kho h, hệ số 1,5 × ROP là quy tắc thực nghiệm (s, S) tạo vùng đệm bổ sung 50% phía trên ROP.",
    )

    _, p_441 = _find_paragraph_by_prefix(children, "Chế độ tham chiếu cơ sở theo trung bình lịch sử 90 ngày:")
    _set_plain_text_preserve_format(
        p_441,
        "Đối chiếu hai đường tham chiếu cơ sở và lý giải Hình 4.11: (1) Trong biểu đồ tĩnh Hình 4.11 (xuất từ main_train_and_evaluate.py), ngưỡng SS = 312 đơn vị và ROP = 707 đơn vị được tính từ nhu cầu dự báo trung bình của Random Forest trên 141 ngày kiểm thử là D_bar_test = 131,50 đơn vị/ngày cùng độ lệch chuẩn sigma_D = 109,15, L = 3 ngày và hệ số Z làm tròn là 1,65: SS = ⌈1,65 × 109,15 × sqrt(3)⌉ = 312 và ROP = ⌈131,50 × 3 + 311,95⌉ = 707 đơn vị. (2) Trên Web Studio, chế độ cơ sở theo trung bình lịch sử 90 ngày cuối dùng D_90 = 136,50 đơn vị/ngày và Z = 1,645, cho SS = 311 đơn vị và ROP = ⌈136,50 × 3 + 310,97⌉ = 721 đơn vị (Q = 858 đơn vị). Trên Hình 4.11, đường tồn kho lịch sử (55–480 đơn vị) nằm hoàn toàn bên dưới cả hai ngưỡng ROP (707 và 777) trong suốt 60 ngày vì trong dữ liệu mô phỏng Kaggle, cột Inventory Level được sinh ngẫu nhiên trong đoạn [50, 500] cho quy mô 1 ngày chứ không phản ánh tồn kho cho chuỗi cung ứng có Lead Time L = 3 ngày (đòi hỏi 3 × D + SS ≈ 707–777 đơn vị). Đồ án hiện mới dừng ở tính toán công thức tĩnh, chưa chạy mô phỏng tồn kho đóng vòng theo ngày để đo lường tỷ lệ thiếu hàng thực tế, chi phí lưu kho hay Service Level đạt được.",
    )

    _, p_444 = _find_paragraph_by_prefix(children, "Đồng thời, hệ thống quét toàn bộ 20 mã sản phẩm")
    _set_plain_text_preserve_format(
        p_444,
        "Đồng thời, hệ thống lập Ma trận phân loại ABC - XYZ cho 20 mã sản phẩm tại cửa hàng S001 trong 731 ngày (tổng doanh thu $97.426.189, Hình 4.12 và Bảng 4.7). Với bộ ngưỡng cố định (X <= 0,40; 0,40 < Y <= 0,70; Z > 0,70), cả 20 mã đều có hệ số biến thiên CV tập trung rất hẹp trong khoảng 0,758–0,834 (trung vị 0,807), khiến toàn bộ danh mục dồn vào duy nhất cột Z (15 AZ, 3 BZ, 2 CZ). Như vậy, ở thiết lập ngưỡng hiện tại, chiều phân loại XYZ bị suy biến và không tạo ra giá trị phân biệt giữa các mặt hàng do quy luật sinh dữ liệu mô phỏng Kaggle tạo độ phân tán tương đối đồng nhất cho mọi SKU. Nếu áp dụng ngưỡng phổ biến hơn cho nhu cầu thất thường trong y văn như CV > 1,0 (Syntetos và cộng sự, 2016) thì toàn bộ 20 mã sẽ rơi vào nhóm Y, hoặc nếu hiệu chỉnh ngưỡng thực nghiệm tại trung vị CV = 0,80 mới tách được danh mục thành 7 mã (CV <= 0,80) và 13 mã (CV > 0,80).",
    )

    _, tbl_47 = _find_table_after_header(children, "Nhóm | Số lượng SKU | Tỷ trọng DT")
    t47_rows = tbl_47.findall(_w("tr"))
    _set_cell_text(
        t47_rows[4].findall(_w("tc"))[5],
        "Cả 20 SKU dồn vào cột Z (ma trận XYZ bị suy biến dưới ngưỡng CV > 0,70); cần hiệu chỉnh ngưỡng hoặc dùng CV > 1,0.",
    )

    # ------------------------------------------------------------------ #
    # 6. CHƯƠNG 6 & BẢNG 6.1: SỬA TRUNG THỰC, BỎ DIỄN ĐẠT QUÁ MỨC
    # ------------------------------------------------------------------ #
    _, p_61_1 = _find_paragraph_by_prefix(children, "Đề tài đã hoàn thành mục tiêu đặt ra")
    _set_plain_text_preserve_format(
        p_61_1,
        "Đề tài đã xây dựng được một quy trình nguyên mẫu từ phân tích dữ liệu (EDA), trích xuất đặc trưng, huấn luyện đối chứng mô hình dự báo đến giao diện hỗ trợ ra quyết định tồn kho trên bộ dữ liệu mô phỏng công khai Kaggle (Retail Store Inventory Forecasting Dataset, 73.100 bản ghi). Trên cặp SKU S001–P0001 (560 ngày huấn luyện / 141 ngày kiểm thử), Random Forest (28 đặc trưng) đạt RMSE = 92,42 và MAE = 73,71 khi đánh giá theo giao thức một bước với biến tồn kho đầu ngày Inventory Level đã biết — thấp hơn 13,65% về RMSE so với SARIMA 141 bước (RMSE = 107,03). Tuy nhiên, kết luận này mang tính có điều kiện và chưa thể khẳng định ưu thế tuyệt đối: hai nhóm mô hình được đo ở hai giao thức khác nhau; khi bỏ biến Inventory Level (chiếm 42,33% tầm quan trọng), RMSE của Random Forest tăng lên 113,19; và khi chạy dự báo đệ quy thuần 141 bước (cùng giao thức với nhóm thống kê và là cơ chế tạo ra D_hat = 155,1 đưa vào tính tồn kho), RMSE của Random Forest tăng lên 119,09, cao hơn SARIMA (107,03), ARIMA (107,07) và Naive (107,52).",
    )

    _, p_61_2 = _find_paragraph_by_prefix(children, "Quan trọng hơn, nghiên cứu không chỉ dừng ở bước dự báo")
    _set_plain_text_preserve_format(
        p_61_2,
        "Ở phân hệ tồn kho và phần mềm, hệ thống đã cài đặt các công thức định mức tính Tồn kho an toàn (SS = 311 đơn vị), Điểm đặt hàng lại (ROP = 777 đơn vị theo D_hat = 155,1 ở L = 3 ngày, SL = 95%) và lượng đặt bổ sung đề xuất Q = 943 đơn vị theo quy tắc thực nghiệm mục tiêu S = 1,5 × ROP trên Web Studio Retail ML Hub (6 màn hình tương tác, cho phép xuất bảng dữ liệu CSV). Tuy vậy, đề tài chưa mô phỏng tồn kho đóng vòng qua thời gian nên chưa định lượng được chi phí lưu kho, tỷ lệ thiếu hàng hay Service Level đạt được thực tế, đồng thời ma trận ABC-XYZ đang bị suy biến do toàn bộ 20 SKU đều rơi vào cột Z (CV từ 0,758 đến 0,834 > 0,70).",
    )

    _, tbl_61 = _find_table_after_header(children, "STT | Mục tiêu nghiên cứu | Kết quả đạt được")
    t61_rows = tbl_61.findall(_w("tr"))
    _set_cell_text(t61_rows[1].findall(_w("tc"))[2], "Phân tích 73.100 giao dịch mô phỏng Kaggle, 7 biểu đồ EDA, Sơ đồ 1.1 & Bảng 1.1–1.3.")
    _set_cell_text(t61_rows[1].findall(_w("tc"))[3], "Hoàn thành (dữ liệu mô phỏng)")

    _set_cell_text(t61_rows[2].findall(_w("tc"))[2], "Triển khai 5 mô hình chính + 2 cơ sở Naive + thử nghiệm bỏ Inventory Level & đệ quy 141 bước trên 1 SKU (S001–P0001).")
    _set_cell_text(t61_rows[2].findall(_w("tc"))[3], "Hoàn thành ở mức 1 SKU (chưa chạy hết 100 SKU)")

    _set_cell_text(t61_rows[3].findall(_w("tc"))[2], "RF đạt RMSE = 92,42 ở 1 bước có Inventory Level; tăng lên 113,19 khi bỏ Inventory Level và 119,09 ở đệ quy thuần 141 bước (kém SARIMA 107,03).")
    _set_cell_text(t61_rows[3].findall(_w("tc"))[3], "Còn lệch giao thức (1 bước vs 141 bước)")

    _set_cell_text(t61_rows[4].findall(_w("tc"))[2], "Tính được SS = 311, ROP = 777, Q = 943 theo heuristic S = 1,5 × ROP; ma trận ABC-XYZ suy biến 100% vào cột Z (CV > 0,70).")
    _set_cell_text(t61_rows[4].findall(_w("tc"))[3], "Hoàn thành công thức tĩnh (XYZ chưa phân biệt)")

    _set_cell_text(t61_rows[5].findall(_w("tc"))[2], "Xây dựng ứng dụng nguyên mẫu 6 màn hình (mô phỏng độ nhạy L và SL, xuất CSV); chưa đo chi phí lưu kho, thiếu hàng hay vốn lưu động thực tế.")
    _set_cell_text(t61_rows[5].findall(_w("tc"))[3], "Đạt nguyên mẫu Web (chưa đo hiệu quả kinh doanh)")

    _, p_62_1 = _find_paragraph_by_prefix(children, "Về mặt khoa học, nghiên cứu cung cấp")
    _set_plain_text_preserve_format(
        p_62_1,
        "Về mặt khoa học, thực nghiệm đối chiếu cho thấy vai trò quyết định của giao thức đánh giá và lựa chọn đặc trưng trong dự báo chuỗi thời gian bán lẻ: các thuật toán học máy tập hợp chỉ đạt sai số thấp hơn mô hình thống kê khi được cung cấp tín hiệu ngoại sinh và biến trễ gần thực tế theo bước 1 ngày, trong khi dễ bị nhạy cảm với biến rào cản nguồn cung (Inventory Level) và tích lũy sai số khi dự báo đệ quy nhiều bước.",
    )

    _, p_62_2 = _find_paragraph_by_prefix(children, "Về mặt thực tiễn, giải pháp mang lại")
    _set_plain_text_preserve_format(
        p_62_2,
        "Về mặt thực tiễn, ứng dụng nguyên mẫu Retail ML Hub minh họa quy trình nối kết đầu ra dự báo sang công thức định mức tồn kho (SS, ROP) và phân tích độ nhạy What-If, làm tiền đề cho các bước tích hợp chi phí thực tế (chi phí đặt hàng K, chi phí lưu kho h) và kiểm thử đóng vòng trước khi triển khai doanh nghiệp.",
    )

    # ------------------------------------------------------------------ #
    # 7. TÁI CẤU TRÚC TOÀN BỘ MỤC TÀI LIỆU THAM KHẢO THEO CHUẨN APA (7th ed.)
    # ------------------------------------------------------------------ #
    ref_heading_idx, _ = _find_paragraph_by_prefix(children, "TÀI LIỆU THAM KHẢO", start_idx=139)
    existing_refs = [
        p for p in children[ref_heading_idx + 1:]
        if p.tag == _w("p") and _p_text(p).strip()
    ]
    # Bỏ đoạn giới thiệu cũ "Danh mục các công trình nghiên cứu khoa học..." nếu có ở ref_heading_idx + 1
    apa_list = _apa_references()  # 15 mục APA
    # Lọc các đoạn [1]..[14] cũ (bỏ dòng giới thiệu đầu)
    old_ref_entries = [p for p in existing_refs if re.match(r"^\[\d+\]", _p_text(p).strip())]
    tpl_ref = copy.deepcopy(old_ref_entries[0])

    # Xóa toàn bộ danh sách cũ [1]..[14]
    for p in old_ref_entries:
        body.remove(p)

    # Ghi 15 tài liệu APA 7th mới ngay sau đoạn giới thiệu (hoặc sau TÀI LIỆU THAM KHẢO)
    insert_pos = list(body).index(existing_refs[0]) + 1 if existing_refs and not re.match(r"^\[\d+\]", _p_text(existing_refs[0]).strip()) else ref_heading_idx + 1
    w14_pid = f"{{{W14_NS}}}paraId"
    for offset, parts in enumerate(apa_list):
        new_ref = copy.deepcopy(tpl_ref)
        if w14_pid in new_ref.attrib:
            new_ref.set(w14_pid, _next_para_id())
        _set_apa_reference_runs(new_ref, parts)
        body.insert(insert_pos + offset, new_ref)

    # ------------------------------------------------------------------ #
    # 8. CHUẨN HÓA DẤU THẬP PHÂN '.' -> ',' TRONG NỘI DUNG CHÍNH (KHÔNG ĐỤNG HEADINGS/TOC/LOF/LOT/REFS)
    # ------------------------------------------------------------------ #
    children_after = list(body)
    ref_heading_idx_after, _ = _find_paragraph_by_prefix(children_after, "TÀI LIỆU THAM KHẢO", start_idx=139)
    skip_styles = {"Heading1", "Heading2", "Heading3", "TOC1", "TOC2", "TOC3", "Caption", "TIeuDeanh", "TableofFigures"}
    for i in range(139, ref_heading_idx_after):
        el = children_after[i]
        if el.tag == _w("p"):
            if _p_style(el) in skip_styles:
                continue
            for t in el.iter(_w("t")):
                if t.text:
                    t.text = _convert_decimal_dots_in_text(t.text)
            for mt in el.iter(f"{{{M_NS}}}t"):
                if mt.text:
                    mt.text = _convert_decimal_dots_in_text(mt.text)
        elif el.tag == _w("tbl"):
            for t in el.iter(_w("t")):
                if t.text:
                    t.text = _convert_decimal_dots_in_text(t.text)
            for mt in el.iter(f"{{{M_NS}}}t"):
                if mt.text:
                    mt.text = _convert_decimal_dots_in_text(mt.text)

    # Serialize và ghi an toàn qua file .tmp
    new_doc_bytes = ET.tostring(doc_root, encoding="utf-8", xml_declaration=True)
    zip_files["word/document.xml"] = new_doc_bytes

    tmp_path = DOCX_PATH + ".tmp"
    with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in zip_files.items():
            zout.writestr(name, data)
    os.replace(tmp_path, DOCX_PATH)
    print(f"--> Đã cập nhật toàn diện (APA 7th + đối chứng đầy đủ): {DOCX_PATH}")


if __name__ == "__main__":
    main()
