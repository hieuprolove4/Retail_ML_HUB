r"""
polish_khoa1_word.py
Đồng bộ hoàn chỉnh 3 chi tiết cuối trên D:\HK7\Machine_Learning\doan\KHOA HỆ THỐNG THÔNG TIN (1).docx:
  1. Thống nhất TS. Trần Duy Thanh tại Lời cam kết (khớp Trang bìa).
  2. Thống nhất tỷ lệ 20% (Hoàn thành tốt) tại bảng Phân công nhiệm vụ (khớp bảng Thành viên).
  3. Gắn style Caption + bookmark _Toc cho Bảng 4.2 và chèn Bảng 4.2 vào DANH SÁCH BẢNG.
Đồng bộ bản sao ra reports/KHOA HỆ THỐNG THÔNG TIN (1).docx và MLBA (3).docx.
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

TARGET_DOCX = r"D:\HK7\Machine_Learning\doan\KHOA HỆ THỐNG THÔNG TIN (1).docx"
SYNC_COPIES = [
    r"d:\HK7\Machine_Learning\doan\ML_Project\reports\KHOA HỆ THỐNG THÔNG TIN (1).docx",
    r"D:\HK7\Machine_Learning\doan\MLBA (3).docx",
]


def _w(tag):
    return f"{{{W_NS}}}{tag}"


def _wt(el):
    return "".join(n.text or "" for n in el.iter(_w("t"))).strip()


def _set_cell_text(tc, text):
    ps = tc.findall(_w("p"))
    if not ps:
        return
    p = ps[0]
    first_r = p.find(_w("r"))
    rpr = copy.deepcopy(first_r.find(_w("rPr"))) if first_r is not None and first_r.find(_w("rPr")) is not None else None
    for r in list(p.findall(_w("r"))):
        p.remove(r)
    nr = ET.SubElement(p, _w("r"))
    if rpr is not None:
        nr.append(rpr)
    t = ET.SubElement(nr, _w("t"))
    t.text = text


def main():
    with zipfile.ZipFile(TARGET_DOCX, "r") as zin:
        zip_files = {n: zin.read(n) for n in zin.namelist()}

    raw = zip_files["word/document.xml"]
    namespaces = dict([node for _, node in ET.iterparse(io.BytesIO(raw), events=["start-ns"])])
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)

    root = ET.fromstring(raw)
    body = root.find(_w("body"))
    children = list(body)

    # 1. Thống nhất TS. Trần Duy Thanh tại Lời cam kết
    for el in children[:170]:
        if el.tag == _w("p") and "dưới sự hướng dẫn của" in _wt(el):
            for t in el.iter(_w("t")):
                if t.text:
                    t.text = t.text.replace("ThS. Trần Duy Thanh", "TS. Trần Duy Thanh")
                    t.text = t.text.replace("TS.Trần Duy Thanh", "TS. Trần Duy Thanh")

    # 2. Thống nhất cột Đánh giá trong bảng PHÂN CÔNG NHIỆM VỤ -> 20% (Hoàn thành tốt)
    for el in children[:160]:
        if el.tag == _w("tbl"):
            rows = el.findall(_w("tr"))
            if rows and "Vai trò" in " ".join(_wt(tc) for tc in rows[0].findall(_w("tc"))):
                for r in rows[1:]:
                    tcs = r.findall(_w("tc"))
                    if len(tcs) >= 5:
                        _set_cell_text(tcs[4], "20% (Hoàn thành tốt)")

    # 3. Gắn style Caption + bookmark cho Bảng 4.2 và thêm vào DANH SÁCH BẢNG
    has_42_in_lot = any(
        el.tag == _w("p") and _wt(el).startswith("Bảng 4.2:")
        for el in children[:145]
    )
    cap_42 = next(
        (el for el in children[150:] if el.tag == _w("p") and _wt(el).startswith("Bảng 4.2:")),
        None,
    )
    if cap_42 is not None:
        ppr = cap_42.find(_w("pPr"))
        if ppr is None:
            ppr = ET.Element(_w("pPr"))
            cap_42.insert(0, ppr)
        ps = ppr.find(_w("pStyle"))
        if ps is None:
            ps = ET.Element(_w("pStyle"))
            ppr.insert(0, ps)
        ps.set(_w("val"), "Caption")

        # Kiểm tra xem cap_42 đã có bookmark _Toc chưa
        bm = cap_42.find(_w("bookmarkStart"))
        if bm is None:
            max_id = max(
                (int(b.attrib.get(_w("id"), "0")) for b in root.iter(_w("bookmarkStart"))),
                default=100,
            )
            bm_id = str(max_id + 1)
            bm_name = "_Toc245763899"
            bm_s = ET.Element(_w("bookmarkStart"))
            bm_s.set(_w("id"), bm_id)
            bm_s.set(_w("name"), bm_name)
            bm_e = ET.Element(_w("bookmarkEnd"))
            bm_e.set(_w("id"), bm_id)
            # Chèn bookmarkStart ngay sau pPr, bookmarkEnd ở cuối đoạn
            insert_at = 1 if cap_42[0].tag == _w("pPr") else 0
            cap_42.insert(insert_at, bm_s)
            cap_42.append(bm_e)
        else:
            bm_name = bm.attrib.get(_w("name"), "_Toc245763899")

        if not has_42_in_lot:
            idx_41 = next(
                (i for i, el in enumerate(children[:145]) if el.tag == _w("p") and _wt(el).startswith("Bảng 4.1:")),
                None,
            )
            if idx_41 is not None:
                lot_41 = children[idx_41]
                lot_42 = copy.deepcopy(lot_41)
                w14_pid = f"{{{W14_NS}}}paraId"
                if w14_pid in lot_42.attrib:
                    lot_42.set(w14_pid, "7C420001")
                hl = lot_42.find(_w("hyperlink"))
                if hl is not None:
                    hl.set(_w("anchor"), bm_name)
                    ts = list(hl.iter(_w("t")))
                    if ts:
                        ts[0].text = _wt(cap_42)
                        for extra_t in ts[1:-1]:
                            extra_t.text = ""
                        if len(ts) >= 2:
                            ts[-1].text = "46"
                body.insert(idx_41 + 1, lot_42)

    zip_files["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    tmp = TARGET_DOCX + ".tmp"
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for k, v in zip_files.items():
            zout.writestr(k, v)
    os.replace(tmp, TARGET_DOCX)
    print(f"--> Đã đồng bộ hoàn chỉnh: {TARGET_DOCX}")

    for cp in SYNC_COPIES:
        os.makedirs(os.path.dirname(cp), exist_ok=True)
        shutil.copy2(TARGET_DOCX, cp)
        print(f"--> Đã chép bản đồng bộ ra: {cp}")


if __name__ == "__main__":
    main()
