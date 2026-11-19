"""Điền dữ liệu vào ĐÚNG file mẫu của người dùng, giữ nguyên 100% định dạng, xuất ra cùng loại file với mẫu.

Không dựng lại file từ đầu (dựng lại sẽ mất định dạng, công thức, danh sách chọn, macro...):
  * .xlsx / .xlsm / .xltx / .xltm : sửa trực tiếp nội dung các ô trong file (zip XML); mọi phần khác của file
    giữ nguyên từng byte (định dạng, độ rộng cột, ô gộp, danh sách chọn, macro, sheet khác...). Không cần Excel.
  * .csv / .txt : giữ nguyên bảng mã (kể cả BOM), dấu phân cách, kiểu xuống dòng và các dòng tiêu đề.
  * .xls / .ods / file "xls" dạng HTML hoặc XML 2003 : nhờ Microsoft Excel trên máy mở chính file mẫu,
    điền giá trị rồi lưu lại đúng định dạng cũ. Máy không có Excel -> báo rõ và xuất .xlsx thay thế.
Chỉ ghi vào các cột đã nhận ra trong mẫu; dòng ví dụ có sẵn dưới tiêu đề chỉ bị xóa GIÁ TRỊ ở các cột đó (giữ định dạng).
"""
from __future__ import annotations

import csv
import io
import json
import os
import re
import subprocess
import tempfile
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import Any

NUMERIC_KEYS = {"thanh_tien", "tien_thue", "don_gia", "tong_tt", "tien_ck", "so_luong", "ty_gia"}
TEXT_KEYS = {"mst", "ma_doi_tuong", "so_hoa_don", "mau_so", "ky_hieu", "so_chung_tu", "ma_hang", "ma_kho",
             "ma_tra_cuu", "ma_cqt", "tk_no", "tk_co", "tk_kho", "tk_cong_no", "tk_doanh_thu", "tk_thue"}
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


# ------------------------------------------------------------------ chuẩn hóa giá trị
def _val(key: str, v: Any) -> Any:
    """-> None | date | int/float | str"""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return int(v) if isinstance(v, float) and v.is_integer() else v
    s = _ILLEGAL.sub("", str(v))
    if key in NUMERIC_KEYS:
        t = s.strip().replace(" ", "")
        if re.fullmatch(r"-?\d+(\.\d+)?", t):
            f = float(t)
            return int(f) if f.is_integer() else f
    return s


def _num_text(v: int | float) -> str:
    if isinstance(v, int):
        return str(v)
    return format(v, "f").rstrip("0").rstrip(".") if "e" not in repr(v) else repr(v)


def _date_text(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def kind_of(tpl: bytes) -> str:
    """ooxml | ods | xls | xml2003 | html | text"""
    if tpl[:2] == b"PK":
        try:
            names = zipfile.ZipFile(io.BytesIO(tpl)).namelist()
        except zipfile.BadZipFile:
            names = []
        if any(n.startswith("xl/") for n in names):
            return "ooxml"
        if "content.xml" in names:
            return "ods"
    if tpl[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "xls"
    from .misa_import import _decode
    t = _decode(tpl[:200000])
    if "urn:schemas-microsoft-com:office:spreadsheet" in t[:3000].lower():
        return "xml2003"
    if "<table" in t.lower() or "<html" in t[:3000].lower():
        return "html"
    return "text"


def out_ext(tpl: bytes, filename: str) -> str:
    k = kind_of(tpl)
    ext = Path(filename or "").suffix.lower()
    if k == "ooxml":
        ct = zipfile.ZipFile(io.BytesIO(tpl)).read("[Content_Types].xml").decode("utf-8", "ignore")
        macro, templ = "macroEnabled" in ct, "template" in ct
        return ".xltm" if macro and templ else ".xlsm" if macro else ".xltx" if templ else ".xlsx"
    if k == "ods":
        return ".ods"
    if k == "xls":
        return ".xls"
    if k in ("xml2003", "html"):
        return ext if ext in (".xls", ".xml", ".htm", ".html") else ".xls"
    return ext if ext in (".csv", ".txt", ".tsv") else ".csv"


# ------------------------------------------------------------------ .xlsx/.xlsm: sửa thẳng XML
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
M = f"{{{NS_MAIN}}}"


def _col_letter(c: int) -> str:
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _split_ref(ref: str) -> tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return _col_index(m.group(1)), int(m.group(2))


def _resolve(base: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base.rsplit("/", 1)[0] + "/" + target).split("/") if "/" in base else target.split("/")
    out: list[str] = []
    for p in parts:
        if p == "..":
            if out:
                out.pop()
        elif p and p != ".":
            out.append(p)
    return "/".join(out)


def _rels_path(part: str) -> str:
    d, f = part.rsplit("/", 1) if "/" in part else ("", part)
    return (d + "/" if d else "") + "_rels/" + f + ".rels"


def _xml_bytes(root) -> bytes:
    from lxml import etree
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


_DATE_IDS = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))


def _is_date_code(code: str) -> bool:
    c = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", code or "")
    return bool(re.search(r"[dy]", c, re.I) or (re.search(r"m", c, re.I) and not re.search(r"[hs]", c, re.I)))


def fill_ooxml(tpl: bytes, sheet_title: str, header_row: int, cols: dict[int, str],
               rows: list[dict[str, Any]]) -> bytes:
    from lxml import etree
    zin = zipfile.ZipFile(io.BytesIO(tpl))
    parts = {n: zin.read(n) for n in zin.namelist()}
    changed: dict[str, bytes] = {}
    removed: set[str] = set()

    wb_rels = etree.fromstring(parts["xl/_rels/workbook.xml.rels"])
    rel_target = {r.get("Id"): _resolve("xl/workbook.xml", r.get("Target")) for r in wb_rels}
    rel_type = {r.get("Id"): r.get("Type", "") for r in wb_rels}
    wbx = etree.fromstring(parts["xl/workbook.xml"])
    sheet_part = None
    for s in wbx.iter(f"{M}sheet"):
        if s.get("name") == sheet_title:
            sheet_part = rel_target[s.get(f"{{{NS_REL}}}id")]
    if not sheet_part:
        raise ValueError(f"Không thấy sheet '{sheet_title}' trong file mẫu")
    pr = wbx.find(f"{M}workbookPr")
    epoch = date(1904, 1, 1) if pr is not None and pr.get("date1904") in ("1", "true") else date(1899, 12, 30)

    # chuỗi dùng chung
    ss_part = next((rel_target[i] for i, t in rel_type.items() if t.endswith("/sharedStrings")), None)
    sst = etree.fromstring(parts[ss_part]) if ss_part and ss_part in parts else None
    ss_index: dict[str, int] = {}
    if sst is not None:
        for i, si in enumerate(sst.findall(f"{M}si")):
            t = si.find(f"{M}t")
            if t is not None and len(si) == 1:
                ss_index.setdefault(t.text or "", i)

    # định dạng số của từng style -> biết ô là ngày / văn bản
    st_part = next((rel_target[i] for i, t in rel_type.items() if t.endswith("/styles")), None)
    xf_kind: list[str] = []
    if st_part and st_part in parts:
        stx = etree.fromstring(parts[st_part])
        codes = {int(n.get("numFmtId")): n.get("formatCode", "") for n in stx.iter(f"{M}numFmt")}
        cx = stx.find(f"{M}cellXfs")
        for xf in (cx if cx is not None else []):
            fid = int(xf.get("numFmtId") or 0)
            code = codes.get(fid, "")
            xf_kind.append("text" if fid == 49 or code == "@" else
                           "date" if fid in _DATE_IDS or (fid in codes and _is_date_code(code)) else "num")

    def style_kind(s: str | None) -> str:
        try:
            return xf_kind[int(s)] if s is not None else "none"
        except (ValueError, IndexError):
            return "none"

    sh = etree.fromstring(parts[sheet_part])
    sd = sh.find(f"{M}sheetData")
    rows_el = {int(r.get("r")): r for r in sd.findall(f"{M}row") if r.get("r")}
    col_style: dict[int, str] = {}
    cols_el = sh.find(f"{M}cols")
    if cols_el is not None:
        for c in cols_el.findall(f"{M}col"):
            if c.get("style"):
                for k in range(int(c.get("min")), int(c.get("max")) + 1):
                    col_style[k] = c.get("style")

    def cells_of(row_el) -> dict[int, Any]:
        return {_split_ref(c.get("r"))[0]: c for c in row_el.findall(f"{M}c") if c.get("r")}

    # style mẫu cho dòng dữ liệu: lấy từ dòng ngay dưới tiêu đề nếu có (dòng ví dụ của mẫu)
    proto = rows_el.get(header_row + 1)
    proto_cells = cells_of(proto) if proto is not None else {}
    proto_row_attrs = {k: v for k, v in (proto.attrib.items() if proto is not None else [])
                       if k in ("s", "customFormat", "ht", "customHeight")}
    had_formula = False

    def get_row(rn: int):
        el = rows_el.get(rn)
        if el is None:
            el = etree.Element(f"{M}row", r=str(rn))
            for k, v in proto_row_attrs.items():
                el.set(k, v)
            after = [k for k in rows_el if k < rn]
            if after:
                rows_el[max(after)].addnext(el)
            else:
                sd.insert(0, el)
            rows_el[rn] = el
        el.attrib.pop("spans", None)  # tùy chọn; bỏ để không lệch khi thêm ô
        return el

    def get_cell(row_el, rn: int, c: int):
        cur = cells_of(row_el)
        if c in cur:
            return cur[c]
        el = etree.Element(f"{M}c", r=f"{_col_letter(c)}{rn}")
        s = proto_cells[c].get("s") if c in proto_cells else None
        s = s or row_el.get("s") if row_el.get("customFormat") in ("1", "true") else s
        s = s or col_style.get(c)
        if s:
            el.set("s", s)
        nxt = [k for k in cur if k > c]
        if nxt:
            cur[min(nxt)].addprevious(el)
        else:
            row_el.append(el)
        return el

    def clear(el) -> None:
        nonlocal had_formula
        for ch in list(el):
            if ch.tag == f"{M}f":
                had_formula = True
            el.remove(ch)
        el.attrib.pop("t", None)

    def set_text(el, s: str) -> None:
        if sst is not None:
            if s not in ss_index:
                si = etree.SubElement(sst, f"{M}si")
                t = etree.SubElement(si, f"{M}t")
                t.text = s
                if s != s.strip():
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                ss_index[s] = len(sst.findall(f"{M}si")) - 1
            el.set("t", "s")
            etree.SubElement(el, f"{M}v").text = str(ss_index[s])
        else:
            el.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(el, f"{M}is"), f"{M}t")
            t.text = s
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    last = header_row + len(rows)
    # kiểu giá trị dòng ví dụ (trước khi xóa): mẫu ghi ngày dạng chữ thì ghi đúng dạng chữ
    proto_text = {c for c, el in proto_cells.items() if el.get("t") in ("s", "inlineStr", "str")}
    # xóa GIÁ TRỊ mọi ô của các dòng ví dụ dưới tiêu đề (mọi cột) – giữ định dạng, cột, dòng của mẫu
    for rn, rel in rows_el.items():
        if rn > header_row:
            for el in cells_of(rel).values():
                clear(el)
    for i, row in enumerate(rows, header_row + 1):
        rel = get_row(i)
        for c, key in cols.items():
            v = _val(key, row.get(key))
            el = get_cell(rel, i, c)
            clear(el)
            if v is None:
                continue
            kind = style_kind(el.get("s"))
            if isinstance(v, date):
                if kind == "date" and c not in proto_text:
                    etree.SubElement(el, f"{M}v").text = str((v - epoch).days)
                else:
                    set_text(el, _date_text(v))
            elif isinstance(v, (int, float)):
                if kind == "text":
                    set_text(el, _num_text(v))
                else:
                    etree.SubElement(el, f"{M}v").text = _num_text(v)
            else:
                set_text(el, v)
    # dòng ví dụ còn thừa phía dưới: xóa giá trị ở các cột đã điền (giữ định dạng)
    for rn, rel in rows_el.items():
        if rn > last:
            cur = cells_of(rel)
            for c in cols:
                if c in cur:
                    clear(cur[c])

    # vùng dữ liệu (dimension) + bảng Excel (ListObject) nới xuống hết dữ liệu
    max_col = max([max(cols)] + [c for r in rows_el.values() for c in cells_of(r)])
    max_row = max(rows_el) if rows_el else last
    dim = sh.find(f"{M}dimension")
    if dim is not None:
        first = dim.get("ref", "A1").split(":")[0]
        dim.set("ref", f"{first}:{_col_letter(max_col)}{max_row}")
    rels_p = _rels_path(sheet_part)
    if rels_p in parts:
        for r in etree.fromstring(parts[rels_p]):
            if r.get("Type", "").endswith("/table"):
                tp = _resolve(sheet_part, r.get("Target"))
                if tp not in parts:
                    continue
                tb = etree.fromstring(parts[tp])
                a, b = tb.get("ref").split(":")
                (_, r1), (c2, r2) = _split_ref(a), _split_ref(b)
                if r1 == header_row and last > r2:
                    tb.set("ref", f"{a}:{_col_letter(c2)}{last}")
                    af = tb.find(f"{M}autoFilter")
                    if af is not None:
                        af.set("ref", f"{a}:{_col_letter(c2)}{last}")
                    changed[tp] = _xml_bytes(tb)
    changed[sheet_part] = _xml_bytes(sh)
    if sst is not None:
        n = len(sst.findall(f"{M}si"))
        sst.set("uniqueCount", str(n))
        if sst.get("count") is not None:
            sst.set("count", str(max(int(sst.get("count") or 0), n)))
        changed[ss_part] = _xml_bytes(sst)

    if had_formula:  # ô có công thức bị ghi đè -> bỏ chuỗi tính toán (Excel tự dựng lại khi mở)
        cc = next((rel_target[i] for i, t in rel_type.items() if t.endswith("/calcChain")), None)
        if cc and cc in parts:
            removed.add(cc)
            for r in list(wb_rels):
                if r.get("Type", "").endswith("/calcChain"):
                    wb_rels.remove(r)
            changed["xl/_rels/workbook.xml.rels"] = _xml_bytes(wb_rels)
            ct = etree.fromstring(parts["[Content_Types].xml"])
            for o in list(ct):
                if o.get("PartName") == "/" + cc:
                    ct.remove(o)
            changed["[Content_Types].xml"] = _xml_bytes(ct)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zout:
        for info in zin.infolist():
            if info.filename in removed:
                continue
            data = changed.get(info.filename, parts[info.filename])
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type, zi.external_attr = info.compress_type, info.external_attr
            zout.writestr(zi, data)
    return out.getvalue()


# ------------------------------------------------------------------ .csv / .txt
def fill_text(tpl: bytes, header_row: int, cols: dict[int, str], rows: list[dict[str, Any]]) -> bytes:
    bom = b""
    for b in (b"\xef\xbb\xbf", b"\xff\xfe", b"\xfe\xff"):
        if tpl.startswith(b):
            bom = b
    if bom == b"\xef\xbb\xbf":
        enc = "utf-8"
    elif bom:
        enc = "utf-16-le" if bom == b"\xff\xfe" else "utf-16-be"
    else:
        enc = next((e for e in ("utf-8", "cp1258", "cp1252") if _try_decode(tpl, e)), "latin-1")
    text = tpl[len(bom):].decode(enc)
    nl = "\r\n" if "\r\n" in text else "\n"
    try:
        dialect = csv.Sniffer().sniff(text[:20000], delimiters=",;\t|")
        delim, quote = dialect.delimiter, dialect.quotechar
    except csv.Error:
        delim, quote = ("\t" if text[:20000].count("\t") > text[:20000].count(",") else ","), '"'
    recs = list(csv.reader(io.StringIO(text), delimiter=delim, quotechar=quote))
    head_lines = _first_records_text(text, delim, quote, header_row)  # giữ nguyên chữ của các dòng tiêu đề
    width = max([len(r) for r in recs[:header_row]] + [max(cols)])
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=delim, quotechar=quote, lineterminator=nl)
    for row in rows:
        line = [""] * width
        for c, key in cols.items():
            v = _val(key, row.get(key))
            line[c - 1] = "" if v is None else _date_text(v) if isinstance(v, date) else \
                _num_text(v) if isinstance(v, (int, float)) else v
        w.writerow(line)
    body = head_lines + ("" if head_lines.endswith(("\n", "\r")) or not head_lines else nl) + buf.getvalue()
    return bom + body.encode(enc)


def _try_decode(b: bytes, enc: str) -> bool:
    try:
        b.decode(enc)
        return True
    except UnicodeDecodeError:
        return False


def _first_records_text(text: str, delim: str, quote: str, n: int) -> str:
    """Đoạn văn bản gốc của n bản ghi đầu (tính cả xuống dòng trong ô có ngoặc kép)."""
    i, count, inq = 0, 0, False
    while i < len(text) and count < n:
        ch = text[i]
        if ch == quote:
            inq = not inq
        elif ch == "\n" and not inq:
            count += 1
        i += 1
    return text[:i]


# ------------------------------------------------------------------ .xls/.ods/html/xml2003: nhờ Excel
_PS_EXCEL = r"""
param([string]$jobFile)
$ErrorActionPreference = 'Stop'
$job = Get-Content -Raw -Encoding UTF8 $jobFile | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AskToUpdateLinks = $false
try {
  $wb = $xl.Workbooks.Open($job.path, 0, $false)
  if ($job.mode -eq 'probe') {
    $wb.SaveAs($job.probe, 51)   # bản sao .xlsx chỉ để đọc tiêu đề đúng như Excel thấy
  } else {
    $ws = $wb.Worksheets.Item([int]$job.sheet)
    $start = [int]$job.start
    $endUsed = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
    $lastCol = $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1
    # kiểu giá trị của dòng ví dụ từng cột (TRƯỚC khi xóa): mẫu ghi ngày dạng chữ "31/01/2020" thì ghi đúng dạng chữ
    $exText = @{}
    foreach ($col in $job.cols) { $c = [int]$col.c; $exText[$c] = ($ws.Cells.Item($start, $c).Value2 -is [string]) }
    # xóa GIÁ TRỊ mọi dòng ví dụ (mọi cột) – giữ nguyên định dạng, danh sách chọn, cột, dòng của mẫu
    if ($job.clear -ne $false -and $endUsed -ge $start) { $ws.Range($ws.Cells.Item($start, 1), $ws.Cells.Item($endUsed, $lastCol)).ClearContents() | Out-Null }
    foreach ($col in $job.cols) {
      $c = [int]$col.c; $n = $col.vals.Count
      $first = $ws.Cells.Item($start, $c)
      $fmt = [string]$first.NumberFormat
      $isDate = ($fmt -match '[dy]') -or ($fmt -match 'm' -and $fmt -notmatch '[hs]')
      $isText = ($fmt -eq '@')
      if ($n -gt 0) {
        $arr = New-Object 'object[,]' $n, 1
        for ($i = 0; $i -lt $n; $i++) {
          $v = $col.vals[$i]
          if ($null -eq $v) { $arr[$i, 0] = $null; continue }
          if ($v.t -eq 'd') {
            if ($isDate -and -not $exText[$c]) { $arr[$i, 0] = [double]$v.oa }
            elseif ($isText) { $arr[$i, 0] = [string]$v.s }
            else { $arr[$i, 0] = "'" + [string]$v.s }   # giữ là chữ dd/mm/yyyy, Excel không tự đổi
          }
          elseif ($v.t -eq 'n') { if ($isText) { $arr[$i, 0] = [string]$v.s } else { $arr[$i, 0] = [double]$v.v } }
          elseif ($v.t -eq 'x' -and -not $isText -and ([string]$v.s -match '^[\d\s.,/:+-]+$')) { $arr[$i, 0] = "'" + [string]$v.s }  # MST, số HĐ: giữ số 0 đầu
          else { $arr[$i, 0] = [string]$v.s }
        }
        $ws.Range($ws.Cells.Item($start, $c), $ws.Cells.Item($start + $n - 1, $c)).Value2 = $arr
      }
    }
    $wb.Save()
  }
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
"""


def excel_available() -> bool:
    if os.name != "nt":
        return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"))
        return True
    except OSError:
        return False


def _run_excel(job: dict[str, Any], tmp: Path) -> None:
    ps = tmp / "fill.ps1"
    ps.write_text(_PS_EXCEL, encoding="utf-8-sig")
    jf = tmp / "job.json"
    jf.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
    flags = 0x08000000 if os.name == "nt" else 0  # không hiện cửa sổ đen
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(ps),
                        str(jf)], capture_output=True, text=True, timeout=900, creationflags=flags)
    if r.returncode != 0:
        raise RuntimeError("Excel không xử lý được file mẫu: " + (r.stderr or r.stdout).strip()[-400:])


def probe_with_excel(tpl: bytes, ext: str) -> bytes:
    """Excel mở file mẫu và lưu bản sao .xlsx -> đọc tiêu đề đúng như Excel thấy."""
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        src = tmp / ("mau" + ext)
        src.write_bytes(tpl)
        _run_excel({"mode": "probe", "path": str(src), "probe": str(tmp / "probe.xlsx")}, tmp)
        return (tmp / "probe.xlsx").read_bytes()


def fill_with_excel(tpl: bytes, ext: str, sheet_index: int, header_row: int, cols: dict[int, str],
                    rows: list[dict[str, Any]], clear: bool = True) -> bytes:
    """clear=False: chỉ ghi các cột được chọn từ dòng header_row+1, giữ nguyên mọi dữ liệu khác của file."""
    from datetime import date as _d
    jcols = []
    for c, key in cols.items():
        vals = []
        for row in rows:
            v = _val(key, row.get(key))
            if v is None:
                vals.append(None)
            elif isinstance(v, date):
                vals.append({"t": "d", "oa": (v - _d(1899, 12, 30)).days, "s": _date_text(v)})
            elif key in TEXT_KEYS:  # mã / số dạng chữ: không để Excel đổi thành số (mất số 0 đầu)
                vals.append({"t": "x", "s": _num_text(v) if isinstance(v, (int, float)) else str(v)})
            elif isinstance(v, (int, float)):
                vals.append({"t": "n", "v": v, "s": _num_text(v)})
            else:
                vals.append({"t": "s", "s": v})
        jcols.append({"c": c, "vals": vals})
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        dst = tmp / ("ket_qua" + ext)
        dst.write_bytes(tpl)
        _run_excel({"mode": "fill", "path": str(dst), "sheet": sheet_index, "start": header_row + 1, "cols": jcols,
                    "clear": clear}, tmp)
        return dst.read_bytes()
