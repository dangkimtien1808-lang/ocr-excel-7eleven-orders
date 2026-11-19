"""Excel utilities for the TEST workbook (test_orders.xlsx). Never touches a real Excel file."""
from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from ocr_utils import OcrResult, merge_split_711, normalize_phone

BASE_DIR = Path(__file__).resolve().parent
TEST_EXCEL = BASE_DIR / "test_orders.xlsx"
BACKUP_DIR = BASE_DIR / "backups"
OUTPUT_DIR = BASE_DIR / "output"

COLUMNS = [
    "訂單編號", "客戶編號", "*客戶名稱(購買人全名)", "客戶代號", "*手機", "購買日", "*商品",
    "*單價", "*數量", "總價(自動計算)", "*E-mail", "備註", "交貨便代碼", "取件人姓名", "取件人手機",
    "門市代碼", "門市名稱", "門市地址", "包裹代收金額", "匯款金額", "匯款日期", "匯款狀態",
    "對帳狀態", "OCR原文", "來源圖片",
]

SAMPLE_ROWS = [
    {"訂單編號": "T0001", "*客戶名稱(購買人全名)": "王小明", "客戶代號": "C-A", "*手機": "0912345678",
     "購買日": "2026/11/01", "*商品": "測試商品A", "*單價": 350, "*數量": 2},
    {"訂單編號": "T0002", "*客戶名稱(購買人全名)": "陳美麗", "客戶代號": "C-B", "*手機": "0923456789",
     "購買日": "2026/11/02", "*商品": "測試商品B", "*單價": 500, "*數量": 1},
    {"訂單編號": "T0003", "*客戶名稱(購買人全名)": "Nguyễn Văn An", "客戶代號": "C-C", "*手機": "0934567890",
     "購買日": "2026/11/03", "*商品": "測試商品C", "*單價": 120, "*數量": 3},
]

TEXT_COLUMNS = {"客戶編號", "*手機", "取件人手機", "交貨便代碼", "門市代碼", "訂單編號"}
MAX_CELL = 32000  # Excel limit is 32767 chars


# --------------------------------------------------------------------------- #
# Create / reset / backup
# --------------------------------------------------------------------------- #
def _col(name: str) -> int:
    return COLUMNS.index(name) + 1


def _total_formula(row: int) -> str:
    p, q = get_column_letter(_col("*單價")), get_column_letter(_col("*數量"))
    return f'=IF(AND(ISNUMBER({p}{row}),ISNUMBER({q}{row})),{p}{row}*{q}{row},"")'


def create_test_excel(path: Path = TEST_EXCEL) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "訂單"
    ws.append(COLUMNS)
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    for i, name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=i)
        cell.font = Font(bold=True, color="C00000" if name.startswith("*") else "000000")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = 40 if name == "OCR原文" else max(12, len(name) * 2)
    ws.freeze_panes = "A2"
    for r, sample in enumerate(SAMPLE_ROWS, start=2):
        for name, value in sample.items():
            _set(ws, r, name, value)
        _finalize_row(ws, r, r - 1)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def ensure_test_excel(path: Path = TEST_EXCEL) -> Path:
    if not path.exists():
        create_test_excel(path)
    return path


def backup_excel(path: Path = TEST_EXCEL, backup_dir: Path = BACKUP_DIR) -> Path | None:
    if not path.exists():
        return None
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    dest = backup_dir / f"{path.stem}_backup_{stamp}{path.suffix}"
    shutil.copy2(path, dest)
    return dest


def reset_test_excel(path: Path = TEST_EXCEL) -> Path | None:
    """Back up the current test file, then recreate it with header + 3 sample rows."""
    backup = backup_excel(path)
    create_test_excel(path)
    return backup


# --------------------------------------------------------------------------- #
# Row helpers
# --------------------------------------------------------------------------- #
def _clean(value):
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)[:MAX_CELL]
    return value


def _get(ws, row: int, name: str):
    return ws.cell(row=row, column=_col(name)).value


def _set(ws, row: int, name: str, value, only_if_empty: bool = False):
    if value is None or value == "":
        return
    cell = ws.cell(row=row, column=_col(name))
    if only_if_empty and cell.value not in (None, ""):
        return
    if name in TEXT_COLUMNS:
        value = str(value)
        cell.number_format = "@"
    cell.value = _clean(value)
    if name == "OCR原文":
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def _append_text(ws, row: int, name: str, value: str, sep: str):
    if not value:
        return
    old = _get(ws, row, name)
    if old and value in str(old):
        return
    _set(ws, row, name, f"{old}{sep}{value}" if old else value)


def _data_rows(ws) -> list[int]:
    rows = []
    for r in range(2, ws.max_row + 1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, len(COLUMNS) + 1)):
            rows.append(r)
    return rows


def _next_row(ws) -> int:
    rows = _data_rows(ws)
    return (rows[-1] + 1) if rows else 2


def _to_int(value):
    try:
        return int(str(value).replace(",", ""))
    except ValueError:
        return None


def _reconcile(ws, row: int):
    cod, paid = _get(ws, row, "包裹代收金額"), _get(ws, row, "匯款金額")
    status = str(_get(ws, row, "匯款狀態") or "")
    if cod in (None, "") or paid in (None, ""):
        result = "待對帳"
    elif _to_int(cod) == _to_int(paid) and "失敗" not in status:
        result = "已對帳"
    else:
        result = "金額不符"
    _set(ws, row, "對帳狀態", result)


def _finalize_row(ws, row: int, number: int):
    _set(ws, row, "客戶編號", f"{number:03d}")
    _set(ws, row, "備註", "7-11")
    ws.cell(row=row, column=_col("總價(自動計算)")).value = _total_formula(row)


def _renumber(ws):
    """客戶編號 = 001, 002, ... from top to bottom; 備註 always 7-11."""
    for i, r in enumerate(_data_rows(ws), start=1):
        _finalize_row(ws, r, i)


def _find_row(ws, name: str, value: str, normalizer=lambda v: str(v or "").strip()) -> int | None:
    target = normalizer(value)
    if not target:
        return None
    for r in _data_rows(ws):
        if normalizer(_get(ws, r, name)) == target:
            return r
    return None


# --------------------------------------------------------------------------- #
# Main write
# --------------------------------------------------------------------------- #
def write_results(results: list[OcrResult], path: Path = TEST_EXCEL, log=print) -> dict:
    """Back up, apply OCR results to the test workbook, save, and export a copy to output/."""
    ensure_test_excel(path)
    backup = backup_excel(path)
    log(f"Đã backup: {backup}")

    wb = load_workbook(path)
    ws = wb.active
    header = [ws.cell(row=1, column=i).value for i in range(1, len(COLUMNS) + 1)]
    if header != COLUMNS:
        raise ValueError("Header của test_orders.xlsx không đúng cấu trúc. Hãy bấm Reset Excel test.")

    stats = {"updated": 0, "added": 0, "skipped": 0}
    results = merge_split_711(results, log=log)
    # 7-11 first so that email results can be matched by 交貨便代碼 afterwards.
    ordered = [r for r in results if r.doc_type == "711"] + [r for r in results if r.doc_type != "711"]

    for res in ordered:
        f = res.fields
        if res.error or res.doc_type not in ("711", "email"):
            stats["skipped"] += 1
            log(f"[BỎ QUA] {res.filename}: {res.error or 'không nhận diện được loại ảnh'}")
            continue

        if res.doc_type == "711":
            phone = f.get("取件人手機", "")
            row = _find_row(ws, "*手機", phone, normalize_phone)
            if row is None and f.get("交貨便代碼"):
                row = _find_row(ws, "交貨便代碼", f["交貨便代碼"])
            if row:
                stats["updated"] += 1
                log(f"[CẬP NHẬT] {res.filename}: SĐT {phone} khớp dòng {row}")
            else:
                row = _next_row(ws)
                stats["added"] += 1
                log(f"[THÊM MỚI] {res.filename}: SĐT {phone or '(không đọc được)'} -> dòng {row}")
                _set(ws, row, "*手機", phone)
            _set(ws, row, "*客戶名稱(購買人全名)", f.get("取件人姓名"), only_if_empty=True)
            _set(ws, row, "取件人姓名", f.get("取件人姓名"))
            _set(ws, row, "取件人手機", phone)
            _set(ws, row, "門市代碼", f.get("服務代碼"))
            _set(ws, row, "門市名稱", f.get("門市名稱"))
            _set(ws, row, "門市地址", f.get("門市地址"))
            _set(ws, row, "包裹代收金額", f.get("包裹代收金額"))
            _set(ws, row, "交貨便代碼", f.get("交貨便代碼"))

        else:  # email: may contain several 交貨便代碼 rows
            for item in res.items or [f]:
                code = item.get("交貨便代碼", "")
                row = _find_row(ws, "交貨便代碼", code)
                if row:
                    stats["updated"] += 1
                    log(f"[CẬP NHẬT] {res.filename}: 交貨便代碼 {code} khớp dòng {row}")
                else:
                    row = _next_row(ws)
                    stats["added"] += 1
                    log(f"[THÊM MỚI] {res.filename}: email chuyển khoản, 交貨便代碼 {code or '(không đọc được)'} -> dòng {row}")
                    _set(ws, row, "交貨便代碼", code)
                _set(ws, row, "匯款金額", item.get("匯款金額"))
                _set(ws, row, "匯款日期", item.get("匯款日期"))
                _set(ws, row, "匯款狀態", item.get("匯款狀態"))
                _append_text(ws, row, "OCR原文", f"[{res.filename}]\n{res.raw_text.strip()}", "\n\n")
                _append_text(ws, row, "來源圖片", res.filename, "; ")
                _reconcile(ws, row)
            continue

        _append_text(ws, row, "OCR原文", f"[{res.filename}]\n{res.raw_text.strip()}", "\n\n")
        _append_text(ws, row, "來源圖片", res.filename, "; ")
        _reconcile(ws, row)

    _renumber(ws)
    wb.save(path)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output = OUTPUT_DIR / f"test_orders_result_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, output)
    log(f"Đã lưu {path.name} và xuất file kết quả: {output}")
    return {"backup": backup, "output": output, **stats}


def read_as_rows(path: Path = TEST_EXCEL) -> list[dict]:
    """Read the workbook for preview in Streamlit."""
    ensure_test_excel(path)
    ws = load_workbook(path).active
    rows = []
    for r in _data_rows(ws):
        row = {name: _get(ws, r, name) for name in COLUMNS}
        price, qty = _to_int(row["*單價"]), _to_int(row["*數量"])
        row["總價(自動計算)"] = price * qty if price is not None and qty is not None else None
        rows.append(row)
    return rows
