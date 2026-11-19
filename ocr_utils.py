"""OCR utilities: OpenCV preprocessing + pytesseract + regex field extraction.

Supports 7-11 / Open Point 交貨便 screenshots and bank transfer (匯款) emails.
"""
from __future__ import annotations

import os
import re
import shutil
import unicodedata
from dataclasses import dataclass, field

import cv2
import numpy as np
import pytesseract

DEFAULT_LANG = "chi_tra+vie+eng"

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

WINDOWS_TESSERACT_PATHS = [
    # Project-local copy first (does not touch the system / other projects)
    os.path.join(PROJECT_DIR, "tools", "Tesseract-OCR", "tesseract.exe"),
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
]


# --------------------------------------------------------------------------- #
# Tesseract setup
# --------------------------------------------------------------------------- #
def configure_tesseract(custom_path: str | None = None) -> str | None:
    """Locate tesseract.exe and register it with pytesseract. Returns the path or None."""
    # Use project-local language data (chi_tra, vie, eng) if present.
    # Only affects this Python process, not the system or other projects.
    # tessdata_best is slower but much more accurate for chi_tra.
    for sub in ("tessdata_best", "tessdata"):
        local_tessdata = os.path.join(PROJECT_DIR, "tools", sub)
        if os.path.isfile(os.path.join(local_tessdata, "eng.traineddata")):
            os.environ["TESSDATA_PREFIX"] = local_tessdata
            break
    candidates = [custom_path, os.environ.get("TESSERACT_CMD"), shutil.which("tesseract")]
    candidates += WINDOWS_TESSERACT_PATHS
    for path in candidates:
        if path and os.path.isfile(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return path
    return None


def tesseract_info() -> dict:
    """Return version and installed languages (or an error message)."""
    try:
        version = str(pytesseract.get_tesseract_version())
        langs = sorted(pytesseract.get_languages(config=""))
        return {"ok": True, "version": version, "languages": langs}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": str(exc), "languages": []}


def resolve_lang(wanted: str, available: list[str]) -> tuple[str, list[str]]:
    """Keep only the installed languages from `wanted`. Returns (lang_string, missing)."""
    parts = [p for p in wanted.split("+") if p]
    if not available:
        return wanted, []
    present = [p for p in parts if p in available]
    missing = [p for p in parts if p not in available]
    return ("+".join(present) or "eng"), missing


# --------------------------------------------------------------------------- #
# Image preprocessing + OCR
# --------------------------------------------------------------------------- #
def load_image(data: bytes) -> np.ndarray:
    """Decode image bytes (works with non-ASCII Windows paths, unlike cv2.imread)."""
    arr = np.frombuffer(data, dtype=np.uint8)
    img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError("Không đọc được ảnh (định dạng không hỗ trợ hoặc file hỏng).")
    return img


def preprocess(img: np.ndarray) -> np.ndarray:
    """Grayscale, upscale small images, denoise and binarize for better OCR."""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape[:2]
    if max(h, w) < 2000:
        scale = 2000 / max(h, w)
        gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    gray = cv2.fastNlMeansDenoising(gray, h=10)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    # Dark-mode screenshots: make text black on white.
    if np.mean(binary) < 127:
        binary = cv2.bitwise_not(binary)
    return binary


def ocr_image(data: bytes, lang: str = DEFAULT_LANG) -> str:
    img = load_image(data)
    processed = preprocess(img)
    config = "--oem 3 --psm 6 -c preserve_interword_spaces=1"
    text = pytesseract.image_to_string(processed, lang=lang, config=config)
    # Fallback: if binarization destroyed the text, try plain grayscale.
    if len(text.strip()) < 10:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        text = pytesseract.image_to_string(gray, lang=lang, config=config)
    return text


# --------------------------------------------------------------------------- #
# Text normalization + field extraction
# --------------------------------------------------------------------------- #
CJK = r"\u3400-\u9fff\uf900-\ufaff"


def normalize_text(text: str) -> str:
    """NFKC (full-width -> half-width), remove spaces Tesseract puts between CJK chars."""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(rf"(?<=[{CJK}])[ \t]+(?=[{CJK}])", "", text)
    text = re.sub(rf"(?<=[{CJK}])[ \t]+(?=:)", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text


def _find(patterns: list[str], text: str) -> str:
    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE | re.MULTILINE)
        if m:
            return m.group(1).strip(" :|-")
    return ""


def normalize_phone(value: str) -> str:
    """Keep digits only; restore the leading 0 of Taiwan mobiles (Excel often drops it)."""
    digits = re.sub(r"\D", "", str(value or ""))
    if len(digits) == 9 and digits.startswith("9"):
        digits = "0" + digits
    if digits.startswith("886") and len(digits) == 12:
        digits = "0" + digits[3:]
    return digits


def parse_amount(value: str) -> int | None:
    digits = re.sub(r"[^\d]", "", str(value or ""))
    return int(digits) if digits else None


SEP = r"\s*[:]?\s*"
CODE_PAT = r"交貨便(?:代碼|代号|編號|號碼)" + SEP + r"([A-Za-z0-9|][A-Za-z0-9|\- ]{5,20}[A-Za-z0-9|])"
# "金額" is often misread as 金題 / 金頓 / 金額 variants -> accept any char after 金.
AMOUNT_WORD = r"金[額额題頓顏]?"


def clean_code(value: str) -> str:
    """Normalize a 交貨便代碼: drop spaces/dashes, fix OCR l/I/| -> 1 and O -> 0 after the first char."""
    value = re.sub(r"[\s\-]", "", value or "")
    if not value:
        return ""
    head, tail = value[0], value[1:]
    tail = tail.translate(str.maketrans({"l": "1", "I": "1", "|": "1", "O": "0", "o": "0"}))
    return (head + tail).upper()


def valid_phone(value: str) -> str:
    """Return a Taiwan mobile number (09xxxxxxxx) or "" if OCR produced a wrong length."""
    digits = normalize_phone(value)
    return digits if re.fullmatch(r"09\d{8}", digits) else ""


# One row of the myship payout table: "H1194886   6300 2026/11/13 已匯款"
EMAIL_ROW = re.compile(
    r"^\s*([A-Za-z][A-Za-z0-9|]{6,14})\s+(?:NT\$)?([\d,]{2,})\s+(\d{4}/\d{1,2}/\d{1,2})\s*([^\s\d]\S*)?",
    re.MULTILINE,
)


def detect_type(text: str) -> str:
    # myship payout email (table of 交貨便代碼 / 匯款金額 / 匯款日期 / 狀態)
    if EMAIL_ROW.search(text) or re.search(r"myship|撥款", text, re.IGNORECASE):
        return "email"
    # 7-11 order screens. Checked before 匯款 because the order detail page contains "代收/匯款資訊".
    if re.search(r"取件人|門市|服務單|7-?ELEVEN|open ?point|統一超商", text, re.IGNORECASE):
        return "711"
    if re.search(r"匯款|轉帳|转账|入帳|chuyển khoản|transfer", text, re.IGNORECASE):
        return "email"
    if re.search(r"交[貨發]便", text):
        return "711"
    return "unknown"


def extract_711(text: str) -> dict:
    phone = valid_phone(
        _find([r"[取及]件人(?:的)?(?:手機|電話)(?:號碼)?" + SEP + r"(\+?[\d][\d\-\s]{7,14}\d)"], text)
    )
    if not phone:
        phone = valid_phone(_find([r"(09\d{2}[\-\s]?\d{3}[\-\s]?\d{3})"], text))
    data = {
        "取件人手機": phone,
        "取件人姓名": _find(
            [r"[取及]件人(?:的)?姓名" + SEP + r"([^\s:\d][^\s:]{0,15})", r"取件人\s*:\s*([^\s:\d][^\s:]{0,15})"],
            text,
        ),
        "服務代碼": _find([r"服務代碼" + SEP + r"([A-Z0-9]{3,20})", r"(?:門市)?店號" + SEP + r"(\d{6})"], text),
        "門市名稱": _find([r"門市名[稱精]" + SEP + r"([^\n]+?)\s*$", r"([^\s:]{2,10}門市)"], text),
        "門市地址": _find([r"門市地址" + SEP + r"([^\n]+?)\s*$", r"地址" + SEP + r"([^\n]+?)\s*$"], text),
        "包裹代收金額": parse_amount(
            _find([r"代收" + AMOUNT_WORD + r"(?:\s*[(][^)\n]*[)])?" + SEP + r"(?:NT\$|NTD|\$)?\s*([\d,]+)"], text)
        ),
        "交貨便代碼": clean_code(_find([CODE_PAT], text)),
    }
    return data


def extract_email(text: str) -> dict:
    date = _find(
        [
            r"(?:匯款|轉帳|交易|入帳)(?:日期|時間)" + SEP + r"(\d{2,4}[/\-.年]\d{1,2}[/\-.月]\d{1,2}日?)",
            r"(\d{4}[/\-.年]\d{1,2}[/\-.月]\d{1,2}日?)",
        ],
        text,
    )
    status = _find([r"(?:匯款|轉帳|交易)狀態" + SEP + r"([^\s]+)"], text)
    if not status:
        m = re.search(r"(交易成功|匯款成功|轉帳成功|已完成|成功|失敗|處理中)", text)
        status = m.group(1) if m else ""
    data = {
        "交貨便代碼": clean_code(_find([CODE_PAT], text)),
        "匯款金額": parse_amount(
            _find([r"(?:匯款|轉帳|入帳|交易)" + AMOUNT_WORD + SEP + r"(?:NT\$|NTD|TWD|\$)?\s*([\d,]+)"], text)
        ),
        "匯款日期": date.replace("年", "/").replace("月", "/").replace("日", "").replace("-", "/").replace(".", "/"),
        "匯款狀態": status,
    }
    return data


def extract_email_rows(text: str) -> list[dict]:
    """Rows of a payout table (one email can list several 交貨便代碼)."""
    rows = []
    for m in EMAIL_ROW.finditer(text):
        rows.append(
            {
                "交貨便代碼": clean_code(m.group(1)),
                "匯款金額": parse_amount(m.group(2)),
                "匯款日期": m.group(3),
                "匯款狀態": (m.group(4) or "").strip(),
            }
        )
    return rows


@dataclass
class OcrResult:
    filename: str
    doc_type: str
    fields: dict = field(default_factory=dict)
    raw_text: str = ""
    error: str = ""
    # For emails with a table: one dict per 交貨便代碼 (fields then holds a summary).
    items: list = field(default_factory=list)


def merge_split_711(results: list[OcrResult], log=print) -> list[OcrResult]:
    """A 7-11 order is often 2 screenshots: 服務單 (only 交貨便代碼) + detail page (phone, name, store...).

    Merge a code-only screenshot with the detail screenshot right next to it (upload order).
    """
    def code_only(r):
        return r.doc_type == "711" and r.fields.get("交貨便代碼") and not r.fields.get("取件人手機")

    def detail_only(r):
        return r.doc_type == "711" and r.fields.get("取件人手機") and not r.fields.get("交貨便代碼")

    out = list(results)
    i = 0
    while i < len(out):
        r = out[i]
        if code_only(r):
            for j in (i + 1, i - 1):
                if 0 <= j < len(out) and detail_only(out[j]):
                    d = out[j]
                    merged = OcrResult(
                        f"{d.filename} + {r.filename}",
                        "711",
                        {**d.fields, "交貨便代碼": r.fields["交貨便代碼"]},
                        raw_text=f"{d.raw_text.strip()}\n\n[{r.filename}]\n{r.raw_text.strip()}",
                    )
                    log(f"[GHÉP] {r.filename} (交貨便代碼 {r.fields['交貨便代碼']}) + {d.filename} (SĐT {d.fields['取件人手機']})")
                    lo, hi = min(i, j), max(i, j)
                    out[lo:hi + 1] = [merged]
                    i = lo
                    break
        i += 1
    return out


def process_image(filename: str, data: bytes, lang: str = DEFAULT_LANG) -> OcrResult:
    try:
        raw = ocr_image(data, lang)
    except Exception as exc:  # noqa: BLE001
        return OcrResult(filename, "error", error=str(exc))
    return parse_text(filename, raw)


def parse_text(filename: str, raw: str) -> OcrResult:
    text = normalize_text(raw)
    doc_type = detect_type(text)
    if doc_type == "email":
        items = extract_email_rows(text)
        if items:
            fields = items[0] if len(items) == 1 else {"交易明細": items}
            return OcrResult(filename, doc_type, fields, raw_text=raw, items=items)
        fields = extract_email(text)
    elif doc_type == "711":
        fields = extract_711(text)
    else:
        # Unknown layout: try both and keep whatever was found.
        fields = {**extract_email(text), **{k: v for k, v in extract_711(text).items() if v}}
        doc_type = "711" if fields.get("取件人手機") else ("email" if fields.get("匯款金額") else "unknown")
    return OcrResult(filename, doc_type, fields, raw_text=raw)
