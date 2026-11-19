"""Streamlit app: OCR 7-11 / Open Point screenshots + transfer emails into a TEST Excel file.

Run:  streamlit run app.py
"""
from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

import excel_utils as xu
import ocr_utils as ou

LOG_DIR = xu.BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
logging.basicConfig(
    filename=LOG_DIR / "app.log",
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger("ocr_app")

st.set_page_config(page_title="OCR 7-11 → Excel TEST", page_icon="🧾", layout="wide")

if "logs" not in st.session_state:
    st.session_state.logs = []
if "last_output" not in st.session_state:
    st.session_state.last_output = None
if "last_results" not in st.session_state:
    st.session_state.last_results = []


def log(msg: str):
    line = f"{datetime.now():%H:%M:%S}  {msg}"
    st.session_state.logs.append(line)
    logger.info(msg)


xu.ensure_test_excel()

# --------------------------------------------------------------------------- #
# Sidebar: Tesseract
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("⚙️ Tesseract")
    custom_path = st.text_input("Đường dẫn tesseract.exe (để trống = tự tìm)", value="")
    found = ou.configure_tesseract(custom_path or None)
    info = ou.tesseract_info() if found else {"ok": False, "error": "Không tìm thấy tesseract.exe", "languages": []}
    if info["ok"]:
        st.success(f"Tesseract {info['version']}\n\n{found}")
    else:
        st.error(f"Tesseract chưa sẵn sàng: {info['error']}\n\nXem README.md để cài đặt.")
    lang_wanted = st.text_input("Ngôn ngữ OCR", value=ou.DEFAULT_LANG)
    lang, missing = ou.resolve_lang(lang_wanted, info["languages"])
    if missing:
        st.warning(f"Thiếu gói ngôn ngữ: {', '.join(missing)}. Đang dùng: {lang}")
    st.caption(f"File test: `{xu.TEST_EXCEL.name}`")

st.title("🧾 OCR ảnh 7-11 / Open Point → Excel TEST")
st.caption("Chỉ ghi vào file test `test_orders.xlsx`, luôn backup trước khi ghi. Không dùng Excel thật.")

# --------------------------------------------------------------------------- #
# Actions
# --------------------------------------------------------------------------- #
col1, col2 = st.columns([1, 3])
with col1:
    if st.button("♻️ Reset Excel test", width="stretch"):
        backup = xu.reset_test_excel()
        log(f"Reset Excel test (backup: {backup})")
        st.session_state.last_output = None
        st.success("Đã tạo lại test_orders.xlsx với header + 3 dòng mẫu.")

with col2:
    files = st.file_uploader(
        "Upload ảnh (7-11 / Open Point, email chuyển khoản)",
        type=["png", "jpg", "jpeg", "bmp", "webp", "tif", "tiff"],
        accept_multiple_files=True,
    )

start = st.button("🚀 Bắt đầu OCR", type="primary", disabled=not files or not info["ok"])

if start and files:
    results: list[ou.OcrResult] = []
    progress = st.progress(0.0, text="Đang OCR...")
    for i, f in enumerate(files, start=1):
        progress.progress(i / len(files), text=f"OCR {i}/{len(files)}: {f.name}")
        res = ou.process_image(f.name, f.getvalue(), lang)
        if res.error:
            log(f"[LỖI] {f.name}: {res.error}")
        else:
            log(f"[OCR] {f.name}: loại = {res.doc_type}, dữ liệu = {res.fields}")
        results.append(res)
    progress.empty()
    try:
        summary = xu.write_results(results, log=log)
        st.session_state.last_output = summary["output"]
        st.success(
            f"Xong! Cập nhật {summary['updated']} dòng, thêm {summary['added']} dòng, "
            f"bỏ qua {summary['skipped']} ảnh."
        )
    except Exception as exc:  # noqa: BLE001
        log(f"[LỖI GHI EXCEL] {exc}")
        st.error(f"Lỗi khi ghi Excel: {exc}")
    st.session_state.last_results = results

# --------------------------------------------------------------------------- #
# Results
# --------------------------------------------------------------------------- #
if st.session_state.last_results:
    st.subheader("Kết quả OCR")
    for res in st.session_state.last_results:
        label = {"711": "7-11 / Open Point", "email": "Email chuyển khoản"}.get(res.doc_type, res.doc_type)
        with st.expander(f"{res.filename} — {label}"):
            if res.error:
                st.error(res.error)
            else:
                st.json(res.fields)
                st.text_area("OCR原文", res.raw_text, height=200, key=f"raw_{res.filename}")

st.subheader("Dữ liệu trong test_orders.xlsx")
st.dataframe(pd.DataFrame(xu.read_as_rows(), columns=xu.COLUMNS).drop(columns=["OCR原文"]), width="stretch")

output: Path | None = st.session_state.last_output
if output and Path(output).exists():
    st.download_button(
        "⬇️ Tải Excel kết quả",
        data=Path(output).read_bytes(),
        file_name=Path(output).name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.download_button(
        "⬇️ Tải test_orders.xlsx hiện tại",
        data=xu.TEST_EXCEL.read_bytes(),
        file_name=xu.TEST_EXCEL.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with st.expander("📜 Xem log", expanded=False):
    if st.session_state.logs:
        st.code("\n".join(st.session_state.logs), language="text")
    else:
        st.write("Chưa có log trong phiên này.")
    st.caption(f"Log đầy đủ: {LOG_DIR / 'app.log'}")
