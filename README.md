# OCR 7-11 / Open Point → Excel TEST

Ứng dụng Streamlit OCR ảnh 7-11 / Open Point (交貨便) và email chuyển khoản, rồi điền vào file **test** `test_orders.xlsx`. Ứng dụng không đụng tới Excel thật.

## Cấu trúc

```
ocr_711_excel/
├── app.py              # Giao diện Streamlit
├── ocr_utils.py        # OpenCV tiền xử lý + pytesseract + regex trích dữ liệu
├── excel_utils.py      # Tạo / reset / backup / cập nhật test_orders.xlsx
├── requirements.txt
├── test_orders.xlsx    # Tự tạo khi chạy lần đầu (header + 3 dòng mẫu)
├── backups/            # Backup tự động trước mỗi lần ghi / reset
├── output/             # File kết quả test_orders_result_YYYYMMDD_HHMMSS.xlsx
└── logs/app.log        # Log
```

## 1. Cài Tesseract OCR trên Windows

> **Máy này đã cài sẵn:** Tesseract 5.4.0 ở `C:\Program Files\Tesseract-OCR`. Gói ngôn ngữ
> `chi_tra`, `vie`, `eng` (bản `tessdata_best`, chính xác hơn) nằm trong dự án ở `tools\tessdata_best\`.
> App tự dùng thư mục này (chỉ trong tiến trình của app, không đổi biến môi trường hệ thống),
> nên không cần tick ngôn ngữ khi cài. `tools\tessdata\` là bản thường, dùng dự phòng.

1. Tải bản cài đặt Windows (UB Mannheim): https://github.com/UB-Mannheim/tesseract/wiki
   (ví dụ `tesseract-ocr-w64-setup-5.x.x.exe`).
2. Khi cài, ở bước **Choose Components** mở **Additional language data (download)** và tick:
   - **Chinese (Traditional)** → `chi_tra`
   - **Vietnamese** → `vie`
   - (English `eng` có sẵn)
3. Giữ thư mục mặc định `C:\Program Files\Tesseract-OCR`. App tự tìm ở đây.
4. Kiểm tra trong PowerShell:
   ```powershell
   & "C:\Program Files\Tesseract-OCR\tesseract.exe" --list-langs
   ```
   Kết quả phải có `chi_tra`, `vie`, `eng`.

**Nếu quên tick ngôn ngữ:** tải `chi_tra.traineddata` và `vie.traineddata` từ
https://github.com/tesseract-ocr/tessdata (hoặc `tessdata_best` để chính xác hơn) rồi chép vào
`C:\Program Files\Tesseract-OCR\tessdata\`.

**Nếu cài chỗ khác:** nhập đường dẫn `tesseract.exe` ở sidebar, hoặc đặt biến môi trường:
```powershell
setx TESSERACT_CMD "D:\Tools\Tesseract-OCR\tesseract.exe"
```

## 2. Cài Python packages

Cần Python 3.10 trở lên.

```powershell
cd C:\Users\user\dev\ocr_711_excel
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> Nếu PowerShell báo lỗi execution policy: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`

## 3. Chạy

```powershell
streamlit run app.py
```

Trình duyệt mở http://localhost:8501.

## 4. Cách dùng

1. **♻️ Reset Excel test**: backup file hiện tại rồi tạo lại `test_orders.xlsx` với header và 3 dòng mẫu.
2. **Upload ảnh**: chọn nhiều ảnh cùng lúc (png/jpg/...).
3. **🚀 Bắt đầu OCR**: OCR từng ảnh, backup vào `backups/`, ghi vào `test_orders.xlsx`, xuất bản sao ra `output/`.
4. Xem kết quả từng ảnh (dữ liệu trích được + OCR原文), bảng dữ liệu Excel, và **📜 Xem log**.
5. **⬇️ Tải Excel kết quả**.

## 5. Quy tắc điền dữ liệu

| Loại ảnh | Nhận diện khi có | Trường trích |
|---|---|---|
| 7-11 / Open Point | 取件人, 門市, 交貨便, 7-ELEVEN, OPEN POINT | 取件人手機, 取件人姓名, 服務代碼, 門市名稱, 門市地址, 包裹代收金額, 交貨便代碼 |
| Email chuyển khoản | 匯款, 轉帳, 入帳 | 交貨便代碼, 匯款金額, 匯款日期, 匯款狀態 |

- **Ảnh 7-11**: nếu `取件人手機` trùng `*手機` trong Excel thì cập nhật dòng đó (so sánh theo chữ số, nên `0912-345-678` khớp `912345678`). Nếu không trùng thì thêm dòng mới (`*手機` = `取件人手機`, `*客戶名稱` = `取件人姓名`).
- **Email chuyển khoản**: không có số điện thoại, nên được ghép vào dòng có cùng `交貨便代碼`. Ảnh 7-11 luôn được xử lý trước email trong cùng một lần chạy. Nếu không tìm thấy dòng khớp thì thêm dòng mới.
- `服務代碼` được ghi vào cột **門市代碼**.
- `*E-mail` **không** được điền.
- `備註` luôn là `7-11`.
- `客戶編號` được đánh lại 001, 002, 003… theo thứ tự từ trên xuống mỗi lần ghi.
- `總價(自動計算)` = công thức Excel `*單價 × *數量`.
- `對帳狀態`: `已對帳` nếu 包裹代收金額 = 匯款金額, `金額不符` nếu khác, `待對帳` nếu còn thiếu một trong hai.
- `OCR原文` và `來源圖片` được nối thêm khi một dòng nhận dữ liệu từ nhiều ảnh.

## 6. Mẹo tăng độ chính xác OCR

- Dùng ảnh chụp màn hình gốc, không chụp lại màn hình bằng điện thoại.
- Cắt bỏ phần thừa, chỉ giữ vùng có thông tin.
- Dùng `tessdata_best` cho `chi_tra`.
- OCR luôn có thể sai: kiểm tra phần **Kết quả OCR** và log trước khi dùng dữ liệu.
- Nếu mẫu ảnh có nhãn khác (ví dụ `收件人` thay vì `取件人`), sửa regex trong `extract_711()` / `extract_email()` ở `ocr_utils.py`.

## 7. Lỗi thường gặp

| Lỗi | Cách xử lý |
|---|---|
| `TesseractNotFoundError` / sidebar báo không tìm thấy | Cài Tesseract (mục 1) hoặc nhập đường dẫn ở sidebar |
| `Failed loading language 'chi_tra'` | Chép `chi_tra.traineddata` vào thư mục `tessdata` |
| Header không đúng cấu trúc | Bấm **Reset Excel test** |
| `PermissionError` khi lưu | Đóng `test_orders.xlsx` trong Excel rồi thử lại |
