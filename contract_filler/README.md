# Package `contract_filler`

`contract_filler` là module Python trung tâm xử lý toàn bộ tác vụ đọc hóa đơn PDF, phân tích danh sách bên bán/bên mua từ file Word, điền thông tin và bảng dữ liệu vào file hợp đồng mẫu `.docx`, cũng như cung cấp giao diện người dùng Tkinter.

---

## Danh Sách Các Submodules

### 1. `pdf_parser.py`
- **Chức năng**: Sử dụng `pdfplumber` để phân tích file hóa đơn PDF (dạng hóa đơn VAT GTGT).
- **Hàm chính**:
  - `parse_line_items(pdf_path)`: Trích xuất danh sách các dòng hàng hóa/sản phẩm từ bảng hóa đơn PDF. Trả về danh sách dict chứa các trường: `stt`, `name`, `unit`, `quantity`, `price`, `amount`.

### 2. `entity_list_parser.py`
- **Chức năng**: Đọc và phân tích file Word `.docx` chứa danh sách các Hộ kinh doanh / Công ty (Bên Mua).
- **Hàm chính**:
  - `parse_entities(docx_path)`: Phân tích file Word danh sách, trích xuất thông tin các đối tác bao gồm: Tên công ty/hộ kinh doanh, Mã số thuế/Mã số ĐKKD, Địa chỉ, Đại diện, Chức vụ.

### 3. `seller_list.py`
- **Chức năng**: Phân tích file Word chứa danh sách Bên Bán (`List bên bán.docx`).
- **Hàm chính**:
  - `parse_sellers(docx_path)`: Đọc thông tin các đơn vị bên bán (Tên, Mã số thuế, Địa chỉ, Đại diện, Chức vụ, Số tài khoản, Ngân hàng).

### 4. `docx_filler.py`
- **Chức năng**: Điền dữ liệu vào mẫu **Hợp đồng Mua bán** (`.docx`).
- **Hàm chính**:
  - `replace_in_paragraph(paragraph, old_text, new_text)`: Thay thế chuỗi văn bản trong một paragraph của Word mà vẫn giữ nguyên định dạng (kể cả khi text nằm rải rác trên nhiều `run`).
  - `fill_sales_contract(...)`: Thay thế các placeholder bên bán, bên mua, số hợp đồng, ngày tháng và tự động chèn/thay thế bảng sản phẩm trích xuất từ PDF vào file mẫu Word.

### 5. `principle_docx_filler.py`
- **Chức năng**: Điền dữ liệu vào mẫu **Hợp đồng Nguyên tắc** (`.docx`).
- **Hàm chính**:
  - `fill_principle_contract(...)`: Thay thế các thông tin bên bán, bên mua (Hộ kinh doanh/Công ty), số hợp đồng và ngày tháng vào mẫu Hợp đồng Nguyên tắc.

### 6. `generator.py`
- **Chức năng**: Module điều phối quá trình tạo **Hợp đồng Mua bán** hàng loạt.
- **Hàm chính**:
  - `generate_sales_contracts(...)`: Kết hợp `pdf_parser`, `entity_list_parser`, và `docx_filler` để tự động tạo ra các file Hợp đồng Mua bán tương ứng cho từng bên mua trong danh sách và lưu vào thư mục đầu ra.

### 7. `principle_generator.py`
- **Chức năng**: Module điều phối quá trình tạo **Hợp đồng Nguyên tắc** hàng loạt.
- **Hàm chính**:
  - `generate_principle_contracts(...)`: Kết hợp `entity_list_parser` và `principle_docx_filler` để xuất hàng loạt các Hợp đồng Nguyên tắc.

### 8. `gui.py`
- **Chức năng**: Giao diện ứng dụng Tkinter (`tkinter` / `ttk`).
- **Lớp chính**:
  - `SalesContractTab`: Tab quản lý tạo Hợp đồng Mua bán.
  - `PrincipleContractTab`: Tab quản lý tạo Hợp đồng Nguyên tắc.
  - `App`: Cửa sổ ứng dụng chính chứa các tab.
  - `main()`: Hàm khởi chạy ứng dụng GUI.

---

## Đơn Vị Kiểm Thử Tương Ứng

Mỗi submodule trong package này đều có bài test tương ứng trong thư mục `tests/`:
- `pdf_parser.py` -> `tests/test_pdf_parser.py`
- `entity_list_parser.py` -> `tests/test_entity_list_parser.py`
- `seller_list.py` -> `tests/test_seller_list.py`
- `docx_filler.py` -> `tests/test_docx_filler.py`
- `principle_docx_filler.py` -> `tests/test_principle_docx_filler.py`
- `generator.py` -> `tests/test_generator.py`
- `principle_generator.py` -> `tests/test_principle_generator.py`
- `gui.py` -> `tests/test_gui_smoke.py`
