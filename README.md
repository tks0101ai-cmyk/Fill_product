# Ứng Dụng Tự Động Điền Sản Phẩm Hợp Đồng (Contract Filler App)

Ứng dụng desktop bằng Python (Tkinter) hỗ trợ trích xuất dữ liệu từ hóa đơn PDF và danh sách doanh nghiệp/hộ kinh doanh (Word `.docx`), tự động sinh hàng loạt các **Hợp đồng Mua bán** và **Hợp đồng Nguyên tắc** dưới dạng file Word (`.docx`).

---

## Tính Năng Chính

1. **Hợp Đồng Mua Bán (Sales Contract)**:
   - Trích xuất tự động danh sách hàng hóa/sản phẩm (STT, Tên hàng, Đơn vị tính, Số lượng, Đơn giá, Thành tiền) từ file hóa đơn PDF (sử dụng `pdfplumber`).
   - Đọc danh sách Bên Mua (Tên công ty/Hộ kinh doanh, Mã số thuế, Địa chỉ, Người đại diện) từ file Word `.docx`.
   - Chọn thông tin Bên Bán từ file danh sách bên bán (`List bên bán.docx`).
   - Tự động thay thế bảng sản phẩm và các thông tin placeholder trong file mẫu `HỢP ĐỒNG MUA BÁN mẫu.docx`.
   - Xuất hàng loạt từng file hợp đồng tương ứng với từng bên mua trong danh sách.

2. **Hợp Đồng Nguyên Tắc (Principle Contract)**:
   - Đọc danh sách Bên Mua từ file Word `.docx`.
   - Ghép thông tin Bên Bán và các trường thông tin hợp đồng (Số hợp đồng, Ngày/Tháng/Năm ký kết).
   - Điền tự động vào mẫu `HỢP ĐỒNG NGUYÊN TẮC mẫu.docx`.
   - Xuất hàng loạt các file hợp đồng nguyên tắc cho từng doanh nghiệp/hộ kinh doanh.

3. **Giao Diện Người Dùng Thân Thiện (GUI)**:
   - Xây dựng bằng `tkinter` / `ttk` với thiết kế dạng Tab trực quan.
   - Hỗ trợ chọn nhanh file mẫu, file dữ liệu, và thư mục lưu kết quả.
   - Tự động nhận diện thư mục mặc định khi chạy trực tiếp hoặc chạy từ file đóng gói `.exe`.

4. **Đóng Gói Ứng Dụng Windows (.exe)**:
   - Sẵn sàng đóng gói bằng PyInstaller thông qua file `DienHopDong.spec` thành một file thực thi duy nhất (`DienHopDong.exe`).

---

## Cấu Trúc Dự Án

```text
.
├── contract_filler/                  # Core module chứa toàn bộ logic xử lý chính
│   ├── __init__.py
│   ├── docx_filler.py                # Xử lý thay thế text & bảng trong Hợp đồng Mua bán mẫu
│   ├── entity_list_parser.py         # Parser đọc danh sách Bên Mua từ file Word (.docx)
│   ├── generator.py                  # Điều phối sinh Hợp đồng Mua bán hàng loạt
│   ├── gui.py                        # Giao diện người dùng Tkinter (Tabs: Mua bán & Nguyên tắc)
│   ├── pdf_parser.py                 # Parser đọc bảng sản phẩm từ hóa đơn PDF
│   ├── principle_docx_filler.py      # Xử lý thay thế placeholders trong Hợp đồng Nguyên tắc mẫu
│   ├── principle_generator.py        # Điều phối sinh Hợp đồng Nguyên tắc hàng loạt
│   └── seller_list.py                # Parser đọc thông tin Bên Bán từ file Word (.docx)
├── docs/                             # Tài liệu kỹ thuật, yêu cầu & thiết kế quy trình
│   ├── BRD.md                        # Business Requirements Document
│   ├── SRS.md                        # Software Requirements Specification
│   ├── BPMN.md                       # Business Process Model & Notation (Biểu đồ quy trình)
│   └── USE_CASES.md                  # Use Case Specification (Biểu đồ & kịch bản Use Case)
├── tests/                            # Bộ kiểm thử tự động (Unit & Integration tests)
│   ├── fixtures/                     # Các file mẫu test (.pdf, .docx)
│   ├── test_docx_filler.py
│   ├── test_entity_list_parser.py
│   ├── test_generator.py
│   ├── test_gui_smoke.py
│   ├── test_pdf_parser.py
│   ├── test_principle_docx_filler.py
│   ├── test_principle_generator.py
│   └── test_seller_list.py
├── Danh sách cty hoặc hộ kinh doanh.docx  # File mẫu danh sách bên mua
├── DienHopDong.spec                  # Cấu hình đóng gói PyInstaller EXE
├── HỢP ĐỒNG MUA BÁN mẫu.docx         # Template Word Hợp đồng Mua bán
├── HỢP ĐỒNG NGUYÊN TẮC mẫu.docx      # Template Word Hợp đồng Nguyên tắc
├── List bên bán.docx                 # File dữ liệu thông tin Bên Bán
├── mẫu HĐ.pdf                        # File hóa đơn PDF mẫu
├── requirements.txt                  # Thư viện phụ thuộc
├── run.py                            # Entry point khởi chạy ứng dụng GUI
└── README.md                         # Tài liệu hướng dẫn dự án (File này)
```

---

## Tài Liệu Dự Án Chi Tiết

Hệ thống được cung cấp đầy đủ bộ tài liệu thiết kế và yêu cầu chuẩn hóa trong thư mục `docs/`:

1. **[BRD (Business Requirements Document)](file:///d:/Điền%20sản%20phâm%20hợp%20đồng/docs/BRD.md)**: Yêu cầu kinh doanh, mục tiêu tối ưu hóa quy trình chứng từ, phạm vi dự án và tiêu chí nghiệm thu.
2. **[SRS (Software Requirements Specification)](file:///d:/Điền%20sản%20phâm%20hợp%20đồng/docs/SRS.md)**: Đặc tả chi tiết yêu cầu chức năng, phi chức năng, ràng buộc hệ thống và kiến trúc phần mềm.
3. **[BPMN (Business Process Model & Notation)](file:///d:/Điền%20sản%20phâm%20hợp%20đồng/docs/BPMN.md)**: Biểu đồ và mô tả chi tiết các quy trình nghiệp vụ (Sequence diagram, Flowchart, thuật toán xử lý Word engine).
4. **[USE CASES (Use Case Specification)](file:///d:/Điền%20sản%20phâm%20hợp%20đồng/docs/USE_CASES.md)**: Biểu đồ Use Case tổng quan và kịch bản chi tiết (Main flow, Alternative flow, Pre/Post-conditions) cho từng chức năng.

---

## Môi Trường & Yêu Cầu

- **Python**: 3.10 trở lên
- **Hệ điều hành**: Windows (Khuyến nghị), Linux / macOS (Hỗ trợ chạy mã nguồn Python)
- **Thư viện chính**:
  - `pdfplumber==0.11.4` (Đọc dữ liệu PDF)
  - `python-docx==1.1.2` (Đọc và ghi file Word `.docx`)
  - `pytest==8.3.3` (Chạy bộ kiểm thử)

---

## Hướng Dẫn Sử Dụng

### 1. Cài Đặt Thư Viện

Mở terminal/powershell tại thư mục dự án và cài đặt các thư viện cần thiết:

```bash
pip install -r requirements.txt
```

### 2. Khởi Chạy Ứng Dụng GUI

Chạy lệnh sau để khởi động giao diện ứng dụng:

```bash
python run.py
```

### 3. Quy Trình Sử Dụng Trên GUI

#### A. Tạo Hợp Đồng Mua Bán:
1. Chuyển sang tab **Hợp đồng Mua bán**.
2. Chọn **File hóa đơn (PDF)** chứa danh sách sản phẩm.
3. Chọn **File hợp đồng mẫu (Word)** (mặc định đã chọn `HỢP ĐỒNG MUA BÁN mẫu.docx`).
4. Chọn **File danh sách hộ KD/công ty (Word)** chứa thông tin các bên mua.
5. Chọn **Bên bán** từ danh sách thả xuống.
6. Chọn **Thư mục lưu kết quả**.
7. Nhập **Số hợp đồng** và **Ngày, Tháng, Năm**.
8. Bấm **Tạo hợp đồng**.

#### B. Tạo Hợp Đồng Nguyên Tắc:
1. Chuyển sang tab **Hợp đồng Nguyên tắc**.
2. Chọn **File hợp đồng mẫu (Word)** (mặc định đã chọn `HỢP ĐỒNG NGUYÊN TẮC mẫu.docx`).
3. Chọn **File danh sách hộ KD/công ty (Word)**.
4. Chọn **Bên bán**, chọn **Thư mục lưu kết quả**, nhập **Số hợp đồng** và **Ngày/Tháng/Năm**.
5. Bấm **Tạo hợp đồng hàng loạt**.

---

## Kiểm Thử (Testing)

Dự án đi kèm bộ kiểm thử tự động toàn diện phủ toàn bộ các module parser, filler, generator và GUI smoke test.

Chạy kiểm thử bằng lệnh:

```bash
python -m pytest
```

Hoặc xem kết quả chi tiết từng testcase:

```bash
python -m pytest -v
```

---

## Đóng Gói Thành File EXE (Windows)

Để tạo file ứng dụng chạy trực tiếp `.exe` cho người dùng cuối không cần cài Python:

1. Cài đặt PyInstaller (nếu chưa có):
   ```bash
   pip install pyinstaller
   ```
2. Thực hiện đóng gói sử dụng file `DienHopDong.spec`:
   ```bash
   pyinstaller DienHopDong.spec
   ```
3. File thực thi `DienHopDong.exe` sẽ được tạo trong thư mục `dist/`.

---

## Giấy Phép & Tác Giả

- Dự án phát triển cho công việc tự động hóa chứng từ hợp đồng.
- Phát triển bởi Nguyễn Ngọc Việt Thắng.
