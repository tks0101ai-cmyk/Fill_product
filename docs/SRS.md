# Tài Liệu Đặc Tả Yêu Cầu Phần Mềm (Software Requirements Specification - SRS)

## 1. Giới Thiệu

### 1.1 Mục Đích
Tài liệu Đặc tả Yêu cầu Phần mềm (SRS) này mô tả chi tiết các yêu cầu chức năng, yêu cầu phi chức năng, kiến trúc hệ thống và giao diện cho ứng dụng **Tự động Điền và Sinh Hợp đồng (Contract Filler App)**.

### 1.2 Phạm Vi Tài Liệu
Tài liệu dành cho các kỹ sư phần mềm, kiểm thử viên (Tester) và quản lý dự án để định hướng phát triển, bảo trì và kiểm thử ứng dụng.

---

## 2. Mô Tả Tổng Quan Hệ Thống

### 2.1 Kiến Trúc Mô-đun (Module Architecture)
Hệ thống được thiết kế theo kiến trúc mảng mô-đun độc lập (Modular Architecture):

```text
+-------------------------------------------------------------------+
|                           gui.py (GUI)                            |
|             (Tkinter App / SalesTab & PrincipleTab)              |
+---------------------------------+---------------------------------+
                                  |
                                  v
+---------------------------------+---------------------------------+
|       generator.py              |      principle_generator.py     |
+---------------+-----------------+----------------+----------------+
                |                                  |
   +------------+------------+         +-----------+------------+
   |                         |         |                        |
   v                         v         v                        v
pdf_parser.py       seller_list.py  entity_list_parser.py  principle_docx_filler.py
                            |
                            v
                     docx_filler.py
```

---

## 3. Yêu Cầu Chức Năng (Functional Requirements)

### FR-01: Trích xuất Dữ liệu Hóa đơn PDF (`pdf_parser.py`)
- **FR-01.1**: Hệ thống sử dụng thư viện `pdfplumber` để đọc nội dung file hóa đơn PDF.
- **FR-01.2**: Trích xuất bảng danh mục hàng hóa bao gồm các cột: `stt`, `name` (Tên hàng hóa dịch vụ), `unit` (Đơn vị tính), `quantity` (Số lượng), `price` (Đơn giá), `amount` (Thành tiền).
- **FR-01.3**: Xử lý và làm sạch định dạng số (giữ nguyên dấu chấm phân cách hàng nghìn theo chuẩn Việt Nam).

### FR-02: Đọc Danh Sách Bên Mua Từ File Word (`entity_list_parser.py`)
- **FR-02.1**: Đọc file `.docx` chứa danh sách Hộ kinh doanh / Công ty.
- **FR-02.2**: Phân tích cú pháp văn bản hoặc bảng để tách các trường thông tin:
  - Tên đơn vị (Bên Mua)
  - Mã số thuế / Mã số ĐKKD
  - Địa chỉ đăng ký kinh doanh
  - Người đại diện pháp luật & Chức vụ

### FR-03: Đọc Danh Sách Bên Bán (`seller_list.py`)
- **FR-03.1**: Đọc dữ liệu từ file `List bên bán.docx`.
- **FR-03.2**: Trích xuất thông tin các đơn vị bên bán (Tên, MST, Địa chỉ, Đại diện, Chức vụ, Số tài khoản, Ngân hàng) để hiển thị lên ComboBox trên giao diện.

### FR-04: Xử Lý Điền Mẫu Hợp Đồng Mua Bán (`docx_filler.py`)
- **FR-04.1**: Thay thế các chuỗi placeholder đại diện cho Bên Bán, Bên Mua, Số hợp đồng, Ngày/Tháng/Năm.
- **FR-04.2**: Cung cấp hàm `replace_in_paragraph` hỗ trợ thay thế chuỗi văn bản bị phân tách thành nhiều `run` trong `python-docx`.
- **FR-04.3**: Tìm vị trí bảng sản phẩm mẫu trong `HỢP ĐỒNG MUA BÁN mẫu.docx`, tự động nhân bản dòng bảng và điền dữ liệu hàng hóa trích xuất từ PDF.

### FR-05: Xử Lý Điền Mẫu Hợp Đồng Nguyên Tắc (`principle_docx_filler.py`)
- **FR-05.1**: Thay thế thông tin Bên Bán, Bên Mua, Số hợp đồng, Ngày/Tháng/Năm vào file `HỢP ĐỒNG NGUYÊN TẮC mẫu.docx`.

### FR-06: Điều Phối Sinh Hợp Đồng Hàng Loạt (`generator.py` & `principle_generator.py`)
- **FR-06.1**: Duyệt qua từng Bên Mua trong danh sách được đọc từ `entity_list_parser.py`.
- **FR-06.2**: Gọi hàm fill tương ứng và ghi kết quả ra file `.docx` riêng biệt trong thư mục đích với quy tắc đặt tên rõ ràng.

### FR-07: Giao Diện Người Dùng (`gui.py`)
- **FR-07.1**: Xây dựng giao diện Tkinter dạng Tab (Tab 1: Hợp đồng Mua bán, Tab 2: Hợp đồng Nguyên tắc).
- **FR-07.2**: Hỗ trợ bộ chọn file (`filedialog.askopenfilename`) và chọn thư mục (`filedialog.askdirectory`).
- **FR-07.3**: Hiển thị thông báo lỗi (`messagebox.showerror`) khi thiếu trường dữ liệu hoặc file bị lỗi.

---

## 4. Yêu Cầu Phi Chức Năng (Non-Functional Requirements)

### NFR-01: Hiệu Năng (Performance)
- Thời gian xử lý đọc 1 file PDF và tạo 1 file Hợp đồng Mua bán Word không quá 0.5 giây.
- Sinh hàng loạt 20 hợp đồng trong thời gian dưới 3 giây.

### NFR-02: Độ Tương Thích & Môi Trường (Compatibility)
- Hoạt động tốt trên Windows 10, Windows 11.
- Mẫu Word tạo ra tương thích hoàn toàn với Microsoft Word 2016, 2019, 365 và WPS Office.

### NFR-03: An Toàn & Bảo Mật (Security)
- Toàn bộ quá trình đọc, ghi và xử lý file diễn ra hoàn toàn trong bộ nhớ máy cục bộ (Local RAM & Disk).
- Không yêu cầu quyền Administrator để khởi chạy ứng dụng.

### NFR-04: Đóng Gói Khởi Chạy (Portability)
- File thực thi `.exe` được tạo bằng PyInstaller độc lập (Single Executable), không yêu cầu máy cài đặt sẵn Python.

---

## 5. Ràng Buộc Hệ Thống (System Constraints)

1. **Định dạng Hóa đơn PDF**: Phải theo đúng mẫu hóa đơn điện tử GTGT chuẩn (các nhãn trường như "Đơn vị bán hàng", các cột bảng sản phẩm STT/Tên hàng hóa/Đơn vị tính/Số lượng/Đơn giá/Thành tiền).
2. **Cấu trúc File Template Word**: Không thay đổi các từ khóa placeholder trong file mẫu trừ khi cập nhật mã nguồn tương ứng.
