# Tài Liệu Yêu Cầu Kinh Doanh (Business Requirements Document - BRD)

## 1. Thông Tin Tài Liệu
- **Tên dự án**: Hệ thống Tự động Điền và Sinh Hợp đồng (Contract Filler App)
- **Phiên bản**: 1.0.0
- **Tác giả**: Nguyễn Ngọc Việt Thắng
- **Ngày tạo**: 03/08/2026

---

## 2. Tổng Quan Dự Án & Bối Cảnh Kinh Doanh

### 2.1 Vấn Đề Hiện Tại
Trong quá trình quản lý chứng từ và giao dịch thương mại, việc khởi tạo Hợp đồng Mua bán và Hợp đồng Nguyên tắc cho từng đối tác (Doanh nghiệp, Hộ kinh doanh) hiện đang được thực hiện thủ công:
- Nhân viên phải sao chép từng thông tin doanh nghiệp (Mã số thuế, địa chỉ, đại diện pháp luật) từ file danh sách sang từng file Word hợp đồng.
- Nhân viên phải trích xuất thủ công các dòng sản phẩm, đơn giá, số lượng từ file hóa đơn PDF để tạo bảng hàng hóa trong hợp đồng mua bán.
- Quy trình này mất nhiều thời gian, năng suất thấp và tiềm ẩn nguy cơ sai sót dữ liệu (nhầm lẫn số tiền, thông tin đối tác, mã số thuế).

### 2.2 Mục Tiêu Dự Án
Xây dựng một giải pháp phần mềm tự động hóa hoàn toàn quy trình điền dữ liệu và tạo file hợp đồng Word (`.docx`):
- **Tối ưu thời gian**: Giảm thời gian khởi tạo hàng loạt hợp đồng từ nhiều giờ xuống chỉ còn vài giây.
- **Chính xác tuyệt đối**: Tự động trích xuất chính xác 100% dữ liệu hàng hóa từ hóa đơn PDF và thông tin bên mua/bên bán từ file danh sách Word.
- **Chuẩn hóa quy trình**: Giữ nguyên định dạng, phông chữ và bố cục chuẩn của các mẫu hợp đồng doanh nghiệp.

---

## 3. Phạm Vi Dự Án (Project Scope)

### 3.1 Phạm Vi Thực Hiện (In-Scope)
- **Tự động đọc hóa đơn PDF**: Trích xuất bảng sản phẩm (STT, Tên hàng, Đơn vị tính, Số lượng, Đơn giá, Thành tiền) từ hóa đơn GTGT PDF.
- **Tự động đọc danh sách Bên Mua**: Đọc danh sách Doanh nghiệp / Hộ kinh doanh từ file Word `.docx`.
- **Tự động đọc danh sách Bên Bán**: Đọc thông tin pháp lý bên bán từ file Word `List bên bán.docx`.
- **Sinh Hợp đồng Mua bán hàng loạt**: Điền thông tin pháp lý và chèn bảng hàng hóa vào template `HỢP ĐỒNG MUA BÁN mẫu.docx`.
- **Sinh Hợp đồng Nguyên tắc hàng loạt**: Điền thông tin pháp lý vào template `HỢP ĐỒNG NGUYÊN TẮC mẫu.docx`.
- **Giao diện người dùng Desktop (GUI)**: Ứng dụng Tkinter dạng Tab trực quan, hỗ trợ xem và chọn file/thư mục.
- **Đóng gói file EXE**: Khả năng đóng gói thành ứng dụng chạy trực tiếp trên Windows không cần cài môi trường Python.

### 3.2 Ngoài Phạm Vi (Out-of-Scope)
- Quản lý và lưu trữ dữ liệu tập trung trên Cloud hoặc Cơ sở dữ liệu cơ sở (Database SQL/NoSQL).
- Ký số điện tử (Digital Signature / HSM / SmartCard) trên file Word.
- Tích hợp trực tiếp qua API với các phần mềm hóa đơn điện tử (MISA, meInvoice, VNPT, Viettel).

---

## 4. Các Bên Liên Quan (Stakeholders & Users)

| Vai trò | Mô tả trách nhiệm / Nhu cầu |
|---|---|
| **Kế toán viên** | Cần tạo nhanh hợp đồng mua bán khớp chính xác với hóa đơn xuất kho/hóa đơn GTGT. |
| **Nhân viên Hành chính / Pháp chế** | Cần tạo hàng loạt hợp đồng nguyên tắc cho danh sách đối tác mới gia nhập hệ thống. |
| **Quản lý / Ban giám đốc** | Đảm bảo tính chính xác của chứng từ pháp lý, giảm thiểu rủi ro sai sót hợp đồng. |

---

## 5. Yêu Cầu Kinh Doanh Chi Tiết (Business Requirements)

### BR-01: Tự động hóa trích xuất dữ liệu hóa đơn
Hệ thống phải tự động quét và trích xuất chính xác danh mục hàng hóa, số lượng, đơn giá, thành tiền từ file hóa đơn PDF đầu vào mà không cần người dùng nhập tay lại.

### BR-02: Tạo hợp đồng hàng loạt theo danh sách đối tác
Hệ thống phải hỗ trợ chọn một file danh sách chứa nhiều Hộ kinh doanh / Công ty và tự động sinh ra số lượng file hợp đồng tương ứng trong một lần thực thi.

### BR-03: Bảo toàn định dạng văn bản mẫu
Mọi file hợp đồng Word được sinh ra phải giữ nguyên hoàn toàn font chữ, kiểu chữ (bold/italic), căn lề và định dạng bảng của file template gốc.

### BR-04: Độc lập và bảo mật dữ liệu cục bộ
Ứng dụng phải hoạt động offline 100% trên máy tính cá nhân của người dùng, không truyền tải thông tin hợp đồng hay hóa đơn ra internet nhằm bảo đảm an toàn thông tin doanh nghiệp.

---

## 6. Tiêu Chi Nghiệm Thu Kinh Doanh (Acceptance Criteria)

1. Tốc độ sinh 10 hợp đồng mua bán hoặc nguyên tắc không quá 5 giây.
2. Thông tin đối tác (Mã số thuế, địa chỉ, người đại diện) trong hợp đồng đầu ra khớp 100% với danh sách đầu vào.
3. Bảng sản phẩm trong Hợp đồng Mua bán hiển thị đầy đủ các dòng hàng hóa từ hóa đơn PDF.
4. Giao diện đơn giản, người dùng không có kiến thức IT vẫn thao tác thành công ngay lần đầu sử dụng.
