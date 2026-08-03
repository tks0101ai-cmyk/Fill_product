# Tài Liệu Đặc Tả Case Sử Dụng (Use Case Specification)

## 1. Biểu Đồ Tổng Quan Use Case (Use Case Diagram)

```mermaid
graph TD
    User((Kế toán / HR / Admin))
    
    subgraph Contract Filler App
        UC1[UC-01: Chọn Tệp & Thư Mục Đầu Vào]
        UC2[UC-02: Tạo Hợp Đồng Mua Bán Hàng Loạt]
        UC3[UC-03: Tạo Hợp Đồng Nguyên Tắc Hàng Loạt]
        UC4[UC-04: Quản Lý Thông Tin Bên Bán & Số Hợp Đồng]
        UC5[UC-05: Đóng Gói Ứng Dụng Thành EXE]
    end

    User --> UC1
    User --> UC2
    User --> UC3
    User --> UC4
    User --> UC5

    UC2 .-> UC1 : include
    UC2 .-> UC4 : include
    UC3 .-> UC1 : include
    UC3 .-> UC4 : include
```

---

## 2. Chi Tiết Các Use Case

### UC-01: Chọn Tệp & Thư Mục Đầu Vào
- **Tóm tắt**: Người dùng duyệt và chọn các tệp hóa đơn PDF, tệp mẫu Word, tệp danh sách bên mua và thư mục lưu kết quả.
- **Tác nhân (Actor)**: Kế toán viên / Nhân viên hành chính.
- **Tiền điều kiện (Pre-conditions)**: Các tệp đầu vào dạng `.pdf` và `.docx` hợp lệ đã có sẵn trên máy tính.
- **Hậu điều kiện (Post-conditions)**: Đường dẫn tệp được lưu vào bộ nhớ ứng dụng và hiển thị trên các ô nhập liệu.
- **Luồng sự kiện chính (Main Flow)**:
  1. Người dùng bấm vào nút "Chọn..." tại ô tương ứng.
  2. Hệ thống hiển thị cửa sổ duyệt tệp/thư mục của hệ điều hành.
  3. Người dùng chọn tệp/thư mục mong muốn và bấm Open/OK.
  4. Hệ thống cập nhật đường dẫn lên giao diện.

---

### UC-02: Tạo Hợp Đồng Mua Bán Hàng Loạt
- **Tóm tắt**: Hệ thống trích xuất bảng sản phẩm từ hóa đơn PDF, kết hợp với danh sách bên mua và thông tin bên bán để sinh ra các file Hợp đồng Mua bán Word `.docx`.
- **Tác nhân (Actor)**: Kế toán viên / Nhân viên hành chính.
- **Tiền điều kiện (Pre-conditions)**:
  - Đã chọn file PDF hóa đơn.
  - Đã chọn file Hợp đồng Mua bán mẫu `.docx`.
  - Đã chọn file danh sách Bên Mua `.docx`.
  - Đã chọn Bên bán, nhập Số hợp đồng và Ngày/Tháng/Năm.
- **Hậu điều kiện (Post-conditions)**: Các file `.docx` hợp đồng mua bán được tạo ra đầy đủ tại thư mục xuất kết quả.
- **Luồng sự kiện chính (Main Flow)**:
  1. Người dùng chọn tab "Hợp đồng Mua bán".
  2. Người dùng nhập/chọn đầy đủ các thông tin đầu vào.
  3. Người dùng bấm nút "Tạo hợp đồng".
  4. Hệ thống đọc bảng sản phẩm từ hóa đơn PDF.
  5. Hệ thống đọc danh sách bên mua từ file Word.
  6. Hệ thống tiến hành thay thế thông tin và bảng sản phẩm vào mẫu Hợp đồng Mua bán cho từng bên mua.
  7. Hệ thống lưu các file Word kết quả.
  8. Hệ thống thông báo thành công và số lượng hợp đồng đã sinh.
- **Luồng ngoại lệ (Exception Flow)**:
  - Nếu thiếu bất kỳ thông tin nào, hệ thống hiển thị hộp thoại cảnh báo: "Vui lòng nhập đầy đủ thông tin!".
  - Nếu tệp PDF bị lỗi hoặc sai định dạng, hệ thống thông báo lỗi chi tiết và dừng quá trình.

---

### UC-03: Tạo Hợp Đồng Nguyên Tắc Hàng Loạt
- **Tóm tắt**: Hệ thống sinh ra hàng loạt Hợp đồng Nguyên tắc cho từng đối tác công ty/hộ kinh doanh trong file danh sách.
- **Tác nhân (Actor)**: Nhân viên hành chính / Pháp chế.
- **Tiền điều kiện (Pre-conditions)**:
  - Đã chọn file Hợp đồng Nguyên tắc mẫu `.docx`.
  - Đã chọn file danh sách Bên Mua `.docx`.
  - Đã chọn Bên bán, nhập Số hợp đồng và Ngày/Tháng/Năm.
- **Hậu điều kiện (Post-conditions)**: Các file Hợp đồng Nguyên tắc Word được sinh ra trong thư mục lưu kết quả.
- **Luồng sự kiện chính (Main Flow)**:
  1. Người dùng chọn tab "Hợp đồng Nguyên tắc".
  2. Người dùng nhập/chọn đầy đủ các thông tin đầu vào.
  3. Người dùng bấm nút "Tạo hợp đồng hàng loạt".
  4. Hệ thống phân tích danh sách bên mua từ file Word.
  5. Hệ thống điền thông tin bên bán, bên mua, số HĐ và ngày tháng vào mẫu Hợp đồng Nguyên tắc.
  6. Hệ thống ghi kết quả ra các file Word tương ứng.
  7. Hệ thống hiển thị hộp thoại thông báo thành công.

---

### UC-04: Quản Lý Thông Tin Bên Bán & Số Hợp Đồng
- **Tóm tắt**: Cho phép người dùng chọn đơn vị Bên Bán từ danh sách có sẵn (`List bên bán.docx`) và tùy chỉnh Số HĐ, Ngày/Tháng/Năm.
- **Tác nhân (Actor)**: Người dùng ứng dụng.
- **Tiền điều kiện (Pre-conditions)**: Tệp `List bên bán.docx` tồn tại trong thư mục ứng dụng.
- **Hậu điều kiện (Post-conditions)**: Thông tin bên bán được chọn sẽ được áp dụng cho toàn bộ hợp đồng sinh ra.
- **Luồng sự kiện chính (Main Flow)**:
  1. Hệ thống tự động nạp danh sách bên bán khi khởi động ứng dụng.
  2. Người dùng chọn một bên bán từ danh sách ComboBox thả xuống.
  3. Người dùng nhập số hợp đồng và ngày, tháng, năm lập hợp đồng.

---

### UC-05: Đóng Gói Ứng Dụng Thành EXE
- **Tóm tắt**: Kỹ sư / Quản trị viên đóng gói toàn bộ mã nguồn Python và dependencies thành file `DienHopDong.exe` chạy độc lập.
- **Tác nhân (Actor)**: Kỹ sư phần mềm / Developer.
- **Tiền điều kiện (Pre-conditions)**: Máy tính đã cài đặt PyInstaller và file `DienHopDong.spec`.
- **Hậu điều kiện (Post-conditions)**: File `dist/DienHopDong.exe` được tạo thành công và chạy mượt mà trên môi trường Windows không cài Python.
- **Luồng sự kiện chính (Main Flow)**:
  1. Developer mở terminal tại thư mục gốc dự án.
  2. Thực thi lệnh `pyinstaller DienHopDong.spec`.
  3. PyInstaller thu thập mã nguồn và thư viện phụ thuộc để đóng gói.
  4. Kiểm tra file `DienHopDong.exe` tạo ra trong thư mục `dist/`.
