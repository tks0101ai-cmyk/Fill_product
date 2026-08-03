# Tài Liệu Quy Trình Nghiệp Vụ BPMN (Business Process Model and Notation)

## 1. Giới Thiệu
Tài liệu này mô tả chi tiết các quy trình nghiệp vụ của ứng dụng **Tự động Điền và Sinh Hợp đồng** dưới dạng biểu đồ sơ đồ dòng dữ liệu và BPMN chuẩn hóa.

---

## 2. Quy Trình 1: Sinh Hợp Đồng Mua Bán Từ Hóa Đơn PDF (BPMN Diagram)

Quy trình này mô tả sự tương tác giữa Người dùng (Kế toán), Giao diện GUI, Bộ phân tích PDF, Bộ phân tích Word và Động cơ sinh Hợp đồng Mua bán.

```mermaid
sequenceDiagram
    autonumber
    actor User as Người dùng (Kế toán)
    participant GUI as Giao diện GUI (gui.py)
    participant PDFParser as Bộ đọc PDF (pdf_parser.py)
    participant WordParser as Bộ đọc Bên mua/bán (entity/seller_list.py)
    participant Generator as Bộ điều phối (generator.py)
    participant DocxFiller as Động cơ điền Word (docx_filler.py)
    participant Disk as Ổ đĩa cục bộ

    User->>GUI: 1. Chọn file PDF hóa đơn, Word mẫu HĐMB, Word danh sách Bên mua & Thư mục lưu
    User->>GUI: 2. Chọn Bên bán, Nhập Số HĐ và Ngày/Tháng/Năm
    User->>GUI: 3. Nhấn nút "Tạo hợp đồng"
    
    GUI->>GUI: 4. Kiểm tra ràng buộc dữ liệu đầu vào (Validation)
    alt Dữ liệu thiếu/Không hợp lệ
        GUI-->>User: Hiển thị thông báo lỗi
    else Dữ liệu hợp lệ
        GUI->>Generator: Gọi generate_sales_contracts(...)
        
        Generator->>PDFParser: Đọc file PDF hóa đơn (parse_line_items)
        PDFParser->>Disk: Đọc nội dung PDF
        Disk-->>PDFParser: Nội dung nhị phân PDF
        PDFParser-->>Generator: Trả về danh sách dòng hàng hóa (line_items)
        
        Generator->>WordParser: Đọc danh sách Bên Mua (parse_entities)
        WordParser->>Disk: Đọc file Word danh sách Bên Mua
        Disk-->>WordParser: Nội dung .docx
        WordParser-->>Generator: Trả về danh sách Bên Mua (entities)
        
        loop Đối với từng Bên Mua (Entity) trong danh sách
            Generator->>DocxFiller: Gọi fill_sales_contract(template, entity, seller, line_items, ...)
            DocxFiller->>DocxFiller: Thay thế thông tin Bên A, Bên B, Số HĐ, Ngày/Tháng/Năm
            DocxFiller->>DocxFiller: Chèn & cập nhật bảng hàng hóa từ line_items
            DocxFiller->>Disk: Ghi file hợp đồng Word kết quả (.docx)
        end
        
        Generator-->>GUI: Trả về danh sách đường dẫn file kết quả
        GUI-->>User: Hiển thị thông báo "Tạo hợp đồng thành công!"
    end
```

---

## 3. Quy Trình 2: Sinh Hợp Đồng Nguyên Tắc Hàng Loạt (BPMN Flowchart)

Biểu đồ dưới đây thể hiện luồng xử lý quyết định khi sinh Hợp đồng Nguyên tắc cho danh sách công ty/hộ kinh doanh.

```mermaid
flowchart TD
    Start([Bắt đầu Quy trình]) --> InputSelection[Người dùng chọn File Mẫu HĐNT, File Danh Sách Bên Mua, Thư Mục Lưu]
    InputSelection --> SellerSelect[Chọn Bên Bán & Nhập Thông Tin Hợp Đồng: Số HĐ, Ngày/Tháng/Năm]
    SellerSelect --> ClickSubmit[Nhấn 'Tạo hợp đồng hàng loạt']
    
    ClickSubmit --> Validate{Kiểm tra đầy đủ thông tin?}
    Validate -- Không --> ShowError[Hiển thị thông báo lỗi trên GUI]
    ShowError --> InputSelection
    
    Validate -- Có --> ParseEntityList[Đọc & phân tích File Danh sách Bên Mua .docx]
    ParseEntityList --> LoopStart[Duyệt từng Bên Mua trong Danh sách]
    
    LoopStart --> CheckMore{Còn Bên Mua chưa xử lý?}
    CheckMore -- Có --> ProcessEntity[Lấy thông tin Bên Mua hiện tại]
    ProcessEntity --> FillTemplate[Điền thông tin Bên A, Bên B, Số HĐ, Ngày tháng vào Mẫu HĐNT]
    FillTemplate --> SaveDocx[Lưu file Word hợp đồng vào Thư mục kết quả]
    SaveDocx --> LoopStart
    
    CheckMore -- Không --> SuccessMsg[Hiển thị thông báo Hoàn thành]
    SuccessMsg --> End([Kết thúc Quy trình])
```

---

## 4. Quy Trình 3: Xử Lý Điền Văn Bản Trong Paragraph & Table (Word Engine Detail)

Mô tả thuật toán thay thế văn bản giữ nguyên định dạng (Run-preserving text replacement) khi điền hợp đồng.

```mermaid
flowchart LR
    A[Bắt đầu thay thế từ khóa] --> B[Gộp toàn bộ văn bản trong các Run của Paragraph]
    B --> C{Tìm thấy từ khóa Placeholder?}
    C -- Không --> D[Bỏ qua Paragraph]
    C -- Có --> E[Xác định vị trí Run bắt đầu và kết thúc của từ khóa]
    E --> F[Đặt giá trị mới vào Run đầu tiên]
    F --> G[Xóa văn bản trùng lặp ở các Run tiếp theo]
    G --> H[Hoàn tất giữ nguyên Định dạng Font/Bold/Italic]
```
