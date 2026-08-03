# Suite Kiểm Thử (Tests)

Thư mục này chứa toàn bộ bộ kiểm thử tự động (Unit Tests & Integration Tests) cho ứng dụng Điền Sản Phẩm Hợp Đồng.

---

## Cấu Trúc Thư Mục Test

```text
tests/
├── fixtures/                           # Chứa các file dữ liệu mẫu cho test
│   ├── sample_invoice.pdf              # File hóa đơn PDF mẫu
│   └── sample_template.docx            # File hợp đồng Word mẫu
├── test_docx_filler.py                 # Test logic điền Hợp đồng Mua bán mẫu
├── test_entity_list_parser.py          # Test parser danh sách bên mua (Word)
├── test_generator.py                   # Test generator sinh Hợp đồng Mua bán
├── test_gui_smoke.py                   # Smoke test cho giao diện Tkinter GUI
├── test_pdf_parser.py                  # Test trích xuất bảng sản phẩm từ PDF
├── test_principle_docx_filler.py       # Test logic điền Hợp đồng Nguyên tắc mẫu
├── test_principle_generator.py         # Test generator sinh Hợp đồng Nguyên tắc
└── test_seller_list.py                 # Test parser danh sách bên bán (Word)
```

---

## Cách Chạy Test

Chạy tất cả các bài test từ thư mục gốc của dự án:

```bash
python -m pytest
```

Chạy với thông tin hiển thị chi tiết (verbose):

```bash
python -m pytest -v
```

Chạy một file test cụ thể:

```bash
python -m pytest tests/test_pdf_parser.py
```
