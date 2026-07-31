import pdfplumber
from contract_filler.pdf_parser import parse_parties

FIXTURE = "tests/fixtures/sample_invoice.pdf"


def _full_text():
    with pdfplumber.open(FIXTURE) as pdf:
        return "\n".join(page.extract_text() for page in pdf.pages)


def test_parse_parties_extracts_seller_and_buyer():
    seller, buyer = parse_parties(_full_text())

    assert seller["name"] == "CÔNG TY CỔ PHẦN THƯƠNG MẠI XUẤT NHẬP KHẨU TKS GROUP"
    assert seller["tax_code"] == "0110534607"
    assert seller["address"] == (
        "Số 9 đường Lê Văn Huấn, Cụm công nghiệp Cầu Nổi, "
        "Xã Sơn Đồng, Thành phố Hà Nội, Việt Nam"
    )

    assert buyer["name"] == "HỘ KINH DOANH TIỆM 81"
    assert buyer["tax_code"] == "064200012728"
    assert buyer["address"] == (
        "201/65/9 Nguyễn Xí, Phường Bình Thạnh, Thành phố Hồ Chí Minh, Việt Nam"
    )


from contract_filler.pdf_parser import parse_line_items


def test_parse_line_items_reads_all_rows_across_pages():
    items = parse_line_items(FIXTURE)

    assert len(items) == 4
    assert items[0]["stt"] == 1
    assert items[0]["unit"] == "Cái"
    assert items[0]["quantity"] == 7000
    assert items[0]["unit_price"] == 385
    assert items[0]["amount"] == 2695000
    assert "Móc treo đồ gắn tường" in items[0]["description"]

    # item 4 lives on page 2 of the PDF
    assert items[3]["stt"] == 4
    assert items[3]["quantity"] == 100
    assert items[3]["unit_price"] == 12800
    assert items[3]["amount"] == 1280000


from contract_filler.pdf_parser import parse_totals


def test_parse_totals_reads_summary_block():
    totals = parse_totals(_full_text())

    assert totals["subtotal"] == 9315000
    assert totals["vat_rate"] == "8"
    assert totals["vat_amount"] == 745200
    assert totals["total_payment"] == 10060200
    assert totals["amount_in_words"] == (
        "Mười triệu không trăm sáu mươi nghìn hai trăm đồng"
    )
