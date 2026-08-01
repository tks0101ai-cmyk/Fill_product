from contract_filler.seller_list import parse_sellers

FIXTURE = "tests/fixtures/seller_list.docx"


def test_parse_sellers_reads_all_four_companies_in_order():
    sellers = parse_sellers(FIXTURE)
    assert len(sellers) == 4
    assert sellers[0]["name"] == "CÔNG TY CỔ PHẦN THƯƠNG MẠI VC EXPRESS"
    assert sellers[1]["name"] == "CÔNG TY CỔ PHẦN THƯƠNG MẠI XUẤT NHẬP KHẨU TKS GROUP"
    assert sellers[2]["name"] == "CÔNG TY CỔ PHẦN XUẤT NHẬP KHẨU THƯƠNG MẠI LUCKY HOUSE"
    assert sellers[3]["name"] == "CÔNG TY CỔ PHẦN THƯƠNG MẠI TỔNG HỢP SHC GROUP"


def test_parse_sellers_reads_full_field_set_for_one_entry():
    sellers = parse_sellers(FIXTURE)
    tks = sellers[1]
    assert tks["tax_code"] == "0110534607"
    assert tks["address"] == (
        "Số 9 đường Lê Văn Huấn, cụm công nghiệp Cầu Nổi, "
        "Xã Sơn Đồng, Thành phố Hà Nội, Việt Nam"
    )
    assert tks["bank_account"] == "6062.666.88888"
    assert tks["bank_name"] == "Ngân hàng TMCP Quân đội – PGD Đông Đô"
    assert tks["representative"] == "Nguyễn Xuân Khoa"
    assert tks["position"] == "Giám đốc"
