from contract_filler.entity_list_parser import parse_entities

FIXTURE = "tests/fixtures/entity_list.docx"


def test_parse_entities_reads_household_and_company_records():
    entities = parse_entities(FIXTURE)
    assert len(entities) == 25

    hkd = entities[0]
    assert hkd["name"] == "HỘ KINH DOANH TRẦN VĂN HÙNG"
    assert hkd["tax_code"] == "0123456789"
    assert hkd["address"] == "Thôn Đông, Xã Phú Lâm, Hà Nội"
    assert hkd["representative"] == "Trần Văn Hùng"
    assert hkd["position"] == "Chủ hộ"
    assert hkd["bank_account"] == ""

    company = entities[15]
    assert company["name"] == "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ MINH PHÁT"
    assert company["address"] == "Số 12 Nguyễn Trãi, Thanh Xuân, Hà Nội"
    assert company["bank_account"] == "VCB 0123456789"
    assert company["representative"] == "Nguyễn Minh Phát"
    assert company["position"] == "Giám đốc"
    assert company["tax_code"] == ""


def test_parse_entities_household_never_has_bank_account_company_never_has_tax_code():
    entities = parse_entities(FIXTURE)
    for entity in entities:
        if entity["name"].startswith("HỘ KINH DOANH"):
            assert entity["bank_account"] == ""
        elif entity["name"].startswith("CÔNG TY"):
            assert entity["tax_code"] == ""
