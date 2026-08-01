from contract_filler.entity_list_parser import parse_entities

FIXTURE = "tests/fixtures/entity_list.docx"
VARIANTS_FIXTURE = "tests/fixtures/entity_list_variants.docx"


def test_parse_entities_reads_household_and_company_records():
    entities = parse_entities(FIXTURE)
    assert len(entities) == 25

    hkd = entities[0]
    assert hkd["name"] == "HỘ KINH DOANH TRẦN VĂN HÙNG"
    assert hkd["tax_code"] == "0123456789"
    assert hkd["address"] == "Thôn Đông, Xã Phú Lâm, Hà Nội"
    assert hkd["phone"] == "0912345678"
    assert hkd["representative"] == "Trần Văn Hùng"
    assert hkd["position"] == "Chủ hộ"
    assert hkd["bank_account"] == ""
    assert hkd["bank_name"] == ""

    company = entities[15]
    assert company["name"] == "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ MINH PHÁT"
    assert company["address"] == "Số 12 Nguyễn Trãi, Thanh Xuân, Hà Nội"
    assert company["bank_name"] == "VCB"
    assert company["bank_account"] == "0123456789"
    assert company["phone"] == "0911222333"
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


def test_parse_entities_tolerates_scrambled_order_and_synonym_labels():
    entities = parse_entities(VARIANTS_FIXTURE)
    assert len(entities) == 2

    household = entities[0]
    assert household["name"] == "HỘ KINH DOANH TEST ONE"
    assert household["phone"] == "0900000001"
    assert household["address"] == "Số 1, Hà Nội"
    assert household["position"] == "Chủ hộ"
    assert household["tax_code"] == "1111111111"
    assert household["representative"] == "Nguyễn Văn Test"

    company = entities[1]
    assert company["name"] == "CÔNG TY TNHH TEST TWO"
    assert company["position"] == "Giám đốc"
    assert company["bank_name"] == "TCB"
    assert company["bank_account"] == "2222222222"
    assert company["representative"] == "Trần Thị Test"
    assert company["address"] == "Số 2, Hà Nội"
    assert company["phone"] == "0900000002"
