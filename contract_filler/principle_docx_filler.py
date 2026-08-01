import unicodedata

from contract_filler.docx_filler import replace_in_paragraph, replace_in_row


def _normalize_table_runs(table):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.text = unicodedata.normalize("NFC", run.text)


def fill_principle_party_a(document, entity):
    for paragraph in document.paragraphs:
        if "BÊN A: BÊN MUA" in paragraph.text:
            replace_in_paragraph(
                paragraph,
                "CÔNG TY TNHH TƯ VẤN PHÁT TRIỂN THỊ TRƯỜNG M.S.V",
                entity["name"],
            )

    table_a = document.tables[1]
    replace_in_row(table_a.rows[0], "ABC, Hà Nội", entity.get("address", ""))
    replace_in_row(table_a.rows[1], "0123456789", entity.get("tax_code", ""))
    replace_in_row(table_a.rows[2], "686345848", entity.get("bank_account", ""))
    replace_in_row(table_a.rows[3], "Ngân hàng TMCP Quân đội", "")
    replace_in_row(table_a.rows[4], "NGUYỄN VĂN A", entity.get("representative", ""))
    replace_in_row(table_a.rows[5], "Giám đốc", entity.get("position", ""))


def fill_principle_party_b(document, seller):
    # the .doc->.docx conversion left some of this table's runs as NFD
    # Unicode, which breaks plain substring matching against NFC literals
    _normalize_table_runs(document.tables[0])

    for paragraph in document.paragraphs:
        if "BÊN B: BÊN BÁN:" in paragraph.text:
            replace_in_paragraph(
                paragraph,
                "CÔNG TY CỔ PHẦN THƯƠNG MẠI ABCXYZ",
                seller["name"],
            )

    table_b = document.tables[0]
    replace_in_row(
        table_b.rows[0],
        "Số 65A-65B Trần Bình Trọng, Phường Bình Lợi Trung, Thành phố Hồ Chí Minh, Việt Nam",
        seller.get("address", ""),
    )
    replace_in_row(table_b.rows[1], "0302629806", seller.get("tax_code", ""))
    replace_in_row(table_b.rows[2], "PHẠM THẾ PHONG", seller.get("representative", ""))
    replace_in_row(table_b.rows[3], "Giám Đốc", seller.get("position", ""))


def fill_principle_contract_meta(document, contract_no, day, month, year):
    for paragraph in document.paragraphs:
        if paragraph.text.strip().startswith("Số "):
            replace_in_paragraph(paragraph, "07/062026/HĐMB", contract_no)
        if "được lập ngày" in paragraph.text:
            replace_in_paragraph(paragraph, "22", day)
            replace_in_paragraph(paragraph, "06", month)
            replace_in_paragraph(paragraph, "2026", year)
