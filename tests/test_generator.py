import os
import docx
from contract_filler.generator import generate_contract


def test_generate_contract_produces_filled_docx(tmp_path):
    output_path = str(tmp_path / "output.docx")

    generate_contract(
        pdf_path="tests/fixtures/sample_invoice.pdf",
        template_path="tests/fixtures/sample_template.docx",
        output_path=output_path,
        contract_no="0099/2026/PPR/HĐMBHH",
        day="31",
        month="07",
        year="2026",
    )

    assert os.path.exists(output_path)
    document = docx.Document(output_path)
    full_text = "\n".join(p.text for p in document.paragraphs)

    assert "CÔNG TY CỔ PHẦN THƯƠNG MẠI XUẤT NHẬP KHẨU TKS GROUP" in full_text
    assert "HỘ KINH DOANH TIỆM 81" in full_text
    assert "0099/2026/PPR/HĐMBHH" in full_text
    assert "Mười triệu không trăm sáu mươi nghìn hai trăm đồng" in full_text

    table = document.tables[2]
    assert len(table.rows) == 8
    assert "10.060.200" in table.rows[7].cells[-1].text
