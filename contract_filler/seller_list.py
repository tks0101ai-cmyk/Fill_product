import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

SELLER_MARKER = "BÊN BÁN:"

_LABEL_TO_KEY = {
    "địa chỉ": "address",
    "mã số thuế": "tax_code",
    "tài khoản số": "bank_account",
    "mở tại ngân hàng": "bank_name",
    "đại diện": "representative",
    "chức vụ": "position",
}

_EMPTY_SELLER = {
    "name": "",
    "address": "",
    "tax_code": "",
    "bank_account": "",
    "bank_name": "",
    "representative": "",
    "position": "",
}


def _row_value(row):
    cells = row.cells
    last = cells[-1].text.strip()
    if last and last != ":":
        if ":" in last:
            return last.split(":", 1)[1].strip()
        return last

    for cell in cells:
        text = cell.text
        if ":" in text:
            after = text.split(":", 1)[1].strip()
            if after:
                return after
    return ""


def parse_sellers(docx_path):
    document = docx.Document(docx_path)
    sellers = []
    current = None

    for child in document.element.body.iterchildren():
        if child.tag.endswith("}p"):
            paragraph = Paragraph(child, document)
            text = paragraph.text.strip()
            if SELLER_MARKER not in text:
                continue
            if current is not None:
                sellers.append(current)
            current = dict(_EMPTY_SELLER)
            current["name"] = text.split(SELLER_MARKER, 1)[1].strip()
        elif child.tag.endswith("}tbl"):
            if current is None:
                continue
            table = Table(child, document)
            for row in table.rows:
                label = row.cells[0].text.strip().lower()
                key = _LABEL_TO_KEY.get(label)
                if key:
                    current[key] = _row_value(row)

    if current is not None:
        sellers.append(current)

    return sellers
