import re

import pdfplumber

BUYER_MARKER = "Tên đơn vị(Company's name):"


def _find(pattern, text):
    match = re.search(pattern, text)
    return match.group(1).strip() if match else ""


def _parse_number(value):
    if not value:
        return 0
    return int(value.replace(".", "").replace(",", "").strip())


def parse_parties(text):
    split_at = text.index(BUYER_MARKER)
    seller_block, buyer_block = text[:split_at], text[split_at:]

    seller = {
        "name": _find(r"Đơn vị bán hàng[^:]*:\s*(.+)", seller_block),
        "tax_code": _find(r"Mã số thuế[^:]*:\s*(\S+)", seller_block),
        "address": _find(r"Địa chỉ[^:]*:\s*(.+)", seller_block),
    }
    buyer = {
        "name": _find(r"Tên đơn vị[^:]*:\s*(.+)", buyer_block),
        "address": _find(r"Địa chỉ[^:]*:\s*(.+)", buyer_block),
        "tax_code": _find(r"Mã số thuế[^:]*:\s*(\S+)", buyer_block),
    }
    return seller, buyer


def parse_line_items(pdf_path):
    items = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table[1:]:
                stt_cell = (row[0] or "").strip()
                if not stt_cell.isdigit():
                    continue
                items.append({
                    "stt": int(stt_cell),
                    "description": (row[1] or "").replace("\n", " ").strip(),
                    "unit": (row[2] or "").strip(),
                    "quantity": _parse_number(row[3]),
                    "unit_price": _parse_number(row[4]),
                    "amount": _parse_number(row[5]),
                })
    return items


def parse_totals(text):
    return {
        "subtotal": _parse_number(
            _find(r"Cộng tiền hàng[^:]*:\s*([\d\.,]+)", text)
        ),
        "vat_rate": _find(r"Thuế suất GTGT[^:]*:\s*(\d+)%", text),
        "vat_amount": _parse_number(
            _find(r"Tiền thuế GTGT[^:]*:\s*([\d\.,]+)", text)
        ),
        "total_payment": _parse_number(
            _find(r"Tổng cộng tiền thanh toán[^:]*:\s*([\d\.,]+)", text)
        ),
        "amount_in_words": _find(r"Số tiền viết bằng chữ[^:]*:\s*(.+)", text),
    }
