import os
import re

import docx
import pdfplumber

from contract_filler.entity_list_parser import parse_entities
from contract_filler.pdf_parser import parse_line_items, parse_totals
from contract_filler.docx_filler import (
    fill_party_info,
    fill_contract_meta,
    fill_product_table_and_totals,
)

_INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')


def sanitize_filename(name):
    cleaned = _INVALID_FILENAME_CHARS.sub("", name).strip()
    return cleaned or "hop_dong"


def generate_sales_contracts(
    pdf_path, template_path, list_path, seller, output_dir, contract_no, day, month, year
):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() for page in pdf.pages)

    items = parse_line_items(pdf_path)
    totals = parse_totals(full_text)
    buyers = parse_entities(list_path)

    os.makedirs(output_dir, exist_ok=True)

    output_paths = []
    for buyer in buyers:
        document = docx.Document(template_path)
        fill_party_info(document, seller, buyer)
        fill_contract_meta(document, contract_no, day, month, year)
        fill_product_table_and_totals(document, items, totals)

        filename = f"HĐMB {sanitize_filename(buyer['name'])}.docx"
        output_path = os.path.join(output_dir, filename)
        document.save(output_path)
        output_paths.append(output_path)

    return output_paths
