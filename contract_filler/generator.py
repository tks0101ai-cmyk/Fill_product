import docx
import pdfplumber

from contract_filler.pdf_parser import parse_parties, parse_line_items, parse_totals
from contract_filler.docx_filler import (
    fill_party_info,
    fill_contract_meta,
    fill_product_table_and_totals,
)


def generate_contract(pdf_path, template_path, output_path, contract_no, day, month, year):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() for page in pdf.pages)

    seller, buyer = parse_parties(full_text)
    items = parse_line_items(pdf_path)
    totals = parse_totals(full_text)

    document = docx.Document(template_path)
    fill_party_info(document, seller, buyer)
    fill_contract_meta(document, contract_no, day, month, year)
    fill_product_table_and_totals(document, items, totals)

    document.save(output_path)
