import os

import docx

from contract_filler.entity_list_parser import parse_entities
from contract_filler.docx_filler import force_black_text
from contract_filler.generator import sanitize_filename
from contract_filler.principle_docx_filler import (
    fill_principle_party_a,
    fill_principle_party_b,
    fill_principle_contract_meta,
)


def generate_principle_contracts(
    list_path, template_path, output_dir, seller, contract_no, day, month, year
):
    entities = parse_entities(list_path)
    os.makedirs(output_dir, exist_ok=True)

    output_paths = []
    for entity in entities:
        document = docx.Document(template_path)
        fill_principle_party_a(document, entity)
        fill_principle_party_b(document, seller)
        fill_principle_contract_meta(document, contract_no, day, month, year)
        force_black_text(document)

        filename = f"HĐNT {sanitize_filename(entity['name'])}.docx"
        output_path = os.path.join(output_dir, filename)
        document.save(output_path)
        output_paths.append(output_path)

    return output_paths
