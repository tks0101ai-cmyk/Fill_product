import docx

_START_LABELS = {"Tên HKD", "Tên công ty"}

_LABEL_TO_KEY = {
    "Tên HKD": "name",
    "Tên công ty": "name",
    "Mã số HKD": "tax_code",
    "Địa chỉ": "address",
    "STK": "bank_account",
    "Đại diện pháp luật": "representative",
    "Chức vụ": "position",
}

_EMPTY_ENTITY = {
    "name": "",
    "tax_code": "",
    "address": "",
    "bank_account": "",
    "representative": "",
    "position": "",
}


def parse_entities(docx_path):
    document = docx.Document(docx_path)
    entities = []
    current = None

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text or ":" not in text:
            continue

        label, _, value = text.partition(":")
        label = label.strip()
        value = value.strip()

        if label in _START_LABELS:
            if current is not None:
                entities.append(current)
            current = dict(_EMPTY_ENTITY)
            current["name"] = value
            continue

        if current is None:
            continue

        key = _LABEL_TO_KEY.get(label)
        if key:
            current[key] = value

    if current is not None:
        entities.append(current)

    return entities
