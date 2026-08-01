import unicodedata

import docx

_NAME = "name"
_BANK = "bank"

# Each canonical field can show up under several different labels in the
# source list (users retype these by hand and don't stick to one wording).
# Matching is diacritic- and case-insensitive (see _normalize_label), so
# e.g. "MST" and "mst" and "Mst" all resolve the same way.
_FIELD_SYNONYMS = {
    _NAME: [
        "Tên", "Tên HKD", "Tên công ty", "Tên cty", "Tên hộ kinh doanh",
        "Tên doanh nghiệp", "Tên đơn vị",
    ],
    "tax_code": [
        "MST", "Mã số thuế", "Mã số HKD", "Mã số doanh nghiệp",
        "Mã số đăng ký kinh doanh", "Mã số ĐKKD",
    ],
    "address": [
        "ĐC", "Địa chỉ",
    ],
    "phone": [
        "SĐT", "Số điện thoại", "ĐT", "Điện thoại",
    ],
    "representative": [
        "ĐDPL", "Đại diện pháp luật", "Người đại diện",
        "Người đại diện pháp luật", "Đại diện",
    ],
    "position": [
        "CV", "Chức vụ",
    ],
    # Stored as "STK: <ngân hàng> <số tài khoản>" and split below.
    _BANK: [
        "STK", "Số tài khoản", "Tài khoản",
    ],
}


def _normalize_label(label):
    label = label.replace("đ", "d").replace("Đ", "D")
    decomposed = unicodedata.normalize("NFD", label)
    without_accents = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    return " ".join(without_accents.lower().split())


_LABEL_TO_KEY = {
    _normalize_label(synonym): key
    for key, synonyms in _FIELD_SYNONYMS.items()
    for synonym in synonyms
}

_EMPTY_ENTITY = {
    "name": "",
    "tax_code": "",
    "address": "",
    "phone": "",
    "bank_account": "",
    "bank_name": "",
    "representative": "",
    "position": "",
}


def _apply_field(entity, key, value):
    if key == _BANK:
        bank_name, _, bank_account = value.partition(" ")
        entity["bank_name"] = bank_name
        entity["bank_account"] = bank_account.strip()
    else:
        entity[key] = value


def parse_entities(docx_path):
    document = docx.Document(docx_path)
    entities = []
    block = []

    def flush():
        if not block:
            return
        entity = dict(_EMPTY_ENTITY)
        for key, value in block:
            _apply_field(entity, key, value)
        if entity["name"]:
            entities.append(entity)
        block.clear()

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            flush()
            continue
        if ":" not in text:
            continue

        label, _, value = text.partition(":")
        key = _LABEL_TO_KEY.get(_normalize_label(label.strip()))
        if key is None:
            continue

        # A file without blank-line separators still needs a boundary: a
        # second "name" field means a new record has started.
        if key == _NAME and any(existing_key == _NAME for existing_key, _ in block):
            flush()

        block.append((key, value.strip()))

    flush()
    return entities
