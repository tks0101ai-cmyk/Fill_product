# Hợp Đồng Nguyên Tắc Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a second feature, "Tạo hợp đồng nguyên tắc", that reads a template contract and a list of hộ kinh doanh/công ty, and generates one filled `.docx` "Hợp đồng nguyên tắc" per entity with that entity's info in Bên A, leaving any missing field blank.

**Architecture:** Mirrors the existing "Tạo hợp đồng mua bán" feature's layered structure (parser → docx filler → generator → GUI). A new parser reads the "Danh sách" list docx into plain dicts. A new docx-filler module fills Bên A's name paragraph and info table plus the contract number/date in a copy of the template, reusing the existing generic `replace_in_paragraph`/`replace_in_row` helpers. A new generator loops over parsed entities, producing one output file per entity in a chosen folder. The GUI becomes a two-tab `ttk.Notebook`: the existing single-contract tab, and a new batch tab.

**Tech Stack:** Python 3.12, `python-docx` 1.1.2, `pytest` 8.3.3, `tkinter`/`ttk` (already in use). No new dependencies.

## Global Constraints

- No new third-party dependencies — reuse `python-docx` and `pytest` already pinned in [requirements.txt](requirements.txt).
- Missing entity fields must be left blank in the output document (never raise, never write a placeholder string like "N/A").
- Preserve the existing "Tạo hợp đồng mua bán" feature and its tests untouched in behavior.
- Follow the existing code style: small pure functions operating on a `docx.Document`, generic substring-based replacement helpers, no comments unless explaining a non-obvious constraint.
- All new/modified Python files must keep passing `pytest` (run from the repo root).

---

## Background: template structure discovered during investigation

The source template `HỢP ĐỒNG ABC mẫu.doc` is a legacy binary `.doc` (OLE) file that `python-docx` cannot open — it must be converted to `.docx` first (Task 1 does this with a Word COM automation script, since Microsoft Word and `pywin32` are available in this environment).

Once converted, the template has this exact structure (paragraph indices and table indices are stable across `python-docx`'s `.paragraphs` / `.tables` collections):

- Paragraph containing `"Số 07/062026/HĐMB"` → the contract number placeholder (starts with `"Số "`).
- Paragraph containing `"được lập ngày 22/06/2026"` → contains day `"22"`, month `"06"`, year `"2026"` as separate substrings.
- Paragraph containing `"BÊN A: BÊN MUA"` → also contains the placeholder company name `"CÔNG TY TNHH TƯ VẤN PHÁT TRIỂN THỊ TRƯỜNG M.S.V"` to replace with each entity's name.
- `document.tables[0]` → Bên B's (the operating company's) real info — **not touched** by this feature.
- `document.tables[1]` → Bên A's placeholder info table, 6 rows × 3 columns, value always in the last cell of the row:
  - row 0: `"ABC, Hà Nội "` → address
  - row 1: `"0123456789"` → tax code (mã số thuế / mã số HKD)
  - row 2: `"686345848"` → bank account number
  - row 3: `"Ngân hàng TMCP Quân đội"` → bank name (never provided by the source list, always cleared to blank)
  - row 4: `"NGUYỄN VĂN A"` → representative
  - row 5: `"Giám đốc "` → position

The source list `Danh sách cty hoặc hộ kd cần điền (giả định).docx` has no tables — it's a flat sequence of `"Label: value"` paragraphs, one entity after another, each entity starting with either `"Tên HKD: ..."` (hộ kinh doanh) or `"Tên công ty: ..."` (công ty). Observed labels: `Tên HKD` / `Tên công ty`, `Mã số HKD` (HKD only), `Địa chỉ`, `SĐT` (not used by the template, dropped), `STK` (công ty only), `Đại diện pháp luật`, `Chức vụ`. A hộ kinh doanh entity never has `STK`; a công ty entity never has `Mã số HKD` — both cases must resolve to an empty string for that field in the output, which is exactly what "leave blank when missing" means here.

---

### Task 1: Convert the legacy `.doc` template to the production `.docx` template asset

**Files:**
- Create: `HỢP ĐỒNG ABC mẫu.docx` (repo root, committed binary asset — the production template used by the app, analogous to the existing [HĐMB mẫu.docx](HĐMB%20mẫu.docx))
- Read: `HỢP ĐỒNG ABC mẫu.doc` (existing source file, untouched)

**Interfaces:**
- Produces: `HỢP ĐỒNG ABC mẫu.docx`, an OOXML file openable by `docx.Document(...)`, matching the structure documented in "Background" above. Later tasks' production code (not tests — tests use small fixtures) reads this file as the default template path in the GUI.

- [ ] **Step 1: Run the Word COM conversion script**

Save this as a throwaway script (e.g. `scratch_convert.py` in the repo root, not committed) and run it once:

```python
import os
import win32com.client

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
try:
    src = os.path.abspath("HỢP ĐỒNG ABC mẫu.doc")
    dst = os.path.abspath("HỢP ĐỒNG ABC mẫu.docx")
    doc = word.Documents.Open(src)
    doc.SaveAs(dst, FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
    doc.Close()
finally:
    word.Quit()
```

Run: `python scratch_convert.py`
Expected: no exception; `HỢP ĐỒNG ABC mẫu.docx` exists in the repo root.

- [ ] **Step 2: Verify the converted file matches the documented structure**

Run this verification script (also throwaway):

```python
import docx

d = docx.Document("HỢP ĐỒNG ABC mẫu.docx")
full_text = "\n".join(p.text for p in d.paragraphs)
assert "Số 07/062026/HĐMB" in full_text
assert "được lập ngày 22/06/2026" in full_text
assert "BÊN A: BÊN MUA" in full_text
assert "CÔNG TY TNHH TƯ VẤN PHÁT TRIỂN THỊ TRƯỜNG M.S.V" in full_text
assert len(d.tables) >= 2
table_a = d.tables[1]
assert len(table_a.rows) == 6
assert table_a.rows[0].cells[-1].text.strip() == "ABC, Hà Nội"
assert table_a.rows[1].cells[-1].text.strip() == "0123456789"
assert table_a.rows[4].cells[-1].text.strip() == "NGUYỄN VĂN A"
print("OK")
```

Run: `python scratch_verify.py`
Expected: prints `OK` with no assertion errors.

- [ ] **Step 3: Delete the throwaway scripts and commit the asset**

```bash
rm scratch_convert.py scratch_verify.py
git add "HỢP ĐỒNG ABC mẫu.docx"
git commit -m "feat: add converted docx template for hop dong nguyen tac"
```

---

### Task 2: Build small test fixtures for the new feature

**Files:**
- Create: `tests/fixtures/principle_contract_template.docx`
- Create: `tests/fixtures/household_company_list.docx`

**Interfaces:**
- Produces: two small, fast-loading fixture files used by every test in Tasks 3–6. `principle_contract_template.docx` mirrors the real template's structure from Task 1 (same placeholder strings, same paragraph markers, same table shape) but trimmed to a handful of paragraphs. `household_company_list.docx` has 3 entities: one complete hộ kinh doanh, one complete công ty, and one hộ kinh doanh missing `Địa chỉ` and `Chức vụ` (to exercise the "leave blank" behavior).

- [ ] **Step 1: Write and run the fixture-builder script**

Save as a throwaway script (e.g. `scratch_build_fixtures.py` in the repo root, not committed) and run it once:

```python
import docx

# --- principle_contract_template.docx ---
d = docx.Document()
d.add_paragraph("CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM")
d.add_paragraph("Độc lập – Tự do – Hạnh Phúc")
d.add_paragraph("HỢP ĐỒNG NGUYÊN TẮC VỀ MUA BÁN HÀNG HÓA")
d.add_paragraph("Số 07/062026/HĐMB")
d.add_paragraph(
    "HỢP ĐỒNG NGUYÊN TẮC VỀ MUA BÁN HÀNG HÓA này được lập ngày "
    "22/06/2026 bởi và giữa:"
)
d.add_paragraph("BÊN A: BÊN MUA :  CÔNG TY TNHH TƯ VẤN PHÁT TRIỂN THỊ TRƯỜNG M.S.V")
d.add_paragraph("BÊN B: BÊN BÁN: CÔNG TY CỔ PHẦN THƯƠNG MẠI ABCXYZ")

table_b = d.add_table(rows=1, cols=1)
table_b.rows[0].cells[0].text = "Bên B info (not touched by this feature)"

table_a = d.add_table(rows=6, cols=3)
rows_a = [
    (" Địa chỉ", ":", "ABC, Hà Nội "),
    (" Mã số thuế", ":", "0123456789"),
    (" Tài khoản số", ":", "686345848"),
    (" Mở tại ngân hàng", ":", "Ngân hàng TMCP Quân đội"),
    (" Đại diện", ":", "NGUYỄN VĂN A"),
    (" Chức vụ", ":", "Giám đốc "),
]
for row, (label, colon, value) in zip(table_a.rows, rows_a):
    row.cells[0].text = label
    row.cells[1].text = colon
    row.cells[2].text = value

d.save("tests/fixtures/principle_contract_template.docx")

# --- household_company_list.docx ---
d2 = docx.Document()
for line in [
    "Tên HKD: HỘ KINH DOANH NGUYỄN VĂN TEST",
    "Mã số HKD: 0000000001",
    "Địa chỉ: 123 Đường Test, Hà Nội",
    "SĐT: 0900000001",
    "Đại diện pháp luật: Nguyễn Văn Test",
    "Chức vụ: Chủ hộ",
]:
    d2.add_paragraph(line)

for line in [
    "Tên công ty: CÔNG TY TNHH TEST ABC",
    "Địa chỉ: 456 Đường Test, TP.HCM",
    "STK: VCB 0000000002",
    "SĐT: 0900000002",
    "Đại diện pháp luật: Trần Thị Test",
    "Chức vụ: Giám đốc",
]:
    d2.add_paragraph(line)

for line in [
    "Tên HKD: HỘ KINH DOANH THIẾU THÔNG TIN",
    "Mã số HKD: 0000000003",
    "Đại diện pháp luật: Lê Văn Thiếu",
]:
    d2.add_paragraph(line)

d2.save("tests/fixtures/household_company_list.docx")
print("fixtures written")
```

Run: `python scratch_build_fixtures.py`
Expected: prints `fixtures written`; both files exist under `tests/fixtures/`.

- [ ] **Step 2: Delete the throwaway script and commit the fixtures**

```bash
rm scratch_build_fixtures.py
git add tests/fixtures/principle_contract_template.docx tests/fixtures/household_company_list.docx
git commit -m "test: add fixtures for hop dong nguyen tac feature"
```

---

### Task 3: Parse the household/company list into entity dicts

**Files:**
- Create: `contract_filler/principle_list_parser.py`
- Test: `tests/test_principle_list_parser.py`

**Interfaces:**
- Produces: `parse_entities(docx_path: str) -> list[dict]`. Each dict always has the keys `name`, `tax_code`, `address`, `bank_account`, `representative`, `position`, all strings, defaulting to `""` when the source list doesn't provide that label for an entity. Later tasks (4, 6) consume this exact dict shape.

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_principle_list_parser.py
from contract_filler.principle_list_parser import parse_entities

FIXTURE = "tests/fixtures/household_company_list.docx"


def test_parse_entities_reads_household_and_company_records():
    entities = parse_entities(FIXTURE)
    assert len(entities) == 3

    hkd = entities[0]
    assert hkd["name"] == "HỘ KINH DOANH NGUYỄN VĂN TEST"
    assert hkd["tax_code"] == "0000000001"
    assert hkd["address"] == "123 Đường Test, Hà Nội"
    assert hkd["representative"] == "Nguyễn Văn Test"
    assert hkd["position"] == "Chủ hộ"
    assert hkd["bank_account"] == ""

    company = entities[1]
    assert company["name"] == "CÔNG TY TNHH TEST ABC"
    assert company["address"] == "456 Đường Test, TP.HCM"
    assert company["bank_account"] == "VCB 0000000002"
    assert company["representative"] == "Trần Thị Test"
    assert company["position"] == "Giám đốc"
    assert company["tax_code"] == ""


def test_parse_entities_leaves_missing_fields_blank():
    entities = parse_entities(FIXTURE)
    incomplete = entities[2]
    assert incomplete["name"] == "HỘ KINH DOANH THIẾU THÔNG TIN"
    assert incomplete["tax_code"] == "0000000003"
    assert incomplete["representative"] == "Lê Văn Thiếu"
    assert incomplete["address"] == ""
    assert incomplete["position"] == ""
    assert incomplete["bank_account"] == ""
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_principle_list_parser.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'contract_filler.principle_list_parser'`

- [ ] **Step 3: Write the implementation**

```python
# contract_filler/principle_list_parser.py
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_principle_list_parser.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add contract_filler/principle_list_parser.py tests/test_principle_list_parser.py
git commit -m "feat: parse household/company list into entity dicts"
```

---

### Task 4: Expose the row-replacement helper as a public function

**Files:**
- Modify: `contract_filler/docx_filler.py:35-79`
- Test: `tests/test_docx_filler.py` (existing tests must still pass unmodified)

**Interfaces:**
- Consumes: nothing new.
- Produces: `replace_in_row(row, old, new) -> bool` (renamed from the existing private `_replace_in_row`, same behavior — searches every cell's paragraphs for `old` and replaces the first match). Task 5's `principle_docx_filler.py` imports this.

- [ ] **Step 1: Rename the function and update its call sites**

In `contract_filler/docx_filler.py`, rename `_replace_in_row` to `replace_in_row` (drop the leading underscore) and update the two call sites inside `fill_party_info` that currently call `_replace_in_row(...)` to call `replace_in_row(...)` instead. No other logic changes.

- [ ] **Step 2: Run the full existing test suite to confirm no regression**

Run: `pytest tests/test_docx_filler.py tests/test_generator.py -v`
Expected: PASS (all previously-passing tests still pass; nothing depended on the leading underscore).

- [ ] **Step 3: Commit**

```bash
git add contract_filler/docx_filler.py
git commit -m "refactor: expose replace_in_row as a public helper"
```

---

### Task 5: Fill Bên A and the contract meta in the principle-contract template

**Files:**
- Create: `contract_filler/principle_docx_filler.py`
- Test: `tests/test_principle_docx_filler.py`

**Interfaces:**
- Consumes: `replace_in_paragraph(paragraph, old, new) -> bool` and `replace_in_row(row, old, new) -> bool` from `contract_filler.docx_filler` ([docx_filler.py](contract_filler/docx_filler.py)); an entity dict shaped as produced by `parse_entities` from Task 3 (`name`, `tax_code`, `address`, `bank_account`, `representative`, `position`).
- Produces: `fill_principle_party_a(document, entity)` and `fill_principle_contract_meta(document, contract_no, day, month, year)`, both mutating a `docx.Document` in place. Task 6's `principle_generator.py` calls both.

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_principle_docx_filler.py
import docx
from contract_filler.principle_docx_filler import (
    fill_principle_party_a,
    fill_principle_contract_meta,
)

TEMPLATE = "tests/fixtures/principle_contract_template.docx"

ENTITY = {
    "name": "HỘ KINH DOANH NGUYỄN VĂN TEST",
    "tax_code": "0000000001",
    "address": "123 Đường Test, Hà Nội",
    "bank_account": "",
    "representative": "Nguyễn Văn Test",
    "position": "Chủ hộ",
}


def test_fill_principle_party_a_sets_fields_and_clears_placeholder():
    document = docx.Document(TEMPLATE)

    fill_principle_party_a(document, ENTITY)

    full_text = "\n".join(p.text for p in document.paragraphs)
    assert ENTITY["name"] in full_text
    assert "CÔNG TY TNHH TƯ VẤN PHÁT TRIỂN THỊ TRƯỜNG M.S.V" not in full_text

    table_text = "\n".join(
        "\n".join(cell.text for cell in row.cells)
        for row in document.tables[1].rows
    )
    assert ENTITY["address"] in table_text
    assert ENTITY["tax_code"] in table_text
    assert ENTITY["representative"] in table_text
    assert ENTITY["position"] in table_text
    assert "686345848" not in table_text
    assert "Ngân hàng TMCP Quân đội" not in table_text


def test_fill_principle_party_a_leaves_missing_fields_blank():
    document = docx.Document(TEMPLATE)
    entity = dict(ENTITY)
    entity["address"] = ""
    entity["position"] = ""

    fill_principle_party_a(document, entity)

    table_text = "\n".join(
        "\n".join(cell.text for cell in row.cells)
        for row in document.tables[1].rows
    )
    assert "ABC, Hà Nội" not in table_text
    assert "Giám đốc" not in table_text


def test_fill_principle_contract_meta_sets_number_and_date():
    document = docx.Document(TEMPLATE)

    fill_principle_contract_meta(document, "01/082026/HĐNT", "15", "08", "2026")

    full_text = "\n".join(p.text for p in document.paragraphs)
    assert "01/082026/HĐNT" in full_text
    assert "15/08/2026" in full_text
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_principle_docx_filler.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'contract_filler.principle_docx_filler'`

- [ ] **Step 3: Write the implementation**

```python
# contract_filler/principle_docx_filler.py
from contract_filler.docx_filler import replace_in_paragraph, replace_in_row


def fill_principle_party_a(document, entity):
    for paragraph in document.paragraphs:
        if "BÊN A: BÊN MUA" in paragraph.text:
            replace_in_paragraph(
                paragraph,
                "CÔNG TY TNHH TƯ VẤN PHÁT TRIỂN THỊ TRƯỜNG M.S.V",
                entity["name"],
            )

    table_a = document.tables[1]
    replace_in_row(table_a.rows[0], "ABC, Hà Nội", entity["address"])
    replace_in_row(table_a.rows[1], "0123456789", entity["tax_code"])
    replace_in_row(table_a.rows[2], "686345848", entity["bank_account"])
    replace_in_row(table_a.rows[3], "Ngân hàng TMCP Quân đội", "")
    replace_in_row(table_a.rows[4], "NGUYỄN VĂN A", entity["representative"])
    replace_in_row(table_a.rows[5], "Giám đốc", entity["position"])


def fill_principle_contract_meta(document, contract_no, day, month, year):
    for paragraph in document.paragraphs:
        if paragraph.text.strip().startswith("Số "):
            replace_in_paragraph(paragraph, "07/062026/HĐMB", contract_no)
        if "được lập ngày" in paragraph.text:
            replace_in_paragraph(paragraph, "22", day)
            replace_in_paragraph(paragraph, "06", month)
            replace_in_paragraph(paragraph, "2026", year)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_principle_docx_filler.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add contract_filler/principle_docx_filler.py tests/test_principle_docx_filler.py
git commit -m "feat: fill Ben A and contract meta in principle contract template"
```

---

### Task 6: Batch-generate one contract per entity

**Files:**
- Create: `contract_filler/principle_generator.py`
- Test: `tests/test_principle_generator.py`

**Interfaces:**
- Consumes: `parse_entities(docx_path) -> list[dict]` from [principle_list_parser.py](contract_filler/principle_list_parser.py) (Task 3); `fill_principle_party_a(document, entity)` and `fill_principle_contract_meta(document, contract_no, day, month, year)` from [principle_docx_filler.py](contract_filler/principle_docx_filler.py) (Task 5).
- Produces: `generate_principle_contracts(list_path, template_path, output_dir, contract_no, day, month, year) -> list[str]`, returning the absolute output paths in entity order. Task 7's GUI calls this.

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_principle_generator.py
import os
import docx
from contract_filler.principle_generator import generate_principle_contracts

LIST_FIXTURE = "tests/fixtures/household_company_list.docx"
TEMPLATE_FIXTURE = "tests/fixtures/principle_contract_template.docx"


def test_generate_principle_contracts_creates_one_file_per_entity(tmp_path):
    output_dir = str(tmp_path / "out")

    output_paths = generate_principle_contracts(
        list_path=LIST_FIXTURE,
        template_path=TEMPLATE_FIXTURE,
        output_dir=output_dir,
        contract_no="01/082026/HĐNT",
        day="15",
        month="08",
        year="2026",
    )

    assert len(output_paths) == 3
    for path in output_paths:
        assert os.path.exists(path)

    first_doc = docx.Document(output_paths[0])
    full_text = "\n".join(p.text for p in first_doc.paragraphs)
    assert "HỘ KINH DOANH NGUYỄN VĂN TEST" in full_text
    assert "01/082026/HĐNT" in full_text


def test_generate_principle_contracts_leaves_missing_fields_blank(tmp_path):
    output_dir = str(tmp_path / "out")

    output_paths = generate_principle_contracts(
        list_path=LIST_FIXTURE,
        template_path=TEMPLATE_FIXTURE,
        output_dir=output_dir,
        contract_no="01/082026/HĐNT",
        day="15",
        month="08",
        year="2026",
    )

    incomplete_doc = docx.Document(output_paths[2])
    table_text = "\n".join(
        "\n".join(cell.text for cell in row.cells)
        for row in incomplete_doc.tables[1].rows
    )
    assert "0000000003" in table_text
    assert "ABC, Hà Nội" not in table_text
    assert "Lê Văn Thiếu" in table_text


def test_generate_principle_contracts_uses_indexed_filenames(tmp_path):
    output_dir = str(tmp_path / "out")

    output_paths = generate_principle_contracts(
        list_path=LIST_FIXTURE,
        template_path=TEMPLATE_FIXTURE,
        output_dir=output_dir,
        contract_no="01/082026/HĐNT",
        day="15",
        month="08",
        year="2026",
    )

    basenames = [os.path.basename(p) for p in output_paths]
    assert basenames[0].startswith("01.")
    assert basenames[1].startswith("02.")
    assert basenames[2].startswith("03.")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_principle_generator.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'contract_filler.principle_generator'`

- [ ] **Step 3: Write the implementation**

```python
# contract_filler/principle_generator.py
import os
import re

import docx

from contract_filler.principle_list_parser import parse_entities
from contract_filler.principle_docx_filler import (
    fill_principle_party_a,
    fill_principle_contract_meta,
)

_INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')


def _sanitize_filename(name):
    cleaned = _INVALID_FILENAME_CHARS.sub("", name).strip()
    return cleaned or "hop_dong"


def generate_principle_contracts(
    list_path, template_path, output_dir, contract_no, day, month, year
):
    entities = parse_entities(list_path)
    os.makedirs(output_dir, exist_ok=True)

    output_paths = []
    for index, entity in enumerate(entities, start=1):
        document = docx.Document(template_path)
        fill_principle_party_a(document, entity)
        fill_principle_contract_meta(document, contract_no, day, month, year)

        filename = f"{index:02d}. Hop dong nguyen tac - {_sanitize_filename(entity['name'])}.docx"
        output_path = os.path.join(output_dir, filename)
        document.save(output_path)
        output_paths.append(output_path)

    return output_paths
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_principle_generator.py -v`
Expected: PASS

- [ ] **Step 5: Run the full test suite**

Run: `pytest -v`
Expected: all tests PASS, including the pre-existing ones from `test_generator.py`, `test_docx_filler.py`, `test_pdf_parser.py`.

- [ ] **Step 6: Commit**

```bash
git add contract_filler/principle_generator.py tests/test_principle_generator.py
git commit -m "feat: batch-generate one principle contract per entity"
```

---

### Task 7: Add the "Tạo hợp đồng nguyên tắc" tab to the GUI

**Files:**
- Modify: `contract_filler/gui.py` (full rewrite of the module — see below)
- Test: `tests/test_gui_smoke.py`

**Interfaces:**
- Consumes: `generate_contract(pdf_path, template_path, output_path, contract_no, day, month, year)` from [generator.py](contract_filler/generator.py) (existing, unchanged); `generate_principle_contracts(list_path, template_path, output_dir, contract_no, day, month, year) -> list[str]` from [principle_generator.py](contract_filler/principle_generator.py) (Task 6).
- Produces: `App(tk.Tk)` with a `ttk.Notebook` containing exactly two tabs, titled `"Tạo hợp đồng mua bán"` and `"Tạo hợp đồng nguyên tắc"`. `main()` remains the entry point used by [run.py](run.py) — its signature and behavior (constructs `App()` and calls `.mainloop()`) don't change.

- [ ] **Step 1: Write the failing smoke test**

```python
# tests/test_gui_smoke.py
from contract_filler.gui import App


def test_app_has_two_tabs_with_expected_titles():
    app = App()
    try:
        assert app.notebook.index("end") == 2
        assert app.notebook.tab(0, "text") == "Tạo hợp đồng mua bán"
        assert app.notebook.tab(1, "text") == "Tạo hợp đồng nguyên tắc"
    finally:
        app.destroy()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_gui_smoke.py -v`
Expected: FAIL — `App` doesn't yet expose a `.notebook` attribute (current `App` builds the single-contract UI directly on itself).

- [ ] **Step 3: Rewrite the GUI module with two tabs**

```python
# contract_filler/gui.py
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from contract_filler.generator import generate_contract
from contract_filler.principle_generator import generate_principle_contracts

DEFAULT_SALES_TEMPLATE = os.path.join(os.getcwd(), "HĐMB mẫu.docx")
DEFAULT_PRINCIPLE_TEMPLATE = os.path.join(os.getcwd(), "HỢP ĐỒNG ABC mẫu.docx")


def _file_row(parent, label, variable, browse_command):
    tk.Label(parent, text=label).pack(anchor="w", padx=10, pady=(10, 0))
    row = tk.Frame(parent)
    row.pack(anchor="w", fill="x", padx=10)
    tk.Entry(row, textvariable=variable, width=50).pack(side="left")
    tk.Button(row, text="Chọn...", command=browse_command).pack(side="left", padx=(6, 0))


def _date_row(parent, day_var, month_var, year_var):
    date_frame = tk.Frame(parent)
    date_frame.pack(anchor="w", padx=10, pady=(10, 0))
    tk.Label(date_frame, text="Ngày:").pack(side="left")
    tk.Entry(date_frame, textvariable=day_var, width=4).pack(side="left", padx=(4, 10))
    tk.Label(date_frame, text="Tháng:").pack(side="left")
    tk.Entry(date_frame, textvariable=month_var, width=4).pack(side="left", padx=(4, 10))
    tk.Label(date_frame, text="Năm:").pack(side="left")
    tk.Entry(date_frame, textvariable=year_var, width=6).pack(side="left", padx=(4, 0))


class SalesContractTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.pdf_path = tk.StringVar()
        self.template_path = tk.StringVar(value=DEFAULT_SALES_TEMPLATE)
        self.contract_no = tk.StringVar()
        self.day = tk.StringVar()
        self.month = tk.StringVar()
        self.year = tk.StringVar()

        _file_row(self, "File hóa đơn (PDF):", self.pdf_path, self._browse_pdf)
        _file_row(self, "File hợp đồng mẫu (Word):", self.template_path, self._browse_template)

        tk.Label(self, text="Số hợp đồng:").pack(anchor="w", padx=10, pady=(10, 0))
        tk.Entry(self, textvariable=self.contract_no, width=40).pack(anchor="w", padx=10)

        _date_row(self, self.day, self.month, self.year)

        tk.Button(self, text="Tạo hợp đồng", command=self._generate).pack(pady=20)

    def _browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf_path.set(path)

    def _browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if path:
            self.template_path.set(path)

    def _generate(self):
        if not self.pdf_path.get() or not self.template_path.get():
            messagebox.showerror("Thiếu thông tin", "Vui lòng chọn file PDF và file mẫu Word.")
            return
        if not (self.contract_no.get() and self.day.get() and self.month.get() and self.year.get()):
            messagebox.showerror("Thiếu thông tin", "Vui lòng nhập số hợp đồng và ngày tháng năm.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")],
            initialfile="Hop dong mua ban.docx",
        )
        if not output_path:
            return

        try:
            generate_contract(
                pdf_path=self.pdf_path.get(),
                template_path=self.template_path.get(),
                output_path=output_path,
                contract_no=self.contract_no.get(),
                day=self.day.get(),
                month=self.month.get(),
                year=self.year.get(),
            )
        except Exception as exc:  # surfaced to the user, not swallowed
            messagebox.showerror("Lỗi", f"Không tạo được hợp đồng:\n{exc}")
            return

        messagebox.showinfo("Thành công", f"Đã tạo hợp đồng:\n{output_path}")


class PrincipleContractTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.list_path = tk.StringVar()
        self.template_path = tk.StringVar(value=DEFAULT_PRINCIPLE_TEMPLATE)
        self.output_dir = tk.StringVar()
        self.contract_no = tk.StringVar()
        self.day = tk.StringVar()
        self.month = tk.StringVar()
        self.year = tk.StringVar()

        _file_row(self, "File danh sách (Word):", self.list_path, self._browse_list)
        _file_row(self, "File hợp đồng mẫu (Word):", self.template_path, self._browse_template)
        _file_row(self, "Thư mục lưu kết quả:", self.output_dir, self._browse_output_dir)

        tk.Label(self, text="Số hợp đồng:").pack(anchor="w", padx=10, pady=(10, 0))
        tk.Entry(self, textvariable=self.contract_no, width=40).pack(anchor="w", padx=10)

        _date_row(self, self.day, self.month, self.year)

        tk.Button(self, text="Tạo hợp đồng hàng loạt", command=self._generate).pack(pady=20)

    def _browse_list(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if path:
            self.list_path.set(path)

    def _browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if path:
            self.template_path.set(path)

    def _browse_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def _generate(self):
        if not self.list_path.get() or not self.template_path.get() or not self.output_dir.get():
            messagebox.showerror(
                "Thiếu thông tin",
                "Vui lòng chọn file danh sách, file mẫu Word và thư mục lưu kết quả.",
            )
            return
        if not (self.contract_no.get() and self.day.get() and self.month.get() and self.year.get()):
            messagebox.showerror("Thiếu thông tin", "Vui lòng nhập số hợp đồng và ngày tháng năm.")
            return

        try:
            output_paths = generate_principle_contracts(
                list_path=self.list_path.get(),
                template_path=self.template_path.get(),
                output_dir=self.output_dir.get(),
                contract_no=self.contract_no.get(),
                day=self.day.get(),
                month=self.month.get(),
                year=self.year.get(),
            )
        except Exception as exc:  # surfaced to the user, not swallowed
            messagebox.showerror("Lỗi", f"Không tạo được hợp đồng:\n{exc}")
            return

        messagebox.showinfo(
            "Thành công",
            f"Đã tạo {len(output_paths)} hợp đồng trong:\n{self.output_dir.get()}",
        )


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Điền hợp đồng")
        self.geometry("560x420")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.add(SalesContractTab(self.notebook), text="Tạo hợp đồng mua bán")
        self.notebook.add(PrincipleContractTab(self.notebook), text="Tạo hợp đồng nguyên tắc")


def main():
    App().mainloop()
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_gui_smoke.py -v`
Expected: PASS

- [ ] **Step 5: Run the full test suite**

Run: `pytest -v`
Expected: all tests PASS.

- [ ] **Step 6: Manually verify the app in the running GUI**

Run: `python run.py`
Expected: window opens titled "Điền hợp đồng" with two tabs. On the second tab, pick `Danh sách cty hoặc hộ kd cần điền (giả định).docx` as the list file, `HỢP ĐỒNG ABC mẫu.docx` as the template, an empty output folder, fill in a contract number and date, click "Tạo hợp đồng hàng loạt" — confirm a success dialog reports 16 files created and that the output folder contains 16 `.docx` files, each opening in Word with a different Bên A and the same Bên B.

- [ ] **Step 7: Commit**

```bash
git add contract_filler/gui.py tests/test_gui_smoke.py
git commit -m "feat: add Tao hop dong nguyen tac tab to the GUI"
```

---

## Self-Review Notes

- **Spec coverage:** reads 2 input files similar to the named samples (Task 1 template asset + Task 3 parser for the list) ✓; one output `.docx` per hộ kinh doanh/công ty (Task 6) ✓; each entity filled into Bên A (Task 5) ✓; missing fields left blank (Tasks 3, 5, 6 all test this explicitly) ✓; keeps the existing "Tạo hợp đồng mua bán" feature and name intact (Task 7 keeps it as the first tab, unchanged behavior) ✓.
- **Placeholder scan:** no TODOs; every step has runnable code and exact expected output.
- **Type consistency:** entity dict keys (`name`, `tax_code`, `address`, `bank_account`, `representative`, `position`) are identical across Task 3's parser, Task 5's filler, and Task 6's/7's tests. `replace_in_row` is renamed once in Task 4 and used with that exact name from Task 5 onward.
