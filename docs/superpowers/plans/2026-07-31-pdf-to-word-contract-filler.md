# PDF → Word Contract Filler Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Windows desktop tool that reads a VAT-invoice PDF (`mẫu HĐ.pdf` format) and auto-fills the matching sales-contract Word template (`HĐMB mẫu.docx`), producing a ready-to-use `.docx` contract.

**Architecture:** A pure-Python pipeline: `pdf_parser.py` extracts seller/buyer info, line items, and totals from the PDF using `pdfplumber` (tables) and regex over extracted text (labeled fields). `docx_filler.py` writes those values into the existing Word template using `python-docx`, including a generic "replace text that spans multiple runs" helper and table-row cloning for a variable number of product lines. `gui.py` is a minimal `tkinter` window (file pickers + a few text fields + a "Tạo hợp đồng" button) wrapping the pipeline. No web server, no database — everything runs locally from a double-clickable script.

**Tech Stack:** Python 3.10+, `pdfplumber` (PDF text/table extraction), `python-docx` (Word manipulation), `tkinter` (GUI, stdlib — no extra install), `pytest` (tests).

## Global Constraints

- Input PDF format is fixed: the VAT invoice layout produced by vin-hoadon.com, as seen in `mẫu HĐ.pdf` (labeled fields like `Đơn vị bán hàng (Seller):`, table columns STT/Tên hàng hóa/Đơn vị tính/Số lượng/Đơn giá/Thành tiền, totals block at the end). The parser does not need to handle other invoice layouts.
- Output Word template is fixed: `HĐMB mẫu.docx`, exact structure as it exists today. The filler targets this template's specific tables/paragraphs by their current placeholder text, not a generic templating engine.
- Fields NOT present in the PDF (bank account, ngân hàng, đại diện, chức vụ, SĐT for both Bên A and Bên B) must be left **blank** in the output — never leave the template's stale sample values (e.g. "Nguyễn Thị B", "898896886").
- Contract number (`Số hợp đồng`) and contract date (ngày/tháng/năm) are **always typed by the user** in the GUI — never inferred from the PDF.
- Money amounts and the "số tiền viết bằng chữ" (amount in words) line are copied verbatim from the PDF, not recomputed — the PDF is the source of truth for totals.
- Vietnamese number formatting uses `.` as the thousands separator and no decimals (e.g. `2.695.000`) — this must be preserved in the output.

## File Structure

```
contract_filler/
  __init__.py
  pdf_parser.py     # parse_parties(), parse_line_items(), parse_totals()
  docx_filler.py     # replace_in_paragraph(), fill_party_info(), fill_contract_meta(), fill_product_table_and_totals()
  generator.py        # generate_contract() — orchestrates parser + filler
  gui.py               # tkinter App
run.py                 # entry point: launches the GUI
requirements.txt
tests/
  fixtures/
    sample_invoice.pdf     # copy of "mẫu HĐ.pdf"
    sample_template.docx    # copy of "HĐMB mẫu.docx"
  test_pdf_parser.py
  test_docx_filler.py
  test_generator.py
```

---

### Task 1: Project scaffolding and dependencies

**Files:**
- Create: `requirements.txt`
- Create: `contract_filler/__init__.py`
- Create: `tests/fixtures/sample_invoice.pdf`
- Create: `tests/fixtures/sample_template.docx`
- Test: `tests/test_pdf_parser.py` (first test added in Task 2, but fixtures are copied here)

**Interfaces:**
- Produces: a working Python environment with `pdfplumber`, `python-docx`, `pytest` installed, and fixture files every later test relies on.

- [ ] **Step 1: Create `requirements.txt`**

```
pdfplumber==0.11.4
python-docx==1.1.2
pytest==8.3.3
```

- [ ] **Step 2: Create the package init file**

`contract_filler/__init__.py`:
```python
```
(empty file — just marks `contract_filler` as a package)

- [ ] **Step 3: Install dependencies**

Run:
```bash
pip install -r requirements.txt
```
Expected: all three packages install without errors.

- [ ] **Step 4: Copy the sample PDF and Word template into test fixtures**

Run:
```bash
mkdir -p tests/fixtures
cp "mẫu HĐ.pdf" tests/fixtures/sample_invoice.pdf
cp "HĐMB mẫu.docx" tests/fixtures/sample_template.docx
```
Expected: both files exist under `tests/fixtures/`.

- [ ] **Step 5: Commit**

```bash
git init
git add requirements.txt contract_filler/__init__.py tests/fixtures/sample_invoice.pdf tests/fixtures/sample_template.docx
git commit -m "chore: scaffold project and add fixture files"
```
(If the user doesn't want git version control here, skip this step and just keep the files on disk.)

---

### Task 2: Parse seller/buyer info from the PDF

**Files:**
- Create: `contract_filler/pdf_parser.py`
- Test: `tests/test_pdf_parser.py`

**Interfaces:**
- Produces: `parse_parties(text: str) -> tuple[dict, dict]` returning `(seller, buyer)`, each a dict with keys `name`, `tax_code`, `address` (all `str`). Later tasks (docx_filler) consume these exact keys.

- [ ] **Step 1: Write the failing test**

`tests/test_pdf_parser.py`:
```python
import pdfplumber
from contract_filler.pdf_parser import parse_parties

FIXTURE = "tests/fixtures/sample_invoice.pdf"


def _full_text():
    with pdfplumber.open(FIXTURE) as pdf:
        return "\n".join(page.extract_text() for page in pdf.pages)


def test_parse_parties_extracts_seller_and_buyer():
    seller, buyer = parse_parties(_full_text())

    assert seller["name"] == "CÔNG TY CỔ PHẦN THƯƠNG MẠI XUẤT NHẬP KHẨU TKS GROUP"
    assert seller["tax_code"] == "0110534607"
    assert seller["address"] == (
        "Số 9 đường Lê Văn Huấn, Cụm công nghiệp Cầu Nổi, "
        "Xã Sơn Đồng, Thành phố Hà Nội, Việt Nam"
    )

    assert buyer["name"] == "HỘ KINH DOANH TIỆM 81"
    assert buyer["tax_code"] == "064200012728"
    assert buyer["address"] == (
        "201/65/9 Nguyễn Xí, Phường Bình Thạnh, Thành phố Hồ Chí Minh, Việt Nam"
    )
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_pdf_parser.py::test_parse_parties_extracts_seller_and_buyer -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'contract_filler.pdf_parser'`

- [ ] **Step 3: Write the implementation**

`contract_filler/pdf_parser.py`:
```python
import re

BUYER_MARKER = "Tên đơn vị (Company's name):"


def _find(pattern, text):
    match = re.search(pattern, text)
    return match.group(1).strip() if match else ""


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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_pdf_parser.py::test_parse_parties_extracts_seller_and_buyer -v`
Expected: PASS

(If it fails because `pdfplumber`'s line breaks don't match the regex, print `_full_text()` and adjust the regexes to the actual extracted layout — `pdfplumber` sometimes merges the label and value onto slightly different whitespace than the raw PDF text stream.)

- [ ] **Step 5: Commit**

```bash
git add contract_filler/pdf_parser.py tests/test_pdf_parser.py
git commit -m "feat: parse seller/buyer info from invoice PDF"
```

---

### Task 3: Parse product line items from the PDF

**Files:**
- Modify: `contract_filler/pdf_parser.py`
- Modify: `tests/test_pdf_parser.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: `parse_line_items(pdf_path: str) -> list[dict]`, each item dict has keys `stt` (int), `description` (str), `unit` (str), `quantity` (int), `unit_price` (int), `amount` (int). Consumed by `docx_filler.fill_product_table_and_totals` in Task 6.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_pdf_parser.py`:
```python
from contract_filler.pdf_parser import parse_line_items


def test_parse_line_items_reads_all_rows_across_pages():
    items = parse_line_items(FIXTURE)

    assert len(items) == 4
    assert items[0]["stt"] == 1
    assert items[0]["unit"] == "Cái"
    assert items[0]["quantity"] == 7000
    assert items[0]["unit_price"] == 385
    assert items[0]["amount"] == 2695000
    assert "Móc treo đồ gắn tường" in items[0]["description"]

    # item 4 lives on page 2 of the PDF
    assert items[3]["stt"] == 4
    assert items[3]["quantity"] == 100
    assert items[3]["unit_price"] == 12800
    assert items[3]["amount"] == 1280000
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_pdf_parser.py::test_parse_line_items_reads_all_rows_across_pages -v`
Expected: FAIL with `ImportError: cannot import name 'parse_line_items'`

- [ ] **Step 3: Write the implementation**

Add to `contract_filler/pdf_parser.py`:
```python
import pdfplumber


def _parse_number(value):
    if not value:
        return 0
    return int(value.replace(".", "").replace(",", "").strip())


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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_pdf_parser.py -v`
Expected: PASS for both tests in the file.

(If `page.extract_table()` returns `None` or a different column count on some pages, print `table` for each page and adjust — e.g. use `page.extract_tables()[0]` if multiple tables are detected, or pass explicit `table_settings` to `extract_table`.)

- [ ] **Step 5: Commit**

```bash
git add contract_filler/pdf_parser.py tests/test_pdf_parser.py
git commit -m "feat: parse product line items across PDF pages"
```

---

### Task 4: Parse totals from the PDF

**Files:**
- Modify: `contract_filler/pdf_parser.py`
- Modify: `tests/test_pdf_parser.py`

**Interfaces:**
- Produces: `parse_totals(text: str) -> dict` with keys `subtotal` (int), `vat_rate` (str, e.g. `"8"`), `vat_amount` (int), `total_payment` (int), `amount_in_words` (str). Consumed by `docx_filler.fill_product_table_and_totals` in Task 6.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_pdf_parser.py`:
```python
from contract_filler.pdf_parser import parse_totals


def test_parse_totals_reads_summary_block():
    totals = parse_totals(_full_text())

    assert totals["subtotal"] == 9315000
    assert totals["vat_rate"] == "8"
    assert totals["vat_amount"] == 745200
    assert totals["total_payment"] == 10060200
    assert totals["amount_in_words"] == (
        "Mười triệu không trăm sáu mươi nghìn hai trăm đồng"
    )
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_pdf_parser.py::test_parse_totals_reads_summary_block -v`
Expected: FAIL with `ImportError: cannot import name 'parse_totals'`

- [ ] **Step 3: Write the implementation**

Add to `contract_filler/pdf_parser.py`:
```python
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_pdf_parser.py -v`
Expected: PASS for all tests in the file.

- [ ] **Step 5: Commit**

```bash
git add contract_filler/pdf_parser.py tests/test_pdf_parser.py
git commit -m "feat: parse invoice totals and amount-in-words"
```

---

### Task 5: Generic cross-run text replacement helper for docx

**Files:**
- Create: `contract_filler/docx_filler.py`
- Test: `tests/test_docx_filler.py`

**Interfaces:**
- Produces: `replace_in_paragraph(paragraph, old: str, new: str) -> bool`. Returns `True` if a replacement was made, `False` if `old` wasn't found. Used by every later fill function in this task group.

- [ ] **Step 1: Write the failing test**

`tests/test_docx_filler.py`:
```python
import docx
from contract_filler.docx_filler import replace_in_paragraph


def _paragraph_with_runs(texts):
    document = docx.Document()
    paragraph = document.add_paragraph()
    for text in texts:
        paragraph.add_run(text)
    return paragraph


def test_replace_within_single_run():
    paragraph = _paragraph_with_runs(["Địa chỉ: ", "ABCC", " "])
    assert replace_in_paragraph(paragraph, "ABCC", "Số 9 đường X")
    assert paragraph.text == "Địa chỉ: Số 9 đường X "


def test_replace_spanning_multiple_runs():
    paragraph = _paragraph_with_runs(["Mã số thuế: ", "0433859", "844"])
    assert replace_in_paragraph(paragraph, "0433859844", "0110534607")
    assert paragraph.text == "Mã số thuế: 0110534607"


def test_replace_returns_false_when_not_found():
    paragraph = _paragraph_with_runs(["Hello"])
    assert replace_in_paragraph(paragraph, "Goodbye", "Hi") is False
    assert paragraph.text == "Hello"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_docx_filler.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'contract_filler.docx_filler'`

- [ ] **Step 3: Write the implementation**

`contract_filler/docx_filler.py`:
```python
def replace_in_paragraph(paragraph, old, new):
    runs = paragraph.runs
    full_text = "".join(run.text for run in runs)
    start = full_text.find(old)
    if start == -1:
        return False
    end = start + len(old)

    pos = 0
    replaced = False
    for run in runs:
        run_start, run_end = pos, pos + len(run.text)
        pos = run_end

        if run_end <= start or run_start >= end:
            continue

        prefix = run.text[: max(0, start - run_start)]
        suffix = run.text[max(0, end - run_start):]

        if not replaced:
            run.text = prefix + new + suffix
            replaced = True
        else:
            run.text = prefix + suffix

    return True
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_docx_filler.py -v`
Expected: PASS for all three tests.

- [ ] **Step 5: Commit**

```bash
git add contract_filler/docx_filler.py tests/test_docx_filler.py
git commit -m "feat: add cross-run text replacement helper for docx paragraphs"
```

---

### Task 6: Fill party info (Bên A / Bên B) in the Word template

**Files:**
- Modify: `contract_filler/docx_filler.py`
- Modify: `tests/test_docx_filler.py`

**Interfaces:**
- Consumes: `seller`/`buyer` dicts from Task 2 (`name`, `tax_code`, `address`).
- Produces: `fill_party_info(document, seller: dict, buyer: dict) -> None`. Mutates `document` in place.

The template has, in this fixed order: a paragraph `"BÊN A: BÊN BÁN: <company name>"`, then `document.tables[0]` with 6 rows (`Địa chỉ`, `Mã số thuế`, `Tài khoản số`, `Mở tại ngân hàng`, `Đại diện`, `Chức vụ` — each row is one paragraph `"Label: value"`), then a paragraph `"BÊN B: BÊN MUA: <company name>"`, then `document.tables[1]` with 7 rows (`Địa chỉ`, `Mã số thuế`, `SĐT`, `Số tài khoản`, `Ngân hàng`, `Đại diện`, `Chức vụ`).

- [ ] **Step 1: Write the failing test**

Append to `tests/test_docx_filler.py`:
```python
from contract_filler.docx_filler import fill_party_info

SELLER = {
    "name": "CÔNG TY CỔ PHẦN THƯƠNG MẠI XUẤT NHẬP KHẨU TKS GROUP",
    "tax_code": "0110534607",
    "address": "Số 9 đường Lê Văn Huấn, Cụm công nghiệp Cầu Nổi, Xã Sơn Đồng, Thành phố Hà Nội, Việt Nam",
}
BUYER = {
    "name": "HỘ KINH DOANH TIỆM 81",
    "tax_code": "064200012728",
    "address": "201/65/9 Nguyễn Xí, Phường Bình Thạnh, Thành phố Hồ Chí Minh, Việt Nam",
}


def test_fill_party_info_sets_names_and_clears_unknown_fields():
    document = docx.Document("tests/fixtures/sample_template.docx")

    fill_party_info(document, SELLER, BUYER)

    full_text = "\n".join(p.text for p in document.paragraphs)
    assert SELLER["name"] in full_text
    assert BUYER["name"] in full_text
    assert "ABCC" not in full_text  # old placeholder gone

    table_a_text = "\n".join(row.cells[0].text for row in document.tables[0].rows)
    assert SELLER["address"] in table_a_text
    assert SELLER["tax_code"] in table_a_text
    assert "Nguyễn Thị B" not in table_a_text  # rep cleared
    assert "898896886" not in table_a_text  # bank account cleared

    table_b_text = "\n".join(row.cells[0].text for row in document.tables[1].rows)
    assert BUYER["address"] in table_b_text
    assert BUYER["tax_code"] in table_b_text
    assert "Huỳnh Tấn Hải" not in table_b_text  # rep cleared
    assert "113003051756" not in table_b_text  # account cleared
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_docx_filler.py::test_fill_party_info_sets_names_and_clears_unknown_fields -v`
Expected: FAIL with `ImportError: cannot import name 'fill_party_info'`

- [ ] **Step 3: Write the implementation**

Add to `contract_filler/docx_filler.py`:
```python
def _row_paragraph(row):
    seen_cells = set()
    for cell in row.cells:
        if id(cell._tc) in seen_cells:
            continue
        seen_cells.add(id(cell._tc))
        for paragraph in cell.paragraphs:
            if paragraph.text.strip():
                return paragraph
    return None


def fill_party_info(document, seller, buyer):
    for paragraph in document.paragraphs:
        if "BÊN A: BÊN BÁN:" in paragraph.text:
            replace_in_paragraph(
                paragraph,
                "CÔNG TY CỔ PHẦN XUẤT NHẬP KHẨU THƯƠNG MẠI ABC",
                seller["name"],
            )
        if "BÊN B: BÊN MUA" in paragraph.text:
            replace_in_paragraph(
                paragraph, "CÔNG TY TNHH TÓC XINH MỸ LỆ", buyer["name"]
            )

    table_a = document.tables[0]
    replace_in_paragraph(_row_paragraph(table_a.rows[0]), "ABCC", seller["address"])
    replace_in_paragraph(_row_paragraph(table_a.rows[1]), "0433859844", seller["tax_code"])
    replace_in_paragraph(_row_paragraph(table_a.rows[2]), "898896886", "")
    replace_in_paragraph(
        _row_paragraph(table_a.rows[3]), "Ngân hàng TMCP QUÂN ĐỘI - MBANK", ""
    )
    replace_in_paragraph(_row_paragraph(table_a.rows[4]), "Nguyễn Thị B", "")
    replace_in_paragraph(_row_paragraph(table_a.rows[5]), "Giám đốc", "")

    table_b = document.tables[1]
    replace_in_paragraph(
        _row_paragraph(table_b.rows[0]),
        "05 đường M2, Dự án Khu dân cư và Công viên Phước Thiện, Khu phố 28, Phường Long Bình, TP Hồ Chí Minh",
        buyer["address"],
    )
    replace_in_paragraph(_row_paragraph(table_b.rows[1]), "0319437919", buyer["tax_code"])
    replace_in_paragraph(_row_paragraph(table_b.rows[2]), "0", "")
    replace_in_paragraph(_row_paragraph(table_b.rows[3]), "113003051756", "")
    replace_in_paragraph(
        _row_paragraph(table_b.rows[4]),
        "Ngân hàng Thương mại Cổ phần Công Thương Việt Nam – Đồng Nai",
        "",
    )
    replace_in_paragraph(_row_paragraph(table_b.rows[5]), "Huỳnh Tấn Hải", "")
    replace_in_paragraph(_row_paragraph(table_b.rows[6]), "Giám đốc", "")
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_docx_filler.py -v`
Expected: PASS for all tests in the file.

(If any `replace_in_paragraph` call returns `False` unexpectedly, print `table_a.rows[i].cells[0].text` / `table_b.rows[i].cells[0].text` to confirm the exact current placeholder text in your copy of `sample_template.docx` and adjust the `old` string to match exactly — trailing spaces or an en-dash vs hyphen will break an exact match.)

- [ ] **Step 5: Commit**

```bash
git add contract_filler/docx_filler.py tests/test_docx_filler.py
git commit -m "feat: fill Bên A / Bên B party info in Word template"
```

---

### Task 7: Fill contract number and date

**Files:**
- Modify: `contract_filler/docx_filler.py`
- Modify: `tests/test_docx_filler.py`

**Interfaces:**
- Produces: `fill_contract_meta(document, contract_no: str, day: str, month: str, year: str) -> None`.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_docx_filler.py`:
```python
from contract_filler.docx_filler import fill_contract_meta


def test_fill_contract_meta_sets_number_and_date():
    document = docx.Document("tests/fixtures/sample_template.docx")

    fill_contract_meta(document, "0099/2026/PPR/HĐMBHH", "31", "07", "2026")

    full_text = "\n".join(p.text for p in document.paragraphs)
    assert "0099/2026/PPR/HĐMBHH" in full_text
    assert "ngày 31 tháng 07 năm 2026" in full_text
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_docx_filler.py::test_fill_contract_meta_sets_number_and_date -v`
Expected: FAIL with `ImportError: cannot import name 'fill_contract_meta'`

- [ ] **Step 3: Write the implementation**

Add to `contract_filler/docx_filler.py`:
```python
def fill_contract_meta(document, contract_no, day, month, year):
    for paragraph in document.paragraphs:
        if "Số :" in paragraph.text or "Số:" in paragraph.text:
            if replace_in_paragraph(paragraph, "02062026/PPR/HĐMBHH", contract_no):
                continue
        if "Hôm nay, ngày" in paragraph.text:
            replace_in_paragraph(paragraph, "13", day)
            replace_in_paragraph(paragraph, "06", month)
            replace_in_paragraph(paragraph, "2026", year)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_docx_filler.py -v`
Expected: PASS for all tests in the file.

(The `"Số :"` paragraph's placeholder is the concatenation of several runs — `"0206" + "2026" + "/" + "PPR/" + "HĐMBHH"`. If `replace_in_paragraph` returns `False` there, print the paragraph's `[r.text for r in paragraph.runs]` to see the exact run split in your copy of the template and fix the `old` string to match the exact concatenation.)

- [ ] **Step 5: Commit**

```bash
git add contract_filler/docx_filler.py tests/test_docx_filler.py
git commit -m "feat: fill contract number and signing date"
```

---

### Task 8: Fill product table rows and totals

**Files:**
- Modify: `contract_filler/docx_filler.py`
- Modify: `tests/test_docx_filler.py`

**Interfaces:**
- Consumes: `items` list from Task 3, `totals` dict from Task 4.
- Produces: `fill_product_table_and_totals(document, items: list[dict], totals: dict) -> None`.

The product table is `document.tables[2]`: row 0 is the header, row 1 is a sample data row (STT=1) that acts as the row template, rows 2–4 are `TIỀN HÀNG` / `TIỀN THUẾ VAT 8%` / `TỔNG THANH TOÁN`. After a paragraph following this table holds `"...viết bằng chữ: Sáu mươi triệu không trăm sáu mươi chín nghìn sáu trăm đồng./.)"`.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_docx_filler.py`:
```python
from contract_filler.docx_filler import fill_product_table_and_totals

ITEMS = [
    {"stt": 1, "description": "Móc treo đồ gắn tường", "unit": "Cái",
     "quantity": 7000, "unit_price": 385, "amount": 2695000},
    {"stt": 2, "description": "Khay đựng bàn chải", "unit": "Cái",
     "quantity": 360, "unit_price": 2500, "amount": 900000},
    {"stt": 3, "description": "Bộ bàn chải đánh răng", "unit": "Bộ",
     "quantity": 300, "unit_price": 14800, "amount": 4440000},
    {"stt": 4, "description": "Dụng cụ tách sò", "unit": "Cái",
     "quantity": 100, "unit_price": 12800, "amount": 1280000},
]
TOTALS = {
    "subtotal": 9315000,
    "vat_rate": "8",
    "vat_amount": 745200,
    "total_payment": 10060200,
    "amount_in_words": "Mười triệu không trăm sáu mươi nghìn hai trăm đồng",
}


def test_fill_product_table_and_totals_writes_all_rows():
    document = docx.Document("tests/fixtures/sample_template.docx")

    fill_product_table_and_totals(document, ITEMS, TOTALS)

    table = document.tables[2]
    # header + 4 item rows + 3 totals rows
    assert len(table.rows) == 8

    for row, item in zip(table.rows[1:5], ITEMS):
        cells_text = [c.text for c in row.cells]
        assert str(item["stt"]) in cells_text[0]
        assert item["description"] in cells_text[1]
        assert item["unit"] in cells_text[2]

    assert "9.315.000" in table.rows[5].cells[-1].text
    assert "8" in table.rows[6].cells[0].text
    assert "745.200" in table.rows[6].cells[-1].text
    assert "10.060.200" in table.rows[7].cells[-1].text

    full_text = "\n".join(p.text for p in document.paragraphs)
    assert TOTALS["amount_in_words"] in full_text
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_docx_filler.py::test_fill_product_table_and_totals_writes_all_rows -v`
Expected: FAIL with `ImportError: cannot import name 'fill_product_table_and_totals'`

- [ ] **Step 3: Write the implementation**

Add to `contract_filler/docx_filler.py`:
```python
from copy import deepcopy
from docx.table import _Row


def _format_vnd(amount):
    return f"{amount:,}".replace(",", ".")


def _set_cell_text(cell, text):
    paragraph = cell.paragraphs[0]
    for extra in cell.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)
    runs = paragraph.runs
    if runs:
        runs[0].text = text
        for run in runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def _set_item_row(row, item):
    cells = row.cells
    _set_cell_text(cells[0], str(item["stt"]))
    _set_cell_text(cells[1], item["description"])
    _set_cell_text(cells[2], item["unit"])
    _set_cell_text(cells[3], _format_vnd(item["quantity"]))
    _set_cell_text(cells[4], _format_vnd(item["unit_price"]))
    _set_cell_text(cells[5], _format_vnd(item["amount"]))


def fill_product_table_and_totals(document, items, totals):
    table = document.tables[2]
    rows = table.rows
    template_row = rows[1]
    subtotal_row, vat_row, total_row = rows[2], rows[3], rows[4]
    anchor_tr = subtotal_row._tr

    _set_item_row(template_row, items[0])
    for item in items[1:]:
        new_tr = deepcopy(template_row._tr)
        anchor_tr.addprevious(new_tr)
        _set_item_row(_Row(new_tr, table), item)

    _set_cell_text(subtotal_row.cells[-1], _format_vnd(totals["subtotal"]))
    vat_paragraph = _row_paragraph(vat_row)
    replace_in_paragraph(vat_paragraph, "8", totals["vat_rate"])
    _set_cell_text(vat_row.cells[-1], _format_vnd(totals["vat_amount"]))
    _set_cell_text(total_row.cells[-1], _format_vnd(totals["total_payment"]))

    for paragraph in document.paragraphs:
        if "viết bằng chữ" in paragraph.text:
            replace_in_paragraph(
                paragraph,
                "Sáu mươi triệu không trăm sáu mươi chín nghìn sáu trăm đồng",
                totals["amount_in_words"],
            )
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_docx_filler.py -v`
Expected: PASS for all tests in the file.

(If row insertion produces rows in the wrong order, print `[[c.text for c in r.cells] for r in table.rows]` right after the loop to see the actual order — `tr.addprevious()` inserts immediately before the anchor each time, so items must be appended in original order, which the loop already does.)

- [ ] **Step 5: Commit**

```bash
git add contract_filler/docx_filler.py tests/test_docx_filler.py
git commit -m "feat: fill product table rows and totals in Word template"
```

---

### Task 9: End-to-end orchestration

**Files:**
- Create: `contract_filler/generator.py`
- Test: `tests/test_generator.py`

**Interfaces:**
- Consumes: `parse_parties`, `parse_line_items`, `parse_totals` (Tasks 2–4); `fill_party_info`, `fill_contract_meta`, `fill_product_table_and_totals` (Tasks 6–8).
- Produces: `generate_contract(pdf_path: str, template_path: str, output_path: str, contract_no: str, day: str, month: str, year: str) -> None`. This is what `gui.py` (Task 10) calls directly.

- [ ] **Step 1: Write the failing test**

`tests/test_generator.py`:
```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_generator.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'contract_filler.generator'`

- [ ] **Step 3: Write the implementation**

`contract_filler/generator.py`:
```python
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_generator.py -v`
Expected: PASS

- [ ] **Step 5: Run the full test suite**

Run: `pytest -v`
Expected: all tests across `tests/test_pdf_parser.py`, `tests/test_docx_filler.py`, and `tests/test_generator.py` PASS.

- [ ] **Step 6: Commit**

```bash
git add contract_filler/generator.py tests/test_generator.py
git commit -m "feat: add end-to-end PDF-to-contract orchestration"
```

---

### Task 10: Tkinter GUI

**Files:**
- Create: `contract_filler/gui.py`
- Create: `run.py`

**Interfaces:**
- Consumes: `generate_contract` from Task 9.
- Produces: a runnable desktop window; no automated test (GUI event loops aren't unit-tested here — verified manually in Step 3).

- [ ] **Step 1: Write the GUI**

`contract_filler/gui.py`:
```python
import os
import tkinter as tk
from tkinter import filedialog, messagebox

from contract_filler.generator import generate_contract

DEFAULT_TEMPLATE = os.path.join(os.getcwd(), "HĐMB mẫu.docx")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Điền hợp đồng mua bán từ hóa đơn PDF")
        self.geometry("560x360")

        self.pdf_path = tk.StringVar()
        self.template_path = tk.StringVar(value=DEFAULT_TEMPLATE)
        self.contract_no = tk.StringVar()
        self.day = tk.StringVar()
        self.month = tk.StringVar()
        self.year = tk.StringVar()

        self._build_layout()

    def _build_layout(self):
        self._file_row("File hóa đơn (PDF):", self.pdf_path, self._browse_pdf)
        self._file_row("File hợp đồng mẫu (Word):", self.template_path, self._browse_template)

        tk.Label(self, text="Số hợp đồng:").pack(anchor="w", padx=10, pady=(10, 0))
        tk.Entry(self, textvariable=self.contract_no, width=40).pack(anchor="w", padx=10)

        date_frame = tk.Frame(self)
        date_frame.pack(anchor="w", padx=10, pady=(10, 0))
        tk.Label(date_frame, text="Ngày:").pack(side="left")
        tk.Entry(date_frame, textvariable=self.day, width=4).pack(side="left", padx=(4, 10))
        tk.Label(date_frame, text="Tháng:").pack(side="left")
        tk.Entry(date_frame, textvariable=self.month, width=4).pack(side="left", padx=(4, 10))
        tk.Label(date_frame, text="Năm:").pack(side="left")
        tk.Entry(date_frame, textvariable=self.year, width=6).pack(side="left", padx=(4, 0))

        tk.Button(self, text="Tạo hợp đồng", command=self._generate).pack(pady=20)

    def _file_row(self, label, variable, browse_command):
        tk.Label(self, text=label).pack(anchor="w", padx=10, pady=(10, 0))
        row = tk.Frame(self)
        row.pack(anchor="w", fill="x", padx=10)
        tk.Entry(row, textvariable=variable, width=50).pack(side="left")
        tk.Button(row, text="Chọn...", command=browse_command).pack(side="left", padx=(6, 0))

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


def main():
    App().mainloop()
```

- [ ] **Step 2: Write the entry point**

`run.py` (repo root):
```python
from contract_filler.gui import main

if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Manual test**

Run:
```bash
python run.py
```
1. Click "Chọn..." next to "File hóa đơn (PDF)" and pick `mẫu HĐ.pdf`.
2. Confirm "File hợp đồng mẫu (Word)" already points at `HĐMB mẫu.docx` (or browse to it).
3. Enter a contract number (e.g. `0099/2026/PPR/HĐMBHH`) and a date (31 / 07 / 2026).
4. Click "Tạo hợp đồng", choose a save location, confirm the success dialog appears.
5. Open the generated `.docx` in Word and verify: seller/buyer names, addresses, tax codes are filled; bank/rep/SĐT fields are blank; all 4 product rows appear with correct quantities/prices; totals and "số tiền viết bằng chữ" match the PDF; contract number and date match what you typed.

Expected: the window opens, the flow completes without errors, and the output `.docx` matches all of the above by eye.

- [ ] **Step 4: Commit**

```bash
git add contract_filler/gui.py run.py
git commit -m "feat: add tkinter GUI for contract generation"
```

---

## Self-Review Notes

- **Spec coverage:** PDF input → Word output ✓ (Tasks 2–4 parse, Tasks 6–9 fill/orchestrate). Full-contract fill (party info, not just the product table) ✓ (Task 6). Missing fields left blank rather than stale ✓ (Task 6, explicit clear calls). Contract number/date always manual ✓ (Task 7 + GUI fields, never derived from PDF). Simple button-based file-picking UI ✓ (Task 10).
- **Known fragility to flag to the executor:** every `old` string in Tasks 6–8 is copied from this specific `HĐMB mẫu.docx`'s current run structure. If the user edits the template's wording/spacing before running this tool, those exact-match replacements will silently return `False` (no crash, but the field stays unfilled) — Step 4 of each task includes a debugging note for exactly this failure mode.
