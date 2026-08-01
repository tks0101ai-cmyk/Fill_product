import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from contract_filler.generator import generate_sales_contracts
from contract_filler.principle_generator import generate_principle_contracts
from contract_filler.seller_list import parse_sellers

DEFAULT_SALES_TEMPLATE = os.path.join(os.getcwd(), "HỢP ĐỒNG MUA BÁN mẫu.docx")
DEFAULT_PRINCIPLE_TEMPLATE = os.path.join(os.getcwd(), "HỢP ĐỒNG NGUYÊN TẮC mẫu.docx")
DEFAULT_SELLER_LIST = os.path.join(os.getcwd(), "List bên bán.docx")


def _load_sellers():
    try:
        return parse_sellers(DEFAULT_SELLER_LIST)
    except Exception:
        return []


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


def _seller_row(parent, sellers, seller_var):
    tk.Label(parent, text="Bên bán:").pack(anchor="w", padx=10, pady=(10, 0))
    names = [seller["name"] for seller in sellers]
    combo = ttk.Combobox(parent, textvariable=seller_var, values=names, width=60, state="readonly")
    combo.pack(anchor="w", padx=10)
    if names:
        combo.current(0)
    return combo


class SalesContractTab(ttk.Frame):
    def __init__(self, parent, sellers):
        super().__init__(parent)
        self.sellers = sellers

        self.pdf_path = tk.StringVar()
        self.template_path = tk.StringVar(value=DEFAULT_SALES_TEMPLATE)
        self.list_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.seller_name = tk.StringVar()
        self.contract_no = tk.StringVar()
        self.day = tk.StringVar()
        self.month = tk.StringVar()
        self.year = tk.StringVar()

        _file_row(self, "File hóa đơn (PDF):", self.pdf_path, self._browse_pdf)
        _file_row(self, "File hợp đồng mẫu (Word):", self.template_path, self._browse_template)
        _file_row(self, "File danh sách hộ KD/công ty (Word):", self.list_path, self._browse_list)
        _seller_row(self, sellers, self.seller_name)
        _file_row(self, "Thư mục lưu kết quả:", self.output_dir, self._browse_output_dir)

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

    def _browse_list(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if path:
            self.list_path.set(path)

    def _browse_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def _selected_seller(self):
        for seller in self.sellers:
            if seller["name"] == self.seller_name.get():
                return seller
        return None

    def _generate(self):
        if not (self.pdf_path.get() and self.template_path.get() and self.list_path.get() and self.output_dir.get()):
            messagebox.showerror(
                "Thiếu thông tin",
                "Vui lòng chọn file PDF, file mẫu Word, file danh sách và thư mục lưu kết quả.",
            )
            return
        if not (self.contract_no.get() and self.day.get() and self.month.get() and self.year.get()):
            messagebox.showerror("Thiếu thông tin", "Vui lòng nhập số hợp đồng và ngày tháng năm.")
            return

        seller = self._selected_seller()
        if seller is None:
            messagebox.showerror("Thiếu thông tin", "Vui lòng chọn bên bán.")
            return

        try:
            output_paths = generate_sales_contracts(
                pdf_path=self.pdf_path.get(),
                template_path=self.template_path.get(),
                list_path=self.list_path.get(),
                seller=seller,
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


class PrincipleContractTab(ttk.Frame):
    def __init__(self, parent, sellers):
        super().__init__(parent)
        self.sellers = sellers

        self.template_path = tk.StringVar(value=DEFAULT_PRINCIPLE_TEMPLATE)
        self.list_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.seller_name = tk.StringVar()
        self.contract_no = tk.StringVar()
        self.day = tk.StringVar()
        self.month = tk.StringVar()
        self.year = tk.StringVar()

        _file_row(self, "File hợp đồng mẫu (Word):", self.template_path, self._browse_template)
        _file_row(self, "File danh sách hộ KD/công ty (Word):", self.list_path, self._browse_list)
        _seller_row(self, sellers, self.seller_name)
        _file_row(self, "Thư mục lưu kết quả:", self.output_dir, self._browse_output_dir)

        tk.Label(self, text="Số hợp đồng:").pack(anchor="w", padx=10, pady=(10, 0))
        tk.Entry(self, textvariable=self.contract_no, width=40).pack(anchor="w", padx=10)

        _date_row(self, self.day, self.month, self.year)

        tk.Button(self, text="Tạo hợp đồng hàng loạt", command=self._generate).pack(pady=20)

    def _browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if path:
            self.template_path.set(path)

    def _browse_list(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if path:
            self.list_path.set(path)

    def _browse_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def _selected_seller(self):
        for seller in self.sellers:
            if seller["name"] == self.seller_name.get():
                return seller
        return None

    def _generate(self):
        if not (self.template_path.get() and self.list_path.get() and self.output_dir.get()):
            messagebox.showerror(
                "Thiếu thông tin",
                "Vui lòng chọn file mẫu Word, file danh sách và thư mục lưu kết quả.",
            )
            return
        if not (self.contract_no.get() and self.day.get() and self.month.get() and self.year.get()):
            messagebox.showerror("Thiếu thông tin", "Vui lòng nhập số hợp đồng và ngày tháng năm.")
            return

        seller = self._selected_seller()
        if seller is None:
            messagebox.showerror("Thiếu thông tin", "Vui lòng chọn bên bán.")
            return

        try:
            output_paths = generate_principle_contracts(
                list_path=self.list_path.get(),
                template_path=self.template_path.get(),
                output_dir=self.output_dir.get(),
                seller=seller,
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
        self.geometry("620x480")

        sellers = _load_sellers()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.add(SalesContractTab(self.notebook, sellers), text="Tạo hợp đồng mua bán")
        self.notebook.add(PrincipleContractTab(self.notebook, sellers), text="Tạo hợp đồng nguyên tắc")


def main():
    App().mainloop()
