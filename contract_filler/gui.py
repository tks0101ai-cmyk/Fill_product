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
