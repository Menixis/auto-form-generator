# -*- coding: utf-8 -*-
"""
ระบบสร้างเอกสารอัตโนมัติ (Auto Form Generator) - ปตท.
=====================================================
สร้างหนังสือราชการอัตโนมัติจากแบบฟอร์มต้นฉบับ 5 ประเภท
+ ดึงข้อมูล PO อัตโนมัติจากไฟล์ Excel (รวมถึงไฟล์บน OneDrive/SharePoint)

วิธีใช้งาน:
  1. ติดตั้ง:  pip install python-docx openpyxl
  2. ตั้งค่า  file_path  ใน config.ini ให้ชี้ไปที่ไฟล์ Excel
  3. วางแบบฟอร์มต้นฉบับใน templates/form1.docx ... form5.docx
  4. รัน:  python auto_form_generator.py
"""

import os
import sys
import datetime
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Import pure logic จาก form_logic.py (ใช้ร่วมกับ streamlit_app.py)
try:
    from form_logic import (
        FORM_TYPES, DOC_TYPE_OPTIONS, DISPLAY_FIELDS,
        generate_form,
    )
except ImportError as e:
    print(f"กรุณาตรวจสอบว่ามีไฟล์ form_logic.py และติดตั้ง python-docx แล้ว: {e}")
    sys.exit(1)

# excel_lookup เป็น optional — ถ้า import ไม่ได้ ปุ่มดึงข้อมูลจะ disabled
try:
    import excel_lookup
    HAS_LOOKUP = True
except ImportError:
    HAS_LOOKUP = False


# =============================================================================
# GUI Application
# =============================================================================
class AutoFormApp:
    BG = "#F5F7FA"
    PRIMARY = "#1F4E79"
    ACCENT = "#2E75B6"
    SUCCESS = "#2E7D32"
    TEXT = "#222222"
    MUTED = "#666666"

    def __init__(self, root):
        self.root = root
        self.root.title("ระบบสร้างเอกสารอัตโนมัติ - ปตท. (เชื่อมต่อ Excel)")
        self.root.geometry("780x800")
        self.root.configure(bg=self.BG)
        self.root.minsize(720, 740)

        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.template_dir = tk.StringVar(value=os.path.join(script_dir, "templates"))
        self.output_dir = tk.StringVar(value=os.path.join(script_dir, "output"))

        self.form_choice = tk.StringVar(value=FORM_TYPES[0]["label"])
        self.doc_number = tk.StringVar()
        self.po_number = tk.StringVar()
        self.doc_type = tk.StringVar(value=DOC_TYPE_OPTIONS[0])
        self.signer = tk.StringVar()

        # State สำหรับข้อมูลที่ดึงจาก Excel
        self.lookup_data = None
        self.lookup_status_var = tk.StringVar(value="ยังไม่ได้ดึงข้อมูล")
        self.po_number.trace_add("write", lambda *a: self._reset_lookup_state())

        self._setup_styles()
        self._build_ui()
        self._on_form_change()

    # -------------------- Styles --------------------
    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TLabel", background=self.BG, foreground=self.TEXT,
                        font=("Tahoma", 11))
        style.configure("Header.TLabel", background=self.PRIMARY, foreground="white",
                        font=("Tahoma", 16, "bold"), padding=12)
        style.configure("SubHeader.TLabel", background=self.BG, foreground=self.MUTED,
                        font=("Tahoma", 10))
        style.configure("Field.TLabel", background=self.BG, foreground=self.TEXT,
                        font=("Tahoma", 11, "bold"))
        style.configure("TEntry", padding=6, fieldbackground="white")
        style.configure("TCombobox", padding=4, fieldbackground="white")
        style.configure("Primary.TButton", background=self.ACCENT, foreground="white",
                        font=("Tahoma", 11, "bold"), padding=(20, 10), borderwidth=0)
        style.map("Primary.TButton",
                  background=[("active", self.PRIMARY), ("pressed", self.PRIMARY)])
        style.configure("Secondary.TButton", background="#E0E0E0",
                        foreground=self.TEXT, font=("Tahoma", 10),
                        padding=(14, 8), borderwidth=0)
        style.map("Secondary.TButton", background=[("active", "#CCCCCC")])
        style.configure("Fetch.TButton", background=self.SUCCESS, foreground="white",
                        font=("Tahoma", 10, "bold"), padding=(10, 6), borderwidth=0)
        style.map("Fetch.TButton", background=[("active", "#1B5E20")])
        style.configure("TRadiobutton", background=self.BG, foreground=self.TEXT,
                        font=("Tahoma", 11))

    # -------------------- UI Layout --------------------
    def _build_ui(self):
        header = ttk.Label(self.root,
                           text="ระบบสร้างเอกสารอัตโนมัติ - ปตท.",
                           style="Header.TLabel", anchor="center")
        header.pack(fill="x")
        ttk.Label(self.root,
                  text="กรอกเลข PO แล้วกด \"ดึงข้อมูล\" "
                       "เพื่อเติมข้อมูลอัตโนมัติจาก Excel",
                  style="SubHeader.TLabel", anchor="center").pack(pady=(8, 12))

        main = tk.Frame(self.root, bg=self.BG)
        main.pack(fill="both", expand=True, padx=24, pady=(0, 12))
        main.columnconfigure(1, weight=1)
        row = 0

        # ----- ประเภทแบบฟอร์ม -----
        ttk.Label(main, text="ประเภทแบบฟอร์ม:", style="Field.TLabel")\
            .grid(row=row, column=0, sticky="w", pady=(6, 4))
        form_combo = ttk.Combobox(
            main, textvariable=self.form_choice,
            values=[f["label"] for f in FORM_TYPES],
            state="readonly", width=60, font=("Tahoma", 10))
        form_combo.grid(row=row, column=1, sticky="ew", pady=(6, 4))
        form_combo.bind("<<ComboboxSelected>>", lambda e: self._on_form_change())
        row += 1

        ttk.Separator(main, orient="horizontal")\
            .grid(row=row, column=0, columnspan=2, sticky="ew", pady=10)
        row += 1

        # ----- เลข PO + ปุ่มดึงข้อมูล -----
        ttk.Label(main, text="เลข PO: *", style="Field.TLabel")\
            .grid(row=row, column=0, sticky="w", pady=6)
        po_frame = tk.Frame(main, bg=self.BG)
        po_frame.grid(row=row, column=1, sticky="ew", pady=6)
        po_frame.columnconfigure(0, weight=1)
        self.po_entry = ttk.Entry(po_frame, textvariable=self.po_number,
                                  font=("Tahoma", 11))
        self.po_entry.grid(row=0, column=0, sticky="ew")
        self.fetch_btn = ttk.Button(po_frame, text="ดึงข้อมูลจาก Excel",
                                    style="Fetch.TButton",
                                    command=self._fetch_data)
        self.fetch_btn.grid(row=0, column=1, padx=(8, 0))
        if not HAS_LOOKUP:
            self.fetch_btn.configure(state="disabled")
        row += 1

        ttk.Label(main, textvariable=self.lookup_status_var,
                  style="SubHeader.TLabel")\
            .grid(row=row, column=1, sticky="w")
        row += 1

        # ----- หมายเลขเอกสาร -----
        ttk.Label(main, text="หมายเลขเอกสาร: *", style="Field.TLabel")\
            .grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(main, textvariable=self.doc_number, font=("Tahoma", 11))\
            .grid(row=row, column=1, sticky="ew", pady=6)
        row += 1
        ttk.Label(main, text="ตัวอย่าง: PTT-001/2568",
                  style="SubHeader.TLabel")\
            .grid(row=row, column=1, sticky="w")
        row += 1

        # ----- ประเภทเอกสาร -----
        self.type_label = ttk.Label(main, text="ประเภทเอกสาร:",
                                    style="Field.TLabel")
        self.type_label.grid(row=row, column=0, sticky="nw", pady=6)
        self.type_frame = tk.Frame(main, bg=self.BG)
        self.type_frame.grid(row=row, column=1, sticky="w", pady=6)
        for i, opt in enumerate(DOC_TYPE_OPTIONS):
            ttk.Radiobutton(self.type_frame, text=opt,
                            variable=self.doc_type, value=opt)\
                .grid(row=0, column=i, padx=(0, 20), sticky="w")
        row += 1

        # ----- ผู้ลงนาม -----
        ttk.Label(main, text="ผู้ลงนาม: *", style="Field.TLabel")\
            .grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(main, textvariable=self.signer, font=("Tahoma", 11))\
            .grid(row=row, column=1, sticky="ew", pady=6)
        row += 1
        ttk.Label(main,
                  text="(ผู้มีอำนาจอนุมัติ หรือ ประธานกรรมการตรวจรับ)",
                  style="SubHeader.TLabel")\
            .grid(row=row, column=1, sticky="w")
        row += 1

        ttk.Separator(main, orient="horizontal")\
            .grid(row=row, column=0, columnspan=2, sticky="ew", pady=10)
        row += 1

        # ----- ข้อมูลที่ดึงจาก Excel (preview) -----
        ttk.Label(main, text="ข้อมูลจาก Excel:", style="Field.TLabel")\
            .grid(row=row, column=0, sticky="nw", pady=6)
        preview_frame = tk.Frame(main, bg="white",
                                 highlightbackground="#CCCCCC",
                                 highlightthickness=1)
        preview_frame.grid(row=row, column=1, sticky="ew", pady=6)
        self.preview_text = tk.Text(
            preview_frame, height=8, font=("Tahoma", 10),
            bg="white", fg=self.TEXT, relief="flat",
            wrap="word", padx=10, pady=8, state="disabled")
        self.preview_text.pack(fill="both", expand=True)
        self._update_preview(None)
        row += 1

        # ----- โฟลเดอร์ -----
        ttk.Label(main, text="โฟลเดอร์ Template:", style="Field.TLabel")\
            .grid(row=row, column=0, sticky="w", pady=4)
        tpl_frame = tk.Frame(main, bg=self.BG)
        tpl_frame.grid(row=row, column=1, sticky="ew", pady=4)
        tpl_frame.columnconfigure(0, weight=1)
        ttk.Entry(tpl_frame, textvariable=self.template_dir,
                  font=("Tahoma", 10)).grid(row=0, column=0, sticky="ew")
        ttk.Button(tpl_frame, text="เลือก...", style="Secondary.TButton",
                   command=lambda: self._browse_dir(self.template_dir))\
            .grid(row=0, column=1, padx=(8, 0))
        row += 1

        ttk.Label(main, text="โฟลเดอร์บันทึก:", style="Field.TLabel")\
            .grid(row=row, column=0, sticky="w", pady=4)
        out_frame = tk.Frame(main, bg=self.BG)
        out_frame.grid(row=row, column=1, sticky="ew", pady=4)
        out_frame.columnconfigure(0, weight=1)
        ttk.Entry(out_frame, textvariable=self.output_dir,
                  font=("Tahoma", 10)).grid(row=0, column=0, sticky="ew")
        ttk.Button(out_frame, text="เลือก...", style="Secondary.TButton",
                   command=lambda: self._browse_dir(self.output_dir))\
            .grid(row=0, column=1, padx=(8, 0))
        row += 1

        # ----- ปุ่มดำเนินการ -----
        btn_frame = tk.Frame(self.root, bg=self.BG)
        btn_frame.pack(fill="x", padx=24, pady=(4, 12))
        ttk.Button(btn_frame, text="ล้างข้อมูล", style="Secondary.TButton",
                   command=self._clear_form).pack(side="left")
        ttk.Button(btn_frame, text="สร้างเอกสาร", style="Primary.TButton",
                   command=self._generate).pack(side="right")

        # ----- Status bar -----
        self.status_var = tk.StringVar(value="พร้อมใช้งาน")
        if not HAS_LOOKUP:
            self.status_var.set(
                "พร้อมใช้งาน (Excel: ไม่พบ excel_lookup.py หรือ openpyxl)")
        status_bar = tk.Label(self.root, textvariable=self.status_var,
                              bg="#E8EEF4", fg=self.PRIMARY, anchor="w",
                              font=("Tahoma", 10), padx=12, pady=8)
        status_bar.pack(fill="x", side="bottom")

    # -------------------- Helpers --------------------
    def _get_selected_form(self):
        for f in FORM_TYPES:
            if f["label"] == self.form_choice.get():
                return f
        return FORM_TYPES[0]

    def _on_form_change(self):
        form = self._get_selected_form()
        for child in self.type_frame.winfo_children():
            child.configure(state="normal" if form["uses_doc_type"] else "disabled")
        self.type_label.configure(
            foreground=self.TEXT if form["uses_doc_type"] else "#999999")

    def _browse_dir(self, var):
        initial = var.get() if os.path.isdir(var.get()) else os.getcwd()
        chosen = filedialog.askdirectory(initialdir=initial,
                                         title="เลือกโฟลเดอร์")
        if chosen:
            var.set(chosen)

    def _clear_form(self):
        self.doc_number.set("")
        self.po_number.set("")
        self.signer.set("")
        self.doc_type.set(DOC_TYPE_OPTIONS[0])
        self._reset_lookup_state()
        self.status_var.set("ล้างข้อมูลเรียบร้อย")

    def _reset_lookup_state(self):
        self.lookup_data = None
        self.lookup_status_var.set("ยังไม่ได้ดึงข้อมูล")
        self._update_preview(None)

    def _update_preview(self, data):
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")
        if not data:
            self.preview_text.insert("1.0",
                "(กด \"ดึงข้อมูลจาก Excel\" เพื่อค้นหาข้อมูล PO)")
        else:
            lines = []
            for col, label in DISPLAY_FIELDS:
                val = data.get(col)
                if val not in (None, ""):
                    lines.append(f"• {label}: {val}")
            extras = [k for k in data.keys()
                      if k not in {c for c, _ in DISPLAY_FIELDS}
                      and data[k] not in (None, "")]
            if extras:
                lines.append(f"\n(พบเพิ่มเติมอีก {len(extras)} columns)")
            self.preview_text.insert("1.0",
                "\n".join(lines) if lines else "(พบ row แต่ค่าทั้งหมดว่าง)")
        self.preview_text.configure(state="disabled")

    # -------------------- Data fetch --------------------
    def _fetch_data(self):
        po = self.po_number.get().strip()
        if not po:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกเลข PO ก่อน")
            return
        if not HAS_LOOKUP:
            messagebox.showerror(
                "Excel lookup ไม่พร้อมใช้งาน",
                "ไม่พบโมดูล excel_lookup\n"
                "กรุณาตรวจสอบว่ามีไฟล์ excel_lookup.py "
                "และติดตั้ง:  pip install openpyxl")
            return

        self.fetch_btn.configure(state="disabled")
        self.lookup_status_var.set("⏳ กำลังอ่านไฟล์ Excel...")
        self.status_var.set(f"กำลังค้นหา PO: {po}")
        self.root.update_idletasks()

        threading.Thread(target=self._fetch_worker, args=(po,),
                         daemon=True).start()

    def _fetch_worker(self, po):
        try:
            data = excel_lookup.lookup_po(po)
            self.root.after(0, self._fetch_done, po, data, None)
        except Exception as exc:
            self.root.after(0, self._fetch_done, po, None, str(exc))

    def _fetch_done(self, po, data, error):
        self.fetch_btn.configure(state="normal")
        if error:
            self.lookup_status_var.set("❌ ดึงข้อมูลไม่สำเร็จ")
            self.status_var.set("Excel: ผิดพลาด")
            messagebox.showerror("เกิดข้อผิดพลาด",
                                 f"ไม่สามารถดึงข้อมูลได้:\n\n{error}")
            return
        if not data:
            self.lookup_data = None
            self.lookup_status_var.set(f"⚠ ไม่พบ PO: {po}")
            self.status_var.set(f"ไม่พบ PO {po} ในไฟล์ Excel")
            self._update_preview(None)
            messagebox.showinfo("ไม่พบข้อมูล",
                                f"ไม่พบ PO หมายเลข {po} ในไฟล์ Excel")
            return
        self.lookup_data = data
        vendor = data.get("VendorName", "")
        self.lookup_status_var.set(
            f"✓ พบข้อมูล: {vendor or '(no vendor)'} — "
            f"จะเติม Vendor / ชื่องาน อัตโนมัติตอนสร้างเอกสาร")
        self.status_var.set(f"ดึงข้อมูล PO {po} สำเร็จ")
        self._update_preview(data)

    # -------------------- Generate --------------------
    def _validate(self, form):
        if not self.doc_number.get().strip():
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกหมายเลขเอกสาร")
            return False
        if not self.signer.get().strip():
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกชื่อผู้ลงนาม")
            return False
        if form["uses_po"] and not self.po_number.get().strip():
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกเลข PO")
            return False
        if not os.path.isdir(self.template_dir.get()):
            messagebox.showerror("ไม่พบโฟลเดอร์ Template",
                                 f"ไม่พบโฟลเดอร์: {self.template_dir.get()}")
            return False
        template_path = os.path.join(self.template_dir.get(), form["filename"])
        if not os.path.isfile(template_path):
            messagebox.showerror("ไม่พบไฟล์ Template",
                f"ไม่พบ {form['filename']} ในโฟลเดอร์:\n{self.template_dir.get()}")
            return False
        return True

    def _generate(self):
        form = self._get_selected_form()
        if not self._validate(form):
            return

        os.makedirs(self.output_dir.get(), exist_ok=True)
        template_path = os.path.join(self.template_dir.get(), form["filename"])

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_doc_num = (self.doc_number.get()
                        .replace("/", "-").replace("\\", "-").strip())
        out_filename = (f"form{form['id']}_{safe_doc_num}_{timestamp}.docx"
                        if safe_doc_num else f"form{form['id']}_{timestamp}.docx")
        output_path = os.path.join(self.output_dir.get(), out_filename)

        try:
            log = generate_form(
                form_id=form["id"],
                template_path=template_path,
                output_path=output_path,
                doc_number=self.doc_number.get().strip(),
                po_number=self.po_number.get().strip(),
                doc_type=self.doc_type.get() if form["uses_doc_type"] else "",
                signer=self.signer.get().strip(),
                lookup_data=self.lookup_data,
            )
        except Exception as exc:
            messagebox.showerror("เกิดข้อผิดพลาด",
                                 f"ไม่สามารถสร้างเอกสารได้:\n{exc}")
            self.status_var.set(f"เกิดข้อผิดพลาด: {exc}")
            return

        self.status_var.set(f"สร้างเอกสารเรียบร้อย: {out_filename}")
        if messagebox.askyesno("สร้างเอกสารสำเร็จ",
            f"บันทึกไฟล์ที่:\n{output_path}\n\n"
            f"รายละเอียด:\n• " + "\n• ".join(log) + "\n\n"
            f"ต้องการเปิดโฟลเดอร์ผลลัพธ์หรือไม่?"):
            self._open_folder(self.output_dir.get())

    def _open_folder(self, path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception:
            pass


# =============================================================================
# Entry point
# =============================================================================
def main():
    root = tk.Tk()
    AutoFormApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
