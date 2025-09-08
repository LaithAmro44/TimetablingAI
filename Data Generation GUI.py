import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import Separator
from ttkbootstrap.tooltip import ToolTip
import pandas as pd

# ========= إعدادات خطوط =========
BASE_FONT_FAMILY = "Cairo"
BASE_FONT_SIZE = 12
HEADING_FONT = (BASE_FONT_FAMILY, BASE_FONT_SIZE + 4, "bold")
SECTION_FONT = (BASE_FONT_FAMILY, BASE_FONT_SIZE + 1, "bold")
BOLD_BTN_FONT = (BASE_FONT_FAMILY, BASE_FONT_SIZE, "bold")
TREE_FONT = (BASE_FONT_FAMILY, BASE_FONT_SIZE)  # للـ Treeview
ROW_HEIGHT_TV = 30  # ارتفاع صفوف الجداول لمنع قصّ الحروف العربية

def ar(text: str) -> str:
    return "\u202B" + text + "\u202C"

# ========= مخازن البيانات =========
materials = []      # [{المادة, الفصل, السنة, عدد الشعب, عدد الطلاب, تخصص عام(list)}]
professors = []     # [{الدكتور, المواد(list of names), عدد الساعات, التخصصات العامة(list)}]
rooms = []          # [{القاعة, النوع, المساحة}]

# ========= أدوات مساعدة عامة =========
def join_specs(specs_list):
    # انضمام بفاصلة عربية "،"
    return "،".join([s for s in specs_list if s.strip()])

def split_specs(s):
    if s is None:
        return []
    s = str(s).strip()
    if not s:
        return []
    # دعم الفاصلة العربية والإنجليزية
    return [x.strip() for x in s.replace("،", ",").split(",") if x.strip()]

def show_info(msg):
    messagebox.showinfo(ar("معلومة"), ar(msg))

def show_err(msg):
    messagebox.showerror(ar("خطأ"), ar(msg))

# ----- خرائط عرض/تخزين للفصل -----
SEM_DISPLAY_TO_INTERNAL = {
    "الفصل الأول": "1",
    "الفصل الثاني": "2",
    "الفصل الأول + الفصل الثاني": "1,2",
}
SEM_INTERNAL_TO_DISPLAY = {
    "1": "الفصل الأول",
    "2": "الفصل الثاني",
    "1,2": "الفصل الأول + الفصل الثاني",
    "1،2": "الفصل الأول + الفصل الثاني",  # دعم الفاصلة العربية إن وجدت
}
def sem_to_internal(display_value: str) -> str:
    return SEM_DISPLAY_TO_INTERNAL.get(display_value.strip(), display_value.strip())

def sem_to_display(internal_value: str) -> str:
    return SEM_INTERNAL_TO_DISPLAY.get((internal_value or "").strip(), internal_value)

# ========= عنصر مركّب: مدير قوائم تخصصات للمواد (إضافة/تعديل/حذف) =========
class SpecListEditor(tb.Frame):
    def __init__(self, master, label_text=ar("التخصصات:"), **kwargs):
        super().__init__(master, **kwargs)
        self.specs = []
        self._build_ui(label_text)

    def _build_ui(self, label_text):
        tb.Label(self, text=label_text, anchor="e", justify="right").grid(row=0, column=1, sticky="e", padx=4)
        self.entry = tb.Entry(self)
        self.entry.grid(row=0, column=0, sticky="ew", padx=4)
        self.grid_columnconfigure(0, weight=1)

        btns = tb.Frame(self)
        btns.grid(row=1, column=0, columnspan=3, sticky="e", pady=2)

        self.btn_add = tb.Button(btns, text=ar("حفظ/إضافة"), bootstyle="secondary", command=self.add_spec)
        self.btn_add.pack(side="right", padx=3)
        ToolTip(self.btn_add, ar("أضف العنصر الحالي إلى القائمة"))

        self.btn_edit = tb.Button(btns, text=ar("تعديل المحدد"), bootstyle="warning", command=self.edit_selected)
        self.btn_edit.pack(side="right", padx=3)
        ToolTip(self.btn_edit, ar("يعدّل العنصر المحدد إلى نص الإدخال الحالي"))

        self.btn_delete = tb.Button(btns, text=ar("حذف المحدد"), bootstyle="danger", command=self.delete_selected)
        self.btn_delete.pack(side="right", padx=3)
        ToolTip(self.btn_delete, ar("يحذف العنصر المحدد من القائمة"))

        # حاوية القائمة + سكروول
        list_wrap = tb.Frame(self)
        list_wrap.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=(2, 4))
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.listbox = tk.Listbox(list_wrap, exportselection=False, height=6, justify="right",
                                activestyle="dotbox", font=(BASE_FONT_FAMILY, BASE_FONT_SIZE))
        self.listbox.grid(row=0, column=0, sticky="nsew")
        # Scrollbar عمودي واضح
        vbar = tb.Scrollbar(list_wrap, orient="vertical", command=self.listbox.yview, bootstyle="primary")
        vbar.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=vbar.set)
        list_wrap.grid_rowconfigure(0, weight=1)
        list_wrap.grid_columnconfigure(0, weight=1)

        # عرض النص المنضم تلقائياً
        tb.Label(self, text=ar("القيمة المنضمّة (تلقائي):"), anchor="e", justify="right").grid(row=3, column=1, sticky="e", padx=4)
        self.joined_var = tk.StringVar()
        self.joined_entry = tb.Entry(self, textvariable=self.joined_var, state="readonly")
        self.joined_entry.grid(row=3, column=0, sticky="ew", padx=4)

    def add_spec(self):
        txt = self.entry.get().strip()
        if not txt:
            return
        self.specs.append(txt)
        self.entry.delete(0, tk.END)
        self._refresh()

    def edit_selected(self):
        idx = self.listbox.curselection()
        if not idx:
            show_err("اختر عنصرًا لتعديله من القائمة.")
            return
        txt = self.entry.get().strip()
        if not txt:
            show_err("أدخل النص الجديد في حقل الإدخال قبل الضغط على تعديل.")
            return
        i = idx[0]
        self.specs[i] = txt
        self._refresh()

    def delete_selected(self):
        idx = self.listbox.curselection()
        if not idx:
            show_err("اختر عنصرًا لحذفه من القائمة.")
            return
        i = idx[0]
        del self.specs[i]
        self._refresh()

    def _refresh(self):
        self.listbox.delete(0, tk.END)
        for s in self.specs:
            self.listbox.insert(tk.END, s)
        self.joined_var.set(join_specs(self.specs))

    def set_specs(self, items):
        self.specs = list(items or [])
        self._refresh()

    def get_specs(self):
        return list(self.specs)

# ========= تبويب المواد =========
class MaterialsTab(tb.Frame):
    def __init__(self, master, on_materials_change, **kwargs):
        super().__init__(master, **kwargs)
        self.on_materials_change = on_materials_change
        self.edit_index = None
        self._build_ui()

    def _build_ui(self):
        # نموذج الإدخال
        card = tb.Labelframe(self, text=ar("إدخال مادة"), padding=10, bootstyle="info")
        card.pack(fill="both", expand=False, pady=6)

        frm = tb.Frame(card)
        frm.pack(fill="x")

        # المادة
        tb.Label(frm, text=ar("اسم المادة:"), anchor="e", justify="right", font=SECTION_FONT)\
            .grid(row=0, column=1, sticky="e", padx=4, pady=2)
        self.name = tb.Entry(frm)
        self.name.grid(row=0, column=0, sticky="ew", padx=4, pady=2)

        # الفصل (ثلاث قيم فقط) - ارتفاع كبير ليظهر Scrollbar الافتراضي عند كثرة العناصر
        tb.Label(frm, text=ar("الفصل:"), anchor="e", justify="right", font=SECTION_FONT)\
            .grid(row=1, column=1, sticky="e", padx=4, pady=2)
        self.sem = tb.Combobox(frm, values=["الفصل الأول", "الفصل الثاني", "الفصل الأول + الفصل الثاني"],
                            state="readonly", height=10)
        self.sem.grid(row=1, column=0, sticky="ew", padx=4, pady=2)

        # السنة
        tb.Label(frm, text=ar("السنة:"), anchor="e", justify="right", font=SECTION_FONT)\
            .grid(row=2, column=1, sticky="e", padx=4, pady=2)
        self.year = tb.Combobox(frm, values=["1", "2", "3", "4", "5"], state="readonly", height=10)
        self.year.grid(row=2, column=0, sticky="ew", padx=4, pady=2)

        # عدد الشعب
        tb.Label(frm, text=ar("عدد الشعب (للتوزيع حسب الشعب):"), anchor="e", justify="right", font=SECTION_FONT)\
            .grid(row=3, column=1, sticky="e", padx=4, pady=2)
        self.num_sections = tb.Entry(frm)
        self.num_sections.grid(row=3, column=0, sticky="ew", padx=4, pady=2)

        # عدد الطلاب
        tb.Label(frm, text=ar("عدد الطلاب (للتوزيع حسب الطلاب):"), anchor="e", justify="right", font=SECTION_FONT)\
            .grid(row=4, column=1, sticky="e", padx=4, pady=2)
        self.num_students = tb.Entry(frm)
        self.num_students.grid(row=4, column=0, sticky="ew", padx=4, pady=2)

        frm.grid_columnconfigure(0, weight=1)

        # محرّر التخصص العام (إضافة/تعديل/حذف)
        self.specs_editor = SpecListEditor(card, label_text=ar("أدخل التخصص العام عنصرًا-عنصرًا:"))
        self.specs_editor.pack(fill="both", expand=True, pady=(4, 2))

        # أزرار
        btns = tb.Frame(card)
        btns.pack(fill="x", pady=(4, 0))
        self.btn_new = tb.Button(btns, text=ar("جديد"), bootstyle="secondary", command=self.clear_form)
        self.btn_new.pack(side="right", padx=4)
        self.btn_add = tb.Button(btns, text=ar("إضافة/تحديث مادة"), bootstyle="success", command=self.add_or_update)
        self.btn_add.pack(side="right", padx=4)
        self.btn_save = tb.Button(btns, text=ar("حفظ إلى Excel"), bootstyle="primary", command=self.save_to_excel)
        self.btn_save.pack(side="right", padx=4)
        # استيراد
        self.btn_import = tb.Button(btns, text=ar("استيراد Excel"), bootstyle="info", command=self.import_from_excel)
        self.btn_import.pack(side="right", padx=4)

        # جدول عرض + Scrollbars
        table_card = tb.Labelframe(self, text=ar("المواد المُدخلة"), padding=6, bootstyle="secondary")
        table_card.pack(fill="both", expand=True, pady=4)

        tv_wrap = tb.Frame(table_card)
        tv_wrap.pack(fill="both", expand=True)

        cols = ("المادة", "الفصل", "السنة", "عدد الشعب", "عدد الطلاب", "تخصص عام")
        self.tree = tb.Treeview(tv_wrap, columns=cols, show="headings", height=8, style="App.Treeview")
        for c in cols:
            self.tree.heading(c, text=ar(c))
            self.tree.column(c, width=140 if c != "تخصص عام" else 220, anchor="e", stretch=True)
        self.tree.grid(row=0, column=0, sticky="nsew")

        # Scrollbars واضحة
        ybar = tb.Scrollbar(tv_wrap, orient="vertical", command=self.tree.yview, bootstyle="primary")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar = tb.Scrollbar(tv_wrap, orient="horizontal", command=self.tree.xview, bootstyle="primary")
        xbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)

        tv_wrap.grid_rowconfigure(0, weight=1)
        tv_wrap.grid_columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self.on_select_row)

    def clear_form(self):
        self.edit_index = None
        self.name.delete(0, tk.END)
        self.sem.set("")
        self.year.set("")
        self.num_sections.delete(0, tk.END)
        self.num_students.delete(0, tk.END)
        self.specs_editor.set_specs([])

    def on_select_row(self, _):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        idx = int(self.tree.item(iid, "text"))
        self.edit_index = idx
        row = materials[idx]
        self.name.delete(0, tk.END); self.name.insert(0, row["المادة"])
        self.sem.set(sem_to_display(row["الفصل"]))
        self.year.set(str(row.get("السنة", "")))
        self.num_sections.delete(0, tk.END); self.num_sections.insert(0, str(row.get("عدد الشعب", "")))
        self.num_students.delete(0, tk.END); self.num_students.insert(0, str(row.get("عدد الطلاب", "")))
        self.specs_editor.set_specs(row.get("تخصص عام", []))

    def add_or_update(self):
        name = self.name.get().strip()
        sem  = self.sem.get().strip()
        year = self.year.get().strip()
        nsec = self.num_sections.get().strip()
        nstu = self.num_students.get().strip()
        specs = self.specs_editor.get_specs()

        if not name:
            return show_err("أدخل اسم المادة.")
        if sem not in ("الفصل الأول", "الفصل الثاني", "الفصل الأول + الفصل الثاني"):
            return show_err("اختر الفصل من القيم المتاحة (الفصل الأول، الفصل الثاني، الفصل الأول + الفصل الثاني).")
        if year not in ("1", "2", "3", "4", "5"):
            return show_err("اختر السنة من 1 إلى 5.")

        sem_internal = sem_to_internal(sem)

        row = {
            "المادة": name,
            "الفصل": sem_internal,   # نخزن داخليًا: 1 / 2 / 1,2
            "السنة": int(year),
            "عدد الشعب": int(nsec) if nsec.isdigit() else "",
            "عدد الطلاب": int(nstu) if nstu.isdigit() else "",
            "تخصص عام": list(specs)
        }

        if self.edit_index is None:
            materials.append(row)
        else:
            materials[self.edit_index] = row

        self.refresh_table()
        self.on_materials_change()  # لتحديث تبويب الدكاترة (المواد + التخصصات العامة)
        self.clear_form()

    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for idx, r in enumerate(materials):
            joined_specs = join_specs(r.get("تخصص عام", []))
            vals = (
                r["المادة"],
                sem_to_display(r["الفصل"]),   # عرض عربي للمستخدم
                r["السنة"],
                r.get("عدد الشعب", ""),
                r.get("عدد الطلاب", ""),
                joined_specs
            )
            self.tree.insert("", "end", text=str(idx), values=vals)

    def save_to_excel(self):
        if not materials:
            return show_err("لا توجد مواد لحفظها.")
        # نحفظ القيم الداخلية كما هي (1 / 2 / 1,2)
        df = pd.DataFrame([
            {
                "المادة": r["المادة"],
                "الفصل": r["الفصل"],
                "السنة": r["السنة"],
                "عدد الطلاب": r.get("عدد الطلاب", ""),
                "عدد الشعب": r.get("عدد الشعب", ""),
                "تخصص عام": join_specs(r.get("تخصص عام", [])),
            } for r in materials
        ])
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="terms.xlsx",
            title=ar("حفظ ملف المواد")
        )
        if not path:
            return
        df.to_excel(path, index=False)
        show_info(f"تم الحفظ في:\n{path}")

    def import_from_excel(self):
        """استيراد مواد من ملف بنفس صيغة الحفظ. فقط معالجة بسيطة للفاصلة العربية في (الفصل) و (تخصص عام)."""
        global materials
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title=ar("اختر ملف المواد (terms.xlsx)")
        )
        if not path:
            return
        try:
            df = pd.read_excel(path)
        except Exception as e:
            return show_err(f"تعذر قراءة الملف:\n{e}")

        new_rows = []
        for _, r in df.iterrows():
            name = str(r.get("المادة", "")).strip()
            if not name:
                continue
            sem_val = str(r.get("الفصل", "")).strip().replace("،", ",")  # 1 / 2 / 1,2
            year_val = r.get("السنة", "")
            try:
                year_val = int(str(year_val).strip())
            except:
                continue

            num_students = r.get("عدد الطلاب", "")
            num_sections = r.get("عدد الشعب", "")

            specs_list = split_specs(r.get("تخصص عام", ""))

            new_rows.append({
                "المادة": name,
                "الفصل": sem_val,
                "السنة": year_val,
                "عدد الطلاب": int(num_students) if str(num_students).strip().isdigit() else "",
                "عدد الشعب": int(num_sections) if str(num_sections).strip().isdigit() else "",
                "تخصص عام": specs_list
            })

        materials = new_rows
        self.refresh_table()
        self.on_materials_change()
        show_info("تم استيراد المواد بنجاح.")

# ========= مُلتقط تخصصات عامة من المواد (ليبني قائمة الدكاترة) =========
def all_general_specs_from_materials():
    specs = []
    for r in materials:
        specs.extend(r.get("تخصص عام", []))
    # إزالة التكرار مع الحفاظ على ترتيب الظهور
    seen = set()
    uniq = []
    for s in specs:
        if s not in seen:
            seen.add(s)
            uniq.append(s)
    return uniq

# ========= تبويب الدكاترة =========
class ProfessorsTab(tb.Frame):
    def __init__(self, master, get_material_names, get_general_specs, **kwargs):
        super().__init__(master, **kwargs)
        self.get_material_names = get_material_names
        self.get_general_specs = get_general_specs
        self.edit_index = None
        self._build_ui()

    def _build_ui(self):
        card = tb.Labelframe(self, text=ar("إدخال دكتور"), padding=10, bootstyle="info")
        card.pack(fill="both", expand=False, pady=6)

        frm = tb.Frame(card)
        frm.pack(fill="x")

        tb.Label(frm, text=ar("اسم الدكتور:"), anchor="e", justify="right", font=SECTION_FONT)\
            .grid(row=0, column=1, sticky="e", padx=4, pady=2)
        self.name = tb.Entry(frm)
        self.name.grid(row=0, column=0, sticky="ew", padx=4, pady=2)

        tb.Label(frm, text=ar("عدد الساعات:"), anchor="e", justify="right", font=SECTION_FONT)\
            .grid(row=1, column=1, sticky="e", padx=4, pady=2)
        self.hours = tb.Entry(frm)
        self.hours.grid(row=1, column=0, sticky="ew", padx=4, pady=2)

        frm.grid_columnconfigure(0, weight=1)

        # اختيار المواد من قائمة المواد المُدخلة (Listbox + Scrollbar)
        pick = tb.Labelframe(card, text=ar("اختر المواد (متعددة) من المواد المُدخلة"), padding=6, bootstyle="secondary")
        pick.pack(fill="both", expand=True, pady=4)

        mats_wrap = tb.Frame(pick)
        mats_wrap.pack(fill="both", expand=True)

        self.materials_listbox = tk.Listbox(mats_wrap, selectmode=tk.MULTIPLE, exportselection=False,
                                            height=7, justify="right", font=(BASE_FONT_FAMILY, BASE_FONT_SIZE))
        self.materials_listbox.grid(row=0, column=0, sticky="nsew", padx=(0,0), pady=(0,0))
        mats_vbar = tb.Scrollbar(mats_wrap, orient="vertical", command=self.materials_listbox.yview, bootstyle="primary")
        mats_vbar.grid(row=0, column=1, sticky="ns")
        self.materials_listbox.configure(yscrollcommand=mats_vbar.set)
        mats_wrap.grid_rowconfigure(0, weight=1)
        mats_wrap.grid_columnconfigure(0, weight=1)

        # تخصصات عامة للدكتور (Listbox + Scrollbar) مبنية من المواد
        spec_pick = tb.Labelframe(card, text=ar("اختر التخصصات العامة (أكثر من واحد ممكن)"), padding=6, bootstyle="secondary")
        spec_pick.pack(fill="both", expand=True, pady=4)

        specs_wrap = tb.Frame(spec_pick)
        specs_wrap.pack(fill="both", expand=True)

        self.general_specs_listbox = tk.Listbox(specs_wrap, selectmode=tk.MULTIPLE, exportselection=False,
                                                height=7, justify="right", activestyle="dotbox",
                                                font=(BASE_FONT_FAMILY, BASE_FONT_SIZE))
        self.general_specs_listbox.grid(row=0, column=0, sticky="nsew")
        specs_vbar = tb.Scrollbar(specs_wrap, orient="vertical", command=self.general_specs_listbox.yview, bootstyle="primary")
        specs_vbar.grid(row=0, column=1, sticky="ns")
        self.general_specs_listbox.configure(yscrollcommand=specs_vbar.set)
        specs_wrap.grid_rowconfigure(0, weight=1)
        specs_wrap.grid_columnconfigure(0, weight=1)

        # أزرار
        btns = tb.Frame(card)
        btns.pack(fill="x", pady=(4, 0))
        self.btn_new = tb.Button(btns, text=ar("جديد"), bootstyle="secondary", command=self.clear_form)
        self.btn_new.pack(side="right", padx=4)
        self.btn_add = tb.Button(btns, text=ar("إضافة/تحديث دكتور"), bootstyle="success", command=self.add_or_update)
        self.btn_add.pack(side="right", padx=4)
        self.btn_save = tb.Button(btns, text=ar("حفظ إلى Excel"), bootstyle="primary", command=self.save_to_excel)
        self.btn_save.pack(side="right", padx=4)
        # استيراد
        self.btn_import = tb.Button(btns, text=ar("استيراد Excel"), bootstyle="info", command=self.import_from_excel)
        self.btn_import.pack(side="right", padx=4)

        # جدول + Scrollbars
        table_card = tb.Labelframe(self, text=ar("الدكاترة المُدخلون"), padding=6, bootstyle="secondary")
        table_card.pack(fill="both", expand=True, pady=4)

        tv_wrap = tb.Frame(table_card)
        tv_wrap.pack(fill="both", expand=True)

        cols = ("الدكتور", "المواد", "عدد الساعات", "التخصصات العامة")
        self.tree = tb.Treeview(tv_wrap, columns=cols, show="headings", height=8, style="App.Treeview")
        for c in cols:
            self.tree.heading(c, text=ar(c))
            self.tree.column(c, width=160 if c != "التخصصات العامة" else 220, anchor="e", stretch=True)
        self.tree.grid(row=0, column=0, sticky="nsew")

        ybar = tb.Scrollbar(tv_wrap, orient="vertical", command=self.tree.yview, bootstyle="primary")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar = tb.Scrollbar(tv_wrap, orient="horizontal", command=self.tree.xview, bootstyle="primary")
        xbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)

        tv_wrap.grid_rowconfigure(0, weight=1)
        tv_wrap.grid_columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self.on_select_row)

        # أول تحميل
        self.reload_materials_names()
        self.reload_general_specs()

    def reload_materials_names(self):
        names = self.get_material_names()
        self.materials_listbox.delete(0, tk.END)
        for n in names:
            self.materials_listbox.insert(tk.END, n)

    def reload_general_specs(self):
        specs = self.get_general_specs()
        self.general_specs_listbox.delete(0, tk.END)
        for s in specs:
            self.general_specs_listbox.insert(tk.END, s)

    def clear_form(self):
        self.edit_index = None
        self.name.delete(0, tk.END)
        self.hours.delete(0, tk.END)
        self.materials_listbox.selection_clear(0, tk.END)
        self.general_specs_listbox.selection_clear(0, tk.END)

    def on_select_row(self, _):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        idx = int(self.tree.item(iid, "text"))
        self.edit_index = idx
        row = professors[idx]

        self.name.delete(0, tk.END); self.name.insert(0, row["الدكتور"])
        self.hours.delete(0, tk.END); self.hours.insert(0, str(row.get("عدد الساعات", "")))

        # المواد المختارة
        self.reload_materials_names()
        names = self.get_material_names()
        self.materials_listbox.selection_clear(0, tk.END)
        for i, nm in enumerate(names):
            if nm in row.get("المواد", []):
                self.materials_listbox.selection_set(i)

        # التخصصات العامة المختارة
        self.reload_general_specs()
        specs = self.get_general_specs()
        self.general_specs_listbox.selection_clear(0, tk.END)
        for i, sp in enumerate(specs):
            if sp in row.get("التخصصات العامة", []):
                self.general_specs_listbox.selection_set(i)

    def add_or_update(self):
        name = self.name.get().strip()
        h = self.hours.get().strip()
        if not name:
            return show_err("أدخل اسم الدكتور.")
        try:
            hours_val = float(h) if h else 0.0
        except:
            return show_err("عدد الساعات يجب أن يكون رقمًا.")

        # مواد مختارة
        names = self.get_material_names()
        sel_idx = list(self.materials_listbox.curselection())
        selected_materials = [names[i] for i in sel_idx]

        # تخصصات عامة مختارة من القائمة
        specs = self.get_general_specs()
        sel_specs_idx = list(self.general_specs_listbox.curselection())
        selected_specs = [specs[i] for i in sel_specs_idx]

        row = {
            "الدكتور": name,
            "المواد": selected_materials,
            "عدد الساعات": hours_val,
            "التخصصات العامة": selected_specs
        }

        if self.edit_index is None:
            professors.append(row)
        else:
            professors[self.edit_index] = row

        self.refresh_table()
        self.clear_form()

    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for idx, r in enumerate(professors):
            vals = (
                r["الدكتور"],
                " | ".join(r.get("المواد", [])),
                r.get("عدد الساعات", ""),
                join_specs(r.get("التخصصات العامة", []))
            )
            self.tree.insert("", "end", text=str(idx), values=vals)

    def save_to_excel(self):
        if not professors:
            return show_err("لا يوجد دكاترة لحفظهم.")
        df = pd.DataFrame([
            {
                "الدكتور": r["الدكتور"],
                "المواد": "،".join(r.get("المواد", [])),
                "عدد الساعات": r.get("عدد الساعات", ""),
                "التخصصات العامة": join_specs(r.get("التخصصات العامة", [])),
            } for r in professors
        ])
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="professors.xlsx", title=ar("حفظ ملف الدكاترة"))
        if not path:
            return
        df.to_excel(path, index=False)
        show_info(f"تم الحفظ في:\n{path}")

    def import_from_excel(self):
        """استيراد دكاترة من ملف بنفس صيغة الحفظ. تقسيم بسيط للفواصل العربية/الإنجليزية."""
        global professors
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title=ar("اختر ملف الدكاترة (professors.xlsx)")
        )
        if not path:
            return
        try:
            df = pd.read_excel(path)
        except Exception as e:
            return show_err(f"تعذر قراءة الملف:\n{e}")

        new_rows = []
        for _, r in df.iterrows():
            name = str(r.get("الدكتور", "")).strip()
            if not name:
                continue

            mats = split_specs(r.get("المواد", ""))
            hours = r.get("عدد الساعات", "")
            try:
                hours = float(hours) if str(hours).strip() != "" else 0.0
            except:
                hours = 0.0

            gens = split_specs(r.get("التخصصات العامة", ""))

            new_rows.append({
                "الدكتور": name,
                "المواد": mats,
                "عدد الساعات": hours,
                "التخصصات العامة": gens
            })

        professors = new_rows
        self.reload_materials_names()
        self.reload_general_specs()
        self.refresh_table()
        show_info("تم استيراد الدكاترة بنجاح.")

# ========= تبويب القاعات =========
class RoomsTab(tb.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.edit_index = None
        self._build_ui()

    def _build_ui(self):
        card = tb.Labelframe(self, text=ar("إدخال قاعة"), padding=10, bootstyle="info")
        card.pack(fill="both", expand=False, pady=6)

        frm = tb.Frame(card)
        frm.pack(fill="x")

        tb.Label(frm, text=ar("اسم/رقم القاعة:"), anchor="e", justify="right", font=SECTION_FONT)\
            .grid(row=0, column=1, sticky="e", padx=4, pady=2)
        self.name = tb.Entry(frm)
        self.name.grid(row=0, column=0, sticky="ew", padx=4, pady=2)

        tb.Label(frm, text=ar("النوع:"), anchor="e", justify="right", font=SECTION_FONT)\
            .grid(row=1, column=1, sticky="e", padx=4, pady=2)
        self.kind = tb.Combobox(frm, values=["نظري", "عملي"], state="readonly", height=10)
        self.kind.grid(row=1, column=0, sticky="ew", padx=4, pady=2)

        tb.Label(frm, text=ar("المساحة/السعة:"), anchor="e", justify="right", font=SECTION_FONT)\
            .grid(row=2, column=1, sticky="e", padx=4, pady=2)
        self.cap = tb.Entry(frm)
        self.cap.grid(row=2, column=0, sticky="ew", padx=4, pady=2)

        frm.grid_columnconfigure(0, weight=1)

        btns = tb.Frame(card)
        btns.pack(fill="x", pady=(4, 0))
        self.btn_new = tb.Button(btns, text=ar("جديد"), bootstyle="secondary", command=self.clear_form)
        self.btn_new.pack(side="right", padx=4)
        self.btn_add = tb.Button(btns, text=ar("إضافة/تحديث قاعة"), bootstyle="success", command=self.add_or_update)
        self.btn_add.pack(side="right", padx=4)
        self.btn_save = tb.Button(btns, text=ar("حفظ إلى Excel"), bootstyle="primary", command=self.save_to_excel)
        self.btn_save.pack(side="right", padx=4)
        # استيراد
        self.btn_import = tb.Button(btns, text=ar("استيراد Excel"), bootstyle="info", command=self.import_from_excel)
        self.btn_import.pack(side="right", padx=4)

        table_card = tb.Labelframe(self, text=ar("القاعات المُدخلة"), padding=6, bootstyle="secondary")
        table_card.pack(fill="both", expand=True, pady=4)

        tv_wrap = tb.Frame(table_card)
        tv_wrap.pack(fill="both", expand=True)

        cols = ("القاعة", "النوع", "المساحة")
        self.tree = tb.Treeview(tv_wrap, columns=cols, show="headings", height=8, style="App.Treeview")
        for c in cols:
            self.tree.heading(c, text=ar(c))
            self.tree.column(c, width=160, anchor="e", stretch=True)
        self.tree.grid(row=0, column=0, sticky="nsew")

        ybar = tb.Scrollbar(tv_wrap, orient="vertical", command=self.tree.yview, bootstyle="primary")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar = tb.Scrollbar(tv_wrap, orient="horizontal", command=self.tree.xview, bootstyle="primary")
        xbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)

        tv_wrap.grid_rowconfigure(0, weight=1)
        tv_wrap.grid_columnconfigure(0, weight=1)

    def clear_form(self):
        self.edit_index = None
        self.name.delete(0, tk.END)
        self.kind.set("")
        self.cap.delete(0, tk.END)

    def on_select_row(self, _):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        idx = int(self.tree.item(iid, "text"))
        self.edit_index = idx
        row = rooms[idx]
        self.name.delete(0, tk.END); self.name.insert(0, row["القاعة"])
        self.kind.set(row["النوع"])
        self.cap.delete(0, tk.END); self.cap.insert(0, str(row.get("المساحة", "")))

    def add_or_update(self):
        name = self.name.get().strip()
        kind = self.kind.get().strip()
        cap  = self.cap.get().strip()

        if not name:
            return show_err("أدخل اسم/رقم القاعة.")
        if kind not in ("نظري", "عملي"):
            return show_err("اختر النوع: نظري أو عملي.")
        if not cap.isdigit():
            return show_err("المساحة/السعة يجب أن تكون رقمًا صحيحًا.")

        row = {"القاعة": name, "النوع": kind, "المساحة": int(cap)}
        if self.edit_index is None:
            rooms.append(row)
        else:
            rooms[self.edit_index] = row

        self.refresh_table()
        self.clear_form()

    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for idx, r in enumerate(rooms):
            vals = (r["القاعة"], r["النوع"], r["المساحة"])
            self.tree.insert("", "end", text=str(idx), values=vals)

    def save_to_excel(self):
        if not rooms:
            return show_err("لا توجد قاعات لحفظها.")
        df = pd.DataFrame(rooms)
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="rooms.xlsx", title=ar("حفظ ملف القاعات"))
        if not path:
            return
        df.to_excel(path, index=False)
        show_info(f"تم الحفظ في:\n{path}")

    def import_from_excel(self):
        """استيراد قاعات من ملف بنفس صيغة الحفظ."""
        global rooms
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title=ar("اختر ملف القاعات (rooms.xlsx)")
        )
        if not path:
            return
        try:
            df = pd.read_excel(path)
        except Exception as e:
            return show_err(f"تعذر قراءة الملف:\n{e}")

        new_rows = []
        for _, r in df.iterrows():
            name = str(r.get("القاعة", "")).strip()
            kind = str(r.get("النوع", "")).strip()
            cap  = r.get("المساحة", "")
            if not name or kind not in ("نظري", "عملي"):
                continue
            try:
                cap_int = int(str(cap).strip())
            except:
                continue
            new_rows.append({"القاعة": name, "النوع": kind, "المساحة": cap_int})

        rooms = new_rows
        self.refresh_table()
        show_info("تم استيراد القاعات بنجاح.")

# ========= التطبيق الرئيسي =========
class App(tb.Window):
    def __init__(self):
        super().__init__(themename="minty")
        self.title(ar("dataGenerate - مُولّد بيانات الجدولة"))
        self.geometry("1200x770")
        self.minsize(820, 770)

        style = tb.Style()
        # خطوط عامة
        style.configure(".", font=(BASE_FONT_FAMILY, BASE_FONT_SIZE))
        style.configure("TButton", font=BOLD_BTN_FONT)
        style.configure("TLabelframe.Label", font=SECTION_FONT)
        # ضبط الـ Treeview لتفادي قصّ الحروف + وضوح أكبر
        style.configure("App.Treeview", rowheight=ROW_HEIGHT_TV, font=TREE_FONT)
        style.configure("App.Treeview.Heading", font=SECTION_FONT)

        # ترويسة
        header = tb.Frame(self, padding=10)
        header.pack(fill="x")
        tb.Label(header, text=ar("🧩 مُولّد بيانات الجدولة (مواد • دكاترة • قاعات)"),
                 font=HEADING_FONT, anchor="e", justify="right").pack(side="right")
        Separator(self).pack(fill="x", pady=(2, 0))

        # تبويبات
        self.notebook = tb.Notebook(self, bootstyle="pills")
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        # المواد
        self.tab_materials = MaterialsTab(self, on_materials_change=self._materials_changed)
        self.notebook.add(self.tab_materials, text=ar("المواد"))

        # الدكاترة
        self.tab_professors = ProfessorsTab(self,
                                            get_material_names=self._get_material_names,
                                            get_general_specs=self._get_general_specs)
        self.notebook.add(self.tab_professors, text=ar("الدكاترة"))

        # القاعات
        self.tab_rooms = RoomsTab(self)
        self.notebook.add(self.tab_rooms, text=ar("القاعات"))

        # تذييل
        footer = tb.Frame(self, padding=8)
        footer.pack(fill="x")
        tb.Label(footer, text=ar("© مُولّد بيانات الجدولة"), anchor="center", justify="center",
                font=(BASE_FONT_FAMILY, BASE_FONT_SIZE-1)).pack(fill="x")

    def _get_material_names(self):
        return [m["المادة"] for m in materials]

    def _get_general_specs(self):
        return all_general_specs_from_materials()

    def _materials_changed(self):
        # حدّث قائمة المواد والتخصصات العامة في تبويب الدكاترة
        self.tab_professors.reload_materials_names()
        self.tab_professors.reload_general_specs()
        self.tab_professors.refresh_table()

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
