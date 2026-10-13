#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Калькулятор потерь напряжения и токов однофазного КЗ
с учетом трансформатора. Экспорт в Excel, сохранение/загрузка проектов.
"""

import json
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from copy import deepcopy
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None

# ====================== Константы и формулы ======================

RHO_CU = 0.0175   # Ом·мм²/м медь (20°C)
RHO_AL = 0.0280   # Ом·мм²/м алюминий
X0_CABLE = 0.08   # Ом/км типичное реактивное сопротивление кабеля

def calc_current(power_kw, voltage, cos_phi, phases=1):
    """Ток нагрузки, А"""
    if cos_phi <= 0 or voltage <= 0:
        return 0.0
    if phases == 3:
        return (power_kw * 1000) / (math.sqrt(3) * voltage * cos_phi)
    else:
        return (power_kw * 1000) / (voltage * cos_phi)

def calc_voltage_drop(I, L_m, S_mm2, cos_phi, material="медь", phases=1):
    """Потеря напряжения ΔU в В и %"""
    if S_mm2 <= 0 or L_m <= 0 or I <= 0:
        return 0.0, 0.0
    rho = RHO_CU if material.lower().startswith("м") else RHO_AL
    R = rho * L_m / S_mm2          # Ом (один проводник)
    X = X0_CABLE * L_m / 1000      # Ом
    sin_phi = math.sqrt(max(0.0, 1 - cos_phi**2))
    if phases == 3:
        dU = math.sqrt(3) * I * (R * cos_phi + X * sin_phi)
        U_nom = 400.0
    else:
        dU = 2 * I * (R * cos_phi + X * sin_phi)
        U_nom = 230.0
    dU_pct = 100.0 * dU / U_nom
    return dU, dU_pct

def calc_transformer_Z(S_kVA, uk_pct, Pk_W, U_lv=0.4):
    """Сопротивление трансформатора, приведённое к НН (Ом)"""
    if S_kVA <= 0:
        return 0.0, 0.0
    # uk% = 100 * Z * S / U²
    Z = (uk_pct / 100.0) * (U_lv ** 2) / S_kVA
    # R ≈ Pk / (3 * Iном²) но упрощённо
    Inom = S_kVA / (math.sqrt(3) * U_lv)
    R = Pk_W / (3 * Inom ** 2) if Inom > 0 else 0.0
    X = math.sqrt(max(0.0, Z**2 - R**2))
    return R, X

def calc_sc_current_1ph(U_ph, R_tr, X_tr, R_cab, X_cab, L_m, contacts_mohm=15.0):
    """Ток однофазного КЗ (петля фаза-ноль), А"""
    # 2 * кабель (фаза + ноль)
    R_loop = R_tr + 2 * R_cab + contacts_mohm / 1000.0
    X_loop = X_tr + 2 * X_cab
    Z_loop = math.sqrt(R_loop**2 + X_loop**2)
    if Z_loop < 1e-9:
        return 0.0
    return U_ph / Z_loop

# ====================== Модель данных ======================

class CableRow:
    def __init__(self):
        self.num = ""
        self.name = ""
        self.group_num = ""
        self.floor = ""
        self.power = 0.0          # кВт
        self.voltage = 230        # 230 или 400
        self.cos_phi = 0.9
        self.current = 0.0
        self.n_cables = 1
        self.brand = "ВВГнг-LS"
        self.cores = 3
        self.section = 2.5
        self.len_tray = 0.0
        self.len_corrug = 0.0
        self.len_channel = 0.0
        self.len_pvc = 0.0
        self.len_steel = 0.0
        self.len_sum = 0.0
        self.material = "медь"
        self.len_line = 0.0
        self.len_reduced = 0.0
        self.dU = 0.0
        self.dU_pct = 0.0
        self.Isc = 0.0

    def to_dict(self):
        return self.__dict__.copy()

    @classmethod
    def from_dict(cls, d):
        r = cls()
        r.__dict__.update(d)
        return r

class GroupSheet:
    def __init__(self, title="Группа"):
        self.title = title
        self.rows = []

    def to_dict(self):
        return {"title": self.title, "rows": [r.to_dict() for r in self.rows]}

    @classmethod
    def from_dict(cls, d):
        g = cls(d.get("title", "Группа"))
        g.rows = [CableRow.from_dict(r) for r in d.get("rows", [])]
        return g

class Project:
    def __init__(self):
        self.name = "Новый проект"
        self.transformer = {
            "S_kVA": 250.0,
            "uk_pct": 5.5,
            "Pk_W": 3700.0,
            "scheme": "D/Yn",
            "U_lv": 0.4,
        }
        self.sheets = [GroupSheet("ВРУ")]
        for i in range(1, 10):
            self.sheets.append(GroupSheet(f"Группа {i}"))

    def to_dict(self):
        return {
            "name": self.name,
            "transformer": self.transformer,
            "sheets": [s.to_dict() for s in self.sheets],
        }

    @classmethod
    def from_dict(cls, d):
        p = cls()
        p.name = d.get("name", "Проект")
        p.transformer = d.get("transformer", p.transformer)
        p.sheets = [GroupSheet.from_dict(s) for s in d.get("sheets", [])]
        return p

# ====================== GUI ======================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Калькулятор потерь напряжения и токов КЗ")
        self.geometry("1400x800")
        self.project = Project()
        self.current_sheet_idx = 0

        self._build_menu()
        self._build_transformer_frame()
        self._build_notebook()
        self._build_buttons()

        self.refresh_sheet()

    def _build_menu(self):
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Новый проект", command=self.new_project)
        filemenu.add_command(label="Открыть...", command=self.load_project)
        filemenu.add_command(label="Сохранить...", command=self.save_project)
        filemenu.add_separator()
        filemenu.add_command(label="Экспорт в Excel...", command=self.export_excel)
        filemenu.add_separator()
        filemenu.add_command(label="Выход", command=self.quit)
        menubar.add_cascade(label="Файл", menu=filemenu)
        self.config(menu=menubar)

    def _build_transformer_frame(self):
        fr = ttk.LabelFrame(self, text="Трансформатор")
        fr.pack(fill="x", padx=5, pady=5)

        self.tr_vars = {}
        fields = [
            ("S_kVA", "Мощность, кВА", 250),
            ("uk_pct", "uk, %", 5.5),
            ("Pk_W", "Потери КЗ, Вт", 3700),
            ("U_lv", "Uнн, кВ", 0.4),
        ]
        for i, (key, label, default) in enumerate(fields):
            ttk.Label(fr, text=label).grid(row=0, column=i*2, padx=3, pady=2)
            var = tk.DoubleVar(value=self.project.transformer.get(key, default))
            self.tr_vars[key] = var
            ttk.Entry(fr, textvariable=var, width=10).grid(row=0, column=i*2+1, padx=3)

        ttk.Label(fr, text="Схема:").grid(row=0, column=8)
        self.scheme_var = tk.StringVar(value=self.project.transformer.get("scheme", "D/Yn"))
        ttk.Combobox(fr, textvariable=self.scheme_var, values=["D/Yn", "Y/Yn", "Y/Zn"], width=8).grid(row=0, column=9)

    def _build_notebook(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=5, pady=5)
        self.frames = []
        for i, sheet in enumerate(self.project.sheets):
            fr = ttk.Frame(self.nb)
            self.nb.add(fr, text=sheet.title)
            self.frames.append(fr)
            self._build_sheet_ui(fr, i)
        self.nb.bind("<<NotebookTabChanged>>", self.on_tab_changed)

    def _build_sheet_ui(self, parent, idx):
        # Таблица
        cols = [
            "№", "Наименование", "№ гр.", "Этаж", "P, кВт", "U, В", "cosφ",
            "I, А", "N каб.", "Марка", "Жил", "S, мм²",
            "Лоток", "Гофра", "Канал", "ПВХ", "Сталь",
            "Σ L", "Матер.", "L лин.", "L прив.", "ΔU, В", "ΔU, %", "Iкз1, А"
        ]
        tree = ttk.Treeview(parent, columns=cols, show="headings", height=18)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=55, anchor="center")
        tree.column("Наименование", width=120)
        tree.column("Марка", width=80)

        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Сохраняем ссылку
        if not hasattr(self, "trees"):
            self.trees = []
        self.trees.append(tree)

        # Кнопки управления строками
        btn_fr = ttk.Frame(parent)
        btn_fr.pack(fill="x")
        ttk.Button(btn_fr, text="+ строка", command=lambda: self.add_row(idx)).pack(side="left", padx=2)
        ttk.Button(btn_fr, text="Удалить", command=lambda: self.del_row(idx)).pack(side="left", padx=2)
        ttk.Button(btn_fr, text="Рассчитать", command=lambda: self.recalc_sheet(idx)).pack(side="left", padx=2)

    def _build_buttons(self):
        fr = ttk.Frame(self)
        fr.pack(fill="x", padx=5, pady=5)
        ttk.Button(fr, text="Рассчитать всё", command=self.recalc_all).pack(side="left", padx=5)
        ttk.Button(fr, text="Сводная ведомость кабелей", command=self.show_cable_summary).pack(side="left", padx=5)
        ttk.Button(fr, text="Экспорт Excel", command=self.export_excel).pack(side="left", padx=5)

    def on_tab_changed(self, event):
        self.current_sheet_idx = self.nb.index(self.nb.select())

    def refresh_sheet(self, idx=None):
        if idx is None:
            idx = self.current_sheet_idx
        tree = self.trees[idx]
        tree.delete(*tree.get_children())
        sheet = self.project.sheets[idx]
        for r in sheet.rows:
            vals = [
                r.num, r.name, r.group_num, r.floor, f"{r.power:.2f}", r.voltage,
                f"{r.cos_phi:.2f}", f"{r.current:.1f}", r.n_cables, r.brand,
                r.cores, r.section,
                f"{r.len_tray:.1f}", f"{r.len_corrug:.1f}", f"{r.len_channel:.1f}",
                f"{r.len_pvc:.1f}", f"{r.len_steel:.1f}", f"{r.len_sum:.1f}",
                r.material, f"{r.len_line:.1f}", f"{r.len_reduced:.1f}",
                f"{r.dU:.2f}", f"{r.dU_pct:.2f}", f"{r.Isc:.0f}"
            ]
            tree.insert("", "end", values=vals)

    def add_row(self, idx):
        r = CableRow()
        r.num = str(len(self.project.sheets[idx].rows) + 1)
        self.project.sheets[idx].rows.append(r)
        self.refresh_sheet(idx)

    def del_row(self, idx):
        tree = self.trees[idx]
        sel = tree.selection()
        if not sel:
            return
        item = sel[0]
        pos = tree.index(item)
        del self.project.sheets[idx].rows[pos]
        self.refresh_sheet(idx)

    def _sync_transformer(self):
        for k, v in self.tr_vars.items():
            self.project.transformer[k] = v.get()
        self.project.transformer["scheme"] = self.scheme_var.get()

    def recalc_sheet(self, idx):
        self._sync_transformer()
        tr = self.project.transformer
        R_tr, X_tr = calc_transformer_Z(tr["S_kVA"], tr["uk_pct"], tr["Pk_W"], tr["U_lv"])

        sheet = self.project.sheets[idx]
        for r in sheet.rows:
            phases = 3 if r.voltage >= 380 else 1
            r.current = calc_current(r.power, r.voltage, r.cos_phi, phases)
            r.len_sum = r.len_tray + r.len_corrug + r.len_channel + r.len_pvc + r.len_steel
            L = r.len_reduced if r.len_reduced > 0 else r.len_line
            if L <= 0:
                L = r.len_sum
            r.dU, r.dU_pct = calc_voltage_drop(
                r.current, L, r.section, r.cos_phi, r.material, phases
            )
            # Сопротивление кабеля на длину L
            rho = RHO_CU if r.material.lower().startswith("м") else RHO_AL
            R_cab = rho * L / r.section if r.section > 0 else 0
            X_cab = X0_CABLE * L / 1000
            U_ph = 230.0
            r.Isc = calc_sc_current_1ph(U_ph, R_tr, X_tr, R_cab, X_cab, L)
        self.refresh_sheet(idx)

    def recalc_all(self):
        for i in range(len(self.project.sheets)):
            self.recalc_sheet(i)
        messagebox.showinfo("Готово", "Все листы пересчитаны")

    def show_cable_summary(self):
        summary = {}
        for sheet in self.project.sheets:
            for r in sheet.rows:
                key = (r.brand, f"{r.cores}x{r.section}")
                if key not in summary:
                    summary[key] = {
                        "tray": 0.0, "corrug": 0.0, "channel": 0.0,
                        "pvc": 0.0, "steel": 0.0, "total": 0.0
                    }
                summary[key]["tray"] += r.len_tray
                summary[key]["corrug"] += r.len_corrug
                summary[key]["channel"] += r.len_channel
                summary[key]["pvc"] += r.len_pvc
                summary[key]["steel"] += r.len_steel
                summary[key]["total"] += r.len_sum

        win = tk.Toplevel(self)
        win.title("Ведомость используемых кабелей")
        win.geometry("900x400")
        cols = ["Марка", "жил×сечение", "Лоток", "Гофра", "Канал", "ПВХ", "Сталь", "Итого"]
        tree = ttk.Treeview(win, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=100, anchor="center")
        for (brand, sec), v in sorted(summary.items()):
            tree.insert("", "end", values=[
                brand, sec,
                f"{v['tray']:.1f}", f"{v['corrug']:.1f}", f"{v['channel']:.1f}",
                f"{v['pvc']:.1f}", f"{v['steel']:.1f}", f"{v['total']:.1f}"
            ])
        tree.pack(fill="both", expand=True)

    def new_project(self):
        if messagebox.askyesno("Новый", "Создать новый проект? Несохранённые данные будут потеряны."):
            self.project = Project()
            self.refresh_all()

    def save_project(self):
        self._sync_transformer()
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON проект", "*.json")]
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.project.to_dict(), f, ensure_ascii=False, indent=2)
        messagebox.showinfo("Сохранено", f"Проект сохранён:\n{path}")

    def load_project(self):
        path = filedialog.askopenfilename(filetypes=[("JSON проект", "*.json")])
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.project = Project.from_dict(data)
        self.refresh_all()
        messagebox.showinfo("Загружено", f"Проект загружен:\n{path}")

    def refresh_all(self):
        # Обновляем вкладки
        for i, sheet in enumerate(self.project.sheets):
            self.nb.tab(i, text=sheet.title)
            self.refresh_sheet(i)
        # Трансформатор
        for k, var in self.tr_vars.items():
            var.set(self.project.transformer.get(k, 0))
        self.scheme_var.set(self.project.transformer.get("scheme", "D/Yn"))

    def export_excel(self):
        if openpyxl is None:
            messagebox.showerror("Ошибка", "Установите openpyxl: pip install openpyxl")
            return
        self._sync_transformer()
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return

        wb = openpyxl.Workbook()
        # Удаляем дефолтный лист
        default = wb.active
        wb.remove(default)

        thin = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        header_fill = PatternFill("solid", fgColor="4472C4")
        header_font = Font(bold=True, color="FFFFFF")

        for sheet in self.project.sheets:
            ws = wb.create_sheet(title=sheet.title[:31])
            headers = [
                "№", "Наименование", "№ группы", "Этаж", "Мощность, кВт", "Напряжение, В",
                "cos φ", "Ток, А", "Кол-во кабелей", "Марка", "Кол-во жил", "Сечение, мм²",
                "в лотке", "в гофре", "в канале", "в ПВХ", "в стальной трубе",
                "Сумма длин", "Материал", "Длина линии", "Приведенная длина",
                "ΔU, В", "ΔU, %", "Iкз1, А"
            ]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(1, col, h)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = thin
                cell.alignment = Alignment(wrap_text=True, horizontal="center")

            for row_idx, r in enumerate(sheet.rows, 2):
                vals = [
                    r.num, r.name, r.group_num, r.floor, r.power, r.voltage,
                    r.cos_phi, round(r.current, 2), r.n_cables, r.brand,
                    r.cores, r.section,
                    r.len_tray, r.len_corrug, r.len_channel, r.len_pvc, r.len_steel,
                    r.len_sum, r.material, r.len_line, r.len_reduced,
                    round(r.dU, 2), round(r.dU_pct, 2), round(r.Isc, 0)
                ]
                for col, v in enumerate(vals, 1):
                    cell = ws.cell(row_idx, col, v)
                    cell.border = thin

            # Автоширина
            for col in range(1, len(headers) + 1):
                ws.column_dimensions[get_column_letter(col)].width = 12

        # Сводная ведомость
        ws = wb.create_sheet("Ведомость кабелей")
        headers = ["Марка кабеля", "жил×сечение", "в лотке", "в гофре", "в канале",
                   "в ПВХ-трубе", "в стальной трубе", "Итого длина"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(1, col, h)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin

        summary = {}
        for sheet in self.project.sheets:
            for r in sheet.rows:
                key = (r.brand, f"{r.cores}x{r.section}")
                if key not in summary:
                    summary[key] = [0.0]*6
                summary[key][0] += r.len_tray
                summary[key][1] += r.len_corrug
                summary[key][2] += r.len_channel
                summary[key][3] += r.len_pvc
                summary[key][4] += r.len_steel
                summary[key][5] += r.len_sum

        for row_idx, ((brand, sec), lens) in enumerate(sorted(summary.items()), 2):
            ws.cell(row_idx, 1, brand).border = thin
            ws.cell(row_idx, 2, sec).border = thin
            for col, v in enumerate(lens, 3):
                cell = ws.cell(row_idx, col, round(v, 1))
                cell.border = thin

        for col in range(1, 9):
            ws.column_dimensions[get_column_letter(col)].width = 16

        # Параметры трансформатора
        ws = wb.create_sheet("Трансформатор")
        ws["A1"] = "Параметр"
        ws["B1"] = "Значение"
        ws["A1"].font = header_font
        ws["B1"].font = header_font
        params = [
            ("Мощность, кВА", self.project.transformer["S_kVA"]),
            ("uk, %", self.project.transformer["uk_pct"]),
            ("Потери КЗ, Вт", self.project.transformer["Pk_W"]),
            ("Uнн, кВ", self.project.transformer["U_lv"]),
            ("Схема", self.project.transformer["scheme"]),
        ]
        for i, (name, val) in enumerate(params, 2):
            ws.cell(i, 1, name)
            ws.cell(i, 2, val)

        wb.save(path)
        messagebox.showinfo("Экспорт", f"Excel сохранён:\n{path}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
