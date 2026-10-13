#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Портативный калькулятор потерь напряжения и токов однофазного КЗ.
Иерархия: ВРУ → неограниченное число щитов → группы в каждом щите.
Экспорт: Excel + PDF + DXF однолинейных схем каждого щита.
"""

import json
import math
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as pdf_canvas
except ImportError:
    pdf_canvas = None

try:
    import ezdxf
    from ezdxf import units
    from ezdxf.enums import TextEntityAlignment
except ImportError:
    ezdxf = None

# ====================== Константы и формулы ======================

RHO_CU = 0.0175
RHO_AL = 0.0280
X0_CABLE = 0.08  # Ом/км

def calc_current(power_kw, voltage, cos_phi, phases=1):
    if cos_phi <= 0 or voltage <= 0:
        return 0.0
    if phases == 3:
        return (power_kw * 1000) / (math.sqrt(3) * voltage * cos_phi)
    return (power_kw * 1000) / (voltage * cos_phi)

def calc_voltage_drop(I, L_m, S_mm2, cos_phi, material="медь", phases=1):
    if S_mm2 <= 0 or L_m <= 0 or I <= 0:
        return 0.0, 0.0
    rho = RHO_CU if str(material).lower().startswith("м") else RHO_AL
    R = rho * L_m / S_mm2
    X = X0_CABLE * L_m / 1000
    sin_phi = math.sqrt(max(0.0, 1 - cos_phi ** 2))
    if phases == 3:
        dU = math.sqrt(3) * I * (R * cos_phi + X * sin_phi)
        U_nom = 400.0
    else:
        dU = 2 * I * (R * cos_phi + X * sin_phi)
        U_nom = 230.0
    return dU, 100.0 * dU / U_nom

def calc_transformer_Z(S_kVA, uk_pct, Pk_W, U_lv=0.4):
    if S_kVA <= 0:
        return 0.0, 0.0
    Z = (uk_pct / 100.0) * (U_lv ** 2) / S_kVA
    Inom = S_kVA / (math.sqrt(3) * U_lv)
    R = Pk_W / (3 * Inom ** 2) if Inom > 0 else 0.0
    X = math.sqrt(max(0.0, Z ** 2 - R ** 2))
    return R, X

def calc_sc_current_1ph(U_ph, R_tr, X_tr, R_cab, X_cab, contacts_mohm=15.0):
    R_loop = R_tr + 2 * R_cab + contacts_mohm / 1000.0
    X_loop = X_tr + 2 * X_cab
    Z_loop = math.sqrt(R_loop ** 2 + X_loop ** 2)
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
        self.power = 0.0
        self.voltage = 230
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
        r.__dict__.update({k: v for k, v in d.items() if k in r.__dict__})
        return r

class Board:
    def __init__(self, title="Щит", is_vru=False):
        self.title = title
        self.is_vru = is_vru
        self.rows = []

    def to_dict(self):
        return {"title": self.title, "is_vru": self.is_vru, "rows": [r.to_dict() for r in self.rows]}

    @classmethod
    def from_dict(cls, d):
        b = cls(d.get("title", "Щит"), d.get("is_vru", False))
        b.rows = [CableRow.from_dict(r) for r in d.get("rows", [])]
        return b

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
        self.vru = Board("ВРУ", is_vru=True)
        self.boards = []

    def all_boards(self):
        return [self.vru] + self.boards

    def to_dict(self):
        return {
            "name": self.name,
            "transformer": self.transformer,
            "vru": self.vru.to_dict(),
            "boards": [b.to_dict() for b in self.boards],
        }

    @classmethod
    def from_dict(cls, d):
        p = cls()
        p.name = d.get("name", "Проект")
        p.transformer = d.get("transformer", p.transformer)
        if "vru" in d:
            p.vru = Board.from_dict(d["vru"])
        if "boards" in d:
            p.boards = [Board.from_dict(b) for b in d["boards"]]
        if "sheets" in d:  # совместимость
            sheets = d["sheets"]
            if sheets:
                p.vru = Board.from_dict(sheets[0])
                p.vru.is_vru = True
                p.vru.title = "ВРУ"
                p.boards = [Board.from_dict(s) for s in sheets[1:]]
        return p

# ====================== PDF-схемы ======================

def draw_single_line_scheme_pdf(board, project, path):
    if pdf_canvas is None:
        raise RuntimeError("reportlab не установлен")

    c = pdf_canvas.Canvas(path, pagesize=landscape(A4))
    W, H = landscape(A4)
    margin = 10 * mm
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(1.5)
    c.rect(margin, margin, W - 2 * margin, H - 2 * margin)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W / 2, H - 18 * mm, f"Однолинейная схема. {board.title}")
    c.setFont("Helvetica", 9)
    c.drawCentredString(W / 2, H - 24 * mm, f"Проект: {project.name}  |  {datetime.now().strftime('%d.%m.%Y')}")

    tr = project.transformer
    c.setFont("Helvetica", 8)
    c.drawString(margin + 5 * mm, H - 32 * mm,
                 f"Трансформатор: {tr['S_kVA']} кВА, uk={tr['uk_pct']}%, схема {tr['scheme']}")

    bus_y = H - 50 * mm
    bus_x0 = 40 * mm
    bus_x1 = W - 40 * mm
    c.setLineWidth(2.5)
    c.line(bus_x0, bus_y, bus_x1, bus_y)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(bus_x0, bus_y + 3 * mm, f"Шина {board.title}  0,4 кВ")

    n = max(len(board.rows), 1)
    step = (bus_x1 - bus_x0) / (n + 1)
    c.setLineWidth(1.0)

    for i, r in enumerate(board.rows):
        x = bus_x0 + step * (i + 1)
        c.line(x, bus_y, x, bus_y - 15 * mm)
        c.rect(x - 4 * mm, bus_y - 22 * mm, 8 * mm, 7 * mm)
        c.setFont("Helvetica", 6)
        c.drawCentredString(x, bus_y - 19 * mm, "QF")
        c.line(x, bus_y - 22 * mm, x, bus_y - 45 * mm)

        y_txt = bus_y - 50 * mm
        c.setFont("Helvetica", 6)
        lines = [
            f"{r.num or i+1}. {r.name or 'Группа'}",
            f"{r.brand} {r.cores}x{r.section}",
            f"L={r.len_sum or r.len_line:.0f} м",
            f"I={r.current:.1f} А",
            f"ΔU={r.dU_pct:.2f}%",
            f"Iкз={r.Isc:.0f} А",
        ]
        for j, t in enumerate(lines):
            c.drawCentredString(x, y_txt - j * 3.5 * mm, t)

    stamp_w, stamp_h = 110 * mm, 30 * mm
    sx = W - margin - stamp_w
    sy = margin + 2 * mm
    c.setLineWidth(0.7)
    c.rect(sx, sy, stamp_w, stamp_h)
    c.line(sx, sy + 10 * mm, sx + stamp_w, sy + 10 * mm)
    c.setFont("Helvetica", 7)
    c.drawString(sx + 2 * mm, sy + 22 * mm, f"Лист: схема {board.title}")
    c.drawString(sx + 2 * mm, sy + 14 * mm, f"Проект: {project.name}")
    c.drawString(sx + 2 * mm, sy + 4 * mm, f"Дата: {datetime.now().strftime('%d.%m.%Y')}")

    c.setFont("Helvetica", 7)
    c.drawString(margin + 3 * mm, margin + 3 * mm,
                 "Упрощённая однолинейная схема (ориентир ГОСТ 2.702 / 21.608). Для проектной документации доработать в CAD.")
    c.save()

# ====================== DXF-схемы ======================

def draw_single_line_scheme_dxf(board, project, path):
    """Однолинейная схема щита в DXF (открывается в AutoCAD → Save as DWG)"""
    if ezdxf is None:
        raise RuntimeError("ezdxf не установлен")

    doc = ezdxf.new("R2010")  # совместимый формат
    doc.units = units.MM
    msp = doc.modelspace()

    # Слои
    doc.layers.add("РАМКА", color=7)
    doc.layers.add("ШИНА", color=1)       # красный
    doc.layers.add("АППАРАТЫ", color=5)   # синий
    doc.layers.add("ЛИНИИ", color=3)      # зелёный
    doc.layers.add("ТЕКСТ", color=7)
    doc.layers.add("ОСНОВНАЯ_НАДПИСЬ", color=7)

    # Размеры листа (А3 альбом в мм)
    W, H = 420.0, 297.0
    margin = 10.0

    # Рамка
    msp.add_lwpolyline(
        [(margin, margin), (W - margin, margin), (W - margin, H - margin), (margin, H - margin), (margin, margin)],
        dxfattribs={"layer": "РАМКА", "lineweight": 50}
    )

    # Заголовок
    msp.add_text(
        f"Однолинейная схема. {board.title}",
        height=5.0,
        dxfattribs={"layer": "ТЕКСТ", "style": "Standard"}
    ).set_placement((W / 2, H - 18), align=TextEntityAlignment.MIDDLE_CENTER)

    msp.add_text(
        f"Проект: {project.name}  |  {datetime.now().strftime('%d.%m.%Y')}",
        height=3.0,
        dxfattribs={"layer": "ТЕКСТ"}
    ).set_placement((W / 2, H - 25), align=TextEntityAlignment.MIDDLE_CENTER)

    tr = project.transformer
    msp.add_text(
        f"Трансформатор: {tr['S_kVA']} кВА, uk={tr['uk_pct']}%, схема {tr['scheme']}",
        height=2.5,
        dxfattribs={"layer": "ТЕКСТ"}
    ).set_placement((margin + 5, H - 32))

    # Шина
    bus_y = H - 55
    bus_x0 = 50
    bus_x1 = W - 50
    msp.add_line((bus_x0, bus_y), (bus_x1, bus_y), dxfattribs={"layer": "ШИНА", "lineweight": 70})
    msp.add_text(
        f"Шина {board.title}  0,4 кВ",
        height=3.0,
        dxfattribs={"layer": "ТЕКСТ"}
    ).set_placement((bus_x0, bus_y + 4))

    n = max(len(board.rows), 1)
    step = (bus_x1 - bus_x0) / (n + 1)

    for i, r in enumerate(board.rows):
        x = bus_x0 + step * (i + 1)

        # Вертикаль от шины до автомата
        msp.add_line((x, bus_y), (x, bus_y - 18), dxfattribs={"layer": "ЛИНИИ"})

        # Условный автомат (прямоугольник)
        qf_h, qf_w = 10, 8
        msp.add_lwpolyline(
            [(x - qf_w/2, bus_y - 18 - qf_h), (x + qf_w/2, bus_y - 18 - qf_h),
             (x + qf_w/2, bus_y - 18), (x - qf_w/2, bus_y - 18), (x - qf_w/2, bus_y - 18 - qf_h)],
            dxfattribs={"layer": "АППАРАТЫ", "lineweight": 35}
        )
        msp.add_text("QF", height=2.2, dxfattribs={"layer": "ТЕКСТ"}).set_placement(
            (x, bus_y - 18 - qf_h/2), align=TextEntityAlignment.MIDDLE_CENTER
        )

        # Линия ниже автомата
        msp.add_line((x, bus_y - 18 - qf_h), (x, bus_y - 50), dxfattribs={"layer": "ЛИНИИ"})

        # Подписи
        y_txt = bus_y - 55
        lines = [
            f"{r.num or i+1}. {r.name or 'Группа'}",
            f"{r.brand} {r.cores}x{r.section}",
            f"L={r.len_sum or r.len_line:.0f} м",
            f"I={r.current:.1f} А",
            f"ΔU={r.dU_pct:.2f}%",
            f"Iкз={r.Isc:.0f} А",
        ]
        for j, t in enumerate(lines):
            msp.add_text(t, height=2.0, dxfattribs={"layer": "ТЕКСТ"}).set_placement(
                (x, y_txt - j * 4.0), align=TextEntityAlignment.MIDDLE_CENTER
            )

    # Основная надпись (упрощённо)
    stamp_w, stamp_h = 120, 35
    sx = W - margin - stamp_w
    sy = margin + 3
    msp.add_lwpolyline(
        [(sx, sy), (sx + stamp_w, sy), (sx + stamp_w, sy + stamp_h), (sx, sy + stamp_h), (sx, sy)],
        dxfattribs={"layer": "ОСНОВНАЯ_НАДПИСЬ"}
    )
    msp.add_line((sx, sy + 12), (sx + stamp_w, sy + 12), dxfattribs={"layer": "ОСНОВНАЯ_НАДПИСЬ"})
    msp.add_text(f"Лист: схема {board.title}", height=2.5, dxfattribs={"layer": "ТЕКСТ"}).set_placement((sx + 3, sy + 25))
    msp.add_text(f"Проект: {project.name}", height=2.5, dxfattribs={"layer": "ТЕКСТ"}).set_placement((sx + 3, sy + 16))
    msp.add_text(f"Дата: {datetime.now().strftime('%d.%m.%Y')}", height=2.5, dxfattribs={"layer": "ТЕКСТ"}).set_placement((sx + 3, sy + 5))

    msp.add_text(
        "Упрощённая однолинейная схема (ориентир ГОСТ 2.702 / 21.608). Откройте в AutoCAD → Сохранить как DWG.",
        height=2.0,
        dxfattribs={"layer": "ТЕКСТ"}
    ).set_placement((margin + 3, margin + 3))

    doc.saveas(path)

# ====================== GUI ======================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Калькулятор потерь напряжения и токов КЗ (портативный)")
        self.geometry("1450x820")
        self.project = Project()
        self.current_board_idx = 0

        self._build_menu()
        self._build_transformer_frame()
        self._build_board_selector()
        self._build_table_area()
        self._build_buttons()
        self.refresh_table()

    def _build_menu(self):
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Новый проект", command=self.new_project)
        filemenu.add_command(label="Открыть...", command=self.load_project)
        filemenu.add_command(label="Сохранить...", command=self.save_project)
        filemenu.add_separator()
        filemenu.add_command(label="Экспорт в Excel...", command=self.export_excel)
        filemenu.add_command(label="Экспорт схем (PDF)...", command=self.export_schemes_pdf)
        filemenu.add_command(label="Экспорт схем (DXF → DWG)...", command=self.export_schemes_dxf)
        filemenu.add_separator()
        filemenu.add_command(label="Выход", command=self.quit)
        menubar.add_cascade(label="Файл", menu=filemenu)
        self.config(menu=menubar)

    def _build_transformer_frame(self):
        fr = ttk.LabelFrame(self, text="Трансформатор")
        fr.pack(fill="x", padx=5, pady=3)
        self.tr_vars = {}
        fields = [
            ("S_kVA", "Мощность, кВА", 250),
            ("uk_pct", "uk, %", 5.5),
            ("Pk_W", "Потери КЗ, Вт", 3700),
            ("U_lv", "Uнн, кВ", 0.4),
        ]
        for i, (key, label, default) in enumerate(fields):
            ttk.Label(fr, text=label).grid(row=0, column=i * 2, padx=3, pady=2)
            var = tk.DoubleVar(value=self.project.transformer.get(key, default))
            self.tr_vars[key] = var
            ttk.Entry(fr, textvariable=var, width=9).grid(row=0, column=i * 2 + 1, padx=3)
        ttk.Label(fr, text="Схема:").grid(row=0, column=8)
        self.scheme_var = tk.StringVar(value=self.project.transformer.get("scheme", "D/Yn"))
        ttk.Combobox(fr, textvariable=self.scheme_var, values=["D/Yn", "Y/Yn", "Y/Zn"], width=7).grid(row=0, column=9)

    def _build_board_selector(self):
        fr = ttk.Frame(self)
        fr.pack(fill="x", padx=5, pady=3)
        ttk.Label(fr, text="Щит:").pack(side="left")
        self.board_var = tk.StringVar()
        self.board_combo = ttk.Combobox(fr, textvariable=self.board_var, width=40, state="readonly")
        self.board_combo.pack(side="left", padx=5)
        self.board_combo.bind("<<ComboboxSelected>>", self.on_board_changed)
        ttk.Button(fr, text="+ Щит", command=self.add_board).pack(side="left", padx=2)
        ttk.Button(fr, text="Переименовать", command=self.rename_board).pack(side="left", padx=2)
        ttk.Button(fr, text="Удалить щит", command=self.del_board).pack(side="left", padx=2)
        self._update_board_list()

    def _build_table_area(self):
        fr = ttk.Frame(self)
        fr.pack(fill="both", expand=True, padx=5, pady=3)
        cols = [
            "№", "Наименование", "№ гр.", "Этаж", "P, кВт", "U, В", "cosφ",
            "I, А", "N каб.", "Марка", "Жил", "S, мм²",
            "Лоток", "Гофра", "Канал", "ПВХ", "Сталь",
            "Σ L", "Матер.", "L лин.", "L прив.", "ΔU, В", "ΔU, %", "Iкз1, А"
        ]
        self.tree = ttk.Treeview(fr, columns=cols, show="headings", height=20)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=52, anchor="center")
        self.tree.column("Наименование", width=110)
        self.tree.column("Марка", width=75)
        vsb = ttk.Scrollbar(fr, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self.edit_row)

        btn_fr = ttk.Frame(self)
        btn_fr.pack(fill="x", padx=5)
        ttk.Button(btn_fr, text="+ строка", command=self.add_row).pack(side="left", padx=2)
        ttk.Button(btn_fr, text="Удалить строку", command=self.del_row).pack(side="left", padx=2)
        ttk.Button(btn_fr, text="Рассчитать текущий щит", command=self.recalc_current).pack(side="left", padx=2)

    def _build_buttons(self):
        fr = ttk.Frame(self)
        fr.pack(fill="x", padx=5, pady=5)
        ttk.Button(fr, text="Рассчитать всё", command=self.recalc_all).pack(side="left", padx=4)
        ttk.Button(fr, text="Ведомость кабелей", command=self.show_cable_summary).pack(side="left", padx=4)
        ttk.Button(fr, text="Экспорт Excel", command=self.export_excel).pack(side="left", padx=4)
        ttk.Button(fr, text="Экспорт PDF", command=self.export_schemes_pdf).pack(side="left", padx=4)
        ttk.Button(fr, text="Экспорт DXF (→DWG)", command=self.export_schemes_dxf).pack(side="left", padx=4)

    def _current_board(self):
        if self.current_board_idx == 0:
            return self.project.vru
        return self.project.boards[self.current_board_idx - 1]

    def _update_board_list(self):
        names = [self.project.vru.title] + [b.title for b in self.project.boards]
        self.board_combo["values"] = names
        if names:
            self.board_combo.current(min(self.current_board_idx, len(names) - 1))

    def on_board_changed(self, event=None):
        self.current_board_idx = self.board_combo.current()
        self.refresh_table()

    def add_board(self):
        name = simpledialog.askstring("Новый щит", "Название щита:", initialvalue=f"Щит {len(self.project.boards)+1}")
        if name:
            self.project.boards.append(Board(name))
            self.current_board_idx = len(self.project.boards)
            self._update_board_list()
            self.refresh_table()

    def rename_board(self):
        b = self._current_board()
        name = simpledialog.askstring("Переименовать", "Новое название:", initialvalue=b.title)
        if name:
            b.title = name
            self._update_board_list()

    def del_board(self):
        if self.current_board_idx == 0:
            messagebox.showwarning("", "ВРУ удалить нельзя")
            return
        if messagebox.askyesno("Удалить", f"Удалить щит «{self._current_board().title}»?"):
            del self.project.boards[self.current_board_idx - 1]
            self.current_board_idx = 0
            self._update_board_list()
            self.refresh_table()

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        b = self._current_board()
        for r in b.rows:
            vals = [
                r.num, r.name, r.group_num, r.floor, f"{r.power:.2f}", r.voltage,
                f"{r.cos_phi:.2f}", f"{r.current:.1f}", r.n_cables, r.brand,
                r.cores, r.section,
                f"{r.len_tray:.1f}", f"{r.len_corrug:.1f}", f"{r.len_channel:.1f}",
                f"{r.len_pvc:.1f}", f"{r.len_steel:.1f}", f"{r.len_sum:.1f}",
                r.material, f"{r.len_line:.1f}", f"{r.len_reduced:.1f}",
                f"{r.dU:.2f}", f"{r.dU_pct:.2f}", f"{r.Isc:.0f}"
            ]
            self.tree.insert("", "end", values=vals)

    def add_row(self):
        r = CableRow()
        r.num = str(len(self._current_board().rows) + 1)
        self._current_board().rows.append(r)
        self.refresh_table()

    def del_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        pos = self.tree.index(sel[0])
        del self._current_board().rows[pos]
        self.refresh_table()

    def edit_row(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        pos = self.tree.index(sel[0])
        r = self._current_board().rows[pos]

        win = tk.Toplevel(self)
        win.title(f"Редактирование строки {r.num}")
        win.geometry("420x520")
        fields = [
            ("num", "№"), ("name", "Наименование"), ("group_num", "№ группы"),
            ("floor", "Этаж"), ("power", "Мощность, кВт"), ("voltage", "Напряжение (230/400)"),
            ("cos_phi", "cos φ"), ("n_cables", "Кол-во кабелей"), ("brand", "Марка"),
            ("cores", "Кол-во жил"), ("section", "Сечение, мм²"),
            ("len_tray", "Длина в лотке"), ("len_corrug", "в гофре"),
            ("len_channel", "в канале"), ("len_pvc", "в ПВХ"), ("len_steel", "в стали"),
            ("material", "Материал (медь/алюм)"), ("len_line", "Длина линии"),
            ("len_reduced", "Приведённая длина"),
        ]
        vars_ = {}
        for i, (attr, label) in enumerate(fields):
            ttk.Label(win, text=label).grid(row=i, column=0, sticky="e", padx=4, pady=1)
            v = tk.StringVar(value=str(getattr(r, attr)))
            vars_[attr] = v
            ttk.Entry(win, textvariable=v, width=28).grid(row=i, column=1, padx=4, pady=1)

        def save():
            for attr, v in vars_.items():
                val = v.get()
                old = getattr(r, attr)
                if isinstance(old, float):
                    try:
                        setattr(r, attr, float(val.replace(",", ".")))
                    except ValueError:
                        pass
                elif isinstance(old, int):
                    try:
                        setattr(r, attr, int(float(val)))
                    except ValueError:
                        pass
                else:
                    setattr(r, attr, val)
            self.refresh_table()
            win.destroy()

        ttk.Button(win, text="Сохранить", command=save).grid(row=len(fields), column=0, columnspan=2, pady=8)

    def _sync_transformer(self):
        for k, v in self.tr_vars.items():
            self.project.transformer[k] = v.get()
        self.project.transformer["scheme"] = self.scheme_var.get()

    def _recalc_board(self, board):
        self._sync_transformer()
        tr = self.project.transformer
        R_tr, X_tr = calc_transformer_Z(tr["S_kVA"], tr["uk_pct"], tr["Pk_W"], tr["U_lv"])
        for r in board.rows:
            phases = 3 if r.voltage >= 380 else 1
            r.current = calc_current(r.power, r.voltage, r.cos_phi, phases)
            r.len_sum = r.len_tray + r.len_corrug + r.len_channel + r.len_pvc + r.len_steel
            L = r.len_reduced if r.len_reduced > 0 else (r.len_line if r.len_line > 0 else r.len_sum)
            r.dU, r.dU_pct = calc_voltage_drop(r.current, L, r.section, r.cos_phi, r.material, phases)
            rho = RHO_CU if str(r.material).lower().startswith("м") else RHO_AL
            R_cab = rho * L / r.section if r.section > 0 else 0
            X_cab = X0_CABLE * L / 1000
            r.Isc = calc_sc_current_1ph(230.0, R_tr, X_tr, R_cab, X_cab)

    def recalc_current(self):
        self._recalc_board(self._current_board())
        self.refresh_table()

    def recalc_all(self):
        for b in self.project.all_boards():
            self._recalc_board(b)
        self.refresh_table()
        messagebox.showinfo("Готово", "Все щиты пересчитаны")

    def show_cable_summary(self):
        summary = {}
        for board in self.project.all_boards():
            for r in board.rows:
                key = (r.brand, f"{r.cores}x{r.section}")
                if key not in summary:
                    summary[key] = {"tray": 0.0, "corrug": 0.0, "channel": 0.0, "pvc": 0.0, "steel": 0.0, "total": 0.0}
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
                brand, sec, f"{v['tray']:.1f}", f"{v['corrug']:.1f}", f"{v['channel']:.1f}",
                f"{v['pvc']:.1f}", f"{v['steel']:.1f}", f"{v['total']:.1f}"
            ])
        tree.pack(fill="both", expand=True)

    def new_project(self):
        if messagebox.askyesno("Новый", "Создать новый проект?"):
            self.project = Project()
            self.current_board_idx = 0
            self._update_board_list()
            self.refresh_table()

    def save_project(self):
        self._sync_transformer()
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.project.to_dict(), f, ensure_ascii=False, indent=2)
        messagebox.showinfo("Сохранено", path)

    def load_project(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.project = Project.from_dict(data)
        self.current_board_idx = 0
        self._update_board_list()
        for k, var in self.tr_vars.items():
            var.set(self.project.transformer.get(k, 0))
        self.scheme_var.set(self.project.transformer.get("scheme", "D/Yn"))
        self.refresh_table()
        messagebox.showinfo("Загружено", path)

    def export_excel(self):
        if openpyxl is None:
            messagebox.showerror("Ошибка", "Установите openpyxl")
            return
        self._sync_transformer()
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return

        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill("solid", fgColor="4472C4")
        header_font = Font(bold=True, color="FFFFFF")

        headers = [
            "№", "Наименование", "№ группы", "Этаж", "Мощность, кВт", "Напряжение, В",
            "cos φ", "Ток, А", "Кол-во кабелей", "Марка", "Кол-во жил", "Сечение, мм²",
            "в лотке", "в гофре", "в канале", "в ПВХ", "в стальной трубе",
            "Сумма длин", "Материал", "Длина линии", "Приведенная длина",
            "ΔU, В", "ΔU, %", "Iкз1, А"
        ]

        for board in self.project.all_boards():
            ws = wb.create_sheet(title=board.title[:31])
            for col, h in enumerate(headers, 1):
                cell = ws.cell(1, col, h)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = thin
                cell.alignment = Alignment(wrap_text=True, horizontal="center")
            for row_idx, r in enumerate(board.rows, 2):
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
            for col in range(1, len(headers) + 1):
                ws.column_dimensions[get_column_letter(col)].width = 11

        ws = wb.create_sheet("Ведомость кабелей")
        h2 = ["Марка кабеля", "жил×сечение", "в лотке", "в гофре", "в канале", "в ПВХ-трубе", "в стальной трубе", "Итого длина"]
        for col, h in enumerate(h2, 1):
            cell = ws.cell(1, col, h)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin
        summary = {}
        for board in self.project.all_boards():
            for r in board.rows:
                key = (r.brand, f"{r.cores}x{r.section}")
                if key not in summary:
                    summary[key] = [0.0] * 6
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
                ws.cell(row_idx, col, round(v, 1)).border = thin

        ws = wb.create_sheet("Трансформатор")
        ws["A1"] = "Параметр"
        ws["B1"] = "Значение"
        for i, (name, key) in enumerate([
            ("Мощность, кВА", "S_kVA"), ("uk, %", "uk_pct"),
            ("Потери КЗ, Вт", "Pk_W"), ("Uнн, кВ", "U_lv"), ("Схема", "scheme")
        ], 2):
            ws.cell(i, 1, name)
            ws.cell(i, 2, self.project.transformer[key])

        wb.save(path)
        messagebox.showinfo("Экспорт", f"Excel сохранён:\n{path}")

    def export_schemes_pdf(self):
        if pdf_canvas is None:
            messagebox.showerror("Ошибка", "Установите reportlab: pip install reportlab")
            return
        self.recalc_all()
        folder = filedialog.askdirectory(title="Папка для PDF-схем")
        if not folder:
            return
        count = 0
        for board in self.project.all_boards():
            safe = "".join(c if c.isalnum() or c in " -_" else "_" for c in board.title)
            path = os.path.join(folder, f"Схема_{safe}.pdf")
            try:
                draw_single_line_scheme_pdf(board, self.project, path)
                count += 1
            except Exception as e:
                messagebox.showerror("Ошибка", f"{board.title}: {e}")
        messagebox.showinfo("Готово", f"Создано PDF-схем: {count}\n{folder}")

    def export_schemes_dxf(self):
        if ezdxf is None:
            messagebox.showerror("Ошибка", "Установите ezdxf: pip install ezdxf")
            return
        self.recalc_all()
        folder = filedialog.askdirectory(title="Папка для DXF-схем (откройте в AutoCAD → Save as DWG)")
        if not folder:
            return
        count = 0
        for board in self.project.all_boards():
            safe = "".join(c if c.isalnum() or c in " -_" else "_" for c in board.title)
            path = os.path.join(folder, f"Схема_{safe}.dxf")
            try:
                draw_single_line_scheme_dxf(board, self.project, path)
                count += 1
            except Exception as e:
                messagebox.showerror("Ошибка", f"{board.title}: {e}")
        messagebox.showinfo(
            "Готово",
            f"Создано DXF-схем: {count}\n{folder}\n\n"
            "Откройте файл в AutoCAD / nanoCAD / BricsCAD\n"
            "и сохраните как .DWG (Файл → Сохранить как → DWG)."
        )

if __name__ == "__main__":
    app = App()
    app.mainloop()
