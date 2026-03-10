"""
Stock Valuation Desktop App
pip install customtkinter matplotlib requests beautifulsoup4 openpyxl
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import threading, sys, os, requests, math
from datetime import datetime, timedelta

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.patches as mpatches
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from stock_valuation import (
    ddm_price, pe_price, riv_price, dcf_price,
    get_sector_pe, update_summary_sheet,
    moex_dividends, moex_price, moex_name, fetch_smartlab,
    safe, _load_dotenv,
)
_load_dotenv()

try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

# ── Цвета ──────────────────────────────────────────────────────
BG      = "#0D1117"
CARD    = "#161B22"
CARD2   = "#1C2128"
BORDER  = "#30363D"
GREEN   = "#3FB950"
GREEN_D = "#0D2818"
RED     = "#F85149"
RED_D   = "#2D1017"
GOLD    = "#E3B341"
GOLD_D  = "#2D2008"
TEXT    = "#E6EDF3"
MUTED   = "#8B949E"
ACCENT  = "#58A6FF"
PURPLE  = "#BC8CFF"

MPL = {
    "figure.facecolor": CARD, "axes.facecolor": CARD2,
    "axes.edgecolor": BORDER, "axes.labelcolor": MUTED,
    "xtick.color": MUTED, "ytick.color": MUTED,
    "grid.color": BORDER, "grid.alpha": 0.5,
    "text.color": TEXT, "lines.linewidth": 2,
}
plt.rcParams.update(MPL)
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# ── Хелперы ────────────────────────────────────────────────────
def lbl(parent, text, size=12, weight="normal", color=TEXT, **kw):
    return ctk.CTkLabel(parent, text=text,
        font=ctk.CTkFont(family="SF Pro Display", size=size, weight=weight),
        text_color=color, **kw)

def inp(parent, width=140, ph="", **kw):
    return ctk.CTkEntry(parent, width=width, placeholder_text=ph,
        fg_color=CARD2, border_color=BORDER, text_color=TEXT,
        placeholder_text_color=MUTED, font=ctk.CTkFont(size=13), **kw)

def btn(parent, text, cmd, color=ACCENT, width=120, height=34, **kw):
    return ctk.CTkButton(parent, text=text, command=cmd,
        fg_color=color, hover_color=BORDER, text_color=TEXT,
        font=ctk.CTkFont(size=13, weight="bold"),
        corner_radius=8, width=width, height=height, **kw)

class Card(ctk.CTkFrame):
    def __init__(self, p, **kw):
        super().__init__(p, fg_color=CARD, corner_radius=12,
                         border_width=1, border_color=BORDER, **kw)

class Div(ctk.CTkFrame):
    def __init__(self, p, **kw):
        super().__init__(p, fg_color=BORDER, height=1, **kw)


# ═══════════════════════════════════════════════════════════════
#  Всплывающая подсказка для карточек моделей
# ═══════════════════════════════════════════════════════════════
MODEL_INFO = {
    "ddm": ("DDM — Gordon Growth Model",
            "P = D₁ / (k − g)\n\n"
            "D₁ — ожидаемый дивиденд\n"
            "k  — ставка дисконтирования\n"
            "g  — темп роста дивидендов\n\n"
            "Работает только если k − g ≥ 5%.\n"
            "Подходит для зрелых дивидендных компаний."),
    "pe":  ("P/E — Отраслевой метод",
            "P = EPS × P/E_отрасли\n\n"
            "EPS — прибыль на акцию\n"
            "P/E_отрасли — медианный мультипликатор\n"
            "              по сектору компании\n\n"
            "Показывает цену относительно\n"
            "среднерыночной оценки аналогов."),
    "riv": ("RIV — Residual Income Valuation",
            "P = BVPS + Σ (ROE − k) × BVPS / (1+k)ᵗ\n\n"
            "BVPS — балансовая стоимость на акцию\n"
            "ROE  — рентабельность капитала\n"
            "k    — ставка дисконтирования\n\n"
            "Учитывает создание стоимости\n"
            "сверх требуемой доходности."),
    "dcf": ("DCF — Дисконтированные потоки",
            "P = Σ EPS×(1+g)ᵗ/(1+k)ᵗ + TV/(1+k)ᵀ\n\n"
            "T  — горизонт прогноза (7 лет)\n"
            "TV — терминальная стоимость\n"
            "g_term = min(g, 4%)\n\n"
            "Наиболее полная модель: учитывает\n"
            "рост прибыли и долгосрочную стоимость."),
}

class TooltipWindow(tk.Toplevel):
    def __init__(self, parent, title, body):
        super().__init__(parent)
        self.overrideredirect(True)
        self.configure(bg=CARD2)
        self.attributes("-topmost", True)

        frame = tk.Frame(self, bg=CARD2, padx=16, pady=12)
        frame.pack()

        tk.Label(frame, text=title, bg=CARD2, fg=ACCENT,
                 font=("SF Pro Display", 13, "bold")).pack(anchor="w")
        tk.Frame(frame, bg=BORDER, height=1).pack(fill="x", pady=8)
        tk.Label(frame, text=body, bg=CARD2, fg=TEXT,
                 font=("SF Pro Mono", 11), justify="left").pack(anchor="w")

        # Позиционируем рядом с курсором
        x = parent.winfo_pointerx() + 12
        y = parent.winfo_pointery() + 12
        self.geometry(f"+{x}+{y}")

        self.bind("<Button-1>", lambda e: self.destroy())
        self.after(6000, self.destroy)


# ═══════════════════════════════════════════════════════════════
#  Страница 1 — Оценка
# ═══════════════════════════════════════════════════════════════
class ValuationPage(ctk.CTkFrame):
    def __init__(self, parent, app, **kw):
        super().__init__(parent, fg_color=BG, **kw)
        self.app = app
        self.data = None
        self.pdf_path = None
        self.portfolio = {}
        self._loaded_price = 0
        self._loaded_name  = ""
        self._loaded_d0    = 0
        self._tooltip_win  = None
        self._spinner_angle = 0
        self._spinning = False
        self._build()

    def _build(self):
        self.grid_columnconfigure(0, weight=0, minsize=300)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self._build_left()
        self._build_results()
        self._build_portfolio()

    # ── Левая панель ───────────────────────────────────────────
    def _build_left(self):
        p = Card(self)
        p.grid(row=0, column=0, rowspan=2, padx=(16,8), pady=16, sticky="nsew")
        p.grid_columnconfigure(0, weight=1)

        lbl(p, "Тикер акции", size=11, color=MUTED).grid(
            row=0, padx=16, pady=(16,4), sticky="w")

        row0 = ctk.CTkFrame(p, fg_color="transparent")
        row0.grid(row=1, padx=16, sticky="ew")
        row0.grid_columnconfigure(0, weight=1)

        self.ticker_e = inp(row0, ph="SBER, LKOH, AAPL…")
        self.ticker_e.grid(row=0, column=0, sticky="ew", padx=(0,8))
        self.ticker_e.bind("<Return>", lambda e: self._load_data())

        # Кнопка с spinner
        self.load_btn = btn(row0, "Загрузить", self._load_data,
                            color=ACCENT, width=100)
        self.load_btn.grid(row=0, column=1)

        # Spinner canvas
        self.spin_canvas = tk.Canvas(p, width=20, height=20,
                                     bg=CARD, highlightthickness=0)
        self.spin_canvas.grid(row=2, padx=16, pady=(6,0), sticky="w")

        self.pdf_lbl = lbl(p, "PDF: не выбран", size=10, color=MUTED)
        self.pdf_lbl.grid(row=3, padx=16, pady=(4,0), sticky="w")
        btn(p, "Прикрепить PDF", self._pick_pdf,
            color=CARD2, width=200).grid(row=4, padx=16, pady=(4,0), sticky="w")

        Div(p).grid(row=5, padx=16, pady=10, sticky="ew")
        lbl(p, "Фундаментальные данные", size=11, color=MUTED).grid(
            row=6, padx=16, pady=(0,8), sticky="w")

        self.fields = {}
        for i, (key, lb, ph) in enumerate([
            ("eps",  "EPS, ₽",           "Прибыль на акцию"),
            ("bvps", "BVPS, ₽",          "Балансовая стоимость / акция"),
            ("roe",  "ROE (0.20 = 20%)", "Рентабельность капитала"),
            ("g",    "g (0.08 = 8%)",    "Рост дивидендов"),
            ("beta", "β (бета)",          "Бета акции"),
        ]):
            r = 7 + i * 2
            lbl(p, lb, size=11, color=MUTED).grid(
                row=r, padx=16, pady=(4,0), sticky="w")
            e = inp(p, ph=ph)
            e.grid(row=r+1, padx=16, pady=(2,0), sticky="ew")
            self.fields[key] = e

        Div(p).grid(row=17, padx=16, pady=10, sticky="ew")

        self.calc_btn = btn(p, "Рассчитать", self._calculate,
                            color="#238636", width=200, height=40)
        self.calc_btn.grid(row=18, padx=16, pady=(0,4), sticky="ew")

        # Блок добавления в портфель
        add_frame = ctk.CTkFrame(p, fg_color=CARD2, corner_radius=8)
        add_frame.grid(row=19, padx=16, pady=(4,0), sticky="ew")
        add_frame.grid_columnconfigure(0, weight=1)

        lbl(add_frame, "Доля в портфеле, %", size=11,
            color=MUTED).grid(row=0, padx=12, pady=(8,2), sticky="w")
        self.weight_e = inp(add_frame, ph="например 10")
        self.weight_e.grid(row=1, padx=12, pady=(0,4), sticky="ew")
        btn(add_frame, "В портфель", self._add_to_portfolio,
            color="#1C4A2A", width=200).grid(
            row=2, padx=12, pady=(0,8), sticky="ew")

        btn(p, "Экспорт Excel", self._export_excel,
            color=CARD2, width=200).grid(
            row=20, padx=16, pady=(8,4), sticky="ew")

        self.status_lbl = lbl(p, "", size=11, color=MUTED)
        self.status_lbl.grid(row=21, padx=16, pady=(4,12), sticky="w")

    # ── Карточки результатов ───────────────────────────────────
    def _build_results(self):
        f = ctk.CTkFrame(self, fg_color="transparent")
        f.grid(row=0, column=1, padx=(8,16), pady=(16,8), sticky="ew")
        f.grid_columnconfigure((0,1,2,3,4), weight=1)

        self.model_cards = {}
        models = [
            ("ddm", "DDM",  "Gordon Growth Model", ACCENT),
            ("pe",  "P/E",  "Отраслевой метод",    PURPLE),
            ("riv", "RIV",  "Остаточный доход",    GOLD),
            ("dcf", "DCF",  "Дисконт. потоки",     "#79C0FF"),
        ]
        for col, (key, name, sub, color) in enumerate(models):
            card = Card(f)
            card.grid(row=0, column=col, padx=4, sticky="nsew")
            card.grid_columnconfigure(0, weight=1)
            card.configure(cursor="hand2")

            lbl(card, name, size=13, weight="bold", color=color).grid(
                row=0, padx=14, pady=(12,0), sticky="w")
            lbl(card, sub, size=10, color=MUTED).grid(
                row=1, padx=14, pady=(0,4), sticky="w")
            vl = lbl(card, "—", size=22, weight="bold", color=TEXT)
            vl.grid(row=2, padx=14, pady=(0,12), sticky="w")

            # Клик — показываем формулу
            for widget in (card, vl):
                widget.bind("<Button-1>",
                    lambda e, k=key: self._show_model_info(k))
            lbl(card, "нажми для формулы", size=9,
                color=MUTED).grid(row=3, padx=14, pady=(0,10), sticky="w")

            self.model_cards[key] = vl

        # Итог
        self.fair_card = Card(f)
        self.fair_card.grid(row=0, column=4, padx=(8,0), sticky="nsew")
        self.fair_card.grid_columnconfigure(0, weight=1)

        lbl(self.fair_card, "Справедливая цена", size=13,
            weight="bold", color=GREEN).grid(
            row=0, padx=14, pady=(12,0), sticky="w")
        lbl(self.fair_card, "Среднее моделей", size=10,
            color=MUTED).grid(row=1, padx=14, pady=(0,4), sticky="w")
        self.fair_val = lbl(self.fair_card, "—", size=26,
                             weight="bold", color=GREEN)
        self.fair_val.grid(row=2, padx=14, pady=(0,4), sticky="w")
        self.upside_lbl = lbl(self.fair_card, "", size=14, color=MUTED)
        self.upside_lbl.grid(row=3, padx=14, pady=(0,10), sticky="w")
        self.models_used = lbl(self.fair_card, "", size=10, color=MUTED)
        self.models_used.grid(row=4, padx=14, pady=(0,12), sticky="w")

    # ── Таблица портфеля ───────────────────────────────────────
    def _build_portfolio(self):
        f = ctk.CTkFrame(self, fg_color="transparent")
        f.grid(row=1, column=1, padx=(8,16), pady=(0,16), sticky="nsew")
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(2, weight=1)

        lbl(f, "Портфель", size=13, weight="bold").grid(
            row=0, pady=(0,8), sticky="w")

        COLS   = ["Тикер","Компания","Доля %","Цена","DDM","P/E","RIV","DCF","Справедл.","Потенц.%",""]
        WIDTHS = [65, 180, 65, 75, 85, 85, 85, 85, 90, 85, 36]

        hdr = Card(f)
        hdr.grid(row=1, sticky="ew")
        for ci, (c, w) in enumerate(zip(COLS, WIDTHS)):
            lbl(hdr, c, size=11, color=MUTED).grid(
                row=0, column=ci, padx=6, pady=6, sticky="w")
            hdr.grid_columnconfigure(ci, minsize=w)

        self.tbl = ctk.CTkScrollableFrame(
            f, fg_color=CARD, corner_radius=0,
            border_width=1, border_color=BORDER)
        self.tbl.grid(row=2, sticky="nsew")
        for ci, w in enumerate(WIDTHS):
            self.tbl.grid_columnconfigure(ci, minsize=w)

    def _refresh_table(self):
        for w in self.tbl.winfo_children():
            w.destroy()

        sorted_items = sorted(
            self.portfolio.items(),
            key=lambda x: x[1].get("upside", 0), reverse=True)

        for ri, (ticker, d) in enumerate(sorted_items):
            bg = CARD if ri % 2 == 0 else CARD2
            up = d.get("upside", 0)
            uc = GREEN if up > 0 else (RED if up < 0 else MUTED)
            wt = d.get("weight", "—")
            wt_str = f"{wt}%" if isinstance(wt, (int,float)) else "—"

            vals = [
                ticker,
                d.get("name","")[:22],
                wt_str,
                f"{d['price']:.2f}",
                f"{d['ddm']:.2f}" if d['ddm'] else "—",
                f"{d['cmp']:.2f}" if d['cmp'] else "—",
                f"{d['riv']:.2f}" if d['riv'] else "—",
                f"{d['dcf']:.2f}" if d['dcf'] else "—",
                f"{d['avg']:.2f}" if d['avg'] else "—",
                f"{up:+.1f}%",
                "✕",
            ]
            for ci, val in enumerate(vals):
                color = uc if ci == 9 else (RED if ci == 10 else TEXT)
                if ci == 2:  # Доля — редактируемое поле
                    e = ctk.CTkEntry(self.tbl, width=55, height=24,
                        fg_color=CARD2, border_color=BORDER,
                        text_color=TEXT, font=ctk.CTkFont(size=12))
                    e.insert(0, wt_str.replace("%",""))
                    e.grid(row=ri, column=ci, padx=4, pady=2, sticky="w")
                    e.bind("<Return>",
                        lambda ev, t=ticker, en=e: self._update_weight(t, en))
                    e.bind("<FocusOut>",
                        lambda ev, t=ticker, en=e: self._update_weight(t, en))
                elif ci == 10:  # Кнопка удаления
                    b = ctk.CTkButton(
                        self.tbl, text="✕", width=28, height=24,
                        fg_color="transparent", hover_color=RED_D,
                        text_color=MUTED, font=ctk.CTkFont(size=11),
                        command=lambda t=ticker: self._remove(t))
                    b.grid(row=ri, column=ci, padx=4, pady=2)
                else:
                    cl = ctk.CTkLabel(
                        self.tbl, text=val,
                        font=ctk.CTkFont(size=12),
                        text_color=color, fg_color=bg,
                        anchor="w", padx=6)
                    cl.grid(row=ri, column=ci, sticky="ew", ipady=4)

    def _update_weight(self, ticker, entry_widget):
        try:
            val = float(entry_widget.get().replace("%","").strip())
            self.portfolio[ticker]["weight"] = val
        except ValueError:
            pass

    def _remove(self, ticker):
        self.portfolio.pop(ticker, None)
        self._refresh_table()

    # ── Spinner ────────────────────────────────────────────────
    def _start_spinner(self):
        self._spinning = True
        self._spinner_angle = 0
        self._tick_spinner()

    def _tick_spinner(self):
        if not self._spinning:
            self.spin_canvas.delete("all")
            return
        c = self.spin_canvas
        c.delete("all")
        cx, cy, r = 10, 10, 7
        for i in range(8):
            a = math.radians(self._spinner_angle + i * 45)
            alpha = int(255 * (i + 1) / 8)
            color = f"#{alpha:02x}{alpha:02x}{alpha:02x}"
            x1 = cx + (r-2) * math.cos(a)
            y1 = cy + (r-2) * math.sin(a)
            x2 = cx + r * math.cos(a)
            y2 = cy + r * math.sin(a)
            c.create_line(x1,y1,x2,y2, fill=color, width=2)
        self._spinner_angle = (self._spinner_angle + 15) % 360
        self.after(50, self._tick_spinner)

    def _stop_spinner(self):
        self._spinning = False

    # ── Подсказка с формулой ───────────────────────────────────
    def _show_model_info(self, key):
        if self._tooltip_win:
            try:
                self._tooltip_win.destroy()
            except Exception:
                pass
        title, body = MODEL_INFO[key]
        self._tooltip_win = TooltipWindow(self, title, body)

    # ── Статус ─────────────────────────────────────────────────
    def _status(self, msg, color=MUTED):
        self.status_lbl.configure(text=msg, text_color=color)

    def _pick_pdf(self):
        p = filedialog.askopenfilename(
            filetypes=[("PDF", "*.pdf")], title="Выбери МСФО отчёт")
        if p:
            self.pdf_path = p
            self.pdf_lbl.configure(
                text=f"PDF: {os.path.basename(p)}", text_color=GREEN)

    def _load_data(self):
        ticker = self.ticker_e.get().strip().upper()
        if not ticker:
            self._status("Введи тикер", RED)
            return
        self.load_btn.configure(state="disabled", text="…")
        self._status("Загружаю…", GOLD)
        self._start_spinner()

        def worker():
            try:
                is_moex = not ("." in ticker) or ticker.endswith(".ME")
                tick = ticker.replace(".ME", "")
                if is_moex:
                    price = moex_price(tick)
                    name  = moex_name(tick)
                    sl    = fetch_smartlab(tick)
                    divs  = moex_dividends(tick)
                    if divs:
                        cutoff = datetime.now().replace(
                            year=datetime.now().year-1).strftime("%Y-%m-%d")
                        recent = [v for d,v in divs if d >= cutoff]
                        d0 = sum(recent) if recent else sum(v for _,v in divs[-4:])
                    else:
                        d0 = 0.0
                    self.after(0, lambda: self._fill({
                        "price":price,"name":name,"d0":d0,**sl}))
                else:
                    import yfinance as yf
                    t = yf.Ticker(ticker)
                    info = t.info
                    hist = t.history(period="5d")
                    price = float(hist["Close"].iloc[-1]) if not hist.empty else 0
                    self.after(0, lambda: self._fill({
                        "price":price,
                        "name":info.get("longName",ticker),
                        "eps": safe(info.get("trailingEps")),
                        "bvps":safe(info.get("bookValue")),
                        "roe": safe(info.get("returnOnEquity"),0.15),
                        "g":   0.05,
                        "beta":safe(info.get("beta"),1.0),
                    }))
            except Exception as e:
                self.after(0, lambda: self._status(f"Ошибка: {e}", RED))
            finally:
                self.after(0, self._stop_spinner)
                self.after(0, lambda: self.load_btn.configure(
                    state="normal", text="Загрузить"))

        threading.Thread(target=worker, daemon=True).start()

    def _fill(self, d):
        self._loaded_price = d.get("price", 0)
        self._loaded_name  = d.get("name", "")
        self._loaded_d0    = d.get("d0", 0)
        self._status(
            f"✓ {d.get('name','')}  Цена: {d.get('price',0):.2f} ₽", GREEN)
        for key in ("eps","bvps","roe","g","beta"):
            val = d.get(key)
            if val:
                self.fields[key].delete(0, "end")
                self.fields[key].insert(0, str(round(float(val), 4)))
        ticker = self.ticker_e.get().strip().upper().replace(".ME","")
        self.app.analytics_page.load_ticker(ticker)

    def _fval(self, key, default=0.0):
        try:
            return float(self.fields[key].get()) or default
        except (ValueError, TypeError):
            return default

    def _calculate(self):
        ticker = self.ticker_e.get().strip().upper().replace(".ME","")
        if not ticker or self._loaded_price == 0:
            self._status("Сначала загрузи данные", RED)
            return
        price = self._loaded_price
        d0    = self._loaded_d0
        name  = self._loaded_name

        eps  = self._fval("eps")
        bvps = self._fval("bvps")
        roe  = self._fval("roe", 0.15)
        g    = self._fval("g",   0.08)
        beta = self._fval("beta", 1.0)

        r_f    = 0.16; r_m = 0.22
        r_capm = r_f + beta * (r_m - r_f)
        d1     = d0 * (1 + g)
        r_ddm  = (d1/price + g) if price and d1 else r_capm
        r_ddm  = max(min(r_ddm, 0.60), r_f)
        r_avg  = (r_capm + r_ddm) / 2
        ri     = (roe - r_avg) * bvps
        pv_ri  = sum(ri/(1+r_avg)**t for t in range(1,6)) if r_avg else 0

        d = dict(ticker=ticker, name=name, price=price,
                 d0=d0, d1=d1, g=g, k=r_avg,
                 eps=eps, pe=(price/eps if eps>0 else 0),
                 bvps=bvps, pv_ri=pv_ri, roe=roe,
                 beta=beta, r_f=r_f, r_m=r_m,
                 r_capm=r_capm, r_ddm=r_ddm, r_avg=r_avg,
                 currency="₽")

        ddm = ddm_price(d); cmp = pe_price(d)
        riv = riv_price(d); dcf = dcf_price(d)
        models = [v for v in [ddm,cmp,riv,dcf] if v > 0]
        avg    = sum(models)/len(models) if models else 0
        upside = (avg/price - 1)*100 if price and avg else 0

        for key, val in [("ddm",ddm),("pe",cmp),("riv",riv),("dcf",dcf)]:
            self.model_cards[key].configure(
                text=f"{val:.2f} ₽" if val > 0 else "—")

        self.fair_val.configure(text=f"{avg:.2f} ₽" if avg else "—")
        self.models_used.configure(text=f"из {len(models)} моделей")

        if upside > 0:
            self.upside_lbl.configure(
                text=f"▲ +{upside:.1f}% потенциал", text_color=GREEN)
            self.fair_card.configure(border_color=GREEN)
        elif upside < 0:
            self.upside_lbl.configure(
                text=f"▼ {upside:.1f}% переоценена", text_color=RED)
            self.fair_card.configure(border_color=RED)

        self.data = {**d, "ddm":ddm,"cmp":cmp,"riv":riv,"dcf":dcf,
                     "avg":avg,"upside":upside}
        self._status(f"✓ k={r_avg*100:.1f}%  CAPM={r_capm*100:.1f}%", GREEN)

    def _add_to_portfolio(self):
        if not self.data:
            self._status("Сначала рассчитай", RED)
            return
        try:
            w = float(self.weight_e.get().strip()) if self.weight_e.get().strip() else None
        except ValueError:
            w = None
        self.data["weight"] = w
        self.portfolio[self.data["ticker"]] = dict(self.data)
        self._refresh_table()
        self._status(f"✓ {self.data['ticker']} добавлен в портфель", GREEN)

    def _export_excel(self):
        if not self.portfolio and not self.data:
            self._status("Нет данных", RED); return
        if not OPENPYXL_OK:
            self._status("pip install openpyxl", RED); return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile="portfolio.xlsx",
            filetypes=[("Excel","*.xlsx")])
        if not path: return
        import openpyxl
        wb = openpyxl.load_workbook(path) if os.path.exists(path) \
             else openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        items = {**self.portfolio}
        if self.data:
            items[self.data["ticker"]] = self.data
        for d in items.values():
            update_summary_sheet(wb, d)
        wb.save(path)
        self._status(f"✓ {os.path.basename(path)}", GREEN)


# ═══════════════════════════════════════════════════════════════
#  Страница 2 — Аналитика
# ═══════════════════════════════════════════════════════════════
PERIODS = {
    "1Д":   (1,    "60"),
    "1Н":   (7,    "60"),
    "1М":   (31,   "24"),
    "3М":   (90,   "24"),
    "1Г":   (365,  "24"),
    "3Г":   (1095, "7"),
    "5Л":   (1825, "7"),
    "Всё":  (0,    "7"),
}

class AnalyticsPage(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=BG, **kw)
        self.current_ticker = None
        self._all_dates  = []
        self._all_closes = []
        self._divs_cache = []
        self._annot_price = None
        self._vline       = None
        self._pin_idx        = None   # индекс закреплённой точки (клик)
        self._pin_marker     = None   # маркер на графике
        self._price_span_patch = None # зелёный коридор между точками
        self._build()

    def _build(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=3)
        self.grid_rowconfigure(2, weight=2)

        # ── Шапка с периодами ──────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.grid(row=0, column=0, padx=16, pady=(14,6), sticky="ew")

        self.title_lbl = lbl(hdr, "Аналитика — загрузи тикер на странице «Оценка»",
                             size=15, weight="bold")
        self.title_lbl.pack(side="left")

        period_frame = ctk.CTkFrame(hdr, fg_color="transparent")
        period_frame.pack(side="right")
        lbl(period_frame, "Период:", size=11, color=MUTED).pack(
            side="left", padx=(0,8))

        self._period_btns = {}
        self._active_period = "1Г"
        for p in PERIODS:
            b = ctk.CTkButton(
                period_frame, text=p, width=42, height=28,
                fg_color=CARD2 if p != "1Г" else ACCENT,
                hover_color=BORDER, text_color=TEXT,
                font=ctk.CTkFont(size=12), corner_radius=6,
                command=lambda pp=p: self._set_period(pp))
            b.pack(side="left", padx=2)
            self._period_btns[p] = b

        # ── График цены ────────────────────────────────────────
        pc = Card(self)
        pc.grid(row=1, column=0, padx=16, pady=(0,6), sticky="nsew")
        pc.grid_columnconfigure(0, weight=1)
        pc.grid_rowconfigure(1, weight=1)

        price_hdr = ctk.CTkFrame(pc, fg_color="transparent")
        price_hdr.grid(row=0, padx=16, pady=(12,4), sticky="ew")
        lbl(price_hdr, "График цены", size=13,
            weight="bold", color=ACCENT).pack(side="left")
        self.price_info = lbl(price_hdr, "", size=12, color=MUTED)
        self.price_info.pack(side="right")

        self.price_fig = Figure(figsize=(10, 3.5), dpi=96)
        self.price_fig.patch.set_facecolor(CARD)
        self.price_ax  = self.price_fig.add_subplot(111)
        self.price_canvas = FigureCanvasTkAgg(self.price_fig, master=pc)
        self.price_canvas.get_tk_widget().grid(
            row=1, padx=8, pady=(0,10), sticky="nsew")
        self._draw_empty(self.price_ax, "Загрузи тикер на странице «Оценка»")
        self.price_canvas.draw()

        # Интерактивность — движение мыши
        self.price_canvas.mpl_connect("motion_notify_event", self._on_price_hover)
        self.price_canvas.mpl_connect("axes_leave_event",    self._on_price_leave)
        self.price_canvas.mpl_connect("button_press_event",  self._on_price_click)

        # ── График дивидендов ──────────────────────────────────
        dc = Card(self)
        dc.grid(row=2, column=0, padx=16, pady=(0,16), sticky="nsew")
        dc.grid_columnconfigure(0, weight=1)
        dc.grid_rowconfigure(1, weight=1)

        div_hdr = ctk.CTkFrame(dc, fg_color="transparent")
        div_hdr.grid(row=0, padx=16, pady=(12,4), sticky="ew")
        lbl(div_hdr, "История дивидендов", size=13,
            weight="bold", color=GOLD).pack(side="left")
        self.div_info = lbl(div_hdr, "", size=12, color=MUTED)
        self.div_info.pack(side="right")

        self.div_fig = Figure(figsize=(10, 2.8), dpi=96)
        self.div_fig.patch.set_facecolor(CARD)
        self.div_ax  = self.div_fig.add_subplot(111)
        self.div_canvas = FigureCanvasTkAgg(self.div_fig, master=dc)
        self.div_canvas.get_tk_widget().grid(
            row=1, padx=8, pady=(0,10), sticky="nsew")
        self._draw_empty(self.div_ax, "Загрузи тикер на странице «Оценка»")
        self.div_canvas.draw()

        self.div_canvas.mpl_connect("motion_notify_event", self._on_div_hover)
        self.div_canvas.mpl_connect("axes_leave_event",    self._on_div_leave)

        # Аннотации (создаём заранее, скрываем)
        self._div_annot = None

    # ── Пустой график ──────────────────────────────────────────
    def _draw_empty(self, ax, msg=""):
        ax.clear()
        ax.set_facecolor(CARD2)
        if msg:
            ax.text(0.5, 0.5, msg, transform=ax.transAxes,
                    ha="center", va="center", color=MUTED, fontsize=12)
        ax.set_xticks([]); ax.set_yticks([])
        for sp in ax.spines.values():
            sp.set_color(BORDER)

    # ── Загрузка данных ────────────────────────────────────────
    def load_ticker(self, ticker: str):
        self.current_ticker = ticker
        self.title_lbl.configure(text=f"Аналитика — {ticker}")
        self._all_dates = []; self._all_closes = []
        threading.Thread(target=self._fetch, args=(ticker,), daemon=True).start()

    def _fetch_candles_all(self, ticker):
        """Грузим всю историю дневных свечей с пагинацией MOEX ISS."""
        dates, closes = [], []
        start = "2015-01-01"
        till  = datetime.now().strftime("%Y-%m-%d")
        while True:
            url = (f"https://iss.moex.com/iss/engines/stock/markets/shares/"
                   f"boards/TQBR/securities/{ticker}/candles.json"
                   f"?from={start}&till={till}"
                   f"&interval=24&iss.meta=off&iss.json=extended&start={len(dates)}")
            r = requests.get(url, timeout=20,
                             headers={"User-Agent": "StockValuator/1.0"})
            r.raise_for_status()
            batch_dates, batch_closes = [], []
            for block in r.json():
                if not isinstance(block, dict): continue
                for row in block.get("candles", []):
                    try:
                        dt = datetime.strptime(row["begin"][:10], "%Y-%m-%d")
                        batch_dates.append(dt)
                        batch_closes.append(float(row["close"]))
                    except Exception:
                        continue
            if not batch_dates:
                break
            dates.extend(batch_dates)
            closes.extend(batch_closes)
            # MOEX отдаёт max 500 свечей за раз; если меньше — конец
            if len(batch_dates) < 500:
                break
        return dates, closes

    def _fetch(self, ticker):
        # ── 1. Цена (все дневные свечи) ────────────────────────
        try:
            dates, closes = self._fetch_candles_all(ticker)
            # Фильтруем: берём только данные за последние 10 лет
            cutoff10 = datetime.now() - timedelta(days=3650)
            pairs = [(d,c) for d,c in zip(dates,closes) if d >= cutoff10]
            if pairs:
                dates, closes = zip(*pairs)
                dates, closes = list(dates), list(closes)
            self._all_dates  = dates
            self._all_closes = closes
            self.after(0, lambda: self._apply_period(self._active_period))
        except Exception as e:
            self._all_dates  = []
            self._all_closes = []
            self.after(0, lambda: (
                self._draw_empty(self.price_ax, f"Ошибка загрузки цены: {e}"),
                self.price_canvas.draw()))

        # ── 2. Дивиденды — рисуем ПОСЛЕ загрузки цен ──────────
        try:
            divs = moex_dividends(ticker)
            self._divs_cache = divs
            self.after(0, lambda: self._draw_divs(divs, ticker))
        except Exception as e:
            self.after(0, lambda: (
                self._draw_empty(self.div_ax, f"Ошибка дивидендов: {e}"),
                self.div_canvas.draw()))

    # ── Переключение периода ───────────────────────────────────
    def _set_period(self, p):
        self._active_period = p
        for k, b in self._period_btns.items():
            b.configure(fg_color=ACCENT if k==p else CARD2)
        self._apply_period(p)

    def _apply_period(self, p):
        if p == "1Д":
            # Интрадей — грузим отдельно минутки за сегодня
            threading.Thread(
                target=self._fetch_intraday,
                args=(self.current_ticker,), daemon=True).start()
            return
        if not self._all_dates:
            return
        days, _ = PERIODS[p]
        if days == 0:
            dates  = list(self._all_dates)
            closes = list(self._all_closes)
        else:
            cutoff = datetime.now() - timedelta(days=days)
            pairs  = [(d,c) for d,c in zip(self._all_dates, self._all_closes)
                      if d >= cutoff]
            if not pairs:
                # Нет данных за период — берём последние доступные точки
                n = min(30, len(self._all_dates))
                dates  = self._all_dates[-n:]
                closes = self._all_closes[-n:]
            else:
                dates, closes = zip(*pairs)
                dates, closes = list(dates), list(closes)
        self._draw_price(dates, closes)

    def _fetch_intraday(self, ticker):
        """Минутные/10-минутные свечи за последние 2 дня."""
        try:
            since = (datetime.now() - timedelta(days=2)).strftime("%Y-%m-%d")
            till  = datetime.now().strftime("%Y-%m-%d")
            url = (f"https://iss.moex.com/iss/engines/stock/markets/shares/"
                   f"boards/TQBR/securities/{ticker}/candles.json"
                   f"?from={since}&till={till}"
                   f"&interval=10&iss.meta=off&iss.json=extended")
            r = requests.get(url, timeout=15,
                             headers={"User-Agent": "StockValuator/1.0"})
            r.raise_for_status()
            dates, closes = [], []
            for block in r.json():
                if not isinstance(block, dict): continue
                for row in block.get("candles", []):
                    try:
                        dt = datetime.strptime(row["begin"][:16], "%Y-%m-%d %H:%M")
                        dates.append(dt)
                        closes.append(float(row["close"]))
                    except Exception:
                        continue
            if dates:
                self.after(0, lambda: self._draw_price(dates, closes))
            else:
                # Fallback: последний день из дневных свечей
                self.after(0, lambda: self._apply_period("1Н"))
        except Exception as e:
            self.after(0, lambda: self._apply_period("1Н"))

    # ── Рисуем цену ────────────────────────────────────────────
    def _draw_price(self, dates, closes):
        ax = self.price_ax
        ax.clear()
        ax.set_facecolor(CARD2)
        self._annot_price    = None
        self._vline          = None
        self._pin_idx        = None
        self._pin_marker     = None
        self._price_span_patch = None

        if not dates:
            self._draw_empty(ax, "Нет данных")
            self.price_canvas.draw()
            return

        # Линия всегда зелёная
        ax.plot(dates, closes, color=GREEN, linewidth=1.8)
        self._price_span_patch = None

        # min / max маркеры
        mn, mx = min(closes), max(closes)
        mni, mxi = closes.index(mn), closes.index(mx)
        ax.scatter([dates[mni]], [mn], color=RED,   s=40, zorder=5)
        ax.scatter([dates[mxi]], [mx], color=GREEN, s=40, zorder=5)
        ax.annotate(f"мин {mn:.0f}", xy=(dates[mni], mn),
            xytext=(0,-20), textcoords="offset points",
            color=RED, fontsize=8, ha="center",
            arrowprops=dict(arrowstyle="-", color=RED, alpha=0.4, lw=0.8))
        ax.annotate(f"макс {mx:.0f}", xy=(dates[mxi], mx),
            xytext=(0,12), textcoords="offset points",
            color=GREEN, fontsize=8, ha="center",
            arrowprops=dict(arrowstyle="-", color=GREEN, alpha=0.4, lw=0.8))

        # Форматирование осей — автоматическое прореживание
        span = (dates[-1] - dates[0]).days
        if span <= 2:
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b %H:%M"))
            ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        elif span <= 14:
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))
            ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
        elif span <= 90:
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))
            ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=2))
        elif span <= 365:
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
            ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))
        elif span <= 365 * 3:
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%b '%y"))
            ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
        else:
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
            ax.xaxis.set_major_locator(mdates.YearLocator())

        self.price_fig.autofmt_xdate(rotation=30, ha="right")
        ax.set_ylabel("Цена, ₽", color=MUTED, fontsize=10)
        ax.grid(True, axis="y", alpha=0.3, linestyle="--")
        for sp in ax.spines.values():
            sp.set_color(BORDER)

        # Аннотация (скрытая)
        self._annot_price = ax.annotate(
            "", xy=(dates[0], closes[0]),
            xytext=(12, 12), textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.5", fc=CARD, ec=ACCENT, lw=1),
            fontsize=10, color=TEXT, zorder=10)
        self._annot_price.set_visible(False)
        self._vline = ax.axvline(x=dates[0], color=ACCENT,
                                  alpha=0.5, linewidth=1, linestyle="--")
        self._vline.set_visible(False)

        self._price_dates  = dates
        self._price_closes = closes

        # Изменение за период
        chg = (closes[-1]/closes[0] - 1)*100 if closes[0] else 0
        chg_color = GREEN if chg >= 0 else RED
        self.price_info.configure(
            text=f"{closes[-1]:.2f} ₽   {chg:+.1f}%", text_color=chg_color)

        self.price_fig.subplots_adjust(top=0.95, bottom=0.18, left=0.08, right=0.97)
        self.price_canvas.draw()

    # ── Hover на графике цены ──────────────────────────────────
    def _on_price_hover(self, event):
        if (event.inaxes != self.price_ax
                or not hasattr(self, "_price_dates")
                or not self._price_dates
                or self._annot_price is None):
            return
        # Найти ближайшую точку
        import bisect
        dates  = self._price_dates
        closes = self._price_closes
        if not dates: return

        # Конвертируем x в datetime
        try:
            x_dt = mdates.num2date(event.xdata).replace(tzinfo=None)
        except Exception:
            return

        xs = [d for d in dates]
        idx = bisect.bisect_left(xs, x_dt)
        idx = max(0, min(idx, len(dates)-1))

        d  = dates[idx]
        cl = closes[idx]

        span = (self._price_dates[-1] - self._price_dates[0]).days
        fmt  = '%d %b %Y  %H:%M' if span <= 7 else '%d %b %Y'
        
        # Режим сравнения двух точек
        if self._pin_idx is not None:
            # Всегда от более ранней к более поздней дате
            if self._pin_idx <= idx:
                i_from, i_to = self._pin_idx, idx
            else:
                i_from, i_to = idx, self._pin_idx
            d_from  = self._price_dates[i_from]
            d_to    = self._price_dates[i_to]
            c_from  = self._price_closes[i_from]
            c_to    = self._price_closes[i_to]
            diff_rub = c_to - c_from
            diff_pct = (c_to / c_from - 1) * 100 if c_from else 0
            sign = '+' if diff_rub >= 0 else ''
            arrow = '▲' if diff_rub >= 0 else '▼'
            ec_color = GREEN if diff_rub >= 0 else RED
            self._annot_price.set_text(
                f"{d_from.strftime(fmt)}  →  {d_to.strftime(fmt)}\n"
                f"{c_from:.2f} ₽  →  {c_to:.2f} ₽\n"
                f"{arrow} {sign}{diff_rub:.2f} ₽  ({sign}{diff_pct:.1f}%)")
            self._annot_price.set(
                bbox=dict(boxstyle='round,pad=0.5', fc=CARD, ec=ec_color, lw=1.5))

            # Рисуем/обновляем зелёный коридор между двумя точками
            if self._price_span_patch:
                try: self._price_span_patch.remove()
                except: pass
            xmin = mdates.date2num(d_from)
            xmax = mdates.date2num(d_to)
            span_color = GREEN if diff_rub >= 0 else RED
            self._price_span_patch = self.price_ax.axvspan(
                xmin, xmax, alpha=0.15, color=span_color, zorder=1)
        else:
            # Убираем коридор если нет закреплённой точки
            if self._price_span_patch:
                try: self._price_span_patch.remove()
                except: pass
                self._price_span_patch = None
            self._annot_price.set_text(f"{d.strftime(fmt)}\n{cl:.2f} ₽")
            self._annot_price.set(
                bbox=dict(boxstyle='round,pad=0.5', fc=CARD, ec=ACCENT, lw=1))

        self._annot_price.xy = (d, cl)
        self._annot_price.set_visible(True)
        self._vline.set_xdata([d])
        self._vline.set_visible(True)

        # Динамическое позиционирование: если курсор в правой половине →
        # тултип рисуем слева от точки, чтобы не выходил за край
        ax = self.price_ax
        xlim = ax.get_xlim()
        x_frac = (event.xdata - xlim[0]) / (xlim[1] - xlim[0]) if xlim[1] != xlim[0] else 0.5
        ylim = ax.get_ylim()
        y_frac = (cl - ylim[0]) / (ylim[1] - ylim[0]) if ylim[1] != ylim[0] else 0.5
        
        # Горизонталь: правее 55% → рисуем влево
        x_off = -14 if x_frac > 0.55 else 14
        # Вертикаль: выше 70% → рисуем вниз
        y_off = -14 if y_frac > 0.70 else 14
        
        ha = 'right' if x_frac > 0.55 else 'left'
        self._annot_price.set_ha(ha)
        self._annot_price.xyann = (x_off, y_off)
        
        # Маркер закреплённой точки
        if self._pin_marker:
            try: self._pin_marker.remove()
            except: pass
        if self._pin_idx is not None:
            pd_ = self._price_dates[self._pin_idx]
            pc_ = self._price_closes[self._pin_idx]
            self._pin_marker = self.price_ax.scatter(
                [pd_], [pc_], color=GOLD, s=60, zorder=11)
        
        self.price_canvas.draw_idle()

    def _on_price_leave(self, event):
        if self._pin_idx is None:   # только если нет закреплённой точки
            if self._annot_price:
                self._annot_price.set_visible(False)
            if self._vline:
                self._vline.set_visible(False)
            self.price_canvas.draw_idle()

    def _on_price_click(self, event):
        if event.inaxes != self.price_ax or not hasattr(self, '_price_dates'):
            return
        import bisect
        try:
            x_dt = mdates.num2date(event.xdata).replace(tzinfo=None)
        except Exception:
            return
        idx = bisect.bisect_left(self._price_dates, x_dt)
        idx = max(0, min(idx, len(self._price_dates)-1))

        if self._pin_idx is None:
            # Первый клик — закрепляем точку
            self._pin_idx = idx
        else:
            # Второй клик — сбрасываем всё
            self._pin_idx = None
            if self._pin_marker:
                try: self._pin_marker.remove()
                except: pass
                self._pin_marker = None
            if self._price_span_patch:
                try: self._price_span_patch.remove()
                except: pass
                self._price_span_patch = None
            if self._annot_price:
                self._annot_price.set_visible(False)
            if self._vline:
                self._vline.set_visible(False)
        self.price_canvas.draw_idle()

    # ── Рисуем дивиденды ───────────────────────────────────────
    def _draw_divs(self, divs, ticker):
        ax = self.div_ax
        ax.clear()
        ax.set_facecolor(CARD2)
        self._div_annot    = None
        self._div_bars_data = []

        if not divs:
            self._draw_empty(ax, "История дивидендов не найдена")
            self.div_canvas.draw()
            return

        # Группируем по году (суммируем несколько выплат в год)
        from collections import defaultdict
        year_divs = defaultdict(float)
        year_raw  = {}
        for raw_d, val in divs:
            yr = raw_d[:4]
            year_divs[yr] += val
            year_raw[yr]   = raw_d
        divs_by_year = sorted(year_divs.items())
        years     = [yr for yr, _ in divs_by_year]
        values    = [val for _, val in divs_by_year]
        raw_dates = [year_raw[yr] for yr in years]

        # Доходность: дивиденд / средняя цена за год
        yields = []
        for yr, val in zip(years, values):
            try:
                prices_year = [c for dt, c in zip(
                    getattr(self, "_all_dates",  []),
                    getattr(self, "_all_closes", []))
                    if dt.strftime("%Y") == yr]
                avg_p = sum(prices_year) / len(prices_year) if prices_year else 0
                yields.append(round(val / avg_p * 100, 1) if avg_p else 0)
            except Exception:
                yields.append(0)

        n = len(values)
        x = np.arange(n)

        ax.bar(x, values, width=0.6, zorder=3,
               color=GREEN_D, edgecolor=GREEN, linewidth=1.0)

        # Подписи суммы над столбцами
        for i, val in enumerate(values):
            ax.text(i, val + max(values) * 0.03,
                    f"{val:.2f} ₽",
                    ha="center", va="bottom",
                    fontsize=8, color=GREEN,
                    fontweight="bold", zorder=5)

        ax.set_xticks(x)
        ax.set_xticklabels(years, rotation=30, ha="right", fontsize=9)
        ax.set_ylabel("Дивиденд, ₽", color=MUTED, fontsize=10)
        ax.grid(True, axis="y", alpha=0.25, linestyle="--")
        for sp in ax.spines.values():
            sp.set_color(BORDER)

        # Аннотация hover
        self._div_annot = ax.annotate(
            "", xy=(0, 0), xytext=(0, 28), textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.5", fc=CARD, ec=GREEN, lw=1.5),
            fontsize=10, color=TEXT, zorder=10, ha="center")
        self._div_annot.set_visible(False)

        self._div_bars_data = list(zip(x, values, years, yields))

        # Итог: последний год
        last_yr  = years[-1] if years else ""
        last_val = values[-1] if values else 0
        self.div_info.configure(
            text=f"за {last_yr}: {last_val:.2f} ₽", text_color=GREEN)

        self.div_fig.subplots_adjust(top=0.88, bottom=0.22, left=0.09, right=0.97)
        self.div_canvas.draw()


    def _gold_shade(self, t):
        """t от 0 до 1 → от тёмного к золотому"""
        r = int(0x2D + t * (0xE3 - 0x2D))
        g = int(0x20 + t * (0xB3 - 0x20))
        b = int(0x08 + t * (0x41 - 0x08))
        return f"#{r:02x}{g:02x}{b:02x}"

    def _on_div_hover(self, event):
        if event.inaxes != self.div_ax or self._div_annot is None:
            return
        for xi, val, date_s, yld in self._div_bars_data:
            if abs(event.xdata - xi) < 0.4:
                txt = f"{date_s}\n{val:.2f} ₽"
                if yld > 0:
                    txt += f"\nДоходность: {yld:.1f}%"
                self._div_annot.set_text(txt)
                self._div_annot.xy = (xi, val)
                self._div_annot.set_visible(True)
                # Позиционирование: правее 60% → тултип влево, иначе вправо
                ax = self.div_ax
                xlim = ax.get_xlim()
                ylim = ax.get_ylim()
                x_frac = (xi - xlim[0]) / (xlim[1] - xlim[0]) if xlim[1] != xlim[0] else 0.5
                y_frac = (val - ylim[0]) / (ylim[1] - ylim[0]) if ylim[1] != ylim[0] else 0.5
                # Горизонталь
                if x_frac > 0.75:
                    x_off, ha = -60, 'center'
                elif x_frac < 0.25:
                    x_off, ha = 60, 'center'
                else:
                    x_off, ha = 0, 'center'
                # Вертикаль: если столбец высокий → тултип сбоку внизу
                y_off = 28 if y_frac < 0.75 else -45
                self._div_annot.set_ha(ha)
                self._div_annot.xyann = (x_off, y_off)
                self.div_canvas.draw_idle()
                return
        self._div_annot.set_visible(False)
        self.div_canvas.draw_idle()

    def _on_div_leave(self, event):
        if self._div_annot:
            self._div_annot.set_visible(False)
            self.div_canvas.draw_idle()


# ═══════════════════════════════════════════════════════════════
#  Главное окно
# ═══════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Stock Valuation")
        self.geometry("1300x820")
        self.minsize(1100, 700)
        self.configure(fg_color=BG)
        self._build_nav()
        self._build_pages()
        self._show_page("valuation")

    def _build_nav(self):
        nav = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0, height=48)
        nav.pack(side="top", fill="x")
        nav.pack_propagate(False)

        lbl(nav, "Stock Valuation", size=15,
            weight="bold", color=TEXT).pack(side="left", padx=20)

        self.nav_btns = {}
        bf = ctk.CTkFrame(nav, fg_color="transparent")
        bf.pack(side="left", padx=16)
        for key, name in [("valuation","Оценка"),
                           ("analytics","График & Дивиденды")]:
            b = ctk.CTkButton(
                bf, text=name, width=180, height=32,
                fg_color="transparent", hover_color=CARD2,
                text_color=MUTED, font=ctk.CTkFont(size=13),
                corner_radius=6,
                command=lambda k=key: self._show_page(k))
            b.pack(side="left", padx=3)
            self.nav_btns[key] = b

    def _build_pages(self):
        c = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        c.pack(fill="both", expand=True)
        c.grid_columnconfigure(0, weight=1)
        c.grid_rowconfigure(0, weight=1)

        self.analytics_page = AnalyticsPage(c)
        self.valuation_page  = ValuationPage(c, app=self)
        for p in (self.valuation_page, self.analytics_page):
            p.grid(row=0, column=0, sticky="nsew")

    def _show_page(self, key):
        {"valuation": self.valuation_page,
         "analytics": self.analytics_page}[key].tkraise()
        for k, b in self.nav_btns.items():
            b.configure(fg_color=CARD2 if k==key else "transparent",
                        text_color=TEXT  if k==key else MUTED)


if __name__ == "__main__":
    App().mainloop()
