import os
import json
import math
import shutil
import threading
import urllib.request
import urllib.error
from datetime import datetime, date

from openpyxl import Workbook, load_workbook

from kivy.app import App
from kivy.clock import Clock
from kivy.metrics import dp
from kivy.graphics import Color, Rectangle, Line
from kivy.properties import StringProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.widget import Widget

APP_TITLE = "EGX Stocks"
SHEET_NAME = "Stock Data"
EXCEL_NAME = "EGX_Stocks_Data_Log.xlsx"

HEADERS = [
    "Date", "Time", "Symbol", "Company Name", "Price", "Change", "Change %",
    "Volume", "Open", "High", "Low", "Previous Close", "Currency",
    "Time Period", "Data Source", "Support", "Resistance"
]

NAVY = (0.035, 0.12, 0.23, 1)
NAVY_DARK = (0.02, 0.08, 0.16, 1)
BLUE = (0.12, 0.44, 0.70, 1)
BLUE_LIGHT = (0.91, 0.95, 0.98, 1)
WHITE = (1, 1, 1, 1)
BG = (0.96, 0.97, 0.98, 1)
TEXT = (0.12, 0.16, 0.20, 1)
MUTED = (0.38, 0.44, 0.50, 1)
GREEN = (0.08, 0.54, 0.23, 1)
RED = (0.85, 0.12, 0.08, 1)
BORDER = (0.75, 0.81, 0.87, 1)


def number(value, digits=2):
    if value is None or value == "":
        return "—"
    try:
        f = float(value)
        return f"{f:,.{digits}f}"
    except Exception:
        return str(value)


def tradingview_scan(columns, timeout=25):
    url = "https://scanner.tradingview.com/egypt/scan"
    payload = {
        "markets": [{"name": "egypt", "symbols": {"query": {"types": []}, "tickers": []}}],
        "symbols": {"query": {"types": []}, "tickers": []},
        "columns": columns,
    }
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


class StockRow(Button):
    def __init__(self, stock, callback, **kwargs):
        super().__init__(**kwargs)
        self.stock = stock
        self.callback = callback
        self.size_hint_y = None
        self.height = dp(42)
        self.background_normal = ""
        self.background_color = WHITE
        self.color = TEXT
        self.font_size = dp(11)
        self.halign = "left"
        self.valign = "middle"
        self.text_size = (None, None)
        self.bind(on_release=lambda *_: callback(stock))
        self.update_text()

    def update_text(self):
        p = number(self.stock.get("close"))
        ch = self.stock.get("change_abs")
        pct = self.stock.get("change")
        vol = number(self.stock.get("volume"), 0)
        ch_text = number(ch)
        pct_text = number(pct) + "%"
        self.text = f"{self.stock.get('symbol','')}    {p}    {ch_text}    {pct_text}    {vol}    {self.stock.get('company','')}"
        if ch is not None:
            try:
                self.color = GREEN if float(ch) > 0 else RED if float(ch) < 0 else TEXT
            except Exception:
                self.color = TEXT


class TrendChart(Widget):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.rows = []
        self.bind(size=lambda *_: self.draw(), pos=lambda *_: self.draw())

    def set_rows(self, rows):
        self.rows = rows or []
        self.draw()

    def draw(self):
        self.canvas.clear()
        with self.canvas:
            Color(*WHITE)
            Rectangle(pos=self.pos, size=self.size)
            if not self.rows:
                return
            left = self.x + dp(34)
            right = self.right - dp(10)
            bottom = self.y + dp(24)
            top = self.top - dp(18)
            w = max(right - left, 10)
            h = max(top - bottom, 10)

            # Grid
            Color(0.86, 0.89, 0.92, 1)
            for i in range(5):
                yy = bottom + h * i / 4
                Line(points=[left, yy, right, yy], width=0.7)

            series = {
                "Support": ([r["support"] for r in self.rows], (0.12, 0.44, 0.70, 1)),
                "Resistance": ([r["resistance"] for r in self.rows], (0.75, 0.10, 0.08, 1)),
                "Volume": ([r["volume"] for r in self.rows], (0.08, 0.52, 0.29, 1)),
            }
            for _, (vals, color) in series.items():
                nums = []
                for v in vals:
                    try:
                        nums.append(float(v))
                    except Exception:
                        nums.append(None)
                valid = [v for v in nums if v is not None]
                if not valid:
                    continue
                first = valid[0] or 1
                norm = [(v / first * 100) if v is not None and first else None for v in nums]
                lo = min(v for v in norm if v is not None)
                hi = max(v for v in norm if v is not None)
                if math.isclose(lo, hi):
                    lo -= 1
                    hi += 1
                points = []
                for i, v in enumerate(norm):
                    if v is None:
                        continue
                    x = left + w * (i / max(len(norm) - 1, 1))
                    y = bottom + h * ((v - lo) / (hi - lo))
                    points.extend([x, y])
                if len(points) >= 4:
                    Color(*color)
                    Line(points=points, width=1.6)
                    for j in range(0, len(points), 2):
                        x, y = points[j], points[j + 1]
                        Line(circle=(x, y, 2.3), width=1.2)


class EGXStocksApp(App):
    status = StringProperty("Ready")

    def build(self):
        self.title = APP_TITLE
        self.stocks = []
        self.stock_by_symbol = {}
        self.selected_symbol = None
        self.current_company = ""
        self.chart_rows = []
        self.excel_path = self.prepare_excel()
        self.root_box = BoxLayout(orientation="vertical", spacing=0)
        self.show_main_screen()
        return self.root_box

    # ---------- Common UI ----------
    def make_label(self, text, size=11, color=TEXT, bold=False, **kwargs):
        lbl = Label(text=str(text), color=color, font_size=dp(size), bold=bold,
                    halign="left", valign="middle", **kwargs)
        lbl.bind(size=lambda obj, *_: setattr(obj, "text_size", (obj.width, obj.height)))
        return lbl

    def button(self, text, callback, bg=BLUE, height=44):
        b = Button(text=text, size_hint_y=None, height=dp(height), background_normal="",
                   background_color=bg, color=WHITE, font_size=dp(12), bold=True)
        b.bind(on_release=lambda *_: callback())
        return b

    def clear_root(self):
        self.root_box.clear_widgets()

    def header(self, title, subtitle=None, back=False):
        box = BoxLayout(orientation="vertical", size_hint_y=None, height=dp(92), padding=[dp(12), dp(8)])
        with box.canvas.before:
            Color(*NAVY)
            box._bg = Rectangle(pos=box.pos, size=box.size)
        box.bind(pos=lambda obj, *_: setattr(obj._bg, "pos", obj.pos),
                 size=lambda obj, *_: setattr(obj._bg, "size", obj.size))
        row = BoxLayout(orientation="horizontal", size_hint_y=None, height=dp(42), spacing=dp(6))
        if back:
            row.add_widget(self.button("Back", self.show_main_screen, bg=BLUE, height=38))
        title_lbl = self.make_label(title, 18, WHITE, True)
        row.add_widget(title_lbl)
        box.add_widget(row)
        if subtitle:
            box.add_widget(self.make_label(subtitle, 9, WHITE))
        return box

    def set_status(self, text):
        self.status = str(text)
        if hasattr(self, "status_label"):
            self.status_label.text = self.status

    # ---------- Excel ----------
    def prepare_excel(self):
        user_path = os.path.join(self.user_data_dir, EXCEL_NAME)
        if not os.path.exists(user_path):
            bundled = os.path.join(os.path.dirname(__file__), "assets", EXCEL_NAME)
            if os.path.exists(bundled):
                shutil.copy2(bundled, user_path)
            else:
                wb = Workbook()
                ws = wb.active
                ws.title = SHEET_NAME
                ws.append(HEADERS)
                wb.save(user_path)
        self.ensure_excel_schema(user_path)
        return user_path

    def ensure_excel_schema(self, path):
        if not os.path.exists(path):
            return
        wb = load_workbook(path)
        if SHEET_NAME not in wb.sheetnames:
            ws = wb.create_sheet(SHEET_NAME)
        else:
            ws = wb[SHEET_NAME]
        current = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)] if ws.max_row else []
        if current != HEADERS:
            # Migration is explicit: only an English schema is accepted in the APK.
            if ws.max_row <= 1:
                ws.delete_rows(1, ws.max_row)
                ws.append(HEADERS)
            else:
                missing = [h for h in HEADERS if h not in current]
                for h in missing:
                    ws.cell(1, ws.max_column + 1).value = h
        wb.save(path)
        wb.close()

    def read_excel_rows(self):
        wb = load_workbook(self.excel_path, data_only=True)
        if SHEET_NAME not in wb.sheetnames:
            wb.close()
            return []
        ws = wb[SHEET_NAME]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        idx = {str(h): i for i, h in enumerate(headers) if h is not None}
        rows = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not any(v not in (None, "") for v in r):
                continue
            rows.append({h: (r[i] if i < len(r) else None) for h, i in idx.items()})
        wb.close()
        return rows

    def write_excel_record(self, record):
        wb = load_workbook(self.excel_path)
        ws = wb[SHEET_NAME]
        ws.append([record.get(h, "") for h in HEADERS])
        wb.save(self.excel_path)
        wb.close()

    # ---------- Main screen ----------
    def show_main_screen(self):
        self.clear_root()
        self.root_box.add_widget(self.header("Phase 1 - EGX", "Egyptian Stocks List • Prices from TradingView"))

        body = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(7))
        search_row = BoxLayout(size_hint_y=None, height=dp(46), spacing=dp(6))
        self.search_entry = TextInput(hint_text="Search", multiline=False, font_size=dp(14), padding=[dp(8), dp(10)])
        self.search_entry.bind(text=lambda *_: self.filter_stocks())
        search_row.add_widget(self.search_entry)
        body.add_widget(search_row)

        btn_row = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(6))
        btn_row.add_widget(self.button("Load Stocks", self.load_stocks))
        self.update_btn = self.button("Update Stock", self.fetch_selected_quote, bg=NAVY)
        self.update_btn.disabled = True
        btn_row.add_widget(self.update_btn)
        body.add_widget(btn_row)

        table_title = self.make_label("Stock List", 12, NAVY_DARK, True, size_hint_y=None, height=dp(30))
        body.add_widget(table_title)

        scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.stock_rows_box = GridLayout(cols=1, spacing=dp(2), size_hint_y=None, size_hint_x=None,
                                         width=dp(900), padding=[dp(2), dp(2)])
        self.stock_rows_box.bind(minimum_height=self.stock_rows_box.setter("height"))
        scroll.add_widget(self.stock_rows_box)
        body.add_widget(scroll)

        self.status_label = self.make_label(self.status, 9, MUTED, size_hint_y=None, height=dp(30))
        body.add_widget(self.status_label)
        self.root_box.add_widget(body)

        # Open the analysis screen from the header title area by tapping the header.
        # A dedicated button is added at the bottom to keep touch behavior explicit.
        self.root_box.add_widget(self.button("Open Stock Analysis", self.show_menu_screen, bg=NAVY, height=46))

    def display_stocks(self, stocks):
        self.stock_rows_box.clear_widgets()
        for s in stocks:
            self.stock_rows_box.add_widget(StockRow(s, self.select_stock))

    def filter_stocks(self):
        q = self.search_entry.text.strip().lower()
        if not q:
            data = self.stocks
        else:
            data = [s for s in self.stocks if q in s.get("symbol", "").lower() or q in s.get("company", "").lower()]
        self.display_stocks(data)

    def select_stock(self, stock):
        self.selected_symbol = stock.get("symbol")
        self.current_company = stock.get("company") or self.selected_symbol
        self.update_btn.disabled = False
        self.open_analysis_for_symbol(self.selected_symbol)

    def load_stocks(self):
        self.set_status("Loading EGX stocks and live TradingView prices...")
        def worker():
            try:
                columns = ["name", "description", "close", "change", "change_abs", "open", "high", "low", "volume", "currency_code"]
                data = tradingview_scan(columns)
                rows = data.get("data") or []
                stocks = []
                for row in rows:
                    full = row.get("s", "")
                    vals = row.get("d") or []
                    if not full or not vals:
                        continue
                    symbol = str(full).split(":")[-1].upper()
                    item = {"symbol": symbol, "tv_symbol": full}
                    for i, key in enumerate(columns):
                        item[key] = vals[i] if i < len(vals) else None
                    item["company"] = str(item.get("description") or item.get("name") or symbol)
                    stocks.append(item)
                unique = {s["symbol"]: s for s in stocks}
                self.stocks = list(unique.values())
                self.stock_by_symbol = {s["symbol"]: s for s in self.stocks}
                Clock.schedule_once(lambda *_: self.display_stocks(self.stocks))
                Clock.schedule_once(lambda *_: self.set_status(f"Loaded {len(self.stocks)} stocks from TradingView"))
                Clock.schedule_once(lambda *_: setattr(self.update_btn, "disabled", False))
            except Exception as e:
                Clock.schedule_once(lambda *_: self.set_status(f"Could not load EGX data: {type(e).__name__}: {e}"))
        threading.Thread(target=worker, daemon=True).start()

    def fetch_selected_quote(self):
        symbol = self.selected_symbol
        if not symbol:
            return
        self.set_status(f"Updating {symbol} from TradingView...")
        def worker():
            try:
                columns = ["name", "description", "close", "change", "change_abs", "open", "high", "low", "volume", "currency_code", "prev_close_price"]
                data = tradingview_scan(columns)
                match = None
                for row in data.get("data") or []:
                    tv = str(row.get("s") or "")
                    if tv.upper() == str(self.stock_by_symbol.get(symbol, {}).get("tv_symbol") or f"EGX:{symbol}").upper() or tv.split(":")[-1].upper() == symbol.upper():
                        match = row
                        break
                if match is None:
                    raise RuntimeError("The selected stock was not found in TradingView")
                vals = match.get("d") or []
                fresh = {k: (vals[i] if i < len(vals) else None) for i, k in enumerate(columns)}
                if fresh.get("change_abs") is None and fresh.get("close") is not None and fresh.get("prev_close_price") not in (None, 0):
                    fresh["change_abs"] = fresh["close"] - fresh["prev_close_price"]
                if fresh.get("change") is None and fresh.get("close") is not None and fresh.get("prev_close_price") not in (None, 0):
                    fresh["change"] = (fresh["close"] - fresh["prev_close_price"]) / fresh["prev_close_price"] * 100
                stock = self.stock_by_symbol[symbol]
                stock.update(fresh)
                stock["company"] = str(fresh.get("description") or fresh.get("name") or stock.get("company") or symbol)
                self.current_company = stock["company"]
                Clock.schedule_once(lambda *_: self.display_stocks(self.stocks))
                Clock.schedule_once(lambda *_: self.set_status(f"Updated {symbol} successfully"))
                Clock.schedule_once(lambda *_: self.open_analysis_for_symbol(symbol))
            except Exception as e:
                Clock.schedule_once(lambda *_: self.set_status(f"Update failed: {type(e).__name__}: {e}"))
        threading.Thread(target=worker, daemon=True).start()

    # ---------- Analysis screen ----------
    def show_menu_screen(self):
        self.clear_root()
        self.root_box.add_widget(self.header("Stock Analysis", "English interface • English Excel schema", back=True))
        body = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(6))

        picker_row = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(6))
        picker_row.add_widget(self.make_label("Select Stock Code", 10, NAVY_DARK, True, size_hint_x=None, width=dp(125)))
        self.stock_picker = Spinner(text="Select", values=(), size_hint_x=1, font_size=dp(12))
        self.stock_picker.bind(text=self.on_picker_selected)
        picker_row.add_widget(self.stock_picker)
        body.add_widget(picker_row)

        self.detail_title = self.make_label("No stock selected", 12, NAVY_DARK, True, size_hint_y=None, height=dp(30))
        body.add_widget(self.detail_title)

        self.detail_grid = GridLayout(cols=6, rows=2, size_hint_y=None, height=dp(110), spacing=dp(3))
        self.metric_labels = {}
        for key, label in [("price", "Last Price"), ("change", "Change"), ("pct", "Change %"), ("open", "Open"), ("high", "High"), ("low", "Low"),
                           ("prev", "Previous Close"), ("volume", "Volume"), ("currency", "Currency")]:
            box = BoxLayout(orientation="vertical", padding=dp(3))
            box.add_widget(self.make_label(label, 8, NAVY_DARK, True, size_hint_y=None, height=dp(24)))
            value = self.make_label("—", 11, TEXT, True)
            box.add_widget(value)
            self.metric_labels[key] = value
            self.detail_grid.add_widget(box)
        body.add_widget(self.detail_grid)

        btns = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(5))
        btns.add_widget(self.button("Record Stock Data", self.save_selected_stock_to_excel))
        btns.add_widget(self.button("Calculate Stock", self.calculate_stock, bg=NAVY))
        body.add_widget(btns)

        result_scroll = ScrollView(do_scroll_y=True, do_scroll_x=False, size_hint_y=None, height=dp(250))
        results = GridLayout(cols=2, spacing=dp(4), size_hint_y=None, padding=dp(3))
        results.bind(minimum_height=results.setter("height"))
        self.result_values = {}
        for key, label in [
            ("support", "Support"), ("resistance", "Resistance"),
            ("correction", "Correction Scenario"), ("breakout", "Breakout Scenario"),
            ("entry", "Entry Range"), ("stop", "Stop Loss"),
            ("risk", "Risk"), ("reward", "Reward"),
            ("target", "Target"), ("ratio", "Risk / Reward Ratio")
        ]:
            results.add_widget(self.make_label(label, 9, NAVY_DARK, True, size_hint_y=None, height=dp(34)))
            val = self.make_label("—", 9, TEXT, True, size_hint_y=None, height=dp(34))
            self.result_values[key] = val
            results.add_widget(val)
        result_scroll.add_widget(results)
        body.add_widget(result_scroll)

        chart_title = self.make_label("Support, Resistance & Volume Trend", 10, NAVY_DARK, True, size_hint_y=None, height=dp(28))
        body.add_widget(chart_title)
        self.chart = TrendChart(size_hint_y=None, height=dp(240))
        body.add_widget(self.chart)
        self.root_box.add_widget(body)

        self.refresh_stock_picker()
        if self.selected_symbol:
            self.open_analysis_for_symbol(self.selected_symbol)

    def refresh_stock_picker(self):
        try:
            rows = self.read_excel_rows()
            counts = {}
            for r in rows:
                symbol = str(r.get("Symbol") or "").strip().upper()
                if symbol:
                    counts[symbol] = counts.get(symbol, 0) + 1
            values = [f"{n}  {s}" for s, n in sorted(counts.items())]
            self.picker_map = {v: v.split("  ", 1)[1] for v in values}
            if hasattr(self, "stock_picker"):
                self.stock_picker.values = values
                if self.selected_symbol:
                    for v, s in self.picker_map.items():
                        if s == self.selected_symbol:
                            self.stock_picker.text = v
                            break
        except Exception as e:
            self.set_status(f"Excel picker error: {e}")

    def on_picker_selected(self, spinner, text):
        symbol = self.picker_map.get(text) if hasattr(self, "picker_map") else None
        if symbol:
            self.open_analysis_for_symbol(symbol)

    def open_analysis_for_symbol(self, symbol):
        self.selected_symbol = symbol
        if not hasattr(self, "stock_picker"):
            return
        stock = self.stock_by_symbol.get(symbol, {})
        self.current_company = stock.get("company") or symbol
        self.detail_title.text = f"{symbol} — {self.current_company}"
        vals = {
            "price": number(stock.get("close")), "change": number(stock.get("change_abs")),
            "pct": number(stock.get("change")) + "%", "open": number(stock.get("open")),
            "high": number(stock.get("high")), "low": number(stock.get("low")),
            "prev": number(stock.get("prev_close_price")), "volume": number(stock.get("volume"), 0),
            "currency": str(stock.get("currency_code") or "EGP")
        }
        for k, v in vals.items():
            if k in self.metric_labels:
                self.metric_labels[k].text = v
        self.refresh_stock_picker()
        self.update_chart_from_second_tree()

    # ---------- Excel record / chart ----------
    def save_selected_stock_to_excel(self):
        symbol = self.selected_symbol
        if not symbol:
            self.set_status("Select a stock first")
            return
        stock = self.stock_by_symbol.get(symbol, {})
        now = datetime.now()
        try:
            existing = self.read_excel_rows()
            lows = []
            highs = []
            for r in existing:
                if str(r.get("Symbol") or "").strip().upper() != symbol.upper():
                    continue
                d = self.to_date(r.get("Date"))
                if d is None:
                    continue
                if (now.date() - d).days <= 30:
                    try: lows.append(float(r.get("Low")))
                    except Exception: pass
                    try: highs.append(float(r.get("High")))
                    except Exception: pass
            try: lows.append(float(stock.get("low")))
            except Exception: pass
            try: highs.append(float(stock.get("high")))
            except Exception: pass
            record = {
                "Date": now.date(), "Time": now.strftime("%H:%M:%S"), "Symbol": symbol,
                "Company Name": stock.get("company") or symbol, "Price": stock.get("close"),
                "Change": stock.get("change_abs"), "Change %": stock.get("change"),
                "Volume": stock.get("volume"), "Open": stock.get("open"), "High": stock.get("high"),
                "Low": stock.get("low"), "Previous Close": stock.get("prev_close_price"),
                "Currency": stock.get("currency_code") or "EGP", "Time Period": "Intraday",
                "Data Source": "TradingView", "Support": min(lows) if lows else "",
                "Resistance": max(highs) if highs else ""
            }
            self.write_excel_record(record)
            self.set_status(f"Recorded {symbol} in Excel")
            self.refresh_stock_picker()
            self.update_chart_from_second_tree()
        except Exception as e:
            self.set_status(f"Excel record failed: {type(e).__name__}: {e}")

    def to_date(self, value):
        if isinstance(value, datetime): return value.date()
        if isinstance(value, date): return value
        if value is None: return None
        s = str(value).strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try: return datetime.strptime(s, fmt).date()
            except Exception: pass
        return None

    def update_chart_from_second_tree(self):
        if not self.selected_symbol or not hasattr(self, "chart"):
            return
        try:
            rows = self.load_chart_rows(self.selected_symbol, 8)
            self.chart_rows = rows
            self.chart.set_rows(rows)
        except Exception:
            self.chart_rows = []
            self.chart.set_rows([])

    def load_chart_rows(self, symbol, max_days=8):
        rows = self.read_excel_rows()
        filtered = [r for r in rows if str(r.get("Symbol") or "").strip().upper() == symbol.upper()]
        dates = sorted({self.to_date(r.get("Date")) for r in filtered if self.to_date(r.get("Date"))})[-max_days:]
        out = []
        for d in dates:
            same = [r for r in filtered if self.to_date(r.get("Date")) == d]
            r = same[-1]
            out.append({
                "date": d, "support": r.get("Support"), "resistance": r.get("Resistance"), "volume": r.get("Volume")
            })
        return out

    # ---------- Calculation ----------
    def calculate_stock(self):
        symbol = self.selected_symbol
        if not symbol:
            self.set_status("Select a stock first")
            return
        try:
            rows = self.read_excel_rows()
            search_date = datetime.now().date()
            symbol_rows = [r for r in rows if str(r.get("Symbol") or "").strip().upper() == symbol.upper() and self.to_date(r.get("Date")) and self.to_date(r.get("Date")) <= search_date]
            if not symbol_rows:
                raise RuntimeError(f"No saved data found for {symbol} up to {search_date}")
            trading_dates = sorted({self.to_date(r.get("Date")) for r in symbol_rows})[-8:]
            analysis = [r for r in symbol_rows if self.to_date(r.get("Date")) in set(trading_dates)]
            analysis.sort(key=lambda r: (self.to_date(r.get("Date")) or date.min, str(r.get("Time") or "")))
            latest = analysis[-1]
            latest_price = self.fnum(latest.get("Price"))
            if latest_price is None:
                raise RuntimeError("Latest stock record has no valid Price")
            highs = [self.fnum(r.get("High")) for r in analysis]
            lows = [self.fnum(r.get("Low")) for r in analysis]
            highs = [x for x in highs if x is not None]
            lows = [x for x in lows if x is not None]
            support = min(lows) if lows else None
            resistance = max(highs) if highs else None
            if support is None or resistance is None:
                raise RuntimeError("Not enough High/Low data for calculation")
            stop = support * 0.99
            risk = latest_price - stop
            if risk <= 0:
                raise RuntimeError("Current price is not above calculated stop loss")
            target = latest_price + risk * 2
            reward = target - latest_price
            self.result_values["support"].text = number(support)
            self.result_values["resistance"].text = number(resistance)
            self.result_values["correction"].text = f"Correction toward {number(support)}"
            self.result_values["breakout"].text = f"Breakout above {number(resistance)}"
            self.result_values["entry"].text = f"{number(latest_price)}"
            self.result_values["stop"].text = number(stop)
            self.result_values["risk"].text = number(risk)
            self.result_values["reward"].text = number(reward)
            self.result_values["target"].text = number(target)
            self.result_values["ratio"].text = "1 : 2"
            self.set_status(f"Calculated {symbol} using {len(trading_dates)} trading days")
        except Exception as e:
            self.set_status(f"Calculation failed: {type(e).__name__}: {e}")

    @staticmethod
    def fnum(v):
        try:
            if v in (None, "", "—"): return None
            return float(v)
        except Exception:
            return None


if __name__ == "__main__":
    EGXStocksApp().run()
