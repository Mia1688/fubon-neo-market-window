"""看盤視窗：富邦 Neo API 桌面行情與下單工具。"""
from __future__ import annotations

import base64
import ctypes
import json
import os
import queue
import subprocess
import threading
import uuid
import urllib.error
import urllib.parse
import urllib.request
from ctypes import wintypes
from datetime import date, datetime, time as clock_time, timedelta
from pathlib import Path
from typing import Any

from PySide6.QtCharts import (
    QChart, QChartView, QLineSeries, QValueAxis,
)
from PySide6.QtCore import QLineF, QMargins, QRectF, QTimer, Qt
from PySide6.QtGui import QBrush, QColor, QPainter, QPen
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QButtonGroup, QFileDialog, QFrame, QGridLayout,
    QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QPushButton, QComboBox, QSpinBox, QSplitter, QStackedWidget, QTabWidget, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from fubon_neo.constant import (
    BSAction, FutOptMarketType, FutOptOrderType, FutOptPriceType,
    MarketType, OrderType, PriceType, StockType, TimeInForce,
)
from fubon_neo.sdk import FubonSDK, FutOptOrder, Mode, Order


APP_DIR = Path(__file__).resolve().parent
SETTINGS_FILE = APP_DIR / "settings.json"
PENDING_FILE = APP_DIR / "pending_orders.json"
GROUPS_FILE = APP_DIR / "stock_groups.json"
TEST_URL = "wss://neoapitest.fbs.com.tw/TASP/XCPXWS"
INDEX_SYMBOL = "IR0001"
FUTURE_SYMBOL = "TXF1!"
TWSE_COMPANY_URL = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"
YUANTA_PCF_URL = "https://etfapi.yuantaetfs.com/ectranslation/api/bridge"


class DataBlob(ctypes.Structure):
    _fields_ = [("size", wintypes.DWORD), ("data", ctypes.POINTER(ctypes.c_ubyte))]


CRYPT32 = ctypes.WinDLL("crypt32", use_last_error=True)
KERNEL32 = ctypes.WinDLL("kernel32", use_last_error=True)
CRYPT32.CryptProtectData.argtypes = [ctypes.POINTER(DataBlob), wintypes.LPCWSTR, ctypes.POINTER(DataBlob), ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(DataBlob)]
CRYPT32.CryptProtectData.restype = wintypes.BOOL
CRYPT32.CryptUnprotectData.argtypes = [ctypes.POINTER(DataBlob), ctypes.POINTER(wintypes.LPWSTR), ctypes.POINTER(DataBlob), ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(DataBlob)]
CRYPT32.CryptUnprotectData.restype = wintypes.BOOL
KERNEL32.LocalFree.argtypes = [ctypes.c_void_p]
KERNEL32.LocalFree.restype = ctypes.c_void_p


def protect_secret(secret: str) -> str:
    """Encrypt a secret for the current Windows user with DPAPI."""
    raw = secret.encode("utf-8")
    buffer = ctypes.create_string_buffer(raw)
    source = DataBlob(len(raw), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)))
    result = DataBlob()
    if not CRYPT32.CryptProtectData(
        ctypes.byref(source), None, None, None, None, 0, ctypes.byref(result)
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        encrypted = ctypes.string_at(result.data, result.size)
        return base64.b64encode(encrypted).decode("ascii")
    finally:
        KERNEL32.LocalFree(ctypes.cast(result.data, ctypes.c_void_p))


def unprotect_secret(value: str) -> str:
    """Decrypt a DPAPI value for the current Windows user."""
    raw = base64.b64decode(value.encode("ascii"), validate=True)
    buffer = ctypes.create_string_buffer(raw)
    source = DataBlob(len(raw), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)))
    result = DataBlob()
    if not CRYPT32.CryptUnprotectData(
        ctypes.byref(source), None, None, None, None, 0, ctypes.byref(result)
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(result.data, result.size).decode("utf-8")
    finally:
        KERNEL32.LocalFree(ctypes.cast(result.data, ctypes.c_void_p))


def as_float(value: Any) -> float | None:
    try: return float(value)
    except (TypeError, ValueError): return None


def fmt(value: Any) -> str:
    number = as_float(value)
    return "—" if number is None else f"{number:,.2f}".rstrip("0").rstrip(".")


def previous_stock_tick(price: Any) -> float | None:
    """Return the valid stock price exactly one tick below the supplied price."""
    number = as_float(price)
    if number is None or number <= 0: return None
    if number <= 10: tick = .01
    elif number <= 50: tick = .05
    elif number <= 100: tick = .1
    elif number <= 500: tick = .5
    elif number <= 1000: tick = 1
    else: tick = 5
    return round(number - tick, 2)


def previous_stock_ticks(price: Any, count: int) -> float | None:
    """Move down by a configurable number of valid Taiwan stock ticks."""
    current = as_float(price)
    if current is None: return None
    for _ in range(max(int(count), 1)):
        current = previous_stock_tick(current)
        if current is None: return None
    return current


def next_stock_order_window(now: datetime | None = None) -> datetime:
    """Return the next safe regular-session order-entry time (local Taipei clock)."""
    current = now or datetime.now()
    candidate = datetime.combine(current.date(), clock_time(8, 31))
    if current >= candidate or current.weekday() >= 5: candidate += timedelta(days=1)
    while candidate.weekday() >= 5: candidate += timedelta(days=1)
    return candidate


def fetch_json(url: str, params: dict[str, str] | None = None) -> Any:
    """Read a public JSON endpoint with a bounded timeout."""
    if params: url = f"{url}?{urllib.parse.urlencode(params)}"
    request = urllib.request.Request(url, headers={"User-Agent": "MarketDesk/1.0"})
    try:
        with urllib.request.urlopen(request, timeout=12) as response:
            return json.loads(response.read().decode("utf-8-sig"))
    except urllib.error.URLError as exc:
        if "CERTIFICATE_VERIFY_FAILED" not in str(exc): raise
        script = "$ProgressPreference='SilentlyContinue'; [Console]::OutputEncoding=[Text.UTF8Encoding]::new(); (Invoke-WebRequest -Uri $env:MARKET_DESK_FETCH_URL -UseBasicParsing -TimeoutSec 15).Content"
        flags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
        environment = os.environ.copy(); environment["MARKET_DESK_FETCH_URL"] = url
        result = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script], capture_output=True, timeout=20, creationflags=flags, env=environment)
        if result.returncode: raise RuntimeError("Windows HTTPS request failed")
        return json.loads(result.stdout.decode("utf-8-sig"))


def unwrap(payload: Any) -> dict[str, Any]:
    if not isinstance(payload, dict): return {}
    nested = payload.get("data")
    if isinstance(nested, dict) and not any(key in payload for key in ("lastPrice", "closePrice", "index")):
        return nested
    return payload


class LineChart(QChartView):
    def __init__(self) -> None:
        super().__init__(); self.setRenderHint(QPainter.RenderHint.Antialiasing); self.setMinimumHeight(190); self.setStyleSheet("background:transparent;border:0"); self.show_values([])

    def show_values(self, values: list[float]) -> None:
        chart = QChart(); chart.setBackgroundVisible(True); chart.setBackgroundRoundness(0); chart.setBackgroundBrush(QBrush(QColor("#101828"))); chart.legend().hide(); chart.setMargins(QMargins(3, 3, 3, 3))
        series = QLineSeries(); series.setPen(QPen(QColor("#39d7ad"), 2.2))
        for index, value in enumerate(values): series.append(index, value)
        chart.addSeries(series)
        x_axis = QValueAxis(); x_axis.setRange(0, max(len(values) - 1, 1)); x_axis.setLabelsVisible(False); x_axis.setGridLineVisible(False); x_axis.setLineVisible(False)
        y_axis = QValueAxis(); y_axis.setLabelFormat("%.0f"); y_axis.setLabelsColor(QColor("#6f7f99")); y_axis.setGridLineColor(QColor("#263650")); y_axis.setLineVisible(False)
        if values:
            low, high = min(values), max(values); padding = max((high - low) * .12, abs(high) * .001, 1)
            y_axis.setRange(low - padding, high + padding)
        else: y_axis.setRange(0, 1)
        chart.addAxis(x_axis, Qt.AlignmentFlag.AlignBottom); chart.addAxis(y_axis, Qt.AlignmentFlag.AlignLeft)
        series.attachAxis(x_axis); series.attachAxis(y_axis); self.setChart(chart)


class QuotePage(QWidget):
    def __init__(self, eyebrow: str, title: str, symbol: str) -> None:
        super().__init__(); layout = QVBoxLayout(self); layout.setContentsMargins(20, 16, 20, 14); layout.setSpacing(7)
        top = QHBoxLayout(); heading = QVBoxLayout(); eye = QLabel(eyebrow); eye.setObjectName("eyebrow"); name = QLabel(title); name.setObjectName("sectionTitle"); heading.addWidget(eye); heading.addWidget(name); top.addLayout(heading); top.addStretch(); badge = QLabel(symbol); badge.setObjectName("badge"); top.addWidget(badge, alignment=Qt.AlignmentFlag.AlignTop); layout.addLayout(top)
        self.price = QLabel("—"); self.price.setObjectName("heroPrice"); self.change = QLabel("等待行情"); self.change.setObjectName("flatChange"); self.detail = QLabel("開 —   高 —   低 —   前收 —"); self.detail.setObjectName("muted"); self.chart = LineChart(); self.source = QLabel("尚未連線"); self.source.setObjectName("source")
        layout.addWidget(self.price); layout.addWidget(self.change); layout.addWidget(self.detail); layout.addWidget(self.chart, 1); layout.addWidget(self.source)

    def update_quote(self, data: dict[str, Any], source: str) -> None:
        last_trade = data.get("lastTrade") if isinstance(data.get("lastTrade"), dict) else {}
        value = data.get("lastPrice") if data.get("lastPrice") is not None else data.get("closePrice")
        if value is None: value = last_trade.get("price")
        self.price.setText(fmt(value)); self._change(as_float(data.get("change")), as_float(data.get("changePercent")))
        self.detail.setText(f"開 {fmt(data.get('openPrice'))}   高 {fmt(data.get('highPrice'))}   低 {fmt(data.get('lowPrice'))}   前收 {fmt(data.get('previousClose'))}")
        self.source.setText(source)

    def update_index(self, value: Any, previous: Any, source: str) -> None:
        current, prior = as_float(value), as_float(previous); change = current - prior if current is not None and prior is not None else None; percent = change / prior * 100 if change is not None and prior else None
        self.price.setText(fmt(current)); self._change(change, percent); self.detail.setText(f"前收 {fmt(prior)}   代碼 {INDEX_SYMBOL}"); self.source.setText(source)

    def _change(self, change: float | None, percent: float | None) -> None:
        if change is None: text, name = "漲跌 —", "flatChange"
        else: text, name = f"{change:+,.2f}" + (f"  {percent:+.2f}%" if percent is not None else ""), "upChange" if change > 0 else "downChange" if change < 0 else "flatChange"
        self.change.setText(text); self.change.setObjectName(name); self.change.style().unpolish(self.change); self.change.style().polish(self.change)


class CandleChart(QWidget):
    def __init__(self) -> None:
        super().__init__(); self.setMinimumHeight(240); self.symbol = "—"; self.rows: list[dict[str, Any]] = []

    def show_candles(self, symbol: str, rows: list[dict[str, Any]]) -> None:
        cleaned: list[dict[str, Any]] = []
        for row in rows:
            try:
                opening, high, low, close = [float(row[key]) for key in ("open", "high", "low", "close")]
                cleaned.append({"date": str(row.get("date", ""))[:10], "open": opening, "high": high, "low": low, "close": close})
            except (KeyError, TypeError, ValueError): continue
        self.symbol = symbol; self.rows = cleaned; self.update()

    def paintEvent(self, event: Any) -> None:
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing); painter.fillRect(self.rect(), QColor("#101828")); painter.setPen(QColor("#8192af")); painter.drawText(QRectF(68, 6, max(self.width() - 86, 1), 24), Qt.AlignmentFlag.AlignCenter, f"{self.symbol} · 日 K")
        plot = QRectF(68, 38, max(self.width() - 88, 1), max(self.height() - 72, 1)); lows = [row["low"] for row in self.rows]; highs = [row["high"] for row in self.rows]
        low = min(lows) if lows else 0.; high = max(highs) if highs else 1.; padding = max((high - low) * .08, .1); low -= padding; high += padding
        painter.setPen(QPen(QColor("#263650"), 1))
        for index in range(5):
            y = plot.top() + plot.height() * index / 4; painter.drawLine(QLineF(plot.left(), y, plot.right(), y)); value = high - (high - low) * index / 4; painter.setPen(QColor("#70819f")); painter.drawText(QRectF(0, y - 9, 61, 18), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, f"{value:.1f}"); painter.setPen(QPen(QColor("#263650"), 1))
        if not self.rows:
            painter.setPen(QColor("#65738d")); painter.drawText(plot, Qt.AlignmentFlag.AlignCenter, "尚無 K 線資料"); return
        step = plot.width() / len(self.rows); body_width = max(2., min(9., step * .58))
        def price_y(price: float) -> float: return plot.bottom() - (price - low) / (high - low) * plot.height()
        for index, row in enumerate(self.rows):
            x = plot.left() + step * (index + .5); opening, close = row["open"], row["close"]
            painter.setPen(QPen(QColor("#ffffff"), 1.15)); painter.drawLine(QLineF(x, price_y(row["high"]), x, price_y(row["low"])))
            top, bottom = min(price_y(opening), price_y(close)), max(price_y(opening), price_y(close)); painter.setPen(Qt.PenStyle.NoPen); painter.setBrush(QColor("#ff667d") if close >= opening else QColor("#35d3a3")); painter.drawRect(QRectF(x - body_width / 2, top, body_width, max(bottom - top, 2.)))
        label_indexes = sorted({0, len(self.rows) // 4, len(self.rows) // 2, len(self.rows) * 3 // 4, len(self.rows) - 1}); painter.setPen(QColor("#70819f"))
        for index in label_indexes:
            x = plot.left() + step * (index + .5); painter.drawText(QRectF(x - 28, plot.bottom() + 5, 56, 18), Qt.AlignmentFlag.AlignCenter, self.rows[index]["date"][5:])


class MarketWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__(); self.setWindowTitle("看盤視窗"); self.resize(1460, 900); self.setMinimumSize(1100, 700)
        self.sdk: Any = None; self.account: Any = None; self.events: queue.Queue[tuple[str, Any]] = queue.Queue(); self.inventory: list[dict[str, Any]] = []; self.quotes: dict[str, dict[str, Any]] = {}; self.pending_orders: list[dict[str, Any]] = []; self.limit_rows_by_symbol: dict[str, dict[str, Any]] = {}; self.previous_index_close: float | None = None; self.connected = False; self.environment_name = "production"; self.cert_password_override = ""; self.limit_scan_running = False
        self._ui(); self._theme(); self._load_settings(); self._load_pending_orders(); self.timer = QTimer(self); self.timer.timeout.connect(self._drain); self.timer.start(100); self.limit_timer = QTimer(self); self.limit_timer.setInterval(30_000); self.limit_timer.timeout.connect(self.scan_limit_monitor); self.limit_timer.start(); self.pending_timer = QTimer(self); self.pending_timer.setInterval(1_000); self.pending_timer.timeout.connect(self._check_pending_orders); self.pending_timer.start()

    def _ui(self) -> None:
        root = QWidget(); root.setObjectName("root"); self.setCentralWidget(root); page = QVBoxLayout(root); page.setContentsMargins(20, 16, 20, 16); page.setSpacing(12)
        header = QHBoxLayout(); brand = QLabel("看盤視窗"); brand.setObjectName("brand"); header.addWidget(brand); tag = QLabel("FUBON NEO · DESKTOP"); tag.setObjectName("muted"); header.addWidget(tag); header.addStretch(); self.dot = QLabel("●"); self.dot.setObjectName("offlineDot"); self.status = QLabel("尚未登入"); self.status.setObjectName("muted"); header.addWidget(self.dot); header.addWidget(self.status); page.addLayout(header)
        login_frame = QFrame(); login_frame.setObjectName("loginBar"); login = QHBoxLayout(login_frame); login.setContentsMargins(12, 8, 12, 8); login.setSpacing(8)
        account_label = QLabel("登入"); account_label.setObjectName("eyebrow"); self.user_id = QLineEdit(); self.user_id.setPlaceholderText("帳號"); self.user_id.setMaximumWidth(240); self.user_password = QLineEdit(); self.user_password.setEchoMode(QLineEdit.EchoMode.Password); self.user_password.setPlaceholderText("密碼"); self.user_password.setMaximumWidth(240); self.user_password.returnPressed.connect(self.login); self.login_button = QPushButton("登入並連線"); self.login_button.setObjectName("primaryButton"); self.login_button.clicked.connect(self.login)
        login.addWidget(account_label); login.addWidget(self.user_id); login.addWidget(self.user_password); login.addWidget(self.login_button); login.addStretch(); page.addWidget(login_frame)

        workspace = QHBoxLayout(); workspace.setSpacing(12)
        sidebar = QFrame(); sidebar.setObjectName("sidebar"); sidebar.setFixedWidth(150); nav = QVBoxLayout(sidebar); nav.setContentsMargins(8, 10, 8, 10); nav.setSpacing(6)
        self.nav_group = QButtonGroup(self); self.nav_group.setExclusive(True); self.main_pages = QStackedWidget(); self.nav_buttons: list[QPushButton] = []
        for index, label in enumerate(("看盤視窗", "漲停監控", "ETF 折溢價")):
            button = QPushButton(label); button.setObjectName("navButton"); button.setCheckable(True); button.setMinimumHeight(44); button.clicked.connect(lambda checked=False, page_index=index: self._switch_page(page_index)); self.nav_group.addButton(button, index); self.nav_buttons.append(button); nav.addWidget(button)
        nav.addStretch(); self.nav_buttons[0].setChecked(True); workspace.addWidget(sidebar)

        dashboard = QWidget(); dashboard_layout = QVBoxLayout(dashboard); dashboard_layout.setContentsMargins(0, 0, 0, 0); dashboard_layout.setSpacing(9)
        self.tabs = QTabWidget(); self.tabs.setObjectName("marketTabs"); self.future_page = QuotePage("TAIFEX · NEAR MONTH", "台指期", FUTURE_SYMBOL); self.index_page = QuotePage("TWSE · INDEX", "發行量加權股價指數", INDEX_SYMBOL); self.tabs.addTab(self.future_page, "台指期"); self.tabs.addTab(self.index_page, "加權指數")
        holdings = QFrame(); holdings.setObjectName("panel"); hv = QVBoxLayout(holdings); hv.setContentsMargins(15, 13, 15, 12); hh = QHBoxLayout(); title = QLabel("庫存持股"); title.setObjectName("sectionTitle"); hh.addWidget(title); hh.addStretch(); self.count = QLabel("0 檔"); self.count.setObjectName("badge"); hh.addWidget(self.count); hv.addLayout(hh); self.table = QTableWidget(0, 6); self.table.setHorizontalHeaderLabels(["商品", "代碼", "現價", "漲跌幅", "庫存", "可賣"]); self.table.verticalHeader().setVisible(False); self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch); self.table.setAlternatingRowColors(True); hv.addWidget(self.table); self.inventory_source = QLabel("登入後載入庫存"); self.inventory_source.setObjectName("source"); hv.addWidget(self.inventory_source)
        order = self._order_panel(); daily = self._daily_panel()
        self.top_splitter = QSplitter(Qt.Orientation.Horizontal); self.top_splitter.setObjectName("panelSplitter"); self.top_splitter.setChildrenCollapsible(False); self.top_splitter.addWidget(holdings); self.top_splitter.addWidget(self.tabs); self.top_splitter.setStretchFactor(0, 2); self.top_splitter.setStretchFactor(1, 3); self.top_splitter.setSizes([460, 700])
        self.bottom_splitter = QSplitter(Qt.Orientation.Horizontal); self.bottom_splitter.setObjectName("panelSplitter"); self.bottom_splitter.setChildrenCollapsible(False); self.bottom_splitter.addWidget(order); self.bottom_splitter.addWidget(daily); self.bottom_splitter.setStretchFactor(0, 2); self.bottom_splitter.setStretchFactor(1, 3); self.bottom_splitter.setSizes([460, 700])
        self.dashboard_splitter = QSplitter(Qt.Orientation.Vertical); self.dashboard_splitter.setObjectName("panelSplitter"); se…8733 tokens truncated…c:
            self.events.put(("limit_scan_error", str(exc)))

    def _chase_limit_order(self, row: dict[str, Any]) -> None:
        if not self.connected: QMessageBox.warning(self, "尚未登入", "請先登入後再送單。"); return
        symbol, lots = str(row.get("symbol", "")), self.limit_order_quantity.value(); unit = int(row.get("unit") or 1000); quantity = lots * unit
        warning = f"商品：{row.get('name', symbol)} ({symbol})\n方向：買進\n價格：市價 ROD\n數量：{lots:,} 張（{quantity:,} 股／單位）\n目前成交：{fmt(row.get('last'))}\n漲停價：{fmt(row.get('limit'))}\n\n市價單不保證成交價格，若漲停打開可能以其他價格成交。確定送出？"
        if QMessageBox.question(self, "市價追漲停確認", warning, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes: return
        self.limit_error.setText(f"程式執行訊息：正在送出 {symbol} 市價買進 {lots:,} 張……"); self.limit_table.setEnabled(False); self.limit_chase_selected_button.setEnabled(False); threading.Thread(target=self._order_worker, args=("股票現股", symbol, "買進", "市價", "", quantity, "追漲停", None), daemon=True).start()

    def _limit_selection_changed(self) -> None:
        self.limit_chase_selected_button.setEnabled(self.connected and self.limit_table.currentRow() >= 0 and self.limit_table.isEnabled())

    def _chase_selected_limit(self) -> None:
        row_index = self.limit_table.currentRow(); symbol_item = self.limit_table.item(row_index, 1) if row_index >= 0 else None
        if symbol_item is None: QMessageBox.warning(self, "尚未選取標的", "請先點選表格中的一檔股票。"); return
        row = self.limit_rows_by_symbol.get(symbol_item.text())
        if row is None: QMessageBox.warning(self, "資料已更新", "選取資料已變更，請重新選取後再送單。"); return
        self._chase_limit_order(row)

    def _http_worker(self) -> None:
        rest = self.sdk.marketdata.rest_client; errors = []
        for row in self.inventory:
            try: self.events.put(("stock", unwrap(rest.stock.intraday.quote(symbol=row["symbol"]))))
            except Exception as exc: errors.append(f"{row['symbol']}: {exc}")
        try: self.events.put(("future", (unwrap(rest.futopt.intraday.quote(symbol=FUTURE_SYMBOL)), "HTTP 當日行情")))
        except Exception as exc: errors.append(f"台指期: {exc}")
        try:
            candles = rest.futopt.intraday.candles(symbol=FUTURE_SYMBOL, timeframe="1"); rows = candles.get("data", []) if isinstance(candles, dict) else []; self.events.put(("future_chart", [value for value in (as_float(row.get("close")) for row in rows) if value is not None]))
        except Exception: pass
        self._load_index_http(rest, errors)
        symbol = self.inventory[0]["symbol"] if self.inventory else self.chart_symbol.text().strip() or "2330"; self.events.put(("auto_chart", symbol)); self.events.put(("http_done", errors))

    def _load_index_http(self, rest: Any, errors: list[str]) -> None:
        end = date.today(); start = end - timedelta(days=30); daily_rows = []
        try:
            try: historical = rest.stock.historical.candles(**{"symbol": INDEX_SYMBOL, "from": start.isoformat(), "to": end.isoformat(), "timeframe": "D", "sort": "desc", "fields": "open,high,low,close,change"})
            except Exception: historical = rest.stock.historical.candles(**{"symbol": INDEX_SYMBOL, "from": start.isoformat(), "to": end.isoformat(), "timeframe": "D", "sort": "desc"})
            daily_rows = historical.get("data", []) if isinstance(historical, dict) else []
            latest = daily_rows[0] if daily_rows else {}; close = as_float(latest.get("close")); change = as_float(latest.get("change")); latest_date = str(latest.get("date", ""))[:10]; historical_previous = close - change if close is not None and change is not None else (as_float(daily_rows[1].get("close")) if len(daily_rows) > 1 else None); live_previous = historical_previous if latest_date == end.isoformat() else close; self.events.put(("index_reference", live_previous))
            if close is not None: self.events.put(("index_fallback", (close, historical_previous, "歷史日 K 收盤")))
        except Exception as exc: errors.append(f"指數歷史: {exc}")
        intraday_values: list[float] = []
        try:
            intraday = rest.stock.intraday.candles(symbol=INDEX_SYMBOL, timeframe="1", sort="asc"); rows = intraday.get("data", []) if isinstance(intraday, dict) else []; values = [value for value in (as_float(row.get("close")) for row in rows) if value is not None]
            if values: intraday_values = values; self.events.put(("index_chart", values)); self.events.put(("index", (values[-1], "當日 1 分 K")))
        except Exception: pass
        try:
            quote = unwrap(rest.stock.intraday.quote(symbol=INDEX_SYMBOL)); candidate = quote.get("index")
            if candidate is None and quote.get("type") == "INDEX": candidate = quote.get("lastPrice") if quote.get("lastPrice") is not None else quote.get("closePrice")
            if candidate is not None: self.events.put(("index", (candidate, "HTTP 指數行情")))
        except Exception: pass
        if daily_rows and not intraday_values: self.events.put(("index_chart", [value for value in (as_float(row.get("close")) for row in reversed(daily_rows)) if value is not None]))

    def _ws_worker(self) -> None:
        try:
            stock = self.sdk.marketdata.websocket_client.stock; future = self.sdk.marketdata.websocket_client.futopt; stock.on("message", lambda message: self._ws_message("stock", message)); future.on("message", lambda message: self._ws_message("future", message)); stock.connect(); future.connect()
            for row in self.inventory:
                if row["symbol"]: stock.subscribe({"channel": "aggregates", "symbol": row["symbol"]})
            stock.subscribe({"channel": "indices", "symbol": INDEX_SYMBOL}); future.subscribe({"channel": "aggregates", "symbol": FUTURE_SYMBOL})
            try: future.subscribe({"channel": "aggregates", "symbol": FUTURE_SYMBOL, "afterHours": True})
            except Exception: pass
            self.events.put(("status", "WebSocket 即時行情已連線"))
        except Exception as exc: self.events.put(("status", f"WebSocket：{exc}"))

    def _ws_message(self, market: str, message: Any) -> None:
        try:
            envelope = json.loads(message) if isinstance(message, str) else message
            if not isinstance(envelope, dict) or envelope.get("event") != "data" or not isinstance(envelope.get("data"), dict): return
            data = envelope["data"]
            if market == "stock" and (envelope.get("channel") == "indices" or data.get("type") == "INDEX"): self.events.put(("index", (data.get("index"), "WebSocket indices 即時")))
            elif market == "stock": self.events.put(("stock", data))
            else: self.events.put(("future", (data, "WebSocket aggregates 即時")))
        except (TypeError, ValueError, json.JSONDecodeError): pass

    def load_daily_chart(self) -> None:
        if not self.connected: QMessageBox.information(self, "尚未登入", "請先登入。"); return
        symbol = self.chart_symbol.text().strip().upper()
        if symbol: threading.Thread(target=self._daily_worker, args=(symbol,), daemon=True).start()

    def _daily_worker(self, symbol: str) -> None:
        try:
            end = date.today(); start = end - timedelta(days=180); payload = self.sdk.marketdata.rest_client.stock.historical.candles(**{"symbol": symbol, "from": start.isoformat(), "to": end.isoformat(), "timeframe": "D", "sort": "asc", "fields": "open,high,low,close,volume,change"}); rows = payload.get("data", []) if isinstance(payload, dict) else []; self.events.put(("daily", (symbol, rows)))
        except Exception as exc: self.events.put(("status", f"日 K 載入失敗：{exc}"))

    def _order_market_changed(self, market: str) -> None:
        future = market == "台指期貨"; self.order_symbol.setText(FUTURE_SYMBOL if future else "2330"); self.order_quantity.setValue(1 if future else 1000); self.preorder_button.setEnabled(not future)
        if future and self.order_mode_group.checkedButton().text() == "預掛": self.order_mode_group.buttons()[0].click()

    def _order_price_type_changed(self, price_type: str) -> None:
        self.order_price.setEnabled(price_type == "限價")
        if price_type == "市價": self.order_price.clear()

    def _order_mode_changed(self, mode: str) -> None:
        preorder = mode == "預掛"; self.market_price_button.setEnabled(not preorder)
        if preorder:
            self.limit_price_button.click(); self.order_button.setText("加入預掛佇列"); self.order_button.setToolTip("先保存在本機，下一個工作日 08:31 登入後自動送出")
        else:
            self.order_button.setText("送出委託"); self.order_button.setToolTip("")

    def submit_order(self) -> None:
        if not self.connected: QMessageBox.warning(self, "尚未登入", "請先登入。"); return
        market_button, side_button, price_type_button = self.order_market_group.checkedButton(), self.order_side_group.checkedButton(), self.order_price_type_group.checkedButton()
        market, side, price_type, mode = market_button.text(), side_button.text(), price_type_button.text(), self.order_mode_group.checkedButton().text()
        symbol, order_price, quantity = self.order_symbol.text().strip().upper(), self.order_price.text().strip(), self.order_quantity.value()
        if mode == "預掛":
            if market != "股票現股": QMessageBox.warning(self, "不支援預掛", "富邦期貨不提供預約單，請改選股票現股或一般委託。"); return
            if price_type != "限價": QMessageBox.warning(self, "預掛資料錯誤", "預掛單必須使用限價。"); return
        if price_type == "限價" and (as_float(order_price) is None or as_float(order_price) <= 0): QMessageBox.warning(self, "委託資料錯誤", "限價單請輸入有效價格。"); return
        summary = f"環境：{self.environment_name}\n送單方式：{mode}\n市場：{market}\n商品：{symbol}\n方向：{side}\n價格：{price_type} {order_price or '市價'}\n數量：{quantity}"
        question = "\n\n確定加入本機預掛佇列？" if mode == "預掛" else "\n\n確定送出委託？"
        if QMessageBox.question(self, "預掛確認" if mode == "預掛" else "送單前確認", summary + question, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes: return
        if mode == "預掛":
            self._queue_preorder(market, symbol, side, price_type, order_price, quantity); return
        self.order_button.setEnabled(False); threading.Thread(target=self._order_worker, args=(market, symbol, side, price_type, order_price, quantity, "一般", None), daemon=True).start()

    def _order_worker(self, market: str, symbol: str, side: str, price_type: str, order_price: str, quantity: int, mode: str, pending_id: str | None = None) -> None:
        try:
            action = BSAction.Buy if side == "買進" else BSAction.Sell
            if market == "股票現股":
                order = Order(buy_sell=action, symbol=symbol, quantity=quantity, market_type=MarketType.Common, price_type=PriceType.Limit if price_type == "限價" else PriceType.Market, time_in_force=TimeInForce.ROD, order_type=OrderType.Stock, price=order_price or None, user_def="QueuedUI" if mode == "預掛執行" else "MarketUI"); result = self.sdk.stock.place_order(self.account, order)
            else:
                order = FutOptOrder(market_type=FutOptMarketType.Future, price_type=FutOptPriceType.Limit if price_type == "限價" else FutOptPriceType.Market, time_in_force=TimeInForce.ROD, order_type=FutOptOrderType.New, buy_sell=action, symbol=symbol, lot=quantity, price=order_price or None, user_def="MarketUI"); result = self.sdk.futopt.place_order(self.account, order)
            self.events.put(("order_result", (result, mode, pending_id)))
        except Exception as exc: self.events.put(("order_error", (str(exc), pending_id, mode)))

    def _show_order_result(self, result: Any, mode: str, pending_id: str | None) -> None:
        if not getattr(result, "is_success", False):
            message = str(getattr(result, "message", None) or result)
            if pending_id and "時間未到" in message:
                self._reschedule_pending_order(pending_id); message += "\n\n券商尚未開放收單，已自動順延至下一個工作日 08:31。"
            elif pending_id:
                self._fail_pending_order(pending_id, message)
            if mode == "追漲停": self.limit_error.setText(f"程式執行失敗原因：{message}")
            QMessageBox.warning(self, "委託未受理", message); return
        payload = getattr(result, "data", None); status = getattr(payload, "status", None); order_no = getattr(payload, "order_no", None) or "—"
        if pending_id:
            self._remove_pending_order(pending_id); label = "預掛委託已送出"
        else:
            label = "追漲停委託成功" if mode == "追漲停" else "委託成功"
        if mode == "追漲停": self.limit_error.setText(f"程式執行訊息：券商已受理市價追單，委託書號 {order_no}。")
        QMessageBox.information(self, label, f"{label}\n狀態：{status}\n委託書號：{order_no}")

    def _drain(self) -> None:
        try:
            while True:
                kind, data = self.events.get_nowait()
                if kind == "login_ok":
                    self.sdk, self.account, self.inventory, environment = data; self.connected = True; self.login_button.setEnabled(True); self.order_button.setEnabled(True); self.limit_scan_button.setEnabled(True); self._limit_selection_changed(); self.dot.setObjectName("onlineDot"); self.dot.style().unpolish(self.dot); self.dot.style().polish(self.dot); self.status.setText(f"{environment} · 已登入"); self._render_inventory(); self._start_data(); self.scan_limit_monitor(); self._check_pending_orders()
                elif kind == "login_error":
                    self.login_button.setEnabled(True); self.status.setText("登入失敗"); message = str(data)
                    if "certificate key error" in message.lower(): self._clear_certificate_password(); message = "憑證密碼不正確，已清除先前保存的憑證密碼。\n請再次登入並輸入正確的憑證密碼。"
                    QMessageBox.critical(self, "登入失敗", message)
                elif kind == "stock":
                    quote = unwrap(data); symbol = str(quote.get("symbol", ""));
                    if symbol: self.quotes[symbol] = quote; self._render_inventory()
                elif kind == "future": self.future_page.update_quote(unwrap(data[0]), data[1])
                elif kind == "future_chart": self.future_page.chart.show_values(data)
                elif kind == "index_reference": self.previous_index_close = as_float(data)
                elif kind == "index_fallback": self.index_page.update_index(data[0], data[1], data[2])
                elif kind == "index": self.index_page.update_index(data[0], self.previous_index_close, data[1])
                elif kind == "index_chart": self.index_page.chart.show_values(data)
                elif kind == "auto_chart": self.chart_symbol.setText(data); self.load_daily_chart()
                elif kind == "daily": self.daily_chart.show_candles(data[0], data[1]); self.daily_source.setText(f"富邦 historical/candles · {len(data[1])} 根日 K")
                elif kind == "http_done": self.inventory_source.setText("HTTP 當日行情 + WebSocket aggregates"); self.footer.setText("部分 HTTP 查詢失敗，WebSocket 仍持續更新" if data else "HTTP 初始資料完成，WebSocket 持續更新")
                elif kind == "status": self.footer.setText(str(data))
                elif kind == "limit_scan": self.limit_scan_running = False; self.limit_scan_button.setEnabled(True); self._render_limit_monitor(*data)
                elif kind == "limit_scan_error": self.limit_scan_running = False; self.limit_scan_button.setEnabled(True); self.limit_status.setText("掃描未完成"); self.limit_error.setText(f"程式執行失敗原因：{data}")
                elif kind == "order_result": self.order_button.setEnabled(True); self.limit_table.setEnabled(True); self._limit_selection_changed(); self._show_order_result(*data)
                elif kind == "order_error":
                    self.order_button.setEnabled(True); self.limit_table.setEnabled(True); message, pending_id, mode = data; self._limit_selection_changed()
                    if pending_id: self._fail_pending_order(pending_id, message)
                    if mode == "追漲停": self.limit_error.setText(f"程式執行失敗原因：{message}")
                    QMessageBox.critical(self, "下單失敗", message)
        except queue.Empty: pass

    def _render_inventory(self) -> None:
        self.table.setRowCount(len(self.inventory)); self.count.setText(f"{len(self.inventory)} 檔")
        for row_index, row in enumerate(self.inventory):
            quote = self.quotes.get(row["symbol"], {}); change = as_float(quote.get("change")); percent = as_float(quote.get("changePercent")); values = [str(quote.get("name") or "—"), row["symbol"], fmt(quote.get("lastPrice") if quote.get("lastPrice") is not None else quote.get("closePrice")), "—" if percent is None else f"{percent:+.2f}%", str(row["quantity"]), str(row["tradable"])]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value); item.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | (Qt.AlignmentFlag.AlignLeft if column < 2 else Qt.AlignmentFlag.AlignRight))
                if column in (2, 3) and change is not None: item.setForeground(QColor("#ff667d" if change > 0 else "#35d3a3" if change < 0 else "#9aa8bf"))
                self.table.setItem(row_index, column, item)

    def _render_limit_monitor(self, rows: list[dict[str, Any]], scanned: int, updated: str, ticks: int, min_volume: int, category: str) -> None:
        self.limit_rows_by_symbol = {str(row["symbol"]): dict(row) for row in rows}; self.limit_table.setSortingEnabled(False); self.limit_table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            percent = row.get("percent"); values = [row["name"], row["symbol"], row["market"], fmt(row["last"]), fmt(row["limit"]), f'{row.get("ticks", ticks)} tick', f'{row["volume"]:,}', "—" if percent is None else f"{percent:+.2f}%", row["time"] or updated]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value)); item.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | (Qt.AlignmentFlag.AlignLeft if column < 3 else Qt.AlignmentFlag.AlignRight))
                if column in (3, 4, 7): item.setForeground(QColor("#ff667d"))
                self.limit_table.setItem(row_index, column, item)
            chase = QPushButton("市價買進"); chase.setObjectName("chaseButton"); chase.setToolTip(f"以市價 ROD 買進 {row['symbol']}"); chase.clicked.connect(lambda checked=False, payload=dict(row): self._chase_limit_order(payload)); self.limit_table.setCellWidget(row_index, 9, chase)
        self.limit_table.setSortingEnabled(True); self._limit_selection_changed(); self.limit_status.setText(f"{updated} 完成 · {category}掃描 {scanned:,} 檔 · 距漲停 {ticks} tick · 量 > {min_volume:,} 張 · 符合 {len(rows)} 檔 · 每 30 秒更新"); self.limit_error.setText("程式執行訊息：最近一次掃描正常，無錯誤。")


if __name__ == "__main__":
    application = QApplication([]); window = MarketWindow(); window.show(); application.exec()

