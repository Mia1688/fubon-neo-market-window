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
        self.dashboard_splitter = QSplitter(Qt.Orientation.Vertical); self.dashboard_splitter.setObjectName("panelSplitter"); self.dashboard_splitter.setChildrenCollapsible(False); self.dashboard_splitter.addWidget(self.top_splitter); self.dashboard_splitter.addWidget(self.bottom_splitter); self.dashboard_splitter.setStretchFactor(0, 3); self.dashboard_splitter.setStretchFactor(1, 2); self.dashboard_splitter.setSizes([480, 320]); dashboard_layout.addWidget(self.dashboard_splitter, 1)
        bottom = QHBoxLayout(); self.footer = QLabel("資料來源：尚未連線"); self.footer.setObjectName("muted"); bottom.addWidget(self.footer); bottom.addStretch(); refresh = QPushButton("重新整理"); refresh.clicked.connect(self.refresh); bottom.addWidget(refresh)
        dashboard_layout.addLayout(bottom); self.main_pages.addWidget(dashboard); self.main_pages.addWidget(self._limit_monitor_page()); self.main_pages.addWidget(self._monitor_page("ETF 折溢價", "比較 ETF 市價、淨值與即時折溢價", ["ETF", "代碼", "市價", "淨值", "折溢價", "更新時間"])); workspace.addWidget(self.main_pages, 1); page.addLayout(workspace, 1)

    def _switch_page(self, page_index: int) -> None:
        self.main_pages.setCurrentIndex(page_index)
        if page_index == 1: self.scan_limit_monitor()

    def _limit_monitor_page(self) -> QWidget:
        page = QWidget(); layout = QVBoxLayout(page); layout.setContentsMargins(0, 0, 0, 0); panel = QFrame(); panel.setObjectName("panel"); body = QVBoxLayout(panel); body.setContentsMargins(20, 18, 20, 18); top = QHBoxLayout(); heading = QVBoxLayout(); title = QLabel("漲停監控"); title.setObjectName("sectionTitle"); self.limit_rule_label = QLabel(); self.limit_rule_label.setObjectName("muted"); heading.addWidget(title); heading.addWidget(self.limit_rule_label); top.addLayout(heading); top.addStretch(); self.limit_chase_selected_button = QPushButton("市價追選取標的"); self.limit_chase_selected_button.setObjectName("chaseButton"); self.limit_chase_selected_button.setEnabled(False); self.limit_chase_selected_button.clicked.connect(self._chase_selected_limit); top.addWidget(self.limit_chase_selected_button); self.limit_scan_button = QPushButton("立即掃描"); self.limit_scan_button.clicked.connect(self.scan_limit_monitor); self.limit_scan_button.setEnabled(False); top.addWidget(self.limit_scan_button); body.addLayout(top)
        filters = QHBoxLayout(); filters.setSpacing(8); self.limit_category = QComboBox(); self.limit_category.setObjectName("limitCategory"); self.limit_category.addItems(["全部一般股", "台灣50", "台灣中型100", "台灣50 + 中型100", "小型股300", "上市一般股", "上櫃一般股", "ETF／ETN", "可轉債", "全部商品"]); self.limit_ticks_input = QSpinBox(); self.limit_ticks_input.setRange(1, 20); self.limit_ticks_input.setValue(1); self.limit_ticks_input.setSuffix(" tick"); self.limit_volume_input = QSpinBox(); self.limit_volume_input.setRange(0, 10_000_000); self.limit_volume_input.setValue(1000); self.limit_volume_input.setSingleStep(100); self.limit_volume_input.setSuffix(" 張"); self.limit_order_quantity = QSpinBox(); self.limit_order_quantity.setRange(1, 9_999); self.limit_order_quantity.setValue(1); self.limit_order_quantity.setSingleStep(1); self.limit_order_quantity.setSuffix(" 張")
        for label, widget in (("股票分類", self.limit_category), ("距漲停", self.limit_ticks_input), ("成交量大於", self.limit_volume_input), ("追單數量", self.limit_order_quantity)):
            filter_label = QLabel(label); filter_label.setObjectName("filterLabel"); filters.addWidget(filter_label); filters.addWidget(widget)
        filters.addStretch(); body.addLayout(filters); self.limit_category.currentTextChanged.connect(self._update_limit_rule_text); self.limit_ticks_input.valueChanged.connect(self._update_limit_rule_text); self.limit_volume_input.valueChanged.connect(self._update_limit_rule_text); self._update_limit_rule_text()
        self.limit_table = QTableWidget(0, 10); self.limit_table.setHorizontalHeaderLabels(["商品", "代碼", "市場", "成交價", "漲停價", "距離", "成交量（張）", "漲幅", "更新時間", "操作"]); self.limit_table.verticalHeader().setVisible(False); self.limit_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch); self.limit_table.setAlternatingRowColors(True); self.limit_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); self.limit_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); self.limit_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection); self.limit_table.itemSelectionChanged.connect(self._limit_selection_changed); self.limit_table.setSortingEnabled(True); body.addWidget(self.limit_table, 1); self.limit_status = QLabel("登入後開始掃描；每 30 秒自動更新"); self.limit_status.setObjectName("source"); body.addWidget(self.limit_status); self.limit_error = QLabel("程式執行訊息：尚無錯誤"); self.limit_error.setObjectName("executionMessage"); self.limit_error.setWordWrap(True); self.limit_error.setMinimumHeight(52); body.addWidget(self.limit_error); layout.addWidget(panel); return page

    def _update_limit_rule_text(self, *_: Any) -> None:
        self.limit_rule_label.setText(f"{self.limit_category.currentText()} · 成交價距漲停 {self.limit_ticks_input.value()} tick · 成交量 > {self.limit_volume_input.value():,} 張")

    def _monitor_page(self, title: str, subtitle: str, columns: list[str]) -> QWidget:
        page = QWidget(); layout = QVBoxLayout(page); layout.setContentsMargins(0, 0, 0, 0); panel = QFrame(); panel.setObjectName("panel"); body = QVBoxLayout(panel); body.setContentsMargins(20, 18, 20, 18); heading = QLabel(title); heading.setObjectName("sectionTitle"); description = QLabel(subtitle); description.setObjectName("muted"); table = QTableWidget(0, len(columns)); table.setHorizontalHeaderLabels(columns); table.verticalHeader().setVisible(False); table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch); table.setAlternatingRowColors(True); empty = QLabel("登入後等待即時資料"); empty.setObjectName("emptyState"); empty.setAlignment(Qt.AlignmentFlag.AlignCenter); body.addWidget(heading); body.addWidget(description); body.addWidget(table, 1); body.addWidget(empty); layout.addWidget(panel); return page

    def _order_panel(self) -> QFrame:
        panel = QFrame(); panel.setObjectName("orderPanel"); layout = QVBoxLayout(panel); layout.setContentsMargins(0, 0, 0, 0)
        self.order_tabs = QTabWidget(); self.order_tabs.setObjectName("orderTabs"); layout.addWidget(self.order_tabs)
        trade_page = QWidget(); trade_layout = QVBoxLayout(trade_page); trade_layout.setContentsMargins(16, 10, 16, 13); form = QGridLayout(); form.setSpacing(8)
        market_row = QWidget(); market_row.setObjectName("choiceRow"); market_layout = QHBoxLayout(market_row); market_layout.setContentsMargins(0, 0, 0, 0); market_layout.setSpacing(6); self.order_market_group = QButtonGroup(self); self.order_market_group.setExclusive(True)
        for index, text in enumerate(("股票現股", "台指期貨")):
            button = QPushButton(text); button.setObjectName("tradeOption"); button.setCheckable(True); button.setChecked(index == 0); self.order_market_group.addButton(button); market_layout.addWidget(button)
        self.order_market_group.buttonClicked.connect(lambda button: self._order_market_changed(button.text()))
        side_row = QWidget(); side_row.setObjectName("choiceRow"); side_layout = QHBoxLayout(side_row); side_layout.setContentsMargins(0, 0, 0, 0); side_layout.setSpacing(6); self.order_side_group = QButtonGroup(self); self.order_side_group.setExclusive(True)
        for index, (text, name) in enumerate((("買進", "buyButton"), ("賣出", "sellButton"))):
            button = QPushButton(text); button.setObjectName(name); button.setCheckable(True); button.setChecked(index == 0); self.order_side_group.addButton(button); side_layout.addWidget(button)
        price_type_row = QWidget(); price_type_row.setObjectName("choiceRow"); price_type_layout = QHBoxLayout(price_type_row); price_type_layout.setContentsMargins(0, 0, 0, 0); price_type_layout.setSpacing(6); self.order_price_type_group = QButtonGroup(self); self.order_price_type_group.setExclusive(True)
        for index, text in enumerate(("限價", "市價")):
            button = QPushButton(text); button.setObjectName("tradeOption"); button.setCheckable(True); button.setChecked(index == 0); self.order_price_type_group.addButton(button); price_type_layout.addWidget(button)
            if text == "限價": self.limit_price_button = button
            else: self.market_price_button = button
        self.order_price_type_group.buttonClicked.connect(lambda button: self._order_price_type_changed(button.text()))
        mode_row = QWidget(); mode_row.setObjectName("choiceRow"); mode_layout = QHBoxLayout(mode_row); mode_layout.setContentsMargins(0, 0, 0, 0); mode_layout.setSpacing(6); self.order_mode_group = QButtonGroup(self); self.order_mode_group.setExclusive(True)
        for index, text in enumerate(("一般", "預掛")):
            button = QPushButton(text); button.setObjectName("preorderButton" if text == "預掛" else "tradeOption"); button.setCheckable(True); button.setChecked(index == 0); self.order_mode_group.addButton(button); mode_layout.addWidget(button)
            if text == "預掛": self.preorder_button = button
        self.order_mode_group.buttonClicked.connect(lambda button: self._order_mode_changed(button.text()))
        self.order_symbol = QLineEdit("2330"); self.order_price = QLineEdit(); self.order_price.setPlaceholderText("價格"); self.order_quantity = QSpinBox(); self.order_quantity.setRange(1, 999999); self.order_quantity.setValue(1000)
        fields = [("市場", market_row), ("商品", self.order_symbol), ("方向", side_row), ("價格類型", price_type_row), ("送單方式", mode_row), ("委託價", self.order_price), ("數量／口數", self.order_quantity)]
        for index, (label, widget) in enumerate(fields): form.addWidget(QLabel(label), index // 2 * 2, index % 2 * 2); form.addWidget(widget, index // 2 * 2 + 1, index % 2 * 2)
        trade_layout.addLayout(form); self.order_button = QPushButton("送出委託"); self.order_button.setObjectName("primaryButton"); self.order_button.setEnabled(False); self.order_button.clicked.connect(self.submit_order); trade_layout.addWidget(self.order_button)

        pending_page = QWidget(); pending_layout = QVBoxLayout(pending_page); pending_layout.setContentsMargins(12, 10, 12, 12); pending_layout.setSpacing(8)
        self.pending_table = QTableWidget(0, 7); self.pending_table.setHorizontalHeaderLabels(["勾選", "方向", "商品", "價格", "數量", "預計送出", "狀態"]); self.pending_table.verticalHeader().setVisible(False); self.pending_table.setAlternatingRowColors(True); self.pending_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); self.pending_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        header = self.pending_table.horizontalHeader()
        for column in (0, 1, 2, 4, 6): header.setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch); header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        pending_layout.addWidget(self.pending_table, 1)
        pending_row = QHBoxLayout(); self.pending_label = QLabel("本機預掛 0 筆"); self.pending_label.setObjectName("source"); self.select_all_pending_button = QPushButton("全選可取消"); self.select_all_pending_button.clicked.connect(self._select_all_pending); self.cancel_selected_button = QPushButton("取消勾選掛單"); self.cancel_selected_button.setObjectName("dangerButton"); self.cancel_selected_button.setEnabled(False); self.cancel_selected_button.clicked.connect(self._cancel_selected_pending_orders); pending_row.addWidget(self.pending_label); pending_row.addStretch(); pending_row.addWidget(self.select_all_pending_button); pending_row.addWidget(self.cancel_selected_button); pending_layout.addLayout(pending_row)
        self.order_tabs.addTab(trade_page, "快速下單"); self.order_tabs.addTab(pending_page, "預掛查詢 (0)")
        return panel

    def _daily_panel(self) -> QFrame:
        panel = QFrame(); panel.setObjectName("panel"); layout = QVBoxLayout(panel); layout.setContentsMargins(16, 13, 16, 13); bar = QHBoxLayout(); title = QLabel("歷史日 K"); title.setObjectName("sectionTitle"); bar.addWidget(title); bar.addStretch(); self.chart_symbol = QLineEdit("2330"); self.chart_symbol.setMaximumWidth(90); bar.addWidget(self.chart_symbol); load = QPushButton("載入"); load.clicked.connect(self.load_daily_chart); bar.addWidget(load); layout.addLayout(bar); self.daily_chart = CandleChart(); layout.addWidget(self.daily_chart, 1); self.daily_source = QLabel("登入後自動載入第一檔持股"); self.daily_source.setObjectName("source"); layout.addWidget(self.daily_source); return panel

    def _theme(self) -> None:
        self.setStyleSheet("""
        QWidget#root{background:#070b16;color:#e8eef9;font-family:'Microsoft JhengHei','Segoe UI';font-size:10pt} QLabel#brand{font-size:25pt;font-weight:800;color:#f7f9ff} QLabel#muted{color:#7f8da8} QLabel#eyebrow{color:#6f8bb8;font-size:8pt;font-weight:700} QLabel#filterLabel{color:#ffffff;font-weight:700} QLabel#sectionTitle{font-size:15pt;font-weight:700;color:#edf3ff} QLabel#heroPrice{font-size:34pt;font-weight:800;color:#f8fbff} QLabel#upChange{color:#ff667d;font-size:13pt;font-weight:700} QLabel#downChange{color:#35d3a3;font-size:13pt;font-weight:700} QLabel#flatChange{color:#8b99b2;font-size:13pt;font-weight:700} QLabel#source{color:#65738d;font-size:9pt} QLabel#executionMessage{background:#160f19;color:#ffb0bf;border:1px solid #633142;border-radius:7px;padding:9px 11px;font-weight:600} QLabel#badge{background:#1b2b4a;color:#8fb2ff;padding:5px 10px;border-radius:6px;font-weight:700} QLabel#offlineDot{color:#58657a} QLabel#onlineDot{color:#35d3a3}
        QFrame#loginBar,QFrame#panel,QFrame#orderPanel,QTabWidget#marketTabs::pane{background:#101828;border:1px solid #22314d;border-radius:11px} QTabWidget#orderTabs::pane{background:#101828;border:0;border-top:1px solid #22314d} QTabBar::tab{background:#0c1423;color:#7f91ad;padding:9px 22px;border:1px solid #22314d} QTabBar::tab:selected{background:#1a2b49;color:#ddebff;border-bottom:2px solid #4d7ff3}
        QFrame#sidebar{background:#0c1322;border:1px solid #22314d;border-radius:11px} QPushButton#navButton{background:transparent;color:#8999b4;border:0;border-radius:7px;text-align:left;padding:10px 13px} QPushButton#navButton:hover{background:#14213a;color:#dbe8ff} QPushButton#navButton:checked{background:#1c3156;color:#ffffff;border-left:3px solid #4d7ff3} QLabel#emptyState{color:#61708b;padding:14px}
        QFrame#orderPanel QLabel{color:#ffffff} QFrame#orderPanel QLineEdit,QFrame#orderPanel QSpinBox{color:#ffffff} QWidget#choiceRow{background:transparent}
        QPushButton#tradeOption{background:#16243c;color:#dce8fa;border:1px solid #2b3e60} QPushButton#tradeOption:checked{background:#315fae;color:#ffffff;border:1px solid #6b9aff}
        QPushButton#preorderButton{background:#382a15;color:#ffd38a;border:1px solid #705326} QPushButton#preorderButton:checked{background:#a96b17;color:#ffffff;border:1px solid #ffc65e}
        QPushButton#buyButton{background:#3a1822;color:#ff9cac;border:1px solid #6e2a3a} QPushButton#buyButton:checked{background:#a92f46;color:#ffffff;border:1px solid #ff7187}
        QPushButton#sellButton{background:#12352d;color:#77dfbf;border:1px solid #236c58} QPushButton#sellButton:checked{background:#17795e;color:#ffffff;border:1px solid #48d4ac}
        QSplitter#panelSplitter::handle{background:#070b16} QSplitter#panelSplitter::handle:horizontal{width:9px} QSplitter#panelSplitter::handle:vertical{height:9px} QSplitter#panelSplitter::handle:hover{background:#315fae}
        QLineEdit,QSpinBox,QComboBox{background:#0a1120;border:1px solid #293957;border-radius:6px;color:#edf3ff;padding:7px 9px} QComboBox QAbstractItemView{background:#101828;color:#edf3ff;selection-background-color:#315fae} QComboBox#limitCategory{background:#0a1120;color:#ffffff;font-weight:700} QComboBox#limitCategory QAbstractItemView{background:#101828;color:#ffffff;selection-background-color:#315fae} QPushButton{background:#182640;color:#bed1f2;border:1px solid #2c4167;border-radius:6px;padding:7px 12px;font-weight:600} QPushButton:hover{background:#23395d} QPushButton:disabled{color:#536078;border-color:#202a3d} QPushButton#primaryButton{background:#3d73ed;color:white;border:0} QPushButton#primaryButton:hover{background:#5388fb} QPushButton#dangerButton{background:#46202a;color:#ff9cac;border:1px solid #7e3142} QPushButton#dangerButton:hover{background:#692a39;color:#ffffff} QPushButton#chaseButton{background:#8d263b;color:#ffffff;border:1px solid #e45b72;padding:5px 9px} QPushButton#chaseButton:hover{background:#b7354f}
        QTableWidget{background:#0b1322;alternate-background-color:#0e192b;color:#dce7f7;border:0;gridline-color:#1e2c44} QHeaderView::section{background:#182641;color:#9fb6da;border:0;padding:7px;font-weight:700}
        """)

    def _load_settings(self) -> None:
        try:
            settings = json.loads(SETTINGS_FILE.read_text(encoding="utf-8")); self.cert_path = str(settings.get("cert_path", "")); protected = str(settings.get("protected_cert_password", "")); self.cert_password_override = unprotect_secret(protected) if protected else ""
        except (OSError, ValueError, TypeError):
            self.cert_path = ""; self.cert_password_override = ""

    def _save_settings(self) -> None:
        try:
            settings = {"cert_path": self.cert_path}
            if self.cert_password_override: settings["protected_cert_password"] = protect_secret(self.cert_password_override)
            SETTINGS_FILE.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError: pass

    def _setup_certificate(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "一次性憑證設定", self.cert_path, "PFX 憑證 (*.pfx);;所有檔案 (*.*)")
        if not path: return
        password, ok = QInputDialog.getText(self, "一次性憑證設定", "請輸入憑證密碼（由 Windows 加密保存）：", QLineEdit.EchoMode.Password)
        if ok:
            self.cert_path = path; self.cert_password_override = password; self._save_settings(); self.status.setText("憑證路徑已設定")

    def _ask_certificate_password(self) -> bool:
        if self.cert_password_override: return True
        password, ok = QInputDialog.getText(self, "一次性憑證設定", "請輸入憑證密碼（由 Windows 加密保存）：", QLineEdit.EchoMode.Password)
        if not ok or not password: return False
        self.cert_password_override = password; self._save_settings(); return True

    def _clear_certificate_password(self) -> None:
        self.cert_password_override = ""
        try: SETTINGS_FILE.write_text(json.dumps({"cert_path": self.cert_path}, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError: pass

    def _load_pending_orders(self) -> None:
        try:
            payload = json.loads(PENDING_FILE.read_text(encoding="utf-8")); self.pending_orders = [row for row in payload if isinstance(row, dict)] if isinstance(payload, list) else []
        except (OSError, ValueError, TypeError): self.pending_orders = []
        self._update_pending_label()

    def _save_pending_orders(self) -> None:
        try: PENDING_FILE.write_text(json.dumps(self.pending_orders, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError: pass
        self._update_pending_label()

    def _update_pending_label(self) -> None:
        queued = sum(row.get("status") == "queued" for row in self.pending_orders); errors = sum(row.get("status") == "error" for row in self.pending_orders); sending = sum(row.get("status") == "sending" for row in self.pending_orders); text = f"等待送出 {queued} 筆"
        if sending: text += f" · 送出中 {sending} 筆"
        if errors: text += f" · 待確認 {errors} 筆"
        self.pending_label.setText(text); self.order_tabs.setTabText(1, f"預掛查詢 ({len(self.pending_orders)})"); self.select_all_pending_button.setEnabled(any(row.get("status") != "sending" for row in self.pending_orders)); self.cancel_selected_button.setEnabled(any(row.get("status") != "sending" for row in self.pending_orders)); self._render_pending_orders()

    def _render_pending_orders(self) -> None:
        self.pending_table.setRowCount(len(self.pending_orders))
        status_labels = {"queued": "等待送出", "sending": "送出中", "error": "待確認"}
        for row_index, row in enumerate(self.pending_orders):
            status = str(row.get("status", "queued")); cancellable = status != "sending"; checkbox = QTableWidgetItem(); checkbox.setData(Qt.ItemDataRole.UserRole, str(row.get("id", ""))); checkbox.setCheckState(Qt.CheckState.Unchecked)
            flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
            if cancellable: flags |= Qt.ItemFlag.ItemIsUserCheckable
            checkbox.setFlags(flags); self.pending_table.setItem(row_index, 0, checkbox)
            try: execute_at = datetime.fromisoformat(str(row.get("execute_at", ""))).strftime("%m/%d %H:%M")
            except ValueError: execute_at = "時間錯誤"
            price = f'{row.get("price_type", "限價")} {row.get("price") or "—"}'
            values = [str(row.get("side", "—")), str(row.get("symbol", "—")), price, str(row.get("quantity", "—")), execute_at, status_labels.get(status, status)]
            for column, value in enumerate(values, 1):
                item = QTableWidgetItem(value); item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if column == 1: item.setForeground(QColor("#ff667d" if value == "買進" else "#35d3a3"))
                if column == 6 and status == "error": item.setForeground(QColor("#ffb35c")); item.setToolTip(str(row.get("error", "未知錯誤")))
                self.pending_table.setItem(row_index, column, item)

    def _select_all_pending(self) -> None:
        for row_index in range(self.pending_table.rowCount()):
            item = self.pending_table.item(row_index, 0)
            if item and item.flags() & Qt.ItemFlag.ItemIsUserCheckable: item.setCheckState(Qt.CheckState.Checked)

    def _cancel_selected_pending_orders(self) -> None:
        selected_ids: set[str] = set()
        for row_index in range(self.pending_table.rowCount()):
            item = self.pending_table.item(row_index, 0)
            if item and item.checkState() == Qt.CheckState.Checked: selected_ids.add(str(item.data(Qt.ItemDataRole.UserRole)))
        cancellable_ids = {str(row.get("id")) for row in self.pending_orders if row.get("status") != "sending"}
        selected_ids &= cancellable_ids
        if not selected_ids: QMessageBox.information(self, "尚未勾選", "請先勾選要取消的本機預掛。"); return
        if QMessageBox.question(self, "取消勾選掛單", f"確定取消勾選的 {len(selected_ids)} 筆本機預掛？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes: return
        self.pending_orders = [row for row in self.pending_orders if str(row.get("id")) not in selected_ids]; self._save_pending_orders()

    def _queue_preorder(self, market: str, symbol: str, side: str, price_type: str, order_price: str, quantity: int) -> None:
        execute_at = next_stock_order_window(); self.pending_orders.append({"id": uuid.uuid4().hex, "market": market, "symbol": symbol, "side": side, "price_type": price_type, "price": order_price, "quantity": quantity, "execute_at": execute_at.isoformat(timespec="seconds"), "status": "queued"}); self._save_pending_orders(); self.order_tabs.setCurrentIndex(1); QMessageBox.information(self, "已加入本機預掛", f"委託尚未送到券商。\n預計送出：{execute_at:%Y-%m-%d %H:%M}\n\n可在「預掛查詢」勾選並取消。程式必須保持開啟並維持登入。")

    def _check_pending_orders(self) -> None:
        if not self.connected or any(row.get("status") == "sending" for row in self.pending_orders): return
        now = datetime.now()
        for row in self.pending_orders:
            if row.get("status") != "queued": continue
            try: due = datetime.fromisoformat(str(row["execute_at"]))
            except (KeyError, TypeError, ValueError): row["status"] = "error"; row["error"] = "預掛時間格式錯誤"; self._save_pending_orders(); continue
            if due > now: continue
            row["status"] = "sending"; self._save_pending_orders(); threading.Thread(target=self._order_worker, args=(row["market"], row["symbol"], row["side"], row["price_type"], row["price"], int(row["quantity"]), "預掛執行", row["id"]), daemon=True).start(); break

    def _remove_pending_order(self, pending_id: str) -> None:
        self.pending_orders = [row for row in self.pending_orders if row.get("id") != pending_id]; self._save_pending_orders()

    def _fail_pending_order(self, pending_id: str, message: str) -> None:
        for row in self.pending_orders:
            if row.get("id") == pending_id: row["status"] = "error"; row["error"] = message; break
        self._save_pending_orders()

    def _reschedule_pending_order(self, pending_id: str) -> None:
        execute_at = next_stock_order_window()
        for row in self.pending_orders:
            if row.get("id") == pending_id: row["status"] = "queued"; row["execute_at"] = execute_at.isoformat(timespec="seconds"); row["error"] = "券商尚未開放，已順延"; break
        self._save_pending_orders()

    def login(self) -> None:
        uid, password = self.user_id.text().strip(), self.user_password.text()
        if not uid or not password: QMessageBox.warning(self, "資料不足", "請輸入帳號與密碼。"); return
        if not self.cert_path or not Path(self.cert_path).is_file(): self._setup_certificate()
        if not self.cert_path or not Path(self.cert_path).is_file(): return
        if not self._ask_certificate_password(): return
        self.login_button.setEnabled(False); self.status.setText("登入中…"); cert_password = self.cert_password_override
        threading.Thread(target=self._login_worker, args=(self.environment_name, uid, password, cert_password), daemon=True).start()

    def _login_worker(self, environment: str, uid: str, password: str, cert_password: str) -> None:
        try:
            sdk = FubonSDK(30, 2, url=TEST_URL) if environment == "test" else FubonSDK(30, 2); result = sdk.login(uid, password, self.cert_path, cert_password)
            if not getattr(result, "is_success", False): raise RuntimeError(getattr(result, "message", None) or "登入失敗")
            accounts = list(getattr(result, "data", []) or [])
            if not accounts: raise RuntimeError("登入成功，但沒有可用帳戶")
            account = accounts[0]; inv = sdk.accounting.inventories(account); rows = []
            if getattr(inv, "is_success", False): rows = [{"symbol": str(getattr(item, "stock_no", "")), "quantity": getattr(item, "today_qty", 0), "tradable": getattr(item, "tradable_qty", 0)} for item in (getattr(inv, "data", []) or [])]
            sdk.init_realtime(Mode.Normal); self.events.put(("login_ok", (sdk, account, rows, environment)))
        except Exception as exc: self.events.put(("login_error", str(exc)))

    def _start_data(self) -> None:
        threading.Thread(target=self._http_worker, daemon=True).start(); threading.Thread(target=self._ws_worker, daemon=True).start()

    def refresh(self) -> None:
        if self.connected: threading.Thread(target=self._http_worker, daemon=True).start()

    def scan_limit_monitor(self) -> None:
        if not self.connected or self.limit_scan_running: return
        ticks, min_volume, category = self.limit_ticks_input.value(), self.limit_volume_input.value(), self.limit_category.currentText()
        self.limit_scan_running = True; self.limit_scan_button.setEnabled(False); self.limit_status.setText(f"正在掃描：{category} · 距漲停 {ticks} tick · 量 > {min_volume:,} 張…")
        threading.Thread(target=self._limit_scan_worker, args=(ticks, min_volume, category), daemon=True).start()

    def _index_groups(self, items: list[Any]) -> dict[str, set[str]]:
        cached: dict[str, Any] = {}
        try:
            loaded = json.loads(GROUPS_FILE.read_text(encoding="utf-8")); cached = loaded if isinstance(loaded, dict) else {}
        except (OSError, ValueError, TypeError): pass
        groups = cached.get("groups") if isinstance(cached.get("groups"), dict) else {}
        if str(cached.get("date")) == date.today().isoformat() and all(name in groups for name in ("台灣50", "台灣中型100", "台灣50 + 中型100", "小型股300")):
            return {name: set(map(str, symbols)) for name, symbols in groups.items()}
        try:
            def etf_constituents(ticker: str) -> set[str]:
                params = {"APIType": "ETFAPI", "CompanyName": "YUANTAFUNDS", "PageName": f"/tradeInfo/pcf/{ticker}", "DeviceId": "00000000-0000-0000-0000-000000000000", "FuncId": "PCF/Daily", "AppName": "ETF", "Device": "3", "Platform": "ETF", "ticker": ticker, "ndate": ""}
                payload = fetch_json(YUANTA_PCF_URL, params); rows = payload.get("InKind", {}).get("FundComposition", []) if isinstance(payload, dict) else []
                return {str(row.get("stkcd")) for row in rows if isinstance(row, dict) and str(row.get("stkcd", "")).isdigit()}
            taiwan50, mid100 = etf_constituents("0050"), etf_constituents("0051")
            company_rows = fetch_json(TWSE_COMPANY_URL); issued_shares: dict[str, int] = {}
            for row in company_rows if isinstance(company_rows, list) else []:
                try: issued_shares[str(row["公司代號"])] = int(str(row["已發行普通股數或TDR原股發行股數"]).replace(",", ""))
                except (KeyError, TypeError, ValueError): continue
            ranked: list[tuple[float, str]] = []
            for snapshot in items:
                symbol = str(getattr(snapshot, "symbol", "") or ""); market = str(getattr(snapshot, "market", "") or ""); reference = as_float(getattr(snapshot, "reference_price", None)); shares = issued_shares.get(symbol)
                if market in ("TAIEX", "TSE") and reference and shares: ranked.append((reference * shares, symbol))
            ranked.sort(reverse=True); small300 = {symbol for _, symbol in ranked[150:450]}
            if len(taiwan50) < 40 or len(mid100) < 80 or len(small300) < 250: raise RuntimeError("指數成分資料筆數不足")
            serializable = {"台灣50": sorted(taiwan50), "台灣中型100": sorted(mid100), "台灣50 + 中型100": sorted(taiwan50 | mid100), "小型股300": sorted(small300)}; GROUPS_FILE.write_text(json.dumps({"date": date.today().isoformat(), "updated": datetime.now().isoformat(timespec="seconds"), "groups": serializable}, ensure_ascii=False, indent=2), encoding="utf-8"); return {name: set(symbols) for name, symbols in serializable.items()}
        except Exception:
            if groups: return {name: set(map(str, symbols)) for name, symbols in groups.items()}
            raise

    def _limit_scan_worker(self, ticks: int, min_volume: int, category: str) -> None:
        try:
            stock_types = [StockType.EtfAndEtn] if category == "ETF／ETN" else [StockType.CovertBond] if category == "可轉債" else [StockType.Stock, StockType.EtfAndEtn, StockType.CovertBond] if category == "全部商品" else [StockType.Stock]
            result = self.sdk.stock.query_symbol_snapshot(self.account, MarketType.Common, stock_types)
            if not getattr(result, "is_success", False): raise RuntimeError(getattr(result, "message", None) or "批次行情查詢失敗")
            payload = getattr(result, "data", None); raw_items = getattr(payload, "symbols", payload); items = list(raw_items or []); matches: list[dict[str, Any]] = []; rest = self.sdk.marketdata.rest_client.stock; index_categories = {"台灣50", "台灣中型100", "台灣50 + 中型100", "小型股300"}; allowed = self._index_groups(items).get(category, set()) if category in index_categories else None
            for snapshot in items:
                symbol = str(getattr(snapshot, "symbol", "") or "")
                market_raw = str(getattr(snapshot, "market", "") or ""); market = "上市" if market_raw in ("TAIEX", "TSE") else "上櫃" if market_raw in ("TAISDAQ", "OTC") else market_raw
                if allowed is not None and symbol not in allowed: continue
                if category == "上市一般股" and market != "上市": continue
                if category == "上櫃一般股" and market != "上櫃": continue
                last_price = as_float(getattr(snapshot, "last_price", None)); limit_up = as_float(getattr(snapshot, "limitup_price", None)); volume = as_float(getattr(snapshot, "total_volume", None)); target = previous_stock_ticks(limit_up, ticks)
                if not symbol or last_price is None or target is None or volume is None: continue
                if abs(last_price - target) > .0001 or volume <= min_volume: continue
                name = symbol; percent = None; update_time = str(getattr(snapshot, "update_time", "") or "")
                try:
                    quote = unwrap(rest.intraday.quote(symbol=symbol)); actual_trade = as_float(quote.get("closePrice")); total = quote.get("total") if isinstance(quote.get("total"), dict) else {}; actual_volume = as_float(total.get("tradeVolume"))
                    if actual_trade is not None: last_price = actual_trade
                    if actual_volume is not None: volume = actual_volume
                    name = str(quote.get("name") or symbol); percent = as_float(quote.get("changePercent"))
                except Exception:
                    reference = as_float(getattr(snapshot, "reference_price", None)); percent = ((last_price / reference) - 1) * 100 if reference else None
                if abs(last_price - target) > .0001 or volume <= min_volume: continue
                unit = int(as_float(getattr(snapshot, "unit", None)) or 1000)
                matches.append({"name": name, "symbol": symbol, "market": market, "last": last_price, "limit": limit_up, "ticks": ticks, "volume": int(volume), "percent": percent, "time": update_time, "unit": unit})
            matches.sort(key=lambda row: row["volume"], reverse=True); self.events.put(("limit_scan", (matches, len(items), datetime.now().strftime("%H:%M:%S"), ticks, min_volume, category)))
        except Exception as exc:
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
