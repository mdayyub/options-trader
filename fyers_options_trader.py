#!/usr/bin/env python3
"""
Fyers Options Trader - simple Windows GUI for placing multiple NIFTY / SENSEX
option orders through the FYERS API v3.

Features
  * Save App ID + Access Token locally (token encrypted with Windows DPAPI)
  * Contracts (symbols, lot sizes, tick sizes, expiries, strikes) from Fyers' symbol master
  * Build a basket of BUY / SELL option orders (lots x lot size)
  * PLACE ALL NOW, or SCHEDULE the whole basket for a time today
  * Optional auto-slice for quantities above the exchange freeze limit
  * Positions with P&L, Available Funds, Order Book + cancel, one-click Exit All

Requires: Python 3.9+ and `pip install requests`

Access token: generate it yourself (Fyers tokens expire every day), paste it
into the app and click Save. App ID looks like  XC4XXXXM-100  (the app_id of
your Fyers API app, including the -100 suffix).
"""
import base64
import csv
import datetime as dt
import json
import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox as _native_messagebox, ttk

import requests

APP_NAME = "FyersOptionsTrader"
BASE = "https://api-t1.fyers.in/api/v3"
DATA_BASE = "https://api-t1.fyers.in/data"
ORDER_TAG = "optionsapp"
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

# one symbol-master file per exchange segment
UNDERLYINGS = {
    "NIFTY": {"url": "https://public.fyers.in/sym_details/NSE_FO.csv", "prefix": "NSE"},
    "SENSEX": {"url": "https://public.fyers.in/sym_details/BSE_FO.csv", "prefix": "BSE"},
}

ICON_NAME = "fyers.ico"

# Window height ~ FIXED + PER_SCALE x display-scale (measured in the taller dark theme).
# Normal layout, and the compact layout used on short screens (log moved into its own tab).
FIXED_N, SCALED_N = 426, 478.0
FIXED_C, SCALED_C = 206, 447.0


def resource_path(name: str) -> str:
    """Path to a bundled file - works both from source and inside the PyInstaller .exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def enable_dpi_awareness():
    """Tell Windows this app handles high-DPI itself, so text/UI are sharp instead of blurry-scaled.
    Must be called before the Tk window is created."""
    if os.name != "nt":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)   # Windows 8.1+: system DPI aware
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()         # Windows Vista / 7 fallback
    except Exception:
        pass


CONFIG_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), APP_NAME)
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
CACHE_FILE = os.path.join(CONFIG_DIR, "contracts.json")

ORDER_STATUS = {1: "CANCELLED", 2: "TRADED", 3: "UNKNOWN", 4: "TRANSIT", 5: "REJECTED", 6: "PENDING", 7: "EXPIRED"}
ORDER_TYPE = {1: "LIMIT", 2: "MARKET", 3: "SL-M", 4: "SL-L"}


# --------------------------------------------------------------------------
# Token storage (Windows DPAPI, falls back to base64 on other platforms)
# --------------------------------------------------------------------------
def _dpapi(data: bytes, encrypt: bool) -> bytes:
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(data, len(data))
    inb = BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    out = BLOB()
    fn = ctypes.windll.crypt32.CryptProtectData if encrypt else ctypes.windll.crypt32.CryptUnprotectData
    if not fn(ctypes.byref(inb), None, None, None, None, 0, ctypes.byref(out)):
        raise OSError("DPAPI call failed")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out.pbData)


def protect(text: str) -> str:
    raw = text.encode("utf-8")
    if os.name == "nt":
        try:
            return "dpapi:" + base64.b64encode(_dpapi(raw, True)).decode()
        except Exception:
            pass
    return "b64:" + base64.b64encode(raw).decode()


def unprotect(blob: str) -> str:
    try:
        kind, _, payload = blob.partition(":")
        raw = base64.b64decode(payload)
        if kind == "dpapi":
            raw = _dpapi(raw, False)
        return raw.decode("utf-8")
    except Exception:
        return ""


def _read_cfg() -> dict:
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def load_config() -> dict:
    cfg = _read_cfg()
    return {"app_id": cfg.get("app_id", ""), "token": unprotect(cfg.get("token", "")), "dark": cfg.get("dark")}


def save_pref(name: str, value) -> None:
    """Store one setting (e.g. the dark-mode choice) without touching the others."""
    cfg = _read_cfg()
    cfg[name] = value
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f)


def save_config(app_id: str, token: str) -> None:
    save_pref("app_id", app_id)
    save_pref("token", protect(token))


# --------------------------------------------------------------------------
# Symbol master
#   CSV columns (no header row in the live file):
#   0 Fytoken | 1 Symbol Details | 2 Instrument type | 3 Min lot size | 4 Tick size |
#   5 ISIN | 6 Session | 7 Last update | 8 Expiry date | 9 Symbol ticker | 10 Exchange |
#   11 Segment | 12 Scrip code | 13 Underlying symbol | 14 Underlying scrip code |
#   15 Strike price | 16 Option type | 17 Underlying FyToken | 18-20 reserved
# --------------------------------------------------------------------------
def expiry_to_date(v: str) -> str:
    v = v.strip()
    try:
        num = float(v)
        if num > 1e8:  # epoch seconds
            return dt.datetime.fromtimestamp(int(num), IST).date().isoformat()
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d-%b-%Y", "%d %b %Y"):
        try:
            return dt.datetime.strptime(v, fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError("bad expiry " + v)


def load_contracts(log, force=False):
    """Return list of NIFTY / SENSEX option contracts (cached per day)."""
    today = dt.date.today().isoformat()
    if not force and os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if data.get("date") == today:
                return data["contracts"]
        except Exception:
            pass

    contracts = []
    for ul, meta in UNDERLYINGS.items():
        log(f"Downloading {ul} symbol master (once a day, may take a minute)...")
        r = requests.get(meta["url"], stream=True, timeout=120)
        r.raise_for_status()
        lines = (ln.decode("utf-8", "replace") for ln in r.iter_lines() if ln)
        ticker_re = re.compile(rf"^{meta['prefix']}:{ul}\d")
        for row in csv.reader(lines):
            if len(row) < 17 or not row[0].strip().isdigit():  # skip header / junk
                continue
            otype = row[16].strip().upper()
            if otype not in ("CE", "PE"):
                continue
            ticker = row[9].strip()
            if row[13].strip().upper() != ul and not ticker_re.match(ticker.upper()):
                continue
            try:
                expiry = expiry_to_date(row[8])
                if expiry < today:
                    continue
                contracts.append({
                    "ul": ul, "symbol": ticker, "expiry": expiry, "strike": float(row[15]),
                    "type": otype, "lot": int(float(row[3])), "tick": float(row[4]) or 0.05,
                })
            except (ValueError, IndexError):
                continue

    if not contracts:
        raise RuntimeError("No NIFTY/SENSEX option contracts found in the symbol master.")
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump({"date": today, "contracts": contracts}, f)
    log(f"Loaded {len(contracts)} option contracts.")
    return contracts


# --------------------------------------------------------------------------
# API client
# --------------------------------------------------------------------------
class Fyers:
    last_ms = None        # duration of the most recent request in ms (None = failed)
    on_latency = None     # optional callback(ms, path) so the GUI can display it

    def __init__(self, app_id, token):
        self.s = requests.Session()
        self.s.headers.update({
            "Authorization": f"{app_id.strip()}:{token.strip()}",
            "Content-Type": "application/json",
            "version": "3",
        })

    def call(self, method, path, body=None, params=None, data_api=False):
        url = (DATA_BASE if data_api else BASE) + path
        t0 = time.perf_counter()
        try:
            r = self.s.request(method, url, json=body, params=params, timeout=15)
        except Exception:
            self.last_ms = None
            if self.on_latency:
                self.on_latency(None, path)
            raise
        self.last_ms = (time.perf_counter() - t0) * 1000   # request + response time in ms
        if self.on_latency:
            self.on_latency(self.last_ms, path)
        try:
            data = r.json()
        except ValueError:
            data = r.text
        return r.status_code, data, r.headers

    def profile(self):
        return self.call("GET", "/profile")[:2]

    def funds(self):
        return self.call("GET", "/funds")[:2]

    def positions(self):
        return self.call("GET", "/positions")[:2]

    def orders(self):
        return self.call("GET", "/orders")[:2]

    def cancel(self, order_id):
        return self.call("DELETE", "/orders/sync", {"id": order_id})[:2]

    def exit_all(self):
        # closes every open position with counter orders and cancels pending orders
        return self.call("DELETE", "/positions", {"pending_orders_cancel": 1})[:2]

    def ltp(self, symbol):
        code, data, _ = self.call("GET", "/quotes", params={"symbols": symbol}, data_api=True)
        if code == 200 and isinstance(data, dict) and data.get("s") == "ok":
            for d in data.get("d", []):
                v = d.get("v", {})
                if "lp" in v:
                    return v["lp"]
        raise RuntimeError(f"LTP failed ({code}): {data}")

    def place(self, it, slice_orders=False):
        body = {
            "symbol": it["symbol"],
            "qty": int(it["qty"]),
            "type": 1 if it["otype"] == "LIMIT" else 2,
            "side": 1 if it["side"] == "BUY" else -1,
            "productType": it["product"],
            "limitPrice": float(it["price"]) if it["otype"] == "LIMIT" else 0,
            "stopPrice": 0,
            "disclosedQty": 0,
            "validity": it["validity"],
            "offlineOrder": False,
            "orderTag": ORDER_TAG,
            "isSliceOrder": bool(slice_orders),
        }
        code, data, hdr = self.call("POST", "/orders/sync", body)
        if code == 429:  # rate-limited => order was NOT accepted, safe to retry once
            try:
                wait = float(hdr.get("X-Retry-After-Ms") or 0) / 1000 or float(hdr.get("Retry-After") or 1)
            except ValueError:
                wait = 1.0
            time.sleep(min(max(wait, 0.3), 5))
            code, data, hdr = self.call("POST", "/orders/sync", body)
        return code, data


# --------------------------------------------------------------------------
# Theme (light / dark)
# --------------------------------------------------------------------------
LIGHT = {"bg": "#f4f5f7", "fg": "#1f2328", "field": "#ffffff", "btn": "#e6e8ec", "hover": "#d8dce2",
         "border": "#cdd2d9", "sel": "#cfe0fb", "selfg": "#111827", "mute": "#8a8f98", "accent": "#2563eb",
         "accent_hover": "#1d4ed8", "alt": "#f5f7fa", "scroll": "#b4bac4", "buy": "#16a34a", "sell": "#dc2626",
         "good": "#15803d", "info": "#15803d", "warn": "#b45309", "loss": "#b91c1c"}
DARK = {"bg": "#1f2023", "fg": "#e6e6e6", "field": "#2b2d31", "btn": "#35373d", "hover": "#45484f",
        "border": "#44474e", "sel": "#2d5a9e", "selfg": "#ffffff", "mute": "#80838b", "accent": "#3b82f6",
        "accent_hover": "#5b9bff", "alt": "#303238", "scroll": "#5a5d66", "buy": "#16a34a", "sell": "#dc2626",
        "good": "#4ade80", "info": "#4ade80", "warn": "#fbbf24", "loss": "#f87171"}


def system_prefers_dark() -> bool:
    """True if Windows is set to dark app mode (used as the default until you pick a mode yourself)."""
    if os.name != "nt":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
            return winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
    except Exception:
        return False


def set_titlebar_dark(window, dark: bool) -> None:
    """Best effort: make the Windows title bar dark/light to match (Windows 10 1809+ / Windows 11)."""
    if os.name != "nt":
        return
    try:
        import ctypes
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        val = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (new / older Windows 10 builds)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(val), ctypes.sizeof(val)) == 0:
                break
    except Exception:
        pass


class _ThemedMessageBox:
    """Drop-in for tkinter.messagebox whose dialogs follow the app's light / dark theme
    (the native Windows message boxes ignore it)."""
    app = None

    def askyesno(self, title, message, icon="question", default="yes", **kw):
        if self.app is None:
            return _native_messagebox.askyesno(title, message, icon=icon, default=default)
        return self.app._dialog(title, message, icon, ("Yes", "No"), 0 if default == "yes" else 1, 1) == 0

    def _notice(self, kind, native, title, message):
        if self.app is None:
            return native(title, message)
        self.app._dialog(title, message, kind, ("OK",), 0, 0)

    def showinfo(self, title, message, **kw):
        return self._notice("info", _native_messagebox.showinfo, title, message)

    def showwarning(self, title, message, **kw):
        return self._notice("warning", _native_messagebox.showwarning, title, message)

    def showerror(self, title, message, **kw):
        return self._notice("error", _native_messagebox.showerror, title, message)


messagebox = _ThemedMessageBox()


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Fyers Options Trader - NIFTY / SENSEX")
        try:  # window / taskbar icon (silently skipped if the file or platform doesn't support it)
            self.iconbitmap(resource_path(ICON_NAME))
        except Exception:
            pass
        dpi = self.winfo_fpixels("1i")                     # real screen DPI (96 = 100% scaling)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        want = max(1.0, dpi / 96.0)                        # scale the system asks for (1.25, 1.5, 2.0 ...)
        # the window (in the taller dark theme) needs roughly 478 px per scale unit + 426 px of fixed padding;
        # shrink the scale a little
        # if the screen is too short, so every button stays visible (e.g. 1080p laptop at 150%)
        self._compact = (sh - 90) < (FIXED_N + SCALED_N)   # normal layout can't fit even at 100% -> compact
        fixed, per = (FIXED_C, SCALED_C) if self._compact else (FIXED_N, SCALED_N)
        fit = max(1.0, (sh - 90 - fixed) / per)
        self.dpi_scale = min(want, fit)
        self.tk.call("tk", "scaling", dpi / 72.0 * (self.dpi_scale / want))   # fonts follow the same scale
        ttk.Style(self).configure("Treeview", rowheight=self.px(20))
        self._setup_fonts()
        self._screen = (sw, sh)
        self._design = (self.px(1100), self.px(850))       # preferred size; grown if the content needs more
        self.q = queue.Queue()
        self.contracts = []
        self.index = {}  # ul -> expiry -> type -> [contracts sorted by strike]
        self.basket = []
        self.positions_data = []
        self.book_status = {}
        self.sched = None
        self._sched_job = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        pref = load_config().get("dark")
        self.v_dark = tk.BooleanVar(value=system_prefers_dark() if pref is None else bool(pref))
        self.v_confirm = tk.BooleanVar(value=bool(_read_cfg().get("confirm", True)))
        self.v_latency = tk.StringVar(value="-- ms")
        self._lat = []
        self._base_theme = ttk.Style(self).theme_use()
        self._build_fitted(dpi, want)
        if load_config().get("token"):             # saved credentials: start with just the slim status bar
            self.cf.pack_forget()
            self.v_conn.set("Checking connection\u2026")
        self.fit_window()
        self.update_info()
        cfg = load_config()
        self.v_client.set(cfg["app_id"])
        self.v_token.set(cfg["token"])
        self.after(100, self._poll)
        self.refresh_contracts()
        if cfg["token"]:
            self.after(300, lambda: self.check_connection(False))
            self.after(800, lambda: self.on_funds(True))
            self.after(1200, self.on_positions)

    # ---- theme ----------------------------------------------------------
    def _all_widgets(self, w=None):
        for c in (w or self).winfo_children():
            yield c
            yield from self._all_widgets(c)

    def _capture_defaults(self):
        """(kept for compatibility - colours now come from the LIGHT / DARK palettes)"""

    def apply_theme(self):
        dark = bool(self.v_dark.get())
        pal = DARK if dark else LIGHT
        st = ttk.Style(self)
        st.theme_use("clam")        # one flat base theme for both modes; the colours come from the palette
        bg, fg, field, btn = pal["bg"], pal["fg"], pal["field"], pal["btn"]
        hover, border, sel, selfg = pal["hover"], pal["border"], pal["sel"], pal["selfg"]
        mute, accent, acc_h = pal["mute"], pal["accent"], pal["accent_hover"]
        st.configure(".", background=bg, foreground=fg, fieldbackground=field, bordercolor=border,
                     darkcolor=bg, lightcolor=bg, troughcolor=field, focuscolor=bg, insertcolor=fg,
                     selectbackground=sel, selectforeground=selfg)
        st.configure("TButton", background=btn, foreground=fg, bordercolor=border, padding=(10, 2 if self._compact else 4))
        st.map("TButton", background=[("active", hover), ("pressed", hover), ("disabled", bg)],
               foreground=[("disabled", mute)])
        for name in ("TEntry", "TSpinbox", "TCombobox"):
            st.configure(name, fieldbackground=field, foreground=fg, background=btn, bordercolor=border,
                         arrowcolor=fg, insertcolor=fg, selectbackground=sel, selectforeground=selfg)
        st.map("TCombobox", fieldbackground=[("readonly", field)], foreground=[("readonly", fg)],
               selectbackground=[("readonly", field)], selectforeground=[("readonly", fg)],
               background=[("active", hover)])
        st.configure("TNotebook", background=bg, bordercolor=border)
        st.configure("TNotebook.Tab", background=btn, foreground=fg, padding=(16, 3 if self._compact else 6), bordercolor=border)
        st.map("TNotebook.Tab", background=[("selected", bg), ("active", hover)], foreground=[("selected", fg)])
        st.configure("Treeview", background=field, fieldbackground=field, foreground=fg, bordercolor=border)
        st.map("Treeview", background=[("selected", sel)], foreground=[("selected", selfg)])
        st.configure("Treeview.Heading", background=btn, foreground=fg, relief="flat", bordercolor=border,
                     font=self._bold_font, padding=(4, 2 if self._compact else 4))
        st.map("Treeview.Heading", background=[("active", hover)])
        st.configure("TLabelframe", background=bg, bordercolor=border)
        st.configure("TLabelframe.Label", background=bg, foreground=fg)
        for name in ("TCheckbutton", "TRadiobutton"):
            st.configure(name, background=bg, foreground=fg, indicatorbackground=field,
                         indicatorforeground="#ffffff", indicatorcolor=field)
            st.map(name, background=[("active", bg)],
                   indicatorcolor=[("selected", accent), ("pressed", accent)],
                   indicatorbackground=[("selected", accent)])
        st.configure("TScrollbar", background=pal["scroll"], troughcolor=field, bordercolor=border, arrowcolor=fg,
                     arrowsize=self.px(15))
        st.map("TScrollbar", background=[("active", mute), ("pressed", mute)])

        # Tk 8.6.9 ignores Treeview row tag colours unless the default style map entries are removed
        def fixed_map(option):
            return [e for e in st.map("Treeview", query_opt=option) if e[:2] != ("!disabled", "!selected")]
        st.map("Treeview", foreground=fixed_map("foreground"), background=fixed_map("background"))
        st.configure("Treeview", rowheight=self.px(20 if self._compact else 26))
        for name, key in (("Good.TLabel", "good"), ("Info.TLabel", "info"), ("Warn.TLabel", "warn"),
                          ("Bad.TLabel", "loss")):
            st.configure(name, foreground=pal[key])
        st.configure("Ping.TButton", padding=(6, 0))
        st.configure("Default.TButton", font=self._bold_font, background=accent, foreground="#ffffff",
                     bordercolor=accent)                       # the dialog's default button
        st.map("Default.TButton", background=[("active", acc_h), ("pressed", acc_h)])

        self.configure(bg=bg)
        self.txt.configure(background=field, foreground=fg, insertbackground=fg, selectbackground=sel,
                           selectforeground=selfg, highlightbackground=border, highlightcolor=border)
        for tree in (self.tv, self.pos, self.bk):
            tree.tag_configure("odd", background=pal["alt"])
            tree.tag_configure("even", background=field)
            tree.tag_configure("buy", foreground=pal["good"])
            tree.tag_configure("sell", foreground=pal["loss"])
        self.pos.tag_configure("profit", foreground=pal["good"])
        self.pos.tag_configure("loss", foreground=pal["loss"])
        self.conn_dot.configure(bg=bg)
        self.btn_place.configure(bg=accent, activebackground=acc_h, fg="#ffffff", activeforeground="#ffffff")
        self._refresh_side()
        for w in self._all_widgets():          # drop-down lists of the combo boxes are plain Tk listboxes
            if w.winfo_class() == "TCombobox":
                try:
                    pd = self.tk.call("ttk::combobox::PopdownWindow", str(w))
                    self.tk.call(f"{pd}.f.l", "configure", "-background", field, "-foreground", fg,
                                 "-selectbackground", sel, "-selectforeground", selfg)
                except tk.TclError:
                    pass
        set_titlebar_dark(self, dark)

    def on_theme(self):
        self.apply_theme()
        try:
            save_pref("dark", bool(self.v_dark.get()))
        except OSError:
            pass
        self.update_idletasks()               # the other theme may need slightly more room
        sw, sh = self._screen
        rw, rh = self.winfo_reqwidth(), self.winfo_reqheight()
        w = min(max(self.winfo_width(), rw), sw - 20)
        h = min(max(self.winfo_height(), rh), sh - 90)
        if (w, h) != (self.winfo_width(), self.winfo_height()):
            self.geometry(f"{w}x{h}")
        self.minsize(min(rw, w), min(rh, h))

    def _dialog(self, title, message, kind, buttons, default, cancel):
        """Themed modal dialog. Returns the index of the button pressed (Esc / closing = `cancel`)."""
        dark = bool(self.v_dark.get())
        bg = (DARK if dark else LIGHT)["bg"]
        dlg = tk.Toplevel(self)
        dlg.withdraw()
        dlg.title(title)
        dlg.configure(bg=bg)
        dlg.resizable(False, False)
        dlg.transient(self)
        try:
            dlg.iconbitmap(resource_path(ICON_NAME))
        except Exception:
            pass
        result = {"v": cancel}
        glyph, colour = {"question": ("?", "#3b82f6"), "info": ("i", "#3b82f6"),
                         "warning": ("!", "#f59e0b"), "error": ("\u2715", "#ef4444")}.get(kind, ("i", "#3b82f6"))
        body = ttk.Frame(dlg, padding=(self.px(18), self.px(16), self.px(18), self.px(8)))
        body.pack(fill="both", expand=True)
        size = self.px(36)
        cv = tk.Canvas(body, width=size, height=size, bg=bg, highlightthickness=0)
        cv.create_oval(1, 1, size - 1, size - 1, fill=colour, outline=colour)
        cv.create_text(size // 2, size // 2, text=glyph, fill="#ffffff", font=("Segoe UI", 14, "bold"))
        cv.pack(side="left", anchor="n", padx=(0, self.px(14)))
        ttk.Label(body, text=message, justify="left", wraplength=self.px(460)).pack(side="left", fill="both", expand=True)
        bar = ttk.Frame(dlg, padding=(self.px(18), self.px(6), self.px(18), self.px(14)))
        bar.pack(fill="x")

        def finish(v):
            result["v"] = v
            dlg.destroy()
        default_btn = None
        for i in reversed(range(len(buttons))):
            b = ttk.Button(bar, text=buttons[i], width=9, command=lambda i=i: finish(i),
                           style="Default.TButton" if i == default else "TButton")
            b.pack(side="right", padx=(self.px(8), 0))
            if i == default:
                default_btn = b
        dlg.bind("<Return>", lambda e: finish(default))
        dlg.bind("<Escape>", lambda e: finish(cancel))
        dlg.protocol("WM_DELETE_WINDOW", lambda: finish(cancel))
        set_titlebar_dark(dlg, dark)
        dlg.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_reqwidth()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_reqheight()) // 3
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")
        dlg.deiconify()
        try:
            dlg.wait_visibility()
            dlg.grab_set()
        except tk.TclError:
            pass
        dlg.focus_force()
        if default_btn is not None:
            default_btn.focus_set()
        self.wait_window(dlg)
        return result["v"]

    # ---- latency --------------------------------------------------------
    def _latency_cb(self, ms, what):
        """Called from worker threads after every API request (ms is None if there was no response)."""
        self.ui(self._show_latency, ms)

    def _show_latency(self, ms):
        if not hasattr(self, "lat_lbl"):
            return
        if ms is None:
            self.v_latency.set("no response")
            self.lat_lbl.configure(style="Bad.TLabel")
            return
        self._lat = (self._lat + [ms])[-10:]
        avg = sum(self._lat) / len(self._lat)
        self.v_latency.set(f"{ms:.0f} ms   (avg {avg:.0f} ms of last {len(self._lat)})")
        self.lat_lbl.configure(style="Good.TLabel" if ms < 300 else "Warn.TLabel" if ms < 800 else "Bad.TLabel")

    def on_ping(self):
        api = self.get_api()
        if not api:
            return

        def work():
            code, _ = api.profile()          # light authenticated request
            self.log_t(f"Ping (profile request): HTTP {code} in {api.last_ms:.0f} ms")
        self.run_bg(work)

    # ---- tables: stripes, BUY/SELL colours, click-to-sort ---------------
    def _tree_with_scrollbar(self, parent, padx, pady, **kw):
        """A Treeview with a vertical scrollbar (always visible) inside its own frame."""
        box = ttk.Frame(parent)
        box.pack(fill="both", expand=True, padx=padx, pady=pady)
        tree = ttk.Treeview(box, **kw)
        bar = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        return tree

    def _restripe(self, tree):
        for n, iid in enumerate(tree.get_children()):
            tags = [t for t in tree.item(iid, "tags") if t not in ("odd", "even")]
            tags.append("odd" if n % 2 else "even")
            tree.item(iid, tags=tags)

    def _tag_sides(self, tree):
        """Colour BUY rows green and SELL rows red (reads the 'side' column)."""
        for iid in tree.get_children():
            side = str(tree.set(iid, "side")).upper()
            tags = [t for t in tree.item(iid, "tags") if t not in ("buy", "sell")]
            if side in ("BUY", "SELL"):
                tags.append(side.lower())
            tree.item(iid, tags=tags)

    def _make_sortable(self, tree):
        """Click a column heading to sort (click again to reverse). Only for read-only tables."""
        self._sort_state = getattr(self, "_sort_state", {})
        self._base_head = getattr(self, "_base_head", {})
        self._sort_state.pop(str(tree), None)
        for c in tree["columns"]:
            self._base_head[(str(tree), c)] = tree.heading(c, "text")
            tree.heading(c, command=lambda c=c, t=tree: self._sort_by(t, c))

    def _sort_by(self, tree, col):
        prev, rev = self._sort_state.get(str(tree), (None, False))
        self._sort_state[str(tree)] = (col, (not rev) if prev == col else False)
        self._apply_sort(tree)

    def _apply_sort(self, tree):
        col, rev = self._sort_state.get(str(tree), (None, False))
        if col is not None:
            def key(iid):
                v = str(tree.set(iid, col)).replace(",", "").replace("\u20b9", "").strip()
                try:
                    return (0, float(v), "")
                except ValueError:
                    return (1, 0.0, v.lower())
            for i, iid in enumerate(sorted(tree.get_children(), key=key, reverse=rev)):
                tree.move(iid, "", i)
            for c in tree["columns"]:
                base = self._base_head.get((str(tree), c), c.upper())
                tree.heading(c, text=base + ((" \u25BC" if rev else " \u25B2") if c == col else ""))
        self._restripe(tree)

    def _refresh_side(self, *a):
        """BUY = green, SELL = red when selected; muted when not."""
        pal = DARK if self.v_dark.get() else LIGHT
        for side, b in getattr(self, "_side_btns", {}).items():
            on = self.v_side.get() == side
            col = pal["buy"] if side == "BUY" else pal["sell"]
            try:
                b.configure(bg=col if on else pal["btn"], fg="#ffffff" if on else pal["mute"],
                            activebackground=col if on else pal["hover"], activeforeground="#ffffff")
            except tk.TclError:
                pass

    # ---- connection status chip -----------------------------------------
    def _grow_window(self):
        self.update_idletasks()
        sw, sh = self._screen
        rw, rh = self.winfo_reqwidth(), self.winfo_reqheight()
        w = min(max(self.winfo_width(), rw), sw - 20)
        h = min(max(self.winfo_height(), rh), sh - 90)
        if (w, h) != (self.winfo_width(), self.winfo_height()):
            self.geometry(f"{w}x{h}")
        self.minsize(min(rw, w), min(rh, h))

    def toggle_creds(self):
        if self.cf.winfo_manager() == "pack":
            self.cf.pack_forget()
        else:
            self.cf.pack(fill="x", padx=8, pady=(0, 4), after=self.conn_bar)
        self.after(30, self._grow_window)

    def set_conn_state(self, state, text=""):
        self.conn_dot.itemconfigure(self._dot, fill={"ok": "#22c55e", "bad": "#ef4444"}.get(state, "#9ca3af"))
        if state == "ok":
            self.v_conn.set("\u2713  Connected  \u00b7  " + text)
            if self.cf.winfo_manager() == "pack":          # connected: tuck the credentials away
                self.cf.pack_forget()
                self.after(30, self._grow_window)
        elif state == "bad":
            self.v_conn.set("Not connected  \u00b7  " + text)
            if self.cf.winfo_manager() != "pack":          # problem: show the credentials again
                self.cf.pack(fill="x", padx=8, pady=(0, 4), after=self.conn_bar)
                self.after(30, self._grow_window)
        else:
            self.v_conn.set(text or "Not connected")

    def _profile_summary(self, code, data):
        if code == 200 and isinstance(data, dict) and data.get("s") == "ok":
            d = data.get("data") or {}
            who = d.get("name") or d.get("fy_id") or self.v_client.get().strip()
            return True, f"{who}  \u00b7  token valid for today"
        return False, ""

    def check_connection(self, manual=False):
        api = self.get_api(silent=not manual)
        if not api:
            return

        def work():
            try:
                code, data = api.profile()
            except Exception as e:
                self.log_t(f"Connection check: no response ({e})")
                self.ui(self.set_conn_state, "bad", "no response - check your internet connection")
                return
            ok, text = self._profile_summary(code, data)
            if ok:
                self.log_t(f"Connection OK - {text}")
                self.ui(self.set_conn_state, "ok", text)
                if manual:
                    self.ui(self.on_funds)
                    self.ui(self.on_positions)
            else:
                self.log_t(f"Connection check failed -> HTTP {code}: {str(data)[:200]}")
                self.ui(self.set_conn_state, "bad", f"HTTP {code} - check your credentials / token")
        self.run_bg(work)

    def on_confirm_toggle(self):
        on = bool(self.v_confirm.get())
        try:
            save_pref("confirm", on)
        except OSError:
            pass
        self.log("Order confirmation ON - you will be asked before orders are placed." if on else
                 "Order confirmation OFF - 'Place this order now' / 'PLACE ALL NOW' send orders immediately.")

    def _setup_fonts(self):
        """Segoe UI for text and Consolas for the log (when installed); a size smaller in the compact layout.
        Also re-applies the point size: Tk only recomputes an in-use font's pixel size when it is reconfigured,
        so this must run again after every change of the display scale."""
        size = 9 if self._compact else 10
        fams = set(tkfont.families(self))
        for fname in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont",
                      "TkTooltipFont"):
            f = tkfont.nametofont(fname)
            if "Segoe UI" in fams:
                f.configure(family="Segoe UI")
            f.configure(size=size)
        fixed = tkfont.nametofont("TkFixedFont")
        if "Consolas" in fams:
            fixed.configure(family="Consolas")
        fixed.configure(size=size)
        fam = tkfont.nametofont("TkDefaultFont").actual("family")
        if not hasattr(self, "_bold_font"):
            self._bold_font = tkfont.Font(self, family=fam, size=size, weight="bold")
        else:
            self._bold_font.configure(family=fam, size=size, weight="bold")

    def _build_fitted(self, dpi, want):
        """Build the UI; if it is larger than the screen, shrink the scale (and, at 100%, switch to the compact
        layout) and rebuild. Measured in BOTH themes, so switching dark / light later never needs more room."""
        avail = self._screen[1] - 90
        availw = self._screen[0] - 20
        start_dark = bool(self.v_dark.get())
        last = None                                   # (scale, height) of the previous attempt
        for attempt in range(8):
            self._build()
            self._capture_defaults()
            messagebox.app = self
            self.v_info.set("BSE:SENSEX26NOV82000PE\nSecurity ID 1234567 | Lot size 75\n"
                            "Quantity = 100 x 75 = 7500 per leg\nWill add 50 leg(s) = 375000 qty in total")
            self.geometry("1x1")                      # measure in a tiny window: nothing can stretch
            req = reqw = 0
            trees = (self.tv, self.pos, self.bk)      # stretchy columns would inflate the measured width
            saved = {(str(t), c): t.column(c, "stretch") for t in trees for c in t["columns"]}
            for t in trees:
                for c in t["columns"]:
                    t.column(c, stretch=False)
            for flag in (True, False):
                self.v_dark.set(flag)
                self.apply_theme()
                self.update_idletasks()
                req = max(req, self.winfo_reqheight())
                reqw = max(reqw, self.winfo_reqwidth())
            for t in trees:
                for c in t["columns"]:
                    t.column(c, stretch=bool(int(saved[(str(t), c)])))
            self.v_dark.set(start_dark)
            self.apply_theme()
            if (req <= avail and reqw <= availw) or attempt == 7:
                break
            old = self.dpi_scale
            if old <= 1.0:
                if self._compact or req <= avail:     # nothing left to shrink (or only the width is a problem)
                    break
                self._compact = True                  # still too tall at 100%: use the compact layout
                self._setup_fonts()
                new, last = 1.0, None
            else:
                new = old
                if req > avail:
                    if last and abs(old - last[0]) > 1e-3 and req != last[1]:
                        slope = (req - last[1]) / (old - last[0])
                    else:
                        slope = req * 0.7 / old
                    new = min(new, max(1.0, old - (req - avail) / max(slope, 50.0) * 1.03))
                if reqw > availw:
                    new = min(new, max(1.0, old - (reqw - availw) / max(reqw * 0.9 / old, 50.0) * 1.03))
                last = (old, req)
            self.dpi_scale = new
            self.tk.call("tk", "scaling", dpi / 72.0 * (self.dpi_scale / want))
            self._setup_fonts()                       # fonts must be re-applied to follow the new scale
            ratio = new / old
            self._design = (int(self._design[0] * ratio), int(self._design[1] * ratio))
            for w in list(self.winfo_children()):
                w.destroy()
        self.geometry("")                             # back to "size to content"

    def fit_window(self):
        """Size the window from what its content really needs (so nothing is clipped at any DPI),
        centred, and never larger than the screen."""
        self.update_idletasks()
        sw, sh = self._screen
        rw, rh = self.winfo_reqwidth(), self.winfo_reqheight()
        w = min(max(self._design[0], rw), sw - 20)
        h = min(max(self._design[1], rh), sh - 90)
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 2 - 20)}")
        self.minsize(min(rw, w), min(rh, h))

    def px(self, n):
        """Scale a pixel size designed for 96 DPI to the current display DPI."""
        return int(round(n * self.dpi_scale))

    # ---- thread helpers -------------------------------------------------
    def ui(self, fn, *a):
        self.q.put((fn, a))

    def _poll(self):
        try:
            while True:
                fn, a = self.q.get_nowait()
                fn(*a)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def run_bg(self, fn):
        def wrap():
            try:
                fn()
            except Exception as e:  # noqa
                self.ui(self.log, f"ERROR: {e}")
        threading.Thread(target=wrap, daemon=True).start()

    def log(self, msg):
        ts = dt.datetime.now().strftime("%H:%M:%S")
        self.txt.configure(state="normal")
        self.txt.insert("end", f"[{ts}] {msg}\n")
        if hasattr(self, "v_status"):
            self.v_status.set(f"[{ts}] {msg}"[:200])
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def log_t(self, msg):
        self.ui(self.log, msg)

    # ---- layout ---------------------------------------------------------
    def _build(self):
        pad = {"padx": 6, "pady": 4}

        self.conn_bar = ttk.Frame(self)           # slim status bar: dot + status text + settings button
        self.conn_bar.pack(fill="x", padx=8, pady=(8, 2))
        dot = self.px(12)
        self.conn_dot = tk.Canvas(self.conn_bar, width=dot, height=dot, highlightthickness=0)
        self._dot = self.conn_dot.create_oval(1, 1, dot - 1, dot - 1, fill="#9ca3af", outline="")
        self.conn_dot.pack(side="left", padx=(6, 8))
        self.v_conn = tk.StringVar(value="Not connected  \u00b7  enter your credentials")
        ttk.Label(self.conn_bar, textvariable=self.v_conn).pack(side="left")
        ttk.Button(self.conn_bar, text="\u2699  Credentials", style="Ping.TButton",
                   command=self.toggle_creds).pack(side="right")
        cf = ttk.LabelFrame(self, text="Fyers credentials")
        self.cf = cf
        cf.pack(fill="x", padx=8, pady=(0, 4))
        self.v_client = tk.StringVar()
        self.v_token = tk.StringVar()
        ttk.Label(cf, text="App ID").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(cf, textvariable=self.v_client, width=18).grid(row=0, column=1, **pad)
        ttk.Label(cf, text="Access Token").grid(row=0, column=2, sticky="e", **pad)
        self.e_token = ttk.Entry(cf, textvariable=self.v_token, width=50, show="*")
        self.e_token.grid(row=0, column=3, **pad)
        self.v_show = tk.BooleanVar()
        ttk.Checkbutton(cf, text="Show", variable=self.v_show,
                        command=lambda: self.e_token.configure(show="" if self.v_show.get() else "*")
                        ).grid(row=0, column=4, **pad)
        ttk.Button(cf, text="Save", command=self.on_save).grid(row=0, column=5, **pad)
        ttk.Button(cf, text="Test connection", command=self.on_test).grid(row=0, column=6, **pad)
        ttk.Button(cf, text="Reload contracts", command=lambda: self.refresh_contracts(True)
                   ).grid(row=0, column=7, **pad)

        ff = ttk.Frame(self)
        ff.pack(fill="x", padx=8, pady=(0, 2))
        self.v_avail = tk.StringVar(value="\u20b9 --")
        self.v_used = tk.StringVar(value="--")
        ttk.Label(ff, text="Available Funds:").pack(side="left", padx=(6, 2))
        ttk.Label(ff, textvariable=self.v_avail, font=("Segoe UI", 13, "bold"),
                  style="Good.TLabel").pack(side="left", padx=(0, 14))
        ttk.Label(ff, text="Utilized:").pack(side="left", padx=(0, 2))
        ttk.Label(ff, textvariable=self.v_used).pack(side="left", padx=(0, 10))
        ttk.Button(ff, text="\u21BB  Refresh funds", command=self.on_funds).pack(side="left", padx=4)
        tk.Button(ff, text="EXIT ALL POSITIONS", bg="#c0392b", fg="white", activebackground="#962d22",
                  activeforeground="white", font=("Segoe UI", 10, "bold"), relief="flat",
                  padx=12, pady=3, command=self.on_exit_all).pack(side="right", padx=6)
        ttk.Checkbutton(ff, text="Dark mode", variable=self.v_dark, command=self.on_theme
                        ).pack(side="right", padx=10)
        ttk.Checkbutton(ff, text="Confirm before placing", variable=self.v_confirm,
                        command=self.on_confirm_toggle).pack(side="right", padx=10)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=4)
        t1, t3, t2 = ttk.Frame(nb), ttk.Frame(nb), ttk.Frame(nb)
        nb.add(t1, text="  \u25B6  Place Orders  ")
        nb.add(t3, text="  \u25A6  Positions  ")
        nb.add(t2, text="  \u2630  Order Book  ")
        self._build_trade(t1)
        self._build_positions(t3)
        self._build_book(t2)

        self.v_status = tk.StringVar(value="")
        if self._compact:       # short screen: full log in its own tab, latest message on a one-line status bar
            lf = ttk.Frame(nb)
            nb.add(lf, text="  \u2261  Log  ")
            self.txt = tk.Text(lf, height=8, state="disabled", wrap="word")
            self.txt.pack(fill="both", expand=True, padx=4, pady=4)
            sb = ttk.Frame(self)
            sb.pack(fill="x", padx=10, pady=(0, 4))
            ttk.Button(sb, text="Ping", style="Ping.TButton", command=self.on_ping).pack(side="right")
            self.lat_lbl = ttk.Label(sb, textvariable=self.v_latency)
            self.lat_lbl.pack(side="right", padx=(4, 8))
            ttk.Label(sb, text="Latency:").pack(side="right")
            ttk.Label(sb, textvariable=self.v_status, anchor="w").pack(side="left", fill="x", expand=True)
        else:
            hdr = ttk.Frame(self)
            ttk.Label(hdr, text="Log").pack(side="left")
            ttk.Label(hdr, text="     Latency:").pack(side="left")
            self.lat_lbl = ttk.Label(hdr, textvariable=self.v_latency)
            self.lat_lbl.pack(side="left", padx=(4, 10))
            ttk.Button(hdr, text="Ping", style="Ping.TButton", command=self.on_ping).pack(side="left")
            lf = ttk.LabelFrame(self, labelwidget=hdr)
            lf.pack(fill="x", padx=8, pady=(4, 8))
            self.txt = tk.Text(lf, height=8, state="disabled", wrap="word")
            self.txt.pack(fill="x", padx=4, pady=4)

    def _build_trade(self, parent):
        left = ttk.LabelFrame(parent, text="Order builder")
        left.pack(side="left", fill="y", padx=6, pady=6)
        # tighter row spacing on short screens (e.g. 1366x768) so the whole builder stays visible
        pad = {"padx": 6, "pady": 0 if self._compact else 3}

        self.v_ul = tk.StringVar(value="NIFTY")
        self.v_exp = tk.StringVar()
        self.v_type = tk.StringVar(value="CE")
        self.v_strike = tk.StringVar()
        self.v_side = tk.StringVar(value="BUY")
        self.v_lots = tk.StringVar(value="1")
        self.v_otype = tk.StringVar(value="MARKET")
        self.v_price = tk.StringVar(value="0")
        self.v_prod = tk.StringVar(value="INTRADAY")
        self.v_valid = tk.StringVar(value="DAY")
        self.v_slice = tk.BooleanVar(value=False)
        self.v_info = tk.StringVar(value="-\n\n\n")
        self.v_legs = tk.StringVar(value="1")
        self.v_legs.trace_add("write", lambda *a: self.update_info())

        def row(r, label, widget):
            ttk.Label(left, text=label).grid(row=r, column=0, sticky="e", **pad)
            widget.grid(row=r, column=1, sticky="w", **pad)

        cb = ttk.Combobox(left, textvariable=self.v_ul, values=list(UNDERLYINGS), state="readonly", width=18)
        cb.bind("<<ComboboxSelected>>", lambda e: self.on_ul())
        row(0, "Underlying", cb)
        self.cb_exp = ttk.Combobox(left, textvariable=self.v_exp, state="readonly", width=18)
        self.cb_exp.bind("<<ComboboxSelected>>", lambda e: self.on_exp())
        row(1, "Expiry", self.cb_exp)

        fr = ttk.Frame(left)
        for t in ("CE", "PE"):
            ttk.Radiobutton(fr, text=t, value=t, variable=self.v_type, command=self.on_exp).pack(side="left", padx=4)
        row(2, "Option type", fr)

        self.cb_strike = ttk.Combobox(left, textvariable=self.v_strike, width=18)
        self.cb_strike.bind("<<ComboboxSelected>>", lambda e: self.update_info())
        self.cb_strike.bind("<FocusOut>", lambda e: self.update_info())
        row(3, "Strike", self.cb_strike)

        fr = ttk.Frame(left)
        self._side_btns = {}
        for t in ("BUY", "SELL"):               # green / red toggle buttons
            b = tk.Button(fr, text=t, width=8, relief="flat", bd=0, highlightthickness=0, cursor="hand2",
                          font=self._bold_font, pady=3, command=lambda t=t: self.v_side.set(t))
            b.pack(side="left", padx=(0, 6))
            self._side_btns[t] = b
        self.v_side.trace_add("write", self._refresh_side)
        row(4, "Action", fr)

        sp = ttk.Spinbox(left, from_=1, to=10000, textvariable=self.v_lots, width=10, command=self.update_info)
        sp.bind("<KeyRelease>", lambda e: self.update_info())
        row(5, "Lots", sp)

        row(6, "Order type", ttk.Combobox(left, textvariable=self.v_otype, values=["MARKET", "LIMIT"],
                                          state="readonly", width=18))
        fr = ttk.Frame(left)
        ttk.Entry(fr, textvariable=self.v_price, width=10).pack(side="left")
        ttk.Button(fr, text="Fetch LTP", command=self.on_ltp).pack(side="left", padx=6)
        row(7, "Limit price", fr)
        row(8, "Product", ttk.Combobox(left, textvariable=self.v_prod, values=["INTRADAY", "MARGIN"],
                                       state="readonly", width=18))
        row(9, "Validity", ttk.Combobox(left, textvariable=self.v_valid, values=["DAY", "IOC"],
                                        state="readonly", width=18))
        ttk.Checkbutton(left, text="Auto-slice above freeze limit", variable=self.v_slice
                        ).grid(row=10, column=0, columnspan=2, sticky="w", **pad)
        ttk.Label(left, textvariable=self.v_info, style="Info.TLabel", justify="left", wraplength=self.px(250)
                  ).grid(row=11, column=0, columnspan=2, sticky="w", **pad)
        fr = ttk.Frame(left)
        ttk.Label(fr, text="No. of legs").pack(side="left")
        ttk.Spinbox(fr, from_=1, to=50, textvariable=self.v_legs, width=4).pack(side="left", padx=4)
        ttk.Label(fr, text="(same strike, added N times)").pack(side="left", padx=4)
        fr.grid(row=12, column=0, columnspan=2, sticky="w", **pad)
        ttk.Button(left, text="Add to basket  \u2192", command=self.on_add
                   ).grid(row=13, column=0, columnspan=2, sticky="we", **pad)
        ttk.Button(left, text="Place this order now", command=self.on_place_single
                   ).grid(row=14, column=0, columnspan=2, sticky="we", **pad)

        right = ttk.LabelFrame(parent, text="Basket")
        right.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        cols = ("no", "side", "symbol", "lots", "qty", "otype", "price", "product")
        self.tv = self._tree_with_scrollbar(right, 4, 4, columns=cols, show="headings", selectmode="extended", height=8)
        widths = {"no": 40, "side": 62, "symbol": 230, "lots": 58, "qty": 72, "otype": 92, "price": 78, "product": 104}
        for c in cols:
            anc = {"symbol": "w", "lots": "e", "qty": "e", "price": "e"}.get(c, "center")
            self.tv.heading(c, text="#" if c == "no" else c.upper(), anchor=anc)
            self.tv.column(c, width=self.px(widths[c]), anchor=anc, stretch=(c == "symbol"))
        self.v_basket_total = tk.StringVar(value="Basket empty")
        ttk.Label(right, textvariable=self.v_basket_total, font=("Segoe UI", 10, "bold")
                  ).pack(fill="x", padx=8, pady=(0, 2))
        bf = ttk.Frame(right)
        bf.pack(fill="x", padx=4, pady=4)
        ttk.Button(bf, text="Remove selected", command=self.on_remove).pack(side="left", padx=4)
        ttk.Button(bf, text="Clear", command=self.on_clear).pack(side="left", padx=4)
        self.btn_place = tk.Button(bf, text="\u25B6  PLACE ALL NOW", font=self._bold_font, relief="flat", bd=0,
                                   fg="#ffffff", bg="#2563eb", activeforeground="#ffffff", cursor="hand2",
                                   padx=18, pady=6, command=self.on_place_all)    # the one accent button
        self.btn_place.pack(side="right", padx=4)

        sf = ttk.LabelFrame(right, text="Schedule 'Place all orders' (today, PC clock)")
        sf.pack(fill="x", padx=4, pady=(2, 6))
        self.v_sched_time = tk.StringVar(value="09:15:00")
        self.v_sched_status = tk.StringVar(value="No schedule set.")
        ttk.Label(sf, text="Time (HH:MM:SS)").pack(side="left", padx=(8, 4), pady=6)
        ttk.Entry(sf, textvariable=self.v_sched_time, width=10).pack(side="left", padx=4)
        ttk.Button(sf, text="Schedule", command=self.on_schedule).pack(side="left", padx=4)
        ttk.Button(sf, text="Cancel schedule", command=self.on_sched_cancel).pack(side="left", padx=4)
        ttk.Label(sf, textvariable=self.v_sched_status, style="Warn.TLabel").pack(side="left", padx=10)

    def _build_positions(self, parent):
        bf = ttk.Frame(parent)
        bf.pack(fill="x", padx=6, pady=6)
        ttk.Button(bf, text="\u21BB  Refresh", command=self.on_positions).pack(side="left", padx=4)
        self.v_open_only = tk.BooleanVar(value=True)
        ttk.Checkbutton(bf, text="Open positions only", variable=self.v_open_only,
                        command=self.render_positions).pack(side="left", padx=8)
        self.v_auto = tk.BooleanVar(value=False)
        ttk.Checkbutton(bf, text="Auto-refresh (5s)", variable=self.v_auto,
                        command=self._auto_tick).pack(side="left", padx=8)
        self.v_pnl = tk.StringVar(value="")
        ttk.Label(bf, textvariable=self.v_pnl, font=("Segoe UI", 10, "bold")).pack(side="right", padx=8)
        cols = ("symbol", "product", "type", "net", "buyavg", "sellavg", "ltp", "realized", "unrealized", "total")
        self.pos = self._tree_with_scrollbar(parent, 6, 6, columns=cols, show="headings", height=8)
        widths = (230, 80, 70, 60, 80, 80, 80, 90, 90, 90)
        for c, w in zip(cols, widths):
            self.pos.heading(c, text=c.upper())
            self.pos.column(c, width=self.px(w), anchor="w" if c == "symbol" else "e")
        self.pos.tag_configure("profit", foreground="#0a7a2f")
        self.pos.tag_configure("loss", foreground="#c0392b")
        self._make_sortable(self.pos)

    def _build_book(self, parent):
        bf = ttk.Frame(parent)
        bf.pack(fill="x", padx=6, pady=6)
        ttk.Button(bf, text="\u21BB  Refresh", command=self.on_book).pack(side="left", padx=4)
        ttk.Button(bf, text="Cancel selected (pending)", command=self.on_cancel).pack(side="left", padx=4)
        cols = ("id", "time", "symbol", "side", "qty", "price", "type", "status", "message")
        self.bk = self._tree_with_scrollbar(parent, 6, 6, columns=cols, show="headings", selectmode="extended", height=8)
        for c, w in zip(cols, (130, 130, 220, 60, 70, 80, 80, 100, 240)):
            self.bk.heading(c, text=c.upper())
            self.bk.column(c, width=self.px(w), anchor="e" if c in ("qty", "price") else "w")
        self._make_sortable(self.bk)

    # ---- credentials ----------------------------------------------------
    def get_api(self, silent=False):
        aid, tok = self.v_client.get().strip(), self.v_token.get().strip()
        if not aid or not tok:
            if not silent:
                messagebox.showwarning("Credentials", "Enter App ID and Access Token first.")
            return None
        api = Fyers(aid, tok)
        api.on_latency = self._latency_cb
        return api

    def on_save(self):
        try:
            save_config(self.v_client.get().strip(), self.v_token.get().strip())
            self.log("Credentials saved to " + CONFIG_FILE)
            self.check_connection(manual=True)
            messagebox.showinfo("Saved", "App ID and token saved on this PC.\n\n"
                                         "Reminder: Fyers tokens expire daily - paste a fresh one each day.")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))

    def on_test(self):
        self.check_connection(manual=True)

    # ---- funds / positions / exit all -----------------------------------
    @staticmethod
    def money(x):
        try:
            return f"\u20b9 {float(x):,.2f}"
        except (TypeError, ValueError):
            return "\u20b9 --"

    def on_funds(self, silent=False):
        api = self.get_api(silent)
        if not api:
            return

        def work():
            code, data = api.funds()
            if code != 200 or not isinstance(data, dict) or data.get("s") != "ok":
                self.log_t(f"Funds -> HTTP {code}: {data}")
                return
            avail = used = None
            for r in data.get("fund_limit", []):
                t = str(r.get("title", "")).lower()
                if "available" in t and avail is None:
                    avail = r
                elif ("utilized" in t or "utilised" in t) and used is None:
                    used = r

            def fmt(r):
                if not r:
                    return "\u20b9 --"
                s = self.money(r.get("equityAmount"))
                com = r.get("commodityAmount") or 0
                return s + (f"  (+ {self.money(com)} commodity)" if com else "")
            self.ui(self.v_avail.set, fmt(avail))
            self.ui(self.v_used.set, fmt(used))
        self.run_bg(work)

    def on_positions(self):
        api = self.get_api(True)
        if not api:
            return

        def work():
            code, data = api.positions()
            if code != 200 or not isinstance(data, dict) or data.get("s") != "ok":
                self.log_t(f"Positions -> HTTP {code}: {data}")
                return
            self.ui(self.set_positions, data.get("netPositions", []))
        self.run_bg(work)

    def set_positions(self, rows):
        self.positions_data = rows
        self.render_positions()

    def render_positions(self):
        self.pos.delete(*self.pos.get_children())
        real_t = unreal_t = 0.0
        open_n = 0
        for i, p in enumerate(self.positions_data):
            net = p.get("netQty", 0) or 0
            real = float(p.get("realized_profit", 0) or 0)
            total = float(p.get("pl", real) or 0)
            unreal = float(p.get("unrealized_profit", total - real) or 0)
            real_t += real
            unreal_t += unreal
            if net != 0:
                open_n += 1
            if self.v_open_only.get() and net == 0:
                continue
            side = "LONG" if net > 0 else "SHORT" if net < 0 else "CLOSED"
            self.pos.insert("", "end", iid=str(i), tags=("profit" if total >= 0 else "loss",), values=(
                p.get("symbol", ""), p.get("productType", ""), side, net,
                f"{float(p.get('buyAvg', 0) or 0):.2f}", f"{float(p.get('sellAvg', 0) or 0):.2f}",
                f"{float(p.get('ltp', 0) or 0):.2f}", f"{real:,.2f}", f"{unreal:,.2f}", f"{total:,.2f}"))
        self._apply_sort(self.pos)
        self.v_pnl.set(f"Open: {open_n}   Realized: {real_t:,.2f}   Unrealized: {unreal_t:,.2f}   "
                       f"TOTAL: {real_t + unreal_t:,.2f}")

    def _auto_tick(self):
        if self.v_auto.get():
            self.on_positions()
            self.on_funds(True)
            self.after(5000, self._auto_tick)

    def on_exit_all(self):
        api = self.get_api()
        if not api:
            return
        if not messagebox.askyesno(
                "EXIT ALL POSITIONS",
                "This will:\n\n"
                "  \u2022 cancel ALL pending orders, and\n"
                "  \u2022 square off ALL open positions with counter orders\n\n"
                "for your ENTIRE Fyers account (every instrument, not just NIFTY / SENSEX).\n\n"
                "Do you want to continue?", icon="warning", default="no"):
            return

        def work():
            code, data = api.exit_all()
            msg = data.get("message", "") if isinstance(data, dict) else data
            rc = data.get("code") if isinstance(data, dict) else None
            ok = code == 200 and isinstance(data, dict) and data.get("s") == "ok"
            if ok and rc == 201:
                self.log_t(f"EXIT ALL: counter order placed but NOT filled yet (low liquidity?). {msg} "
                           "- check the Positions tab.")
            elif ok:
                self.log_t(f"EXIT ALL accepted. {msg}")
            else:
                self.log_t(f"EXIT ALL failed -> HTTP {code}: {data}")
            time.sleep(2)
            self.ui(self.on_positions)
            self.ui(self.on_funds, True)
            self.ui(self.on_book)
        self.run_bg(work)

    # ---- contracts ------------------------------------------------------
    def refresh_contracts(self, force=False):
        def work():
            c = load_contracts(self.log_t, force)
            self.ui(self.set_contracts, c)
        self.run_bg(work)

    def set_contracts(self, contracts):
        self.contracts = contracts
        idx = {}
        for c in contracts:
            idx.setdefault(c["ul"], {}).setdefault(c["expiry"], {}).setdefault(c["type"], []).append(c)
        for ul in idx.values():
            for ex in ul.values():
                for lst in ex.values():
                    lst.sort(key=lambda x: x["strike"])
        self.index = idx
        self.on_ul()

    def on_ul(self):
        exps = sorted(self.index.get(self.v_ul.get(), {}))
        self.cb_exp["values"] = exps
        self.v_exp.set(exps[0] if exps else "")
        self.on_exp()

    def on_exp(self):
        lst = self.index.get(self.v_ul.get(), {}).get(self.v_exp.get(), {}).get(self.v_type.get(), [])
        strikes = [f"{c['strike']:g}" for c in lst]
        self.cb_strike["values"] = strikes
        if strikes:
            self.v_strike.set(strikes[len(strikes) // 2])
        self.update_info()

    def current_contract(self):
        try:
            k = float(self.v_strike.get())
        except ValueError:
            return None
        for c in self.index.get(self.v_ul.get(), {}).get(self.v_exp.get(), {}).get(self.v_type.get(), []):
            if abs(c["strike"] - k) < 1e-6:
                return c
        return None

    def leg_count(self):
        """How many legs the next 'Add to basket' click creates (1-50), or None if invalid."""
        try:
            n = int(self.v_legs.get())
        except ValueError:
            return None
        return n if 1 <= n <= 50 else None

    def update_info(self):
        c = self.current_contract()
        if not c:
            self.v_info.set("Pick a valid strike.\n\n\n")
            return
        try:
            lots = int(self.v_lots.get())
        except ValueError:
            lots = 0
        n = self.leg_count() or 0
        self.v_info.set(f"{c['symbol']}\nLot size {c['lot']} | Tick {c['tick']:g}\n"
                        f"Quantity = {lots} x {c['lot']} = {lots * c['lot']} per leg\n"
                        f"Will add {n} leg(s) = {n * lots * c['lot']} qty in total")

    def on_ltp(self):
        c, api = self.current_contract(), self.get_api()
        if not c or not api:
            return

        def work():
            p = api.ltp(c["symbol"])
            self.log_t(f"LTP {c['symbol']}: {p}")
            self.ui(self.v_price.set, str(p))
        self.run_bg(work)

    # ---- basket ---------------------------------------------------------
    def build_item(self, c=None):
        c = c or self.current_contract()
        if not c:
            messagebox.showwarning("Order", "Select a valid underlying / expiry / strike.")
            return None
        try:
            lots = int(self.v_lots.get())
            assert lots > 0
        except Exception:
            messagebox.showwarning("Order", "Lots must be a positive whole number.")
            return None
        price = 0.0
        if self.v_otype.get() == "LIMIT":
            try:
                price = float(self.v_price.get())
                assert price > 0
            except Exception:
                messagebox.showwarning("Order", "Enter a valid limit price.")
                return None
            tick = c["tick"]
            rounded = round(round(price / tick) * tick, 2)  # price must be a multiple of tick size
            if abs(rounded - price) > 1e-9:
                self.log(f"Limit price {price:g} rounded to tick size {tick:g}: {rounded:g}")
            price = rounded
        return {"side": self.v_side.get(), "symbol": c["symbol"], "lots": lots, "qty": lots * c["lot"],
                "otype": self.v_otype.get(), "price": price, "product": self.v_prod.get(),
                "validity": self.v_valid.get()}

    def on_add(self):
        n = self.leg_count()
        if n is None:
            messagebox.showwarning("Order", "No. of legs must be a whole number from 1 to 50.")
            return
        items = []
        for _ in range(n):
            it = self.build_item()   # a fresh, separate order dict for every leg
            if not it:
                return
            items.append(it)
        self.basket.extend(items)
        self.render_basket()
        kids = self.tv.get_children()
        if kids:
            self.tv.see(kids[-1])               # scroll to the newly added legs
        if n > 1:
            self.log(f"Added {n} legs of {items[0]['symbol']} to basket.")

    def _update_basket_total(self):
        b = self.basket
        if not b:
            self.v_basket_total.set("Basket empty")
            return
        buy = sum(i["qty"] for i in b if i["side"] == "BUY")
        sell = sum(i["qty"] for i in b if i["side"] == "SELL")
        self.v_basket_total.set(f"Total:  {len(b)} order(s)  |  {sum(i['lots'] for i in b):,} lots  |  "
                                f"qty {buy + sell:,}   (BUY {buy:,}  /  SELL {sell:,})")

    def render_basket(self):
        self._update_basket_total()
        self.tv.delete(*self.tv.get_children())
        for i, it in enumerate(self.basket):
            price = "MKT" if it["otype"] == "MARKET" else f"{it['price']:g}"
            self.tv.insert("", "end", iid=str(i), values=(
                i + 1, it["side"], it["symbol"], it["lots"], it["qty"], it["otype"], price, it["product"]))
        self._tag_sides(self.tv)
        self._restripe(self.tv)

    def on_remove(self):
        for iid in sorted((int(i) for i in self.tv.selection()), reverse=True):
            del self.basket[iid]
        self.render_basket()

    def on_clear(self):
        self.basket.clear()
        self.render_basket()

    # ---- placing --------------------------------------------------------
    def on_place_single(self):
        it = self.build_item()
        if it:
            self.confirm_and_place([it])

    def on_place_all(self):
        if not self.basket:
            messagebox.showinfo("Basket", "Basket is empty.")
            return
        self.confirm_and_place(list(self.basket), from_basket=True)

    def confirm_and_place(self, items, from_basket=False):
        api = self.get_api()
        if not api:
            return
        lines = [f"{i['side']:4} {i['qty']:>6}  {i['symbol']}  "
                 f"{'MKT' if i['otype'] == 'MARKET' else i['price']}  {i['product']}" for i in items]
        if self.v_confirm.get() and not messagebox.askyesno("Confirm LIVE orders",
                                   f"Place {len(items)} LIVE order(s) on your Fyers account?\n\n" +
                                   "\n".join(lines[:15]) + ("\n..." if len(lines) > 15 else "")):
            return
        self.execute(api, items, self.v_slice.get(), from_basket)

    def execute(self, api, items, slice_on, from_basket=False):
        """Place items one after another, in order. No confirmation here."""
        def work():
            placed = []
            n = len(items)
            lats = []
            for k, it in enumerate(items, 1):
                tag = f"[{k}/{n}]"
                what = f"{it['side']} {it['qty']} {it['symbol']}"
                try:
                    code, data = api.place(it, slice_on)
                except Exception as e:
                    self.log_t(f"{tag} FAILED {what}: {e} - check Order Book before retrying")
                    continue
                lat = api.last_ms
                if lat is not None:
                    lats.append(lat)
                ms = f" ({lat:.0f} ms)" if lat is not None else ""
                body_ok = isinstance(data, dict) and (
                    data.get("s") == "ok" or (data.get("id") and data.get("s") != "error"))
                rc = data.get("code") if isinstance(data, dict) else None
                if code == 200 and body_ok and rc == 201:
                    self.log_t(f"{tag}{ms} UNCONFIRMED {what} -> no acknowledgement. CHECK ORDER BOOK before re-placing.")
                elif code == 200 and body_ok:
                    placed.append(it)
                    self.log_t(f"{tag}{ms} OK  {what} -> id {data.get('id', '')} {data.get('message', '')}")
                else:
                    self.log_t(f"{tag}{ms} REJECTED {what} -> HTTP {code}: {data}")
                time.sleep(0.15)  # stay well under 10 orders/sec
            self.log_t(f"Done: {len(placed)}/{n} accepted."
                       + (f"  Order latency: avg {sum(lats) / len(lats):.0f} ms, max {max(lats):.0f} ms." if lats else ""))
            if from_basket and placed:
                # remove every accepted order; keep only failed / unconfirmed ones in the basket
                self.ui(self.remove_from_basket, placed)
                left = n - len(placed)
                self.log_t("Basket cleared." if not left else
                           f"{len(placed)} placed order(s) removed; {left} left in basket (failed/unconfirmed - "
                           "see log above, check Order Book before re-placing).")
            time.sleep(1)
            self.ui(self.on_positions)
            self.ui(self.on_funds, True)
        self.run_bg(work)

    def remove_from_basket(self, items):
        self.basket = [b for b in self.basket if not any(b is i for i in items)]
        self.render_basket()

    # ---- scheduling -----------------------------------------------------
    def keep_awake(self, on):
        """Stop Windows from sleeping while a schedule is pending."""
        if os.name != "nt":
            return
        try:
            import ctypes
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | (0x00000001 if on else 0))
        except Exception:
            pass

    def on_schedule(self):
        if self.sched:
            messagebox.showinfo("Schedule", "A schedule is already active. Cancel it first.")
            return
        if not self.basket:
            messagebox.showinfo("Basket", "Basket is empty.")
            return
        api = self.get_api()
        if not api:
            return
        txt = self.v_sched_time.get().strip()
        t = None
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                t = dt.datetime.strptime(txt, fmt).time()
                break
            except ValueError:
                pass
        if t is None:
            messagebox.showwarning("Schedule", "Enter the time as HH:MM:SS (24-hour), e.g. 09:15:00.")
            return
        target = dt.datetime.combine(dt.date.today(), t)
        if target <= dt.datetime.now() + dt.timedelta(seconds=3):
            messagebox.showwarning("Schedule", "That time has already passed today. Pick a later time.")
            return
        items = list(self.basket)  # snapshot: later basket edits do not affect it
        lines = [f"{k}. {i['side']:4} {i['qty']:>6}  {i['symbol']}  "
                 f"{'MKT' if i['otype'] == 'MARKET' else i['price']}  {i['product']}"
                 for k, i in enumerate(items, 1)]
        if not messagebox.askyesno(
                "Confirm SCHEDULED live orders",
                f"Place {len(items)} LIVE order(s) automatically at {target:%H:%M:%S} today?\n\n" +
                "\n".join(lines[:15]) + ("\n..." if len(lines) > 15 else "") +
                "\n\nThey will be sent WITHOUT any further confirmation.\n"
                "Keep this app open, the PC on, and the internet connected.\n"
                "Make sure your access token is valid for today."):
            return
        self.sched = {"target": target, "items": items, "slice": self.v_slice.get(), "api": api}
        self.keep_awake(True)
        self.log(f"Scheduled {len(items)} order(s) for {target:%H:%M:%S}.")

        def check():
            code, data = api.profile()
            good = code == 200 and isinstance(data, dict) and data.get("s") == "ok"
            self.log_t("Token check before schedule: " + ("OK" if good else
                       f"WARNING - failed (HTTP {code}). Update your token before the scheduled time."))
        self.run_bg(check)
        self._sched_tick()

    def _sched_tick(self):
        s = self.sched
        if not s:
            return
        remaining = (s["target"] - dt.datetime.now()).total_seconds()
        if remaining <= 0:
            self.sched = None
            self.keep_awake(False)
            if remaining < -30:  # e.g. PC was asleep - prices may have moved, do not fire blindly
                self.v_sched_status.set("Missed schedule - NOT placed.")
                self.log(f"Scheduled time missed by {int(-remaining)}s - orders NOT placed.")
                messagebox.showwarning("Schedule missed",
                                       "The scheduled time passed by more than 30 seconds "
                                       "(PC asleep or app busy). Orders were NOT placed.")
                return
            self.v_sched_status.set("Triggered - placing orders...")
            self.log(f"Scheduled time reached (+{-remaining * 1000:.0f} ms after target) - placing {len(s['items'])} order(s).")
            api = self.get_api(True) or s["api"]  # use latest saved token if changed
            self.execute(api, s["items"], s["slice"], from_basket=True)
            return
        h, rem = divmod(int(remaining), 3600)
        m, sec = divmod(rem, 60)
        self.v_sched_status.set(f"{len(s['items'])} order(s) at {s['target']:%H:%M:%S}  (in {h:02d}:{m:02d}:{sec:02d})")
        delay = 500 if remaining > 2 else max(1, int(remaining * 1000))
        self._sched_job = self.after(delay, self._sched_tick)

    def on_sched_cancel(self):
        if not self.sched:
            self.v_sched_status.set("No schedule set.")
            return
        if self._sched_job:
            self.after_cancel(self._sched_job)
            self._sched_job = None
        self.sched = None
        self.keep_awake(False)
        self.v_sched_status.set("Schedule cancelled.")
        self.log("Schedule cancelled.")

    def on_close(self):
        if self.sched and not messagebox.askyesno(
                "Scheduled orders pending",
                "A schedule is active. If you close the app, those orders will NOT be placed.\n\nExit anyway?",
                icon="warning", default="no"):
            return
        self.keep_awake(False)
        self.destroy()

    # ---- order book -----------------------------------------------------
    def on_book(self):
        api = self.get_api()
        if not api:
            return

        def work():
            code, data = api.orders()
            if code != 200 or not isinstance(data, dict) or data.get("s") != "ok":
                self.log_t(f"Order book -> HTTP {code}: {data}")
                return
            self.ui(self.render_book, data.get("orderBook", []) or [])
        self.run_bg(work)

    def render_book(self, rows):
        self.bk.delete(*self.bk.get_children())
        self.book_status = {}
        for o in rows:
            oid = str(o.get("id"))
            st = o.get("status")
            self.book_status[oid] = st
            price = o.get("tradedPrice") or o.get("limitPrice") or ""
            self.bk.insert("", "end", iid=oid, values=(
                oid, str(o.get("orderDateTime", ""))[-8:], o.get("symbol", ""),
                "BUY" if o.get("side") == 1 else "SELL", o.get("qty", ""), price,
                ORDER_TYPE.get(o.get("type"), o.get("type", "")), ORDER_STATUS.get(st, st),
                str(o.get("message", ""))[:80]))
        self._tag_sides(self.bk)
        self._apply_sort(self.bk)
        self.log(f"Order book: {len(rows)} orders.")

    def on_cancel(self):
        api = self.get_api()
        ids = [i for i in self.bk.selection() if self.book_status.get(i) in (4, 6)]  # transit / pending
        if not api or not ids:
            messagebox.showinfo("Cancel", "Select pending orders to cancel.")
            return
        if not messagebox.askyesno("Cancel", f"Cancel {len(ids)} order(s)?"):
            return

        def work():
            for oid in ids:
                code, data = api.cancel(oid)
                self.log_t(f"Cancel {oid} -> HTTP {code}: {data}")
                time.sleep(0.15)
            self.ui(self.on_book)
        self.run_bg(work)


if __name__ == "__main__":
    enable_dpi_awareness()
    App().mainloop()
