"""COA Request Builder - native Windows app (Tkinter), Fiori-style look.

Uses coa_engine.py, audit.py and audit_store.py unchanged.
Run:        python coa_gui.py
Self-test:  python coa_gui.py --selftest result.txt
Logo:       app_logo.png (optional; next to the exe, in the data folder, or bundled at build time)
"""
import hashlib
import hmac
import io
import json
import os
import queue
import sys
import tempfile
import threading
import traceback
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

APP_NAME = "COA Request Builder"
VERSION = "0.1 alpha"
AUTHOR = "zulkiflime"

# windowed exe has no console: give print/warnings somewhere to go
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

try:                                    # sharp text on high-DPI screens
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

SELFTEST = "--selftest" in sys.argv
DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), APP_NAME)
os.makedirs(DATA_DIR, exist_ok=True)
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
ERROR_LOG = os.path.join(DATA_DIR, "error.log")
DEFAULTS = {"enforce_wednesday": True, "db_path": "", "pin_salt": "", "pin_hash": "", "last_dir": ""}


def load_settings():
    s = dict(DEFAULTS)
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            s.update(json.load(f))
    except (OSError, ValueError):
        pass
    return s


SETTINGS = load_settings()


def save_settings():
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(SETTINGS, f, indent=1)
    except OSError:
        pass


# audit_store reads AUDIT_DB when it is imported, so set it first
if SELFTEST:
    os.environ["AUDIT_DB"] = os.path.join(tempfile.mkdtemp(), "audit.db")
elif "AUDIT_DB" not in os.environ:
    os.environ["AUDIT_DB"] = SETTINGS.get("db_path") or os.path.join(DATA_DIR, "audit.db")

import pandas as pd  # noqa: E402

import audit as aud  # noqa: E402
import audit_store as store  # noqa: E402
import coa_engine as eng  # noqa: E402


def _tz():
    try:
        return ZoneInfo(os.getenv("APP_TZ", "Asia/Jakarta"))
    except Exception:                       # no tz database available: Jakarta has no DST
        return timezone(timedelta(hours=7), "WIB")


TZ = _tz()


def now_local():
    return datetime.now(TZ)


# ---------------- look and feel (edit the colours here) ----------------
C = dict(shell="#354A5F", shell_sub="#C6D0DA",
         bg="#F5F6F7", panel="#FFFFFF", line="#D9D9D9", field_line="#89919A",
         text="#32363A", sub="#6A6D70",
         blue="#0A6ED1", blue_dark="#085CAD", blue_soft="#EBF5FE",
         red="#BB0000", green="#107E3E", amber="#E9730C")
FONT = ("Segoe UI", 10)
FONT_B = ("Segoe UI", 10, "bold")
FONT_S = ("Segoe UI", 9)
FONT_H = ("Segoe UI", 16)
FONT_K = ("Segoe UI", 20)

STRIPS = {   # state -> (icon, background, accent)
    "DONE": ("\u2713", "#F1FDF7", C["green"]),
    "SKIPPED": ("!", "#FEF7F1", C["amber"]),
    "UPCOMING": ("i", C["blue_soft"], C["blue"]),
    "DUE": ("!", "#FEF7F1", C["amber"]),
    "OVERDUE": ("\u2715", "#FFEBEB", C["red"]),
}
COMP_COLORS = {"Done": "#E2F0D9", "Skipped": "#FFF2CC", "Missed": "#F8CBAD",
               "Pending": "#DDEBF7", "Overdue": "#F4B084"}
LOGO_TRANSPARENT_CORNERS = True      # make the white corners of the logo picture transparent on the dark bar


def apply_theme(root):
    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except tk.TclError:
        pass
    root.configure(bg=C["bg"])
    root.option_add("*Font", "{Segoe UI} 10")
    root.option_add("*TCombobox*Listbox.font", "{Segoe UI} 10")
    root.option_add("*TCombobox*Listbox.background", "white")
    root.option_add("*TCombobox*Listbox.selectBackground", C["blue"])
    root.option_add("*TCombobox*Listbox.selectForeground", "white")

    st.configure(".", font=FONT, background=C["bg"], foreground=C["text"], bordercolor=C["line"],
                 focuscolor=C["bg"])
    st.configure("TFrame", background=C["bg"])
    st.configure("P.TFrame", background=C["panel"])
    st.configure("TLabel", background=C["bg"], foreground=C["text"])
    st.configure("P.TLabel", background=C["panel"], foreground=C["text"])
    st.configure("Sub.P.TLabel", background=C["panel"], foreground=C["sub"])
    st.configure("Bold.P.TLabel", background=C["panel"], font=FONT_B)

    for name, bgc in (("TCheckbutton", C["bg"]), ("P.TCheckbutton", C["panel"])):
        st.configure(name, background=bgc, focuscolor=bgc)
        st.map(name, background=[("active", bgc)],
               indicatorbackground=[("selected", C["blue"]), ("!selected", "white")],
               indicatorforeground=[("selected", "white")])

    field = dict(fieldbackground="white", bordercolor=C["field_line"], lightcolor="white",
                 darkcolor="white", padding=(5, 4))
    st.configure("TEntry", **field)
    st.map("TEntry", bordercolor=[("focus", C["blue"])],
           fieldbackground=[("readonly", "#FAFAFA"), ("disabled", "#EDEDED")])
    st.configure("TCombobox", background="white", arrowcolor=C["blue"], **field)
    st.map("TCombobox", fieldbackground=[("readonly", "white")], bordercolor=[("focus", C["blue"])],
           selectbackground=[("readonly", "white")], selectforeground=[("readonly", C["text"])])
    st.configure("TSpinbox", background="white", arrowcolor=C["blue"], **field)
    st.map("TSpinbox", bordercolor=[("focus", C["blue"])])

    # toolbar / secondary button: white with blue outline (like the SAP toolbar)
    st.configure("Tool.TButton", background="white", foreground=C["blue"], bordercolor=C["blue"],
                 lightcolor="white", darkcolor="white", focuscolor="white", padding=(12, 5), relief="flat")
    st.map("Tool.TButton", background=[("active", C["blue_soft"]), ("disabled", C["bg"])],
           foreground=[("disabled", "#A0A8B0")], bordercolor=[("disabled", C["line"])])
    # primary button: solid blue (like Save)
    st.configure("Primary.TButton", background=C["blue"], foreground="white", bordercolor=C["blue"],
                 lightcolor=C["blue"], darkcolor=C["blue"], focuscolor=C["blue"], padding=(16, 5),
                 relief="flat", font=FONT_B)
    st.map("Primary.TButton", background=[("active", C["blue_dark"]), ("disabled", "#C4D3E3")],
           foreground=[("disabled", "white")], bordercolor=[("active", C["blue_dark"]), ("disabled", "#C4D3E3")],
           lightcolor=[("active", C["blue_dark"]), ("disabled", "#C4D3E3")],
           darkcolor=[("active", C["blue_dark"]), ("disabled", "#C4D3E3")])
    # text-style button (Exit)
    st.configure("Link.TButton", background="white", foreground=C["blue"], borderwidth=0, bordercolor="white",
                 lightcolor="white", darkcolor="white", focuscolor="white", padding=(10, 5), relief="flat")
    st.map("Link.TButton", background=[("active", C["blue_soft"])])

    st.configure("TScrollbar", background="#C6CBD0", troughcolor="#F2F2F2", bordercolor="#F2F2F2",
                 arrowcolor=C["text"], lightcolor="#C6CBD0", darkcolor="#C6CBD0")
    st.configure("Treeview", background="white", fieldbackground="white", foreground=C["text"],
                 rowheight=26, borderwidth=0, font=FONT)
    st.configure("Treeview.Heading", background="#F2F2F2", foreground=C["text"], font=FONT_B,
                 relief="flat", padding=(6, 5), bordercolor=C["line"])
    st.map("Treeview", background=[("selected", "#D1E8FF")], foreground=[("selected", C["text"])])
    st.map("Treeview.Heading", background=[("active", "#E4E4E4")])


# ---------------- logo ----------------
def _app_dirs():
    dirs = []
    if getattr(sys, "frozen", False):
        dirs.append(os.path.dirname(sys.executable))
    else:
        dirs.append(os.path.dirname(os.path.abspath(__file__)))
    dirs.append(DATA_DIR)
    mp = getattr(sys, "_MEIPASS", None)
    if mp:
        dirs.append(mp)
    return dirs


def find_logo():
    for d in _app_dirs():
        for n in ("app_logo.png", "app_logo.jpg", "app_logo.jpeg"):
            p = os.path.join(d, n)
            if os.path.isfile(p):
                return p
    return None


def load_logo_image(height):
    """Return a PhotoImage of the logo scaled to `height` pixels, or None."""
    path = find_logo()
    if not path:
        return None
    try:
        from PIL import Image, ImageDraw, ImageTk
        img = Image.open(path).convert("RGBA")
        w, h = img.size
        if LOGO_TRANSPARENT_CORNERS and img.getpixel((0, 0))[3] > 0:
            for xy in ((0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)):
                ImageDraw.floodfill(img, xy, (255, 255, 255, 0), thresh=40)
        img = img.resize((max(1, int(w * height / h)), height), Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception:
        try:                                    # Pillow missing: plain Tk can still read PNG
            ph = tk.PhotoImage(file=path)
            f = max(1, ph.height() // height)
            return ph.subsample(f, f)
        except Exception:
            return None


def draw_logo(parent, height, bg):
    """Fallback badge drawn when app_logo.png is not available."""
    w = int(height * 2.2)
    c = tk.Canvas(parent, width=w, height=height, bg=bg, highlightthickness=0)
    c.create_rectangle(2, 4, w - 2, height - 4, fill="#0A6ED1", outline="")
    c.create_text(w / 2, height * 0.5, text="COA", fill="white",
                  font=("Segoe UI", -max(10, int(height * 0.42)), "bold"))
    return c


# ---------------- supervisor PIN (stored as a salted hash) ----------------
def pin_is_set():
    return bool(SETTINGS.get("pin_hash"))


def _hash_pin(pin, salt_hex):
    return hashlib.pbkdf2_hmac("sha256", pin.encode("utf-8"), bytes.fromhex(salt_hex), 200000).hex()


def set_pin(pin):
    salt = os.urandom(16).hex()
    SETTINGS["pin_salt"], SETTINGS["pin_hash"] = salt, _hash_pin(pin, salt)
    save_settings()


def check_pin(pin):
    return pin_is_set() and hmac.compare_digest(_hash_pin(pin, SETTINGS["pin_salt"]), SETTINGS["pin_hash"])


# ---------------- small helpers ----------------
def log_error(text):
    try:
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write("%s\n%s\n\n" % (now_local().isoformat(timespec="seconds"), text))
    except OSError:
        pass


def last_line(tb_text):
    lines = [l for l in tb_text.strip().splitlines() if l.strip()]
    return lines[-1] if lines else "Unknown error"


def cell(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, float):
        return ("%.3f" % v).rstrip("0").rstrip(".")
    return str(v)


def read_bytes(path):
    try:
        with open(path, "rb") as f:
            return f.read()
    except OSError as e:
        raise ValueError("Cannot read %s\n(%s)\nIf it is open in another program, close it and try again."
                         % (os.path.basename(path), e))


def ask_open(title):
    kw = dict(title=title, filetypes=[("Excel files", "*.xlsx")])
    if SETTINGS.get("last_dir") and os.path.isdir(SETTINGS["last_dir"]):
        kw["initialdir"] = SETTINGS["last_dir"]
    p = filedialog.askopenfilename(**kw)
    if p:
        SETTINGS["last_dir"] = os.path.dirname(p)
        save_settings()
    return p


def save_bytes(data, default_name, title):
    kw = dict(title=title, defaultextension=".xlsx", initialfile=default_name,
              filetypes=[("Excel files", "*.xlsx")])
    if SETTINGS.get("last_dir") and os.path.isdir(SETTINGS["last_dir"]):
        kw["initialdir"] = SETTINGS["last_dir"]
    p = filedialog.asksaveasfilename(**kw)
    if not p:
        return None
    try:
        with open(p, "wb") as f:
            f.write(data)
    except OSError as e:
        messagebox.showerror(APP_NAME, "Could not save the file:\n%s" % e)
        return None
    SETTINGS["last_dir"] = os.path.dirname(p)
    save_settings()
    return p


def make_tree(parent, columns, widths, height=10):
    outer = tk.Frame(parent, bg=C["panel"], highlightbackground=C["line"], highlightthickness=1)
    tree = ttk.Treeview(outer, columns=columns, show="headings", height=height, selectmode="browse")
    for c, w in zip(columns, widths):
        tree.heading(c, text=c)
        tree.column(c, width=w, anchor="w", stretch=True)
    ys = ttk.Scrollbar(outer, orient="vertical", command=tree.yview)
    xs = ttk.Scrollbar(outer, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
    tree.grid(row=0, column=0, sticky="nsew")
    ys.grid(row=0, column=1, sticky="ns")
    xs.grid(row=1, column=0, sticky="ew")
    outer.rowconfigure(0, weight=1)
    outer.columnconfigure(0, weight=1)
    return outer, tree


def fill_rows(tree, rows, tags=None):
    tree.delete(*tree.get_children())
    for i, r in enumerate(rows):
        tree.insert("", "end", values=[cell(v) for v in r], tags=((tags[i],) if tags else ()))


def status_tag(status):
    return "S_" + str(status).replace(" ", "_")


GEN_COLS = ["Block", "Customer", "Truck", "Product", "Lot", "Expired", "Qty", "Unit", "Remarks"]
GEN_W = [50, 220, 90, 240, 120, 90, 80, 50, 220]
AUDIT_COLS = ["Block", "Customer", "Truck", "Product", "Lot", "Auto Qty", "Manual Qty", "Unit",
              "Auto SLED", "Manual SLED", "Status", "Note"]
AUDIT_W = [50, 200, 80, 220, 110, 75, 80, 45, 85, 85, 150, 260]


class Worker:
    """Runs a function in a background thread; results come back on the Tk thread."""
    def __init__(self, root):
        self.root, self.q = root, queue.Queue()
        root.after(100, self._poll)

    def run(self, fn, done, fail):
        def target():
            try:
                self.q.put((done, fn()))
            except Exception:
                self.q.put((fail, traceback.format_exc()))
        threading.Thread(target=target, daemon=True).start()

    def _poll(self):
        try:
            while True:
                cb, payload = self.q.get_nowait()
                cb(payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll)


# ---------------- reusable pieces ----------------
class TabBar(tk.Frame):
    """Fiori-style tabs: text with a blue underline on the selected one."""
    def __init__(self, parent, titles, command):
        super().__init__(parent, bg=C["panel"])
        self.command, self.items = command, []
        for i, t in enumerate(titles):
            f = tk.Frame(self, bg=C["panel"], cursor="hand2")
            lab = tk.Label(f, text=t, bg=C["panel"], fg=C["text"], font=FONT, padx=18, pady=9)
            line = tk.Frame(f, height=3, bg=C["panel"])
            lab.pack()
            line.pack(fill="x")
            for w in (f, lab):
                w.bind("<Button-1>", lambda e, i=i: self.select(i))
            f.pack(side="left")
            self.items.append((lab, line))

    def select(self, i):
        for j, (lab, line) in enumerate(self.items):
            on = j == i
            lab.configure(fg=C["blue"] if on else C["text"], font=FONT_B if on else FONT)
            line.configure(bg=C["blue"] if on else C["panel"])
        self.command(i)


class Strip(tk.Frame):
    """Message strip with a coloured left edge."""
    def __init__(self, parent):
        super().__init__(parent)
        self.bar = tk.Frame(self, width=5)
        self.bar.pack(side="left", fill="y")
        self.icon = tk.Label(self, font=("Segoe UI", 11, "bold"), padx=12)
        self.icon.pack(side="left")
        self.msg = tk.Label(self, anchor="w", font=FONT_B, pady=9)
        self.msg.pack(side="left", fill="x", expand=True)

    def show(self, icon, text, bg, accent):
        self.configure(bg=bg)
        self.bar.configure(bg=accent)
        self.icon.configure(text=icon, bg=bg, fg=accent)
        self.msg.configure(text=text, bg=bg, fg=C["text"])


class Kpi(tk.Frame):
    """Small white tile with a big number and a caption."""
    def __init__(self, parent, caption):
        super().__init__(parent, bg=C["panel"], highlightbackground=C["line"], highlightthickness=1)
        self.val = tk.StringVar(value="\u2013")
        self.lbl = tk.Label(self, textvariable=self.val, bg=C["panel"], fg=C["text"], font=FONT_K,
                            anchor="w", width=9)
        self.lbl.pack(anchor="w", padx=14, pady=(8, 0))
        tk.Label(self, text=caption, bg=C["panel"], fg=C["sub"], font=FONT_S, anchor="w").pack(
            anchor="w", padx=14, pady=(0, 8))

    def set(self, value, color=None):
        self.val.set(str(value))
        self.lbl.configure(fg=color or C["text"])


def page_header(parent, title, subtitle):
    h = tk.Frame(parent, bg=C["bg"])
    h.pack(fill="x", padx=16, pady=(14, 6))
    tk.Label(h, text=title, font=FONT_H, bg=C["bg"], fg=C["text"]).pack(anchor="w")
    tk.Label(h, text=subtitle, font=FONT_S, bg=C["bg"], fg=C["sub"]).pack(anchor="w")


def page_footer(parent):
    """White action bar at the bottom of a page (pack it BEFORE the expanding widgets)."""
    foot = tk.Frame(parent, bg=C["panel"])
    foot.pack(side="bottom", fill="x")
    tk.Frame(parent, height=1, bg=C["line"]).pack(side="bottom", fill="x")
    return foot


# ---------------- file + settings block (shared by Generate and Weekly Audit) ----------------
class Inputs(tk.Frame):
    def __init__(self, parent, with_manual=False):
        super().__init__(parent, bg=C["panel"], highlightbackground=C["line"], highlightthickness=1)
        self.with_manual = with_manual
        self.plan_path, self.exp_path, self.manual_path = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.sheet = tk.StringVar()
        self.use_date = tk.BooleanVar(value=False)
        self.date_text = tk.StringVar(value=now_local().strftime("%d/%m/%Y"))
        self.ref = tk.IntVar(value=1)
        self.excl = tk.BooleanVar(value=True)
        self.slocs = tk.StringVar(value="2002")
        self.mobil = tk.BooleanVar(value=True)

        body = ttk.Frame(self, style="P.TFrame", padding=(14, 10))
        body.pack(fill="x")
        body.columnconfigure(1, weight=1)
        r = 0
        self._file_row(body, r, "Rencana Kirim (.xlsx)", self.plan_path, self.pick_plan); r += 1
        self._file_row(body, r, "EXPORT SAP stock (.xlsx)", self.exp_path,
                       lambda: self._pick(self.exp_path, "Select the EXPORT SAP stock file")); r += 1
        if with_manual:
            self._file_row(body, r, "Manual PERMINTAAN COA (.xlsx)", self.manual_path,
                           lambda: self._pick(self.manual_path, "Select the manual PERMINTAAN COA file")); r += 1

        ttk.Label(body, text="Plan sheet", style="P.TLabel").grid(row=r, column=0, sticky="w", pady=4)
        self.sheet_box = ttk.Combobox(body, textvariable=self.sheet, state="readonly", width=30)
        self.sheet_box.grid(row=r, column=1, sticky="w", padx=8); r += 1

        dfr = ttk.Frame(body, style="P.TFrame")
        dfr.grid(row=r, column=0, columnspan=3, sticky="w", pady=4); r += 1
        ttk.Checkbutton(dfr, variable=self.use_date, command=self._date_state, style="P.TCheckbutton",
                        text="Override plan date (otherwise the date inside the sheet is used):").pack(side="left")
        self.date_entry = ttk.Entry(dfr, textvariable=self.date_text, width=12, state="disabled")
        self.date_entry.pack(side="left", padx=8)
        ttk.Label(dfr, text="DD/MM/YYYY", style="Sub.P.TLabel").pack(side="left")

        sfr = ttk.Frame(body, style="P.TFrame")
        sfr.grid(row=r, column=0, columnspan=3, sticky="w", pady=4)
        ttk.Label(sfr, text="Shelf-life % reference = plan date minus (days):", style="P.TLabel").pack(side="left")
        ttk.Spinbox(sfr, from_=0, to=10, width=4, textvariable=self.ref).pack(side="left", padx=(6, 16))
        ttk.Checkbutton(sfr, variable=self.excl, text="Exclude expired lots", style="P.TCheckbutton").pack(
            side="left", padx=(0, 16))
        ttk.Label(sfr, text="Trading storage locations:", style="P.TLabel").pack(side="left")
        ttk.Entry(sfr, textvariable=self.slocs, width=12).pack(side="left", padx=(6, 16))
        ttk.Checkbutton(sfr, variable=self.mobil, text="Write JENIS MOBIL in column K",
                        style="P.TCheckbutton").pack(side="left")

    def _file_row(self, body, r, label, var, cmd):
        ttk.Label(body, text=label, style="P.TLabel").grid(row=r, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=var, state="readonly").grid(row=r, column=1, sticky="ew", padx=8)
        ttk.Button(body, text="Browse...", style="Tool.TButton", command=cmd).grid(row=r, column=2)

    def _date_state(self):
        self.date_entry.configure(state="normal" if self.use_date.get() else "disabled")

    def _pick(self, var, title):
        p = ask_open(title)
        if p:
            var.set(p)

    def pick_plan(self):
        p = ask_open("Select the Rencana Kirim file")
        if not p:
            return
        try:
            sheets = eng.list_plan_sheets(io.BytesIO(read_bytes(p)))
        except Exception as e:
            messagebox.showerror(APP_NAME, "Cannot read the plan file:\n%s" % e)
            return
        self.plan_path.set(p)
        self.sheet_box.configure(values=sheets)
        self.sheet.set(sheets[eng.guess_sheet_index(sheets, now_local().date())])

    def collect(self):
        """Validate and return (plan_bytes, export_bytes, manual_bytes_or_None, sheet, opts)."""
        if not self.plan_path.get():
            raise ValueError("Choose the Rencana Kirim file.")
        if not self.exp_path.get():
            raise ValueError("Choose the EXPORT SAP stock file.")
        if self.with_manual and not self.manual_path.get():
            raise ValueError("Choose the manual PERMINTAAN COA file.")
        if not self.sheet.get():
            raise ValueError("Choose the plan sheet.")
        try:
            ref = int(self.ref.get())
        except (tk.TclError, ValueError):
            raise ValueError("Shelf-life reference must be a whole number from 0 to 10.")
        if not 0 <= ref <= 10:
            raise ValueError("Shelf-life reference must be a whole number from 0 to 10.")
        override = None
        if self.use_date.get():
            try:
                override = datetime.strptime(self.date_text.get().strip(), "%d/%m/%Y").date()
            except ValueError:
                raise ValueError("Plan date must look like 27/10/2026.")
        opts = dict(ref_offset_days=ref, exclude_expired=bool(self.excl.get()), show_mobil=bool(self.mobil.get()),
                    plan_date_override=override,
                    trading_slocs=tuple(s.strip() for s in self.slocs.get().split(",") if s.strip()))
        manual = read_bytes(self.manual_path.get()) if self.with_manual else None
        return read_bytes(self.plan_path.get()), read_bytes(self.exp_path.get()), manual, self.sheet.get(), opts


# ---------------- main window ----------------
class App:
    def __init__(self, root):
        self.root = root
        store.init()
        apply_theme(root)
        root.title("%s %s by %s" % (APP_NAME, VERSION, AUTHOR))
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w, h = min(1300, sw - 40), min(920, sh - 90)
        root.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw - w) // 2), max(0, (sh - h) // 3)))
        root.minsize(900, 600)
        self.worker = Worker(root)
        self.status = None
        self.gen_result = None
        self.audit_result = None
        self.hist_rows = {}
        self.hist_detail = None

        self._build_shell()
        self._build_toolbar()
        self.strip = Strip(root)
        self.strip.pack(fill="x")
        self.tabs = TabBar(root, ["Generate COA request", "Weekly Audit", "Audit history"], self.show_page)
        self.tabs.pack(fill="x")
        tk.Frame(root, height=1, bg=C["line"]).pack(fill="x")
        self.host = tk.Frame(root, bg=C["bg"])
        self.host.pack(fill="both", expand=True)
        self.pages = [tk.Frame(self.host, bg=C["bg"]) for _ in range(3)]
        self._build_generate(self.pages[0])
        self._build_audit(self.pages[1])
        self._build_history(self.pages[2])
        self.tabs.select(0)
        self.refresh_status(initial=True)
        root.after(60000, self._tick)

    def show_page(self, i):
        for p in self.pages:
            p.pack_forget()
        self.pages[i].pack(fill="both", expand=True)

    def _tick(self):
        self.refresh_status()
        self.root.after(60000, self._tick)

    # ---- shell bar (dark top bar with the logo) ----
    def _build_shell(self):
        shell = tk.Frame(self.root, bg=C["shell"], height=50)
        shell.pack(fill="x")
        shell.pack_propagate(False)
        self.logo_img = load_logo_image(32)
        if self.logo_img is not None:
            tk.Label(shell, image=self.logo_img, bg=C["shell"]).pack(side="left", padx=(14, 14))
            try:
                self.root.iconphoto(True, self.logo_img)
            except tk.TclError:
                pass
        else:
            draw_logo(shell, 32, C["shell"]).pack(side="left", padx=(14, 14))
        tk.Label(shell, text=APP_NAME, bg=C["shell"], fg="white", font=("Segoe UI", 12, "bold")).pack(side="left")
        self.clock = tk.Label(shell, bg=C["shell"], fg=C["shell_sub"], font=FONT_S)
        self.clock.pack(side="right", padx=16)

    # ---- toolbar (white, blue outlined buttons) ----
    def _build_toolbar(self):
        bar = tk.Frame(self.root, bg=C["panel"])
        bar.pack(fill="x")
        tk.Frame(self.root, height=1, bg=C["line"]).pack(fill="x")

        self.menu_main = tk.Menu(self.root, tearoff=0)
        self.menu_main.add_command(label="Open data folder", command=lambda: os.startfile(DATA_DIR))
        self.menu_main.add_separator()
        self.menu_main.add_command(label="Exit", command=self.root.destroy)

        self.enforce_var = tk.BooleanVar(value=bool(SETTINGS["enforce_wednesday"]))
        self.menu_set = tk.Menu(self.root, tearoff=0)
        self.menu_set.add_checkbutton(label="Lock Generate from Wednesday until the weekly audit is saved",
                                      variable=self.enforce_var, command=self.toggle_enforce)
        self.menu_set.add_command(label="Set supervisor PIN...", command=self.set_pin_dialog)

        for text, menu in (("Menu  \u25BE", self.menu_main), ("Settings  \u25BE", self.menu_set)):
            b = ttk.Button(bar, text=text, style="Tool.TButton")
            b.configure(command=lambda b=b, m=menu: self.popup(b, m))
            b.pack(side="left", padx=(10, 0), pady=7)
        ttk.Button(bar, text="About", style="Tool.TButton", command=self.about).pack(side="left", padx=(8, 0))
        ttk.Button(bar, text="Exit", style="Link.TButton", command=self.root.destroy).pack(side="right", padx=12)

    def popup(self, widget, menu):
        try:
            menu.tk_popup(widget.winfo_rootx(), widget.winfo_rooty() + widget.winfo_height())
        finally:
            menu.grab_release()

    def ask_pin(self, prompt="Supervisor PIN:"):
        pin = simpledialog.askstring(APP_NAME, prompt, show="*", parent=self.root)
        if pin is None:
            return False
        if not check_pin(pin):
            messagebox.showerror(APP_NAME, "Wrong PIN.")
            return False
        return True

    def toggle_enforce(self):
        new = bool(self.enforce_var.get())
        if pin_is_set() and not self.ask_pin("Supervisor PIN (needed to change the lock):"):
            self.enforce_var.set(not new)
            return
        SETTINGS["enforce_wednesday"] = new
        save_settings()
        self.refresh_status()

    def set_pin_dialog(self):
        if pin_is_set() and not self.ask_pin("Current supervisor PIN:"):
            return
        p1 = simpledialog.askstring(APP_NAME, "New supervisor PIN (at least 4 characters):",
                                    show="*", parent=self.root)
        if p1 is None:
            return
        if len(p1) < 4:
            messagebox.showwarning(APP_NAME, "The PIN must have at least 4 characters.")
            return
        p2 = simpledialog.askstring(APP_NAME, "Type the new PIN again:", show="*", parent=self.root)
        if p2 != p1:
            messagebox.showwarning(APP_NAME, "The two entries do not match. PIN not changed.")
            return
        set_pin(p1)
        messagebox.showinfo(APP_NAME, "Supervisor PIN saved.")
        self.refresh_status()

    def about(self):
        messagebox.showinfo(APP_NAME, "%s %s\n\nAudit history file:\n%s\n\nSettings and error log:\n%s\n\n"
                            "History is stored on this PC only, unless the history file is moved to a shared folder."
                            % (APP_NAME, VERSION, store.DB_PATH, DATA_DIR))

    # ---- weekly status strip ----
    def refresh_status(self, initial=False):
        now = now_local()
        self.status = store.week_status(now, bool(SETTINGS["enforce_wednesday"]))
        s, w, locked = self.status["state"], self.status["iso_week"], self.status["locked"]
        tail = "  Generate is locked until it is saved." if locked else ""
        msg = {
            "DONE": "Weekly audit for week %d is done." % w,
            "SKIPPED": "Week %d audit was skipped by a supervisor (logged)." % w,
            "UPCOMING": "Weekly audit is due on Wednesday. Prepare your manual picks.",
            "DUE": "Weekly audit is due TODAY (Wednesday)." + tail,
            "OVERDUE": "Weekly audit for week %d is OVERDUE." % w + tail,
        }[s]
        icon, bg, accent = STRIPS[s]
        self.strip.show(icon, msg, bg, accent)
        self.clock.configure(text=now.strftime("%a %d/%m/%Y  %H:%M") + "  (Jakarta)")
        self.btn_generate.configure(state="disabled" if locked else "normal")
        self.lock_label.configure(text="LOCKED: complete this week's Weekly Audit first." if locked else "")
        if initial and locked:
            self.tabs.select(1)
        self.refresh_history()

    # ================= Generate page =================
    def _build_generate(self, t):
        foot = page_footer(t)
        self.btn_generate = ttk.Button(foot, text="Generate", style="Primary.TButton", command=self.do_generate)
        self.btn_generate.pack(side="right", padx=(6, 14), pady=8)
        self.btn_save_gen = ttk.Button(foot, text="Save PERMINTAAN COA...", style="Tool.TButton",
                                       command=self.save_generated, state="disabled")
        self.btn_save_gen.pack(side="right", pady=8)
        self.lock_label = tk.Label(foot, bg=C["panel"], fg=C["red"], font=FONT_B)
        self.lock_label.pack(side="left", padx=14)

        page_header(t, "Generate COA request",
                    "Rencana Kirim + EXPORT SAP stock  \u2192  PERMINTAAN COA (first-expiry-first-out)")
        self.gen_in = Inputs(t)
        self.gen_in.pack(fill="x", padx=16)

        row = tk.Frame(t, bg=C["bg"])
        row.pack(fill="x", padx=16, pady=10)
        self.k_blocks, self.k_lines, self.k_flag = Kpi(row, "Blocks"), Kpi(row, "Plan lines"), Kpi(row, "Lines flagged CEK STOK")
        for k in (self.k_blocks, self.k_lines, self.k_flag):
            k.pack(side="left", padx=(0, 10))
        self.gen_warn = tk.Label(row, bg=C["bg"], fg=C["red"], font=FONT_B)
        self.gen_warn.pack(side="left", padx=10)
        self.gen_busy = tk.Label(row, bg=C["bg"], fg=C["sub"], font=FONT)
        self.gen_busy.pack(side="left")

        fr, self.gen_tree = make_tree(t, GEN_COLS, GEN_W, height=12)
        fr.pack(fill="both", expand=True, padx=16, pady=(0, 12))

    def do_generate(self):
        if self.status["locked"]:
            messagebox.showwarning(APP_NAME, "Locked: complete this week's Weekly Audit first.")
            return
        try:
            plan_b, exp_b, _, sheet, opts = self.gen_in.collect()
        except ValueError as e:
            messagebox.showwarning(APP_NAME, str(e))
            return
        self.btn_generate.configure(state="disabled")
        self.btn_save_gen.configure(state="disabled")
        self.gen_busy.configure(text="Working...")
        self.worker.run(lambda: eng.generate(plan_b, sheet, exp_b, **opts), self._gen_done, self._gen_fail)

    def _gen_done(self, res):
        self.gen_result = res
        st = res["stats"]
        self.gen_busy.configure(text="")
        self.k_blocks.set(st["blocks"])
        self.k_lines.set(st["lines"])
        self.k_flag.set(st["flagged"], C["red"] if st["flagged"] else None)
        self.gen_warn.configure(text="Some lines have no / not enough unrestricted stock. "
                                     "Look for 'CEK STOK' in the file." if st["flagged"] else "")
        df = eng.flatten(res["blocks"])
        fill_rows(self.gen_tree, [] if df.empty else df.reindex(columns=GEN_COLS).values.tolist())
        self.btn_save_gen.configure(state="normal")
        self.refresh_status()

    def _gen_fail(self, tb):
        log_error(tb)
        self.gen_busy.configure(text="")
        self.refresh_status()
        messagebox.showerror(APP_NAME, "Failed: %s" % last_line(tb))

    def save_generated(self):
        res = self.gen_result
        if not res:
            return
        name = "PERMINTAAN COA %s.xlsx" % res["plan_date"].strftime("%d %B %Y").upper()
        p = save_bytes(res["xlsx"], name, "Save PERMINTAAN COA")
        if p and messagebox.askyesno(APP_NAME, "Saved. Open the file now?"):
            os.startfile(p)

    # ================= Weekly Audit page =================
    def _build_audit(self, t):
        foot = page_footer(t)
        self.btn_save_audit = ttk.Button(foot, text="Save this audit (counts for this week)",
                                         style="Primary.TButton", command=self.save_audit, state="disabled")
        self.btn_save_audit.pack(side="right", padx=(6, 14), pady=8)
        self.btn_save_report = ttk.Button(foot, text="Save audit report...", style="Tool.TButton",
                                          command=self.save_audit_report, state="disabled")
        self.btn_save_report.pack(side="right", pady=8)
        ttk.Button(foot, text="Supervisor: skip this week...", style="Tool.TButton",
                   command=self.skip_week).pack(side="left", padx=14, pady=8)

        page_header(t, "Weekly Audit",
                    "Manual picks vs automated picks. Use the same plan and the same SAP EXPORT snapshot "
                    "that you used when picking manually.")
        self.aud_in = Inputs(t, with_manual=True)
        self.aud_in.pack(fill="x", padx=16)

        bar = tk.Frame(t, bg=C["bg"])
        bar.pack(fill="x", padx=16, pady=8)
        ttk.Label(bar, text="Auditor name").pack(side="left")
        self.auditor = tk.StringVar()
        ttk.Entry(bar, textvariable=self.auditor, width=28).pack(side="left", padx=8)
        self.btn_compare = ttk.Button(bar, text="Compare", style="Primary.TButton", command=self.do_compare)
        self.btn_compare.pack(side="left", padx=4)
        self.aud_busy = tk.Label(bar, bg=C["bg"], fg=C["sub"], font=FONT)
        self.aud_busy.pack(side="left", padx=8)

        row = tk.Frame(t, bg=C["bg"])
        row.pack(fill="x", padx=16)
        self.k_match, self.k_qty = Kpi(row, "Row match"), Kpi(row, "Qty agreement")
        self.k_rows, self.k_blk = Kpi(row, "Rows compared"), Kpi(row, "Blocks paired")
        for k in (self.k_match, self.k_qty, self.k_rows, self.k_blk):
            k.pack(side="left", padx=(0, 10))
        self.aud_status = tk.StringVar()
        tk.Label(row, textvariable=self.aud_status, bg=C["bg"], fg=C["sub"], font=FONT_S,
                 justify="left", anchor="w").pack(side="left", padx=10)

        nf = tk.Frame(t, bg=C["bg"])
        nf.pack(side="bottom", fill="x", padx=16, pady=(4, 8))
        ttk.Label(nf, text="Notes / reason for differences (optional)").pack(anchor="w")
        self.notes = tk.Text(nf, height=2, wrap="word", relief="flat", font=FONT, highlightthickness=1,
                             highlightbackground=C["field_line"], highlightcolor=C["blue"])
        self.notes.pack(fill="x")

        self.only_diff = tk.BooleanVar(value=True)
        ttk.Checkbutton(t, variable=self.only_diff, text="Show differences only",
                        command=self.show_audit).pack(anchor="w", padx=16, pady=(8, 2))
        fr, self.aud_tree = make_tree(t, AUDIT_COLS, AUDIT_W, height=6)
        fr.pack(fill="both", expand=True, padx=16, pady=(0, 4))
        for st, col in aud.COLORS.items():
            self.aud_tree.tag_configure(status_tag(st), background="#" + col)

    def do_compare(self):
        try:
            plan_b, exp_b, man_b, sheet, opts = self.aud_in.collect()
        except ValueError as e:
            messagebox.showwarning(APP_NAME, str(e))
            return
        auditor = self.auditor.get().strip()
        if not auditor:
            messagebox.showwarning(APP_NAME, "Enter the auditor name.")
            return

        def job():
            res = eng.generate(plan_b, sheet, exp_b, **opts)
            manual_blocks = aud.parse_manual(man_b)
            if not manual_blocks:
                raise ValueError("No data found in the manual file (expected the PERMINTAAN COA layout).")
            summary, df = aud.compare(eng.blocks_list(res["blocks"]), manual_blocks,
                                      res["snapshot"], res["name2kode"])
            return dict(summary=summary, df=df, plan_date=res["plan_date"], auditor=auditor)

        self.btn_compare.configure(state="disabled")
        self.btn_save_report.configure(state="disabled")
        self.btn_save_audit.configure(state="disabled")
        self.aud_busy.configure(text="Working...")
        self.worker.run(job, self._cmp_done, self._cmp_fail)

    def _cmp_done(self, a):
        self.audit_result = a
        self.aud_busy.configure(text="")
        self.btn_compare.configure(state="normal")
        self.btn_save_report.configure(state="normal")
        self.btn_save_audit.configure(state="normal")
        self.show_audit()

    def _cmp_fail(self, tb):
        log_error(tb)
        self.audit_result = None
        self.aud_busy.configure(text="")
        self.btn_compare.configure(state="normal")
        messagebox.showerror(APP_NAME, "Failed: %s" % last_line(tb))

    def show_audit(self):
        a = self.audit_result
        if not a:
            return
        sm, df = a["summary"], a["df"]
        self.k_match.set("%.1f%%" % sm["match_pct"], C["green"] if sm["match_pct"] >= 100 else None)
        self.k_qty.set("%.1f%%" % sm["qty_agreement"], C["green"] if sm["qty_agreement"] >= 100 else None)
        self.k_rows.set(sm["rows_total"])
        self.k_blk.set("%d / %d" % (sm["blocks_matched"], sm["blocks_auto"]))
        self.aud_status.set("\n".join("%s: %d" % kv for kv in sm["by_status"].items()))
        if df.empty:
            fill_rows(self.aud_tree, [])
            return
        view = df[df["Status"] != aud.STATUS_MATCH] if self.only_diff.get() else df
        fill_rows(self.aud_tree, view.reindex(columns=AUDIT_COLS).values.tolist(),
                  [status_tag(s) for s in view["Status"]])

    def save_audit_report(self):
        a = self.audit_result
        if not a:
            return
        meta = {"Plan date": a["plan_date"].strftime("%d/%m/%Y"), "Auditor": a["auditor"],
                "Audit time": now_local().strftime("%d/%m/%Y %H:%M")}
        data = aud.build_report(a["summary"], a["df"], meta)
        p = save_bytes(data, "AUDIT %s.xlsx" % a["plan_date"].strftime("%Y-%m-%d"), "Save audit report")
        if p and messagebox.askyesno(APP_NAME, "Saved. Open the file now?"):
            os.startfile(p)

    def save_audit(self):
        a = self.audit_result
        if not a:
            return
        try:
            store.add_audit(now_local(), "AUDIT", a["auditor"], a["plan_date"], a["summary"], a["df"],
                            self.notes.get("1.0", "end").strip())
        except Exception:
            tb = traceback.format_exc()
            log_error(tb)
            messagebox.showerror(APP_NAME, "Could not save the audit:\n%s" % last_line(tb))
            return
        self.audit_result = None
        self.btn_save_report.configure(state="disabled")
        self.btn_save_audit.configure(state="disabled")
        fill_rows(self.aud_tree, [])
        for k in (self.k_match, self.k_qty, self.k_rows, self.k_blk):
            k.set("\u2013")
        self.aud_status.set("")
        self.notes.delete("1.0", "end")
        messagebox.showinfo(APP_NAME, "Audit saved. It counts for this week.")
        self.refresh_status()

    def skip_week(self):
        if self.status["state"] not in ("DUE", "OVERDUE"):
            messagebox.showinfo(APP_NAME, "A skip is only possible from Wednesday, in a week with no saved audit.")
            return
        if not pin_is_set():
            messagebox.showinfo(APP_NAME, "No supervisor PIN is set. Use Settings > Set supervisor PIN first.")
            return
        if not self.ask_pin():
            return
        who = simpledialog.askstring(APP_NAME, "Supervisor name:", parent=self.root)
        if not who or not who.strip():
            messagebox.showwarning(APP_NAME, "A name is required.")
            return
        why = simpledialog.askstring(APP_NAME, "Reason for skipping:", parent=self.root)
        if not why or not why.strip():
            messagebox.showwarning(APP_NAME, "A reason is required.")
            return
        store.add_audit(now_local(), "SKIPPED", who.strip(), notes=why.strip())
        self.refresh_status()

    # ================= Audit history page =================
    def _build_history(self, t):
        foot = page_footer(t)
        self.btn_hist_report = ttk.Button(foot, text="Save this report...", style="Tool.TButton",
                                          command=self.save_hist_report, state="disabled")
        self.btn_hist_report.pack(side="right", padx=14, pady=8)

        page_header(t, "Audit history", "Weekly audits saved on this PC. Select a row to open it.")
        ttk.Label(t, text="Last 12 weeks", font=FONT_B).pack(anchor="w", padx=16)
        fr, self.comp_tree = make_tree(t, ["Week", "Wednesday", "Status", "Auditor", "Match %"],
                                       [90, 110, 100, 200, 90], height=5)
        fr.pack(fill="x", padx=16, pady=4)
        for k, col in COMP_COLORS.items():
            self.comp_tree.tag_configure("C_" + k, background=col)
        self.comp_caption = tk.StringVar()
        ttk.Label(t, textvariable=self.comp_caption).pack(anchor="w", padx=16)

        ttk.Label(t, text="All audits", font=FONT_B).pack(anchor="w", padx=16, pady=(8, 0))
        hc = ["#", "Kind", "Week", "Plan date", "Auditor", "Saved", "Row match %", "Qty agr. %", "Rows"]
        fr2, self.hist_tree = make_tree(t, hc, [45, 80, 80, 100, 180, 150, 100, 90, 60], height=5)
        fr2.pack(fill="x", padx=16, pady=4)
        self.hist_tree.bind("<<TreeviewSelect>>", self.open_audit)

        self.hist_note = tk.StringVar()
        ttk.Label(t, textvariable=self.hist_note, wraplength=1100, justify="left").pack(anchor="w", padx=16)
        fr3, self.det_tree = make_tree(t, AUDIT_COLS, AUDIT_W, height=6)
        fr3.pack(fill="both", expand=True, padx=16, pady=(4, 10))
        for st, col in aud.COLORS.items():
            self.det_tree.tag_configure(status_tag(st), background="#" + col)

    def refresh_history(self):
        comp = store.compliance(now_local(), 12)
        fill_rows(self.comp_tree, [[r["Week"], r["Wednesday"], r["Status"], r["Auditor"], r["Match %"]] for r in comp],
                  ["C_" + r["Status"] for r in comp])
        done = sum(1 for r in comp if r["Status"] == "Done")
        self.comp_caption.set("Completed %d of the last %d weeks." % (done, len(comp)))
        hist = store.list_audits()
        self.hist_rows = {r["id"]: r for r in hist}
        fill_rows(self.hist_tree, [[r["id"], r["kind"], "%d-W%02d" % (r["iso_year"], r["iso_week"]),
                                    r["plan_date"] or "-", r["auditor"], r["created_at"], r["match_pct"],
                                    r["qty_agreement"], r["rows_total"]] for r in hist])

    def open_audit(self, _event=None):
        sel = self.hist_tree.selection()
        if not sel:
            return
        audit_id = int(self.hist_tree.item(sel[0], "values")[0])
        det = store.get_detail(audit_id)
        note = self.hist_rows[audit_id]["notes"]
        self.hist_note.set(("Notes: " + note) if note else "")
        self.hist_detail = None
        if det and det["rows"]:
            ddf = pd.DataFrame(det["rows"])
            self.hist_detail = (audit_id, det["summary"], ddf)
            fill_rows(self.det_tree, ddf.reindex(columns=AUDIT_COLS).values.tolist(),
                      [status_tag(s) for s in ddf["Status"]])
            self.btn_hist_report.configure(state="normal")
        else:
            fill_rows(self.det_tree, [])
            self.btn_hist_report.configure(state="disabled")

    def save_hist_report(self):
        if not self.hist_detail:
            return
        audit_id, summary, ddf = self.hist_detail
        p = save_bytes(aud.build_report(summary, ddf), "AUDIT #%d.xlsx" % audit_id, "Save audit report")
        if p and messagebox.askyesno(APP_NAME, "Saved. Open the file now?"):
            os.startfile(p)


# ---------------- self-test (used by build.ps1; also runs against the built exe) ----------------
def selftest(result_path):
    try:
        import pandas as _pd
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "27 Oktober"
        ws.append(["RENCANA KIRIM"])
        ws.append(["TANGGAL 27 - 10 - 2026"])
        ws.append([])
        ws.append(["JENIS MOBIL", "NAMA CUSTOMER", "KODE BRG", "NAMA BARANG", "QTY", "ST", "NOMOR SO", "NOMOR DO"])
        ws.append(["TRUK A", "PT ALPHA", 1001, "PRODUK X", 100, "KG", "1200000001", "8000001"])
        ws.append([None, None, 1002, "PRODUK Y", 50, "KG", "1200000001", "8000001"])
        ws.append(["TRUK B", "PT BETA", 1001, "PRODUK X", 150, "KG", "1200000002", "8000002"])
        ws.append(["TRUK C", "PT ALPHA", 1001, "PRODUK X", 20, "KG", "1200000003", "8000003"])
        ws.append([None, None, 1003, "PRODUK Z", 10, "KG", "1200000003", "8000003"])
        buf = io.BytesIO()
        wb.save(buf)
        plan = buf.getvalue()

        cols = ["Material", "Unrestricted", "SLED/BBD", "Special Stock Number", "Batch", "Storage Location",
                "Supplier Batch", "Date of Manufacture", "SLED Percentage", "Base Unit of Measure"]
        rows = [["1001", "500", "2026-10-20", "", "B0", "1000", "", "2026-01-01", "0", "KG"],
                ["1001", "120", "2027-03-01", "", "B1", "1000", "", "2026-03-01", "0", "KG"],
                ["1001", "200", "2027-06-01", "", "B2", "1000", "", "2026-06-01", "0", "KG"],
                ["1002", "60", "2027-05-01", "", "B3", "1000", "", "2026-05-01", "0", "KG"],
                ["Total", "880", "", "", "", "", "", "", "", ""]]
        buf = io.BytesIO()
        _pd.DataFrame(rows, columns=cols).to_excel(buf, index=False)
        export = buf.getvalue()

        res = eng.generate(plan, "27 Oktober", export)
        assert res["stats"] == dict(blocks=3, lines=5, flagged=1), res["stats"]
        manual = aud.parse_manual(eng.write_workbook(res["blocks"], res["plan_date"]))
        s, df = aud.compare(eng.blocks_list(res["blocks"]), manual, res["snapshot"], res["name2kode"])
        assert s["match_pct"] == 100.0 and s["blocks_matched"] == 3, s
        assert len(aud.build_report(s, df, {"Auditor": "selftest"})) > 1000

        root = tk.Tk()                    # builds the whole window to prove the UI loads
        if find_logo():
            from PIL import Image, ImageTk  # noqa: F401  (the logo needs Pillow inside the exe)
            assert load_logo_image(32) is not None, "logo file found but could not be loaded"
        App(root)
        root.update()
        root.destroy()
        msg, code = "OK", 0
    except Exception:
        msg, code = traceback.format_exc(), 1
    try:
        with open(result_path, "w", encoding="utf-8") as f:
            f.write(msg)
    except OSError:
        pass
    return code


def main():
    if SELFTEST:
        i = sys.argv.index("--selftest")
        path = sys.argv[i + 1] if len(sys.argv) > i + 1 else os.path.join(tempfile.gettempdir(), "coa_selftest.txt")
        sys.exit(selftest(path))
    root = tk.Tk()

    def on_exception(exc, val, tb):
        log_error("".join(traceback.format_exception(exc, val, tb)))
        messagebox.showerror(APP_NAME, "Unexpected error:\n%s\n\nDetails are in:\n%s" % (val, ERROR_LOG))

    root.report_callback_exception = on_exception
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
