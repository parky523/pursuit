"""Pursuit: a desktop + phone app for a job search - daily checklist, coach, and live application pipeline.

Today tab
- Click the box to tick or untick an item; type in the bottom field + Enter to add; hover + x to delete.
- "From your tracker" items (follow-ups due, drafts to send) are generated from the Daily Log.
- Ticks reset every new day; your own items stay until you delete them.

Coach tab
- Job-hunt day count, time since last application, this week vs last week (target 5), routine streak,
  a motivation line and suggested next moves (click "+ Today" to put one on the checklist).

Applications tab
- Reads the tracker workbook > Daily Log (re-reads automatically whenever the file is saved).
- Grouped: Interview/Offer, Follow-up due, Applied (waiting), Ready to send (Drafted), Researching, Closed.
- Left-click a row opens its Applications\\YYYY-MM\\... folder. Right-click sets a new status, written back to the
  Daily Log via Excel COM (never openpyxl writes - that would drop the dashboard charts).

Settings live in config.json (see config.example.json); state lives in checklist.json, both next to this script.
"""
import ctypes
import datetime as dt
import json
import os
import queue
import re
import shutil
import socket
import sys
import tempfile
import threading
import tkinter as tk
from pathlib import Path

# Pursuit.exe (PyInstaller build) sits next to its data; the .py runs from the same folder
BASE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
STATE = BASE / "checklist.json"
CONFIG_FILE = BASE / "config.json"

CONFIG_DEFAULTS = {
    "tracker": "Job_Search_Tracker.xlsx",        # path to the workbook (relative = next to this file)
    "applications_dir": "Applications",          # folder with YYYY-MM\YYYY-MM-DD_Company_Role subfolders
    "weekly_target": 5,
    "follow_up_days": 14,
    "deadline": None,                            # optional {"label", "date": "YYYY-MM-DD", "note"} Coach tile
    "sheets": {"log": "Daily Log", "targets": "Target Companies", "connections": "Connections",
               "content": "Content Calendar"},
    "daily_items": [
        "Pick 1 target company from the Target Companies sheet",
        "Research the company + find 1-2 LinkedIn contacts",
        "Tailor CV + cover letter",
        "Submit the 1 application",
        "Log it in the Daily Log (same day)",
        "Send 1-2 LinkedIn connection requests",
        "LinkedIn 15 min: 3-5 comments",
    ],
}


def load_config():
    cfg = json.loads(json.dumps(CONFIG_DEFAULTS))
    try:
        user = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        cfg.update({k: v for k, v in user.items() if not k.startswith("_")})
        cfg["sheets"] = {**CONFIG_DEFAULTS["sheets"], **user.get("sheets", {})}
    except (OSError, ValueError):
        pass
    return cfg


def cfg_path(p):
    p = Path(os.path.expandvars(str(p)))
    return p if p.is_absolute() else BASE / p


CFG = load_config()
TRACKER = cfg_path(CFG["tracker"])
APPS_DIR = cfg_path(CFG["applications_dir"])
SHEETS = CFG["sheets"]
DEFAULT_ITEMS = list(CFG["daily_items"])
STATUSES = ["Researching", "Drafted", "Applied", "Interview", "Offer", "Rejected", "No reply", "Skipped - Not a Fit"]
FOLLOW_UP_DAYS = int(CFG["follow_up_days"])
WEEKLY_TARGET = int(CFG["weekly_target"])
DEADLINE = CFG.get("deadline") or None
if DEADLINE:
    DEADLINE = {**DEADLINE, "date": dt.date.fromisoformat(DEADLINE["date"])}
SENT = ("Applied", "Interview", "Offer", "Rejected", "No reply")

# ---- Pursuit design system: ink black, warm paper text, acid lime signal ----
BG, CARD, LINE = "#0b0b0c", "#151517", "#2a2a2e"
FG, MUTED, DIM = "#f2efe6", "#8c8a83", "#55534e"
LIME, ORANGE, RED, SKY = "#c8ff3d", "#ff9f1c", "#ff5a4e", "#7cc7ff"
ACCENT = DONE = GOLD = LIME
WARN, BAD = ORANGE, RED

FONT_FILES = BASE / "fonts"
DISPLAY = "Unbounded"            # wide geometric display face, for numbers + wordmark
DISPLAY_M = "Unbounded Medium"
MONO = "JetBrains Mono Medium"   # labels, tags, meta
MONO_B = "JetBrains Mono"
SANS = "Instrument Sans"         # body text
SANS_B = "Instrument Sans SemiBold"
FONT, FONT_SB = (SANS, 10), (SANS_B, 10)
FONT_B, FONT_S = (SANS_B, 11), (MONO, 8)
WIDTH = 410
STATUS_COLOR = {"Interview": LIME, "Offer": LIME, "Applied": SKY, "Drafted": ORANGE, "Researching": MUTED,
                "Rejected": RED, "No reply": DIM}


def load_fonts():
    """Register the bundled TTFs for this process only (FR_PRIVATE) - nothing gets installed system-wide."""
    for f in FONT_FILES.glob("*.ttf"):
        try:
            ctypes.windll.gdi32.AddFontResourceExW(str(f), 0x10, 0)
        except Exception:
            pass


def today():
    return dt.date.today().isoformat()


def load_state():
    try:
        s = json.loads(STATE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        s = {"items": [{"text": t, "done": False} for t in DEFAULT_ITEMS], "date": today(), "history": {}}
    s.setdefault("history", {})
    s.setdefault("auto_done", [])
    s.setdefault("tab", "today")
    return s


def save_state(s):
    tmp = STATE.with_suffix(".tmp")
    tmp.write_text(json.dumps(s, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(STATE)


# ---------- tracker ----------
def as_date(v):
    """Cell value -> date. Rounds to nearest day (COM-written dates can land at 22:00 the day before)."""
    if isinstance(v, dt.datetime):
        return (v + dt.timedelta(hours=12)).date()
    if isinstance(v, dt.date):
        return v
    return None


def read_tracker():
    """Return Daily Log rows as dicts. Reads a temp copy so an open Excel never blocks it."""
    import openpyxl
    tmp = Path(tempfile.gettempdir()) / "job_checklist_tracker_copy.xlsx"
    shutil.copyfile(TRACKER, tmp)
    wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
    apps = []
    for n, r in enumerate(wb[SHEETS["log"]].iter_rows(min_row=2, max_col=13, values_only=True), start=2):
        if not r[1]:
            continue
        apps.append({"row": n, "date": as_date(r[0]), "company": str(r[1]).strip(), "role": str(r[2] or "").strip(),
                     "location": r[3] or "", "status": str(r[6] or "Researching").strip(),
                     "follow": as_date(r[11]), "notes": r[12] or ""})
    # optional sheets: skipped quietly when the workbook doesn't have them
    rows = lambda key, **kw: wb[SHEETS[key]].iter_rows(values_only=True, **kw) if SHEETS[key] in wb.sheetnames else []
    targets = []
    for r in rows("targets", min_row=3, max_col=8):  # row 1 = note, row 2 = headers
        if r[1]:
            targets.append({"prio": r[0], "company": r[1], "status": r[7] or "Not Started"})
    extra = {"people": [], "posts": [], "has_connections": SHEETS["connections"] in wb.sheetnames}
    for r in rows("connections", min_row=2, max_col=16):
        if isinstance(r[0], (int, float)) and r[1]:
            extra["people"].append({"name": r[1], "target": r[5] or r[4], "sent": as_date(r[9]),
                                    "accepted": as_date(r[10]), "status": r[15] or ""})
    for r in rows("content", min_row=2, max_col=10):
        if as_date(r[1]) and r[5]:
            extra["posts"].append({"date": as_date(r[1]), "topic": r[5], "status": r[9] or ""})
    wb.close()
    return apps, targets, extra


REMOVED_LOG = BASE / "removed_rows.json"  # backup of every Daily Log row deleted from the app


def serial(d):
    """Excel date serial. Writing Python datetimes via COM shifts them a day, serials don't."""
    return (d - dt.date(1899, 12, 30)).days


def put_date(cell, d):
    cell.Value = serial(d)
    if cell.NumberFormat == "General":
        cell.NumberFormat = "yyyy-mm-dd"


def add_note(ws, row, text):
    old = ws.Cells(row, 13).Value
    ws.Cells(row, 13).Value = f"{old} | {text}" if old else text


def edit_tracker(row, company, action, status=None):
    """Apply one edit to Daily Log row `row` via Excel COM (attaches to an open Excel if the tracker is open).

    action: "status"  -> set col G; Drafted/Researching -> Applied also moves Date to today + follow-up +14d
            "followed" -> note the follow-up, push the follow-up date 7 days out
            "remove"   -> delete the row (a copy is kept in removed_rows.json)
    Never openpyxl-saves the workbook: that would drop the dashboard charts.
    """
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    xl = wb = None
    opened = False
    t = dt.date.today()
    try:
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
            for b in xl.Workbooks:
                if os.path.normcase(b.FullName) == os.path.normcase(str(TRACKER)):
                    wb = b
        except Exception:
            xl = None
        if wb is None:
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(str(TRACKER))
            opened = True
        ws = wb.Worksheets(SHEETS["log"])
        if str(ws.Cells(row, 2).Value).strip() != company:
            raise RuntimeError(f"row {row} is no longer {company}, press SYNC and retry")

        if action == "status":
            old = str(ws.Cells(row, 7).Value or "").strip()
            ws.Cells(row, 7).Value = status
            if status == "Applied":
                if old in ("Drafted", "Researching", ""):
                    logged = ws.Cells(row, 1).Text
                    put_date(ws.Cells(row, 1), t)
                    add_note(ws, row, f"{old or 'Logged'} {logged}; applied {t.isoformat()} (Pursuit)")
                    put_date(ws.Cells(row, 12), t + dt.timedelta(days=FOLLOW_UP_DAYS))
                elif not ws.Cells(row, 12).Value:
                    put_date(ws.Cells(row, 12), t + dt.timedelta(days=FOLLOW_UP_DAYS))
            elif old and old != status:
                add_note(ws, row, f"{status} {t.isoformat()} (Pursuit)")
        elif action == "followed":
            add_note(ws, row, f"Followed up {t.isoformat()} (Pursuit)")
            put_date(ws.Cells(row, 12), t + dt.timedelta(days=7))
        elif action == "remove":
            values = [ws.Cells(row, c).Text for c in range(1, 14)]
            try:
                log = json.loads(REMOVED_LOG.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                log = []
            log.append({"removed": dt.datetime.now().isoformat(timespec="seconds"), "row": row, "values": values})
            REMOVED_LOG.write_text(json.dumps(log, indent=2, ensure_ascii=False), encoding="utf-8")
            ws.Rows(row).Delete()
        wb.Save()
    finally:
        if opened:
            wb.Close(SaveChanges=False)
            xl.Quit()
        pythoncom.CoUninitialize()


def find_folder(app):
    """Best-matching Applications\\YYYY-MM\\YYYY-MM-DD_Company_* folder for a Daily Log row."""
    key = re.sub(r"[^a-z0-9]", "", re.split(r"[\s(/&,]+", app["company"].lower())[0])
    if not key or not APPS_DIR.exists():
        return None
    best = None
    for month in APPS_DIR.glob("20??-??"):
        for f in month.iterdir():
            if f.is_dir() and key in re.sub(r"[^a-z0-9]", "", f.name.lower()):
                try:
                    d = dt.date.fromisoformat(f.name[:10])
                except ValueError:
                    continue
                gap = abs((d - app["date"]).days) if app["date"] else 999
                if best is None or gap < best[0]:
                    best = (gap, f)
    return best[1] if best and best[0] <= 7 else None


def group_apps(apps):
    t = dt.date.today()
    groups = {k: [] for k in ("hot", "due", "applied", "drafted", "research", "closed")}
    for a in apps:
        s = a["status"]
        if s in ("Interview", "Offer"):
            groups["hot"].append(a)
        elif s == "Applied":
            follow = a["follow"] or (a["date"] + dt.timedelta(days=FOLLOW_UP_DAYS) if a["date"] else None)
            a["due_on"] = follow
            groups["due" if follow and follow <= t else "applied"].append(a)
        elif s == "Drafted":
            groups["drafted"].append(a)
        elif s == "Researching":
            groups["research"].append(a)
        else:
            groups["closed"].append(a)
    newest = lambda a: a["date"] or dt.date.min
    for k in groups:
        groups[k].sort(key=newest, reverse=k != "drafted")  # drafts: oldest first = send these first
    return groups


GROUPS = [("hot", "INTERVIEW / OFFER", LIME), ("due", "FOLLOW-UP DUE", ORANGE), ("applied", "APPLIED · WAITING", SKY),
          ("drafted", "READY TO SEND", ORANGE), ("research", "RESEARCHING", MUTED), ("closed", "CLOSED", DIM)]


def card_meta(key, a):
    """(meta line, colour) shown under an application card - shared by desktop and phone."""
    if key == "hot":
        return f"Logged {ago(a['date'])}", LIME
    if key == "due":
        return (f"Applied {ago(a['date'])} · follow-up overdue {(dt.date.today() - a['due_on']).days}d",
                ORANGE)
    if key == "applied":
        fu = a["due_on"].strftime("%d %b") if a.get("due_on") else "-"
        return f"Applied {ago(a['date'])} · follow up {fu}", MUTED
    if key == "drafted":
        return f"Drafted {ago(a['date'])} · not sent yet", MUTED
    if key == "research":
        return f"Added {ago(a['date'])}", MUTED
    return ago(a["date"]), MUTED


def card_actions(a):
    """[(label, kind, new_status, filled)] push-buttons for a card - shared by desktop and phone."""
    acts = {"Drafted": [("✓ MARK APPLIED", "Applied")],
            "Researching": [("DRAFTED", "Drafted"), ("✓ APPLIED", "Applied")],
            "Applied": [("INTERVIEW", "Interview"), ("REJECTED", "Rejected")],
            "Interview": [("OFFER", "Offer"), ("REJECTED", "Rejected")]}.get(a["status"], [])
    out = [(text, "status", new, i == 0 and new != "Rejected") for i, (text, new) in enumerate(acts)]
    if a["status"] == "Applied":
        out.append(("FOLLOWED UP", "followed", None, False))
    return out


def ago(d):
    if not d:
        return ""
    n = (dt.date.today() - d).days
    return "today" if n == 0 else "yesterday" if n == 1 else f"{n}d ago" if n > 0 else f"in {-n}d"


class Checklist:
    def __init__(self):
        self.s = load_state()
        self.apps, self.targets, self.extra, self.groups = [], [], {"people": [], "posts": []}, group_apps([])
        self.mtime, self.err, self.show_closed = None, "", False

        load_fonts()
        try:  # must happen before the first window exists, or the taskbar groups us under Python
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Pursuit.JobHunt")
        except Exception:
            pass
        self.root = tk.Tk()
        self.root.title("Pursuit")
        try:
            self.root.iconbitmap(default=str(BASE / "pursuit.ico"))
            # big + small icon for taskbar / Alt-Tab (iconbitmap alone only sets the small one)
            self.icons = [tk.PhotoImage(file=str(BASE / f)) for f in ("pursuit.png", "pursuit_28.png")]
            self.root.iconphoto(True, *self.icons)
        except Exception:
            pass
        self.root.option_add("*Menu.background", CARD)
        self.root.option_add("*Menu.foreground", FG)
        self.root.option_add("*Menu.activeBackground", LIME)
        self.root.option_add("*Menu.activeForeground", BG)
        self.root.option_add("*Menu.disabledForeground", MUTED)
        self.root.option_add("*Menu.font", f"{{{SANS}}} 10")
        self.root.option_add("*Menu.relief", "flat")
        self.root.configure(bg=BG)
        self.root.minsize(WIDTH, 300)
        x, y = self.s.get("pos", (self.root.winfo_screenwidth() - WIDTH - 40, 40))
        h = self.s.get("height", 720)
        self.root.geometry(f"{WIDTH}x{h}+{x}+{y}")
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        try:  # dark title bar on Windows 10/11
            self.root.update()
            ctypes.windll.dwmapi.DwmSetWindowAttribute(self.hwnd(), 20, ctypes.byref(ctypes.c_int(1)), 4)
        except Exception:
            pass

        # header: logo + wordmark + live counter
        head = tk.Frame(self.root, bg=BG)
        head.pack(fill="x", padx=16, pady=(14, 10))
        try:
            self.logo = tk.PhotoImage(file=str(BASE / "pursuit_28.png"))
            tk.Label(head, image=self.logo, bg=BG).pack(side="left", padx=(0, 10))
        except tk.TclError:
            pass
        tk.Label(head, text="PURSUIT", bg=BG, fg=FG, font=(DISPLAY, 14, "bold")).pack(side="left")
        phone = tk.Label(head, text="PHONE", bg=BG, fg=LIME, font=(MONO_B, 7, "bold"), cursor="hand2", padx=6,
                         highlightthickness=1, highlightbackground=LINE)
        phone.pack(side="right", padx=(8, 0))
        phone.bind("<Button-1>", lambda e: self.phone_dialog())
        self.count = tk.Label(head, bg=BG, fg=MUTED, font=(MONO, 8))
        self.count.pack(side="right")
        for w in (head, self.count):
            w.bind("<Button-3>", self.menu)

        # segmented tabs
        tabs = tk.Frame(self.root, bg=LINE)
        tabs.pack(fill="x", padx=16)
        self.tab_btn = {}
        for i, (key, label) in enumerate((("today", "TODAY"), ("coach", "COACH"), ("apps", "PIPELINE"))):
            b = tk.Label(tabs, text=f"0{i + 1}  {label}", font=(MONO_B, 8, "bold"), cursor="hand2", pady=7)
            b.pack(side="left", fill="x", expand=True, padx=1, pady=1)
            b.bind("<Button-1>", lambda e, k=key: self.switch(k))
            self.tab_btn[key] = b

        self.bar = tk.Canvas(self.root, height=3, bg=LINE, highlightthickness=0, width=WIDTH - 32)
        self.bar.pack(fill="x", padx=16, pady=(10, 4))

        # footer (packed before the scroll area so it stays visible)
        self.footer = tk.Frame(self.root, bg=BG)
        self.footer.pack(side="bottom", fill="x", padx=16, pady=(6, 14))
        self.add = tk.Frame(self.footer, bg=LINE)
        inner = tk.Frame(self.add, bg=CARD)
        inner.pack(fill="x", padx=1, pady=1)
        self.entry = tk.Entry(inner, bg=CARD, fg=FG, insertbackground=LIME, relief="flat", font=FONT,
                              insertwidth=2)
        self.entry.pack(side="left", fill="x", expand=True, padx=(12, 0), ipady=8)
        self.entry.bind("<Return>", self.add_item)
        self.placeholder(True)
        self.entry.bind("<FocusIn>", lambda e: self.placeholder(False))
        self.entry.bind("<FocusOut>", lambda e: self.placeholder(True))
        plus = tk.Label(inner, text="ADD", bg=LIME, fg=BG, font=(MONO_B, 8, "bold"), cursor="hand2", padx=12)
        plus.pack(side="right", fill="y")
        plus.bind("<Button-1>", self.add_item)
        self.info = tk.Frame(self.footer, bg=BG)
        self.status_lbl = tk.Label(self.info, bg=BG, fg=MUTED, font=(MONO, 8), anchor="w")
        self.status_lbl.pack(side="left")
        for text, cmd in (("OPEN XLSX ↗", lambda: os.startfile(TRACKER)), ("SYNC ↻", self.force_reload)):
            b = tk.Label(self.info, text=text, bg=BG, fg=LIME, font=(MONO_B, 8, "bold"), cursor="hand2", padx=6)
            b.pack(side="right")
            b.bind("<Button-1>", lambda e, c=cmd: c())

        # scrollable body
        wrap = tk.Frame(self.root, bg=BG)
        wrap.pack(fill="both", expand=True, padx=(16, 8))
        self.canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=BG)
        self.body_id = self.canvas.create_window(0, 0, window=self.body, anchor="nw")
        self.body.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.body_id, width=e.width - 6))
        self.root.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(-e.delta // 120, "units"))

        self.rollover()
        self.watch_tracker()
        self.render()
        self.tick()

        # phone app: HTTP server on a background thread; its calls run here on the Tk thread
        self.calls = queue.Queue()
        self.pump()
        from phone_server import PhoneServer
        self.phone = PhoneServer(self, BASE)
        save_state(self.s)  # persists the phone key on first run
        if not self.phone.start():
            self.err = self.phone.error

    def pump(self):
        while True:
            try:
                fn, ev, box = self.calls.get_nowait()
            except queue.Empty:
                break
            try:
                box["result"] = fn()
            except Exception as ex:
                box["error"] = ex
            ev.set()
        self.root.after(50, self.pump)

    def tick(self):
        """Keep the Coach countdowns live."""
        if self.s["tab"] == "coach":
            self.render()
        self.root.after(60_000, self.tick)

    # ---- data ----
    def rollover(self):
        """New day: record yesterday's score, clear ticks."""
        if self.s.get("date") != today():
            items = self.s["items"]
            if items and self.s.get("date"):
                self.s["history"][self.s["date"]] = [sum(i["done"] for i in items), len(items)]
            for i in items:
                i["done"] = False
            self.s["auto_done"] = []
            self.s["date"] = today()
            save_state(self.s)
            self.render()
        self.root.after(60_000, self.rollover)

    def watch_tracker(self):
        """Re-read the tracker whenever Excel saves it."""
        self.load_tracker()
        self.root.after(5000, self.watch_tracker)

    def load_tracker(self):
        try:
            m = TRACKER.stat().st_mtime
            if m != self.mtime:
                self.apps, self.targets, self.extra = read_tracker()
                self.groups = group_apps(self.apps)
                self.mtime, self.err = m, ""
                self.loaded_at = dt.datetime.now()
                self.render()
        except Exception as ex:
            self.err = f"Tracker read failed: {ex}"[:80]
            self.render()

    def force_reload(self):
        self.mtime = None
        self.load_tracker()

    def auto_items(self):
        """Today's pipeline tasks generated from the tracker."""
        out = []
        for a in self.groups["hot"]:
            out.append(f"{a['status']}: prepare for {a['company']}")
        for a in self.groups["due"]:
            out.append(f"Follow up: {a['company']} (applied {ago(a['date'])})")
        for a in self.groups["drafted"][:2]:
            out.append(f"Send drafted application: {a['company']}")
        if not self.groups["drafted"]:
            nxt = next((t for t in self.targets if t["status"] == "Not Started"), None)
            if nxt:
                out.append(f"Next target ({nxt['prio']}): {nxt['company']}")
        return out

    def toggle(self, idx):
        self.s["items"][idx]["done"] = not self.s["items"][idx]["done"]
        save_state(self.s)
        self.render()

    def toggle_auto(self, text):
        done = self.s["auto_done"]
        done.remove(text) if text in done else done.append(text)
        save_state(self.s)
        self.render()

    def delete(self, idx):
        del self.s["items"][idx]
        save_state(self.s)
        self.render()

    def add_item(self, _=None):
        text = self.entry.get().strip()
        if not text or self.entry.cget("fg") == MUTED:
            self.entry.focus_force()
            return
        self.s["items"].append({"text": text, "done": False})
        save_state(self.s)
        self.entry.delete(0, "end")
        self.render()

    def reset_ticks(self):
        for i in self.s["items"]:
            i["done"] = False
        self.s["auto_done"] = []
        save_state(self.s)
        self.render()

    def set_status(self, app, status):
        self.tracker_action(app, "status", status, f"{app['company'][:18]} → {status}")

    def tracker_action(self, app, action, status=None, label="", confirm=True):
        """Run one Excel edit in the background, then re-read the tracker."""
        if getattr(self, "busy", False):
            return
        if action == "remove" and confirm:  # the phone asks on its own screen instead
            from tkinter import messagebox
            if not messagebox.askyesno("Pursuit", f"Remove this row from the Daily Log?\n\n{app['company']}\n"
                                                  f"{app['role']}\n\nA copy is kept in removed_rows.json.",
                                       icon="warning", parent=self.root):
                return
        self.busy = True
        self.err = f"Saving {label or action}…"
        self.render()

        def work():
            try:
                edit_tracker(app["row"], app["company"], action, status)
                msg = f"Saved: {label or action}"
            except Exception as ex:
                msg = f"Save failed: {ex}"[:90]

            def done():
                self.busy = False
                self.force_reload()  # clears err on a clean read
                self.err = msg
                self.render()
                self.root.after(6000, lambda: self.err == msg and (setattr(self, "err", ""), self.render()))
            self.root.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    # ---- view ----
    def placeholder(self, show):
        if show and not self.entry.get():
            self.entry.config(fg=MUTED)
            self.entry.insert(0, "Add a task and press Enter")
        elif not show and self.entry.cget("fg") == MUTED:
            self.entry.delete(0, "end")
            self.entry.config(fg=FG)

    def switch(self, tab):
        self.s["tab"] = tab
        save_state(self.s)
        self.canvas.yview_moveto(0)
        self.render()

    def section(self, text, color=MUTED, click=None, count=None):
        """Mono label, hairline rule, optional count: `FOLLOW-UP DUE ─────── 03`."""
        row = tk.Frame(self.body, bg=BG, cursor="hand2" if click else "")
        row.pack(fill="x", pady=(16, 6))
        tk.Label(row, text=text, bg=BG, fg=color, font=(MONO_B, 8, "bold")).pack(side="left")
        if count is not None:
            tk.Label(row, text=f"{count:02d}", bg=BG, fg=color, font=(MONO_B, 8, "bold")).pack(side="right")
        tk.Frame(row, bg=LINE, height=1).pack(side="left", fill="x", expand=True, padx=8)
        if click:
            for w in [row] + row.winfo_children():
                w.bind("<Button-1>", lambda e: click())

    def checkbox(self, parent, done):
        c = tk.Canvas(parent, width=18, height=18, bg=BG, highlightthickness=0, cursor="hand2")
        if done:
            c.create_rectangle(1, 1, 17, 17, fill=LIME, outline=LIME)
            c.create_line(4.5, 9.5, 7.5, 12.5, 13.5, 5.5, fill=BG, width=2.4, capstyle="round", joinstyle="round")
        else:
            c.create_rectangle(1.5, 1.5, 16.5, 16.5, outline=DIM, width=1.5)
        return c

    def check_row(self, text, done, on_toggle, on_delete=None):
        row = tk.Frame(self.body, bg=BG)
        row.pack(fill="x", pady=3)
        box = self.checkbox(row, done)
        box.pack(side="left", anchor="n", pady=(1, 0), padx=(0, 10))
        txt = tk.Label(row, text=text, bg=BG, fg=DIM if done else FG, anchor="w", justify="left",
                       font=FONT + (("overstrike",) if done else ()), wraplength=WIDTH - 100, cursor="hand2")
        txt.pack(side="left", fill="x", expand=True)
        for w in (box, txt):
            w.bind("<Button-1>", lambda e: on_toggle())
        if on_delete:
            x = tk.Label(row, text="DEL", bg=BG, fg=BG, font=(MONO_B, 7, "bold"), cursor="hand2", padx=4)
            x.pack(side="right", anchor="n")
            x.bind("<Button-1>", lambda e: on_delete())
            for w in (row, box, txt, x):
                w.bind("<Enter>", lambda e, b=x: b.config(fg=RED), add="+")
                w.bind("<Leave>", lambda e, b=x: b.config(fg=BG), add="+")

    def app_line(self, a, meta, meta_color=MUTED):
        """One application card: status stripe, company, filled status tag, role, mono meta line."""
        color = STATUS_COLOR.get(a["status"], DIM)
        outer = tk.Frame(self.body, bg=LINE, cursor="hand2")
        outer.pack(fill="x", pady=3)
        row = tk.Frame(outer, bg=CARD)
        row.pack(fill="x", padx=1, pady=1)
        tk.Frame(row, bg=color, width=3).pack(side="left", fill="y")
        mid = tk.Frame(row, bg=CARD)
        mid.pack(side="left", fill="x", expand=True, padx=12, pady=8)
        top = tk.Frame(mid, bg=CARD)
        top.pack(fill="x")
        company = a["company"] if len(a["company"]) < 32 else a["company"][:30] + "…"
        tk.Label(top, text=company, bg=CARD, fg=FG, font=(SANS_B, 11), anchor="w").pack(side="left")
        tag = a["status"].replace(" - Not a Fit", "").upper()
        tk.Label(top, text=f" {tag} ", bg=color, fg=BG, font=(MONO_B, 7, "bold")).pack(side="right")
        role = a["role"] if len(a["role"]) < 54 else a["role"][:52] + "…"
        tk.Label(mid, text=role, bg=CARD, fg=MUTED, font=(SANS, 9), anchor="w").pack(fill="x", pady=(1, 0))
        tk.Label(mid, text=meta.upper(), bg=CARD, fg=meta_color, font=(MONO, 7), anchor="w").pack(fill="x",
                                                                                                pady=(5, 0))
        hover = lambda c: (lambda e: outer.config(bg=c))
        for w in [outer, row, mid, top] + mid.winfo_children() + top.winfo_children():
            w.bind("<Button-1>", lambda e: self.open_folder(a))
            w.bind("<Button-3>", lambda e: self.app_menu(e, a))
            w.bind("<Enter>", hover(color))
            w.bind("<Leave>", hover(LINE))

        # action buttons - every click writes straight to the Daily Log in Excel
        acts = tk.Frame(mid, bg=CARD)
        acts.pack(fill="x", pady=(8, 0))
        for text, kind, new, filled in card_actions(a):
            if kind == "status":
                self.button(acts, text, lambda new=new: self.set_status(a, new), filled=filled)
            else:
                self.button(acts, text, lambda: self.tracker_action(a, "followed",
                                                                    label=f"followed up {a['company'][:18]}"))
        self.button(acts, "REMOVE", lambda: self.tracker_action(a, "remove", label=f"removed {a['company'][:18]}"),
                    color=RED, side="right")
        self.button(acts, "···", lambda: self.app_menu(None, a, acts), color=MUTED, side="right")

    def button(self, parent, text, cmd, filled=False, color=LIME, side="left"):
        """Small mono push-button: filled lime for the main action, outlined for the rest."""
        frame = tk.Frame(parent, bg=color if filled else LINE)
        frame.pack(side=side, padx=(0, 6) if side == "left" else (6, 0))
        b = tk.Label(frame, text=text, bg=color if filled else CARD, fg=BG if filled else color,
                     font=(MONO_B, 7, "bold"), padx=8, pady=3, cursor="hand2")
        b.pack(padx=1, pady=1)
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: b.config(bg=color, fg=BG))
        b.bind("<Leave>", lambda e: b.config(bg=color if filled else CARD, fg=BG if filled else color))

    def render(self):
        for w in self.body.winfo_children():
            w.destroy()
        tab = self.s["tab"]
        for k, b in self.tab_btn.items():
            b.config(fg=BG if k == tab else MUTED, bg=LIME if k == tab else BG)
        self.bar.delete("all")
        if tab == "today":
            self.info.pack_forget()
            self.add.pack(fill="x")
            self.render_today()
        else:
            self.add.pack_forget()
            self.info.pack(fill="x")
            self.render_apps() if tab == "apps" else self.render_coach()
            when = self.loaded_at.strftime("%H:%M") if getattr(self, "loaded_at", None) else "-"
            self.status_lbl.config(text=(self.err or f"● SYNCED WITH TRACKER {when}").upper()[:52],
                                   fg=RED if "fail" in self.err else MUTED)

    def render_today(self):
        items = self.s["items"]
        auto = self.auto_items()
        done = sum(i["done"] for i in items) + sum(t in self.s["auto_done"] for t in auto)
        total = len(items) + len(auto)

        hero = tk.Frame(self.body, bg=BG)
        hero.pack(fill="x", pady=(8, 0))
        tk.Label(hero, text=dt.date.today().strftime("%A").upper(), bg=BG, fg=MUTED,
                 font=(MONO_B, 8, "bold")).pack(anchor="w")
        line = tk.Frame(hero, bg=BG)
        line.pack(fill="x")
        tk.Label(line, text=dt.date.today().strftime("%d %b"), bg=BG, fg=FG, font=(DISPLAY, 22, "bold")).pack(
            side="left")
        pct = round(100 * done / total) if total else 0
        tk.Label(line, text=f"{pct}%", bg=BG, fg=LIME if pct == 100 else FG, font=(DISPLAY_M, 22)).pack(
            side="right")

        if auto:
            self.section("FROM YOUR PIPELINE", LIME, count=len(auto))
            for t in auto:
                self.check_row(t, t in self.s["auto_done"], lambda t=t: self.toggle_auto(t))
        self.section("DAILY ROUTINE", MUTED, count=len(items))
        for idx, it in enumerate(items):
            self.check_row(it["text"], it["done"], lambda i=idx: self.toggle(i), lambda i=idx: self.delete(i))
        self.count.config(text=f"{done:02d}/{total:02d} DONE" + ("  ✦ CLEARED" if total and done == total else ""))
        self.draw_bar(done / total if total else 0)

    def draw_bar(self, frac, color=LIME):
        self.root.update_idletasks()
        w = self.bar.winfo_width()
        if frac > 0:
            self.bar.create_rectangle(0, 0, max(w * min(frac, 1), 3), 3, fill=color, width=0)

    def render_apps(self):
        g = self.groups
        n = lambda *st: sum(a["status"] in st for a in self.apps)
        sent = n("Applied", "Interview", "Offer", "Rejected", "No reply")
        answered = n("Interview", "Offer", "Rejected")
        self.count.config(text=f"{sent:02d} SENT · {answered:02d} ANSWERED")
        self.draw_bar(answered / sent if sent else 0)  # response rate

        stats = tk.Frame(self.body, bg=LINE)
        stats.pack(fill="x", pady=(8, 0))
        for label, val, col in (("APPLIED", n("Applied"), SKY), ("INTERVW", n("Interview"), LIME),
                                ("OFFER", n("Offer"), LIME), ("TO SEND", len(g["drafted"]), ORANGE),
                                ("REJECT", n("Rejected"), RED)):
            c = tk.Frame(stats, bg=CARD)
            c.pack(side="left", fill="both", expand=True, padx=(1, 0), pady=1)
            tk.Label(c, text=str(val), bg=CARD, fg=col if val else DIM, font=(DISPLAY, 17, "bold")).pack(
                anchor="w", padx=10, pady=(8, 0))
            tk.Label(c, text=label, bg=CARD, fg=MUTED, font=(MONO, 7)).pack(anchor="w", padx=10, pady=(0, 8))

        for key, title, color in GROUPS[:-1]:
            if g[key]:
                self.section(title, color, count=len(g[key]))
                for a in g[key]:
                    self.app_line(a, *card_meta(key, a))
        not_started = [t for t in self.targets if t["status"] == "Not Started"]
        if self.targets:
            self.section(f"TARGET BOOK · {len(not_started)}/{len(self.targets)} UNTOUCHED", MUTED)
            for t in not_started[:3]:
                r = tk.Frame(self.body, bg=BG)
                r.pack(fill="x", pady=2)
                tk.Label(r, text=t["prio"] or "--", bg=BG, fg=LIME, font=(MONO_B, 8, "bold")).pack(side="left")
                tk.Label(r, text=t["company"], bg=BG, fg=FG, font=(SANS, 10)).pack(side="left", padx=10)
        if g["closed"]:
            arrow = "−" if self.show_closed else "+"
            self.section(f"{arrow} CLOSED", DIM, self.toggle_closed, count=len(g["closed"]))
            if self.show_closed:
                for a in g["closed"]:
                    self.app_line(a, *card_meta("closed", a))


    # ---- coach ----
    def week_counts(self):
        t = dt.date.today()
        monday = t - dt.timedelta(days=t.weekday())
        sent = [a["date"] for a in self.apps if a["status"] in SENT and a["date"]]
        this = sum(d >= monday for d in sent)
        last = sum(monday - dt.timedelta(days=7) <= d < monday for d in sent)
        return this, last, monday

    def streak(self):
        """Consecutive days (ending today, or yesterday if today isn't there yet) with 60%+ of the routine ticked."""
        hist = dict(self.s["history"])
        items = self.s["items"]
        if items:
            hist[today()] = [sum(i["done"] for i in items), len(items)]
        ok = lambda d: (lambda v: v[1] > 0 and v[0] / v[1] >= 0.6)(hist.get(d.isoformat(), [0, 0]))
        d = dt.date.today()
        if not ok(d):
            d -= dt.timedelta(days=1)
        n = 0
        while ok(d):
            n += 1
            d -= dt.timedelta(days=1)
        return n

    def suggestions(self):
        t = dt.date.today()
        g, out = self.groups, []
        this, _, _ = self.week_counts()
        days_left = 7 - t.weekday()
        need = max(WEEKLY_TARGET - this, 0)
        for a in g["hot"]:
            out.append(f"Prepare for {a['company']}: research the team, 3 STAR stories, questions to ask")
        for a in g["due"][:2]:
            out.append(f"Send a short follow-up to {a['company']} (applied {ago(a['date'])})")
        if need:
            pace = "1 today" if need <= days_left else f"{need} in {days_left} days"
            if g["drafted"]:
                a = g["drafted"][0]
                out.append(f"Week target: {need} more to send ({pace}). Start with the {a['company']} draft "
                           f"({ago(a['date'])})")
            else:
                out.append(f"Week target: {need} more applications to send ({pace})")
        for a in g["drafted"][1:]:
            if a["date"] and (t - a["date"]).days >= 5:
                out.append(f"{a['company']} draft is {(t - a['date']).days} days old: send it or mark it Skipped")
                break
        people = self.extra["people"]
        for p in people:
            if p["sent"] and not p["accepted"] and (t - p["sent"]).days >= 14:
                out.append(f"{p['name'].split()[0]}'s request ({p['target']}) pending {(t - p['sent']).days}d: "
                           "find a 2nd contact there")
        if self.extra.get("has_connections") and len(people) < 10:
            nxt = next((x for x in self.targets if x["status"] == "Not Started"), None)
            where = f" at {nxt['company']}" if nxt else ""
            out.append(f"Only {len(people)} people in Connections: add 2 contacts{where} today")
        for p in self.extra["posts"]:
            if p["date"] == t and p["status"] != "Posted":
                out.append(f"Post today's LinkedIn post: {p['topic']}")
            elif 0 < (p["date"] - t).days <= 3 and p["status"] == "Idea":
                out.append(f"Draft {p['date'].strftime('%a')}'s post: {p['topic']}")
        if t.weekday() == 5:
            out.append("Saturday: run the Weekly Review + log LinkedIn KPIs")
        if t.weekday() == 6:
            out.append(f"Sunday: pick next week's {WEEKLY_TARGET} target companies")
        sent = sum(a["status"] in SENT for a in self.apps)
        if sent >= 5 and not g["hot"]:
            out.append(f"{sent} sent, no interview yet: ask one contact for a referral before the next application")
        return out

    def motivation(self, this, last, since):
        if self.groups["hot"]:
            return "You've got an interview in play. Prepare like it's the only one."
        if this >= WEEKLY_TARGET:
            return "Weekly target hit. Spend the spare time networking, that's where referrals come from."
        if since is not None and since >= 4:
            return f"{since} days since the last one went out. One good application today restarts the rhythm."
        if this > last:
            return "Already ahead of last week. Keep the rhythm: 1 a day."
        if self.groups["drafted"]:
            return f"{len(self.groups['drafted'])} applications are already written. Sending one takes 10 minutes."
        return "Quality over quantity: 1 tailored application today."

    def stat_card(self, parent, big, label, sub, color, col, row):
        c = tk.Frame(parent, bg=CARD)
        c.grid(row=row, column=col, sticky="nsew", padx=(0 if col else 0, 1 if col == 0 else 0), pady=(0, 1))
        tk.Label(c, text=label.upper(), bg=CARD, fg=MUTED, font=(MONO, 7)).pack(anchor="w", padx=12, pady=(10, 0))
        tk.Label(c, text=big, bg=CARD, fg=color, font=(DISPLAY, 19, "bold")).pack(anchor="w", padx=12, pady=(2, 0))
        tk.Label(c, text=sub, bg=CARD, fg=DIM, font=(SANS, 8)).pack(anchor="w", padx=12, pady=(0, 10))

    def coach_data(self):
        """All Coach numbers in one place - used by the desktop Coach tab and the phone API."""
        t = dt.date.today()
        this, last, monday = self.week_counts()
        dates = [a["date"] for a in self.apps if a["date"]]
        start = min(dates) if dates else t
        sent_dates = [a["date"] for a in self.apps if a["status"] in SENT and a["date"]]
        last_sent = max(sent_dates) if sent_dates else None
        since = (t - last_sent).days if last_sent else None
        left = dt.datetime.combine(monday + dt.timedelta(days=7), dt.time()) - dt.datetime.now()
        return {"day": (t - start).days + 1, "this": this, "last": last, "since": since, "last_sent": last_sent,
                "left_d": left.days, "left_h": left.seconds // 3600, "streak": self.streak(),
                "sent": sum(a["status"] in SENT for a in self.apps),
                "answered": sum(a["status"] in ("Interview", "Offer", "Rejected") for a in self.apps),
                "due": len(self.groups["due"]), "target": WEEKLY_TARGET, "extra_tile": self.extra_tile(),
                "motivation": self.motivation(this, last, since)}

    def extra_tile(self):
        """6th Coach tile: a configured deadline (e.g. visa end) countdown, else drafts waiting to be sent."""
        if DEADLINE:
            d = DEADLINE["date"]
            return {"label": DEADLINE.get("label", "Deadline"), "value": f"{(d - dt.date.today()).days}d",
                    "sub": f"to {d.strftime('%d %b %Y')}" + (f" · {DEADLINE['note']}" if DEADLINE.get("note") else "")}
        return {"label": "Drafts waiting", "value": str(len(self.groups["drafted"])), "sub": "written, not sent yet"}

    def render_coach(self):
        t = dt.date.today()
        cd = self.coach_data()
        this, last, since, last_sent = cd["this"], cd["last"], cd["since"], cd["last_sent"]
        left = dt.timedelta(days=cd["left_d"], hours=cd["left_h"])
        day = cd["day"]
        self.count.config(text=f"DAY {day:03d} OF THE HUNT")
        self.draw_bar(this / WEEKLY_TARGET)

        # hero: big day number + motivation on a lime slab
        hero = tk.Frame(self.body, bg=LIME)
        hero.pack(fill="x", pady=(8, 0))
        top = tk.Frame(hero, bg=LIME)
        top.pack(fill="x", padx=14, pady=(12, 0))
        tk.Label(top, text="DAY", bg=LIME, fg=BG, font=(MONO_B, 8, "bold")).pack(side="left", anchor="s",
                                                                                pady=(0, 7))
        tk.Label(top, text=str(day), bg=LIME, fg=BG, font=(DISPLAY, 30, "bold")).pack(side="left", padx=(6, 0))
        tk.Label(top, text=f"WEEK\n{this}/{WEEKLY_TARGET}", bg=LIME, fg=BG, font=(MONO_B, 9, "bold"),
                 justify="right").pack(side="right", anchor="e")
        tk.Label(hero, text=self.motivation(this, last, since), bg=LIME, fg=BG, font=(SANS_B, 11),
                 wraplength=WIDTH - 70, justify="left", anchor="w").pack(fill="x", padx=14, pady=(2, 14))

        grid = tk.Frame(self.body, bg=LINE)
        grid.pack(fill="x", pady=(10, 0))
        grid.columnconfigure((0, 1), weight=1, uniform="c")
        trend = "▲" if this > last else "▼" if this < last else "="
        self.stat_card(grid, f"{this}/{WEEKLY_TARGET}", "Sent this week", f"last week {last} {trend}",
                       LIME if this >= WEEKLY_TARGET else FG, 0, 0)
        self.stat_card(grid, f"{left.days}d {left.seconds // 3600:02d}h", "Week closes in", "Sunday 24:00",
                       ORANGE if left.days < 2 and this < WEEKLY_TARGET else FG, 1, 0)
        self.stat_card(grid, "-" if since is None else ("0d" if since == 0 else f"{since}d"), "Since last sent",
                       last_sent.strftime("%a %d %b") if last_sent else "none yet",
                       RED if since and since >= 4 else FG, 0, 1)
        self.stat_card(grid, f"{self.streak()}d", "Routine streak", "days at 60%+ checklist", LIME, 1, 1)
        sent = sum(a["status"] in SENT for a in self.apps)
        answered = sum(a["status"] in ("Interview", "Offer", "Rejected") for a in self.apps)
        self.stat_card(grid, str(sent), "Applications out",
                       f"{answered} answered · {len(self.groups['due'])} follow-ups due", FG, 0, 2)
        x = self.extra_tile()
        self.stat_card(grid, x["value"], x["label"], x["sub"], MUTED, 1, 2)

        moves = self.suggestions()
        self.section("NEXT MOVES", LIME, count=len(moves))
        have = {i["text"] for i in self.s["items"]}
        for n, text in enumerate(moves, 1):
            row = tk.Frame(self.body, bg=BG)
            row.pack(fill="x", pady=5)
            tk.Label(row, text=f"{n:02d}", bg=BG, fg=LIME, font=(MONO_B, 9, "bold")).pack(side="left", anchor="n",
                                                                                        padx=(0, 10))
            added = text in have
            btn = tk.Label(row, text="✓ ON LIST" if added else "+ TODAY", bg=BG if added else LIME,
                           fg=DIM if added else BG, font=(MONO_B, 7, "bold"), padx=6, pady=2,
                           cursor="" if added else "hand2")
            btn.pack(side="right", anchor="n", padx=(8, 0))
            if not added:
                btn.bind("<Button-1>", lambda e, x=text: self.add_suggestion(x))
            tk.Label(row, text=text, bg=BG, fg=DIM if added else FG, font=FONT, wraplength=WIDTH - 150,
                     justify="left", anchor="w").pack(side="left", fill="x", expand=True)

    def add_suggestion(self, text):
        self.s["items"].insert(0, {"text": text, "done": False})
        save_state(self.s)
        self.render()

    # ---- phone API (runs on the Tk thread via pump) ----
    def state_json(self):
        t = dt.date.today()
        items = self.s["items"]
        auto = self.auto_items()
        cd = self.coach_data()
        cd["last_sent"] = cd["last_sent"].strftime("%a %d %b") if cd["last_sent"] else None
        have = {i["text"] for i in items}
        n = lambda *st: sum(a["status"] in st for a in self.apps)
        groups = []
        for key, title, color in GROUPS:
            cards = []
            for a in self.groups[key]:
                meta, meta_color = card_meta(key, a)
                cards.append({"row": a["row"], "company": a["company"], "role": a["role"], "status": a["status"],
                              "color": STATUS_COLOR.get(a["status"], DIM), "meta": meta, "meta_color": meta_color,
                              "actions": [{"label": l, "kind": k, "status": s, "filled": f}
                                          for l, k, s, f in card_actions(a)]})
            groups.append({"key": key, "title": title, "color": color, "cards": cards})
        not_started = [x for x in self.targets if x["status"] == "Not Started"]
        return {
            "date": t.isoformat(), "weekday": t.strftime("%A"), "day_month": t.strftime("%d %b"),
            "today": {"auto": [{"text": x, "done": x in self.s["auto_done"]} for x in auto],
                      "items": [{"text": i["text"], "done": i["done"]} for i in items]},
            "coach": cd,
            "moves": [{"text": m, "added": m in have} for m in self.suggestions()],
            "pipeline": {"counts": {"Applied": n("Applied"), "Interview": n("Interview"), "Offer": n("Offer"),
                                    "To send": len(self.groups["drafted"]), "Rejected": n("Rejected")},
                         "groups": groups, "statuses": STATUSES,
                         "targets": {"untouched": len(not_started), "total": len(self.targets),
                                     "next": [{"prio": x["prio"], "company": x["company"]} for x in not_started[:3]]}},
            "busy": bool(getattr(self, "busy", False)), "msg": self.err,
            "synced": self.loaded_at.strftime("%H:%M") if getattr(self, "loaded_at", None) else None,
        }

    def phone_do(self, p):
        """One action from the phone. Items are matched by text, so a stale phone view can't hit the wrong row."""
        act, text = p.get("action"), p.get("text", "")
        idx = next((i for i, it in enumerate(self.s["items"]) if it["text"] == text), None)
        if act == "toggle" and idx is not None:
            self.toggle(idx)
        elif act == "toggle_auto":
            self.toggle_auto(text)
        elif act == "delete" and idx is not None:
            self.delete(idx)
        elif act == "add" and text.strip():
            self.s["items"].append({"text": text.strip()[:200], "done": False})
            save_state(self.s)
            self.render()
        elif act == "suggest" and idx is None and text:
            self.add_suggestion(text)
        elif act == "sync":
            self.force_reload()
        elif act in ("status", "followed", "remove"):
            app = next((a for a in self.apps if a["row"] == p.get("row") and a["company"] == p.get("company")), None)
            if app is None:
                raise RuntimeError("that application changed on the PC, pull to refresh")
            label = {"status": f"{app['company'][:18]} → {p.get('status')}",
                     "followed": f"followed up {app['company'][:18]}",
                     "remove": f"removed {app['company'][:18]}"}[act]
            self.tracker_action(app, act, p.get("status"), label, confirm=False)
        return self.state_json()

    def phone_dialog(self):
        """Window with QR codes to open the phone app."""
        import qrcode
        from PIL import ImageTk
        urls = self.phone.urls()
        win = tk.Toplevel(self.root, bg=BG)
        win.title("Pursuit · Phone")
        win.resizable(False, False)
        try:
            win.update()
            ctypes.windll.dwmapi.DwmSetWindowAttribute(ctypes.windll.user32.GetParent(win.winfo_id()), 20,
                                                       ctypes.byref(ctypes.c_int(1)), 4)
        except Exception:
            pass
        tk.Label(win, text="OPEN ON YOUR PHONE", bg=BG, fg=LIME, font=(MONO_B, 8, "bold")).pack(anchor="w", padx=20,
                                                                                             pady=(18, 2))
        tk.Label(win, text="Scan with the camera, then\nAdd to Home Screen.", bg=BG, fg=FG, font=(SANS_B, 12),
                 justify="left").pack(anchor="w", padx=20)
        self.qr_imgs = []
        shown = [(k, u) for k, u in (("APP · INSTALLABLE (HTTPS)", urls["https"]),
                                     ("ANYWHERE · TAILSCALE", None if urls["https"] else urls["tailscale"]),
                                     ("HOME WI-FI", None if urls["https"] else urls["lan"])) if u]
        for label, url in shown:
            q = qrcode.QRCode(border=2, box_size=5)
            q.add_data(url)
            img = q.make_image(fill_color="#0b0b0c", back_color="#c8ff3d").get_image()
            self.qr_imgs.append(ImageTk.PhotoImage(img))
            tk.Label(win, text=label, bg=BG, fg=MUTED, font=(MONO_B, 7, "bold")).pack(anchor="w", padx=20,
                                                                                     pady=(16, 4))
            tk.Label(win, image=self.qr_imgs[-1], bg=BG).pack(anchor="w", padx=20)
            e = tk.Entry(win, bg=CARD, fg=FG, relief="flat", font=(MONO, 8), readonlybackground=CARD, width=46)
            e.insert(0, url)
            e.config(state="readonly")
            e.pack(anchor="w", padx=20, pady=(6, 0), ipady=4)
        if not urls["tailscale"]:
            tk.Label(win, text="Tailscale isn't connected on this PC yet, so the phone\nonly reaches Pursuit on the "
                               "same Wi-Fi. Install Tailscale on PC +\nphone and sign in with the same account to "
                               "use it anywhere.", bg=BG, fg=ORANGE, font=(SANS, 9), justify="left").pack(
                anchor="w", padx=20, pady=(14, 0))
        tk.Label(win, text="Keep Pursuit running on the PC (minimise, don't close)\nand the PC awake, or the phone "
                           "shows the last saved view.", bg=BG, fg=DIM, font=(SANS, 9), justify="left").pack(
            anchor="w", padx=20, pady=(12, 18))

    def toggle_closed(self):
        self.show_closed = not self.show_closed
        self.render()

    def open_folder(self, a):
        f = find_folder(a)
        if f:
            os.startfile(f)
        else:
            self.err = f"No Applications folder found for {a['company']}"
            self.render()

    def app_menu(self, e, a, anchor=None):
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label=f"{a['company'][:40]}", state="disabled")
        m.add_separator()
        for st in STATUSES:
            m.add_command(label=("● " if st == a["status"] else "   ") + f"Set status: {st}",
                          command=lambda st=st: self.set_status(a, st))
        m.add_separator()
        m.add_command(label="   Open application folder", command=lambda: self.open_folder(a))
        m.add_command(label="   Remove from tracker…", command=lambda: self.tracker_action(a, "remove"))
        if e is not None:
            m.tk_popup(e.x_root, e.y_root)
        else:
            m.tk_popup(anchor.winfo_rootx() + anchor.winfo_width() - 180, anchor.winfo_rooty() + 24)

    # ---- window behaviour ----
    def hwnd(self):
        return ctypes.windll.user32.GetParent(self.root.winfo_id()) or self.root.winfo_id()

    def close(self):
        self.s["pos"] = (self.root.winfo_x(), self.root.winfo_y())
        self.s["height"] = self.root.winfo_height()
        save_state(self.s)
        self.root.destroy()

    def menu(self, e):
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label="Untick all", command=self.reset_ticks)
        m.add_command(label="Restore default items", command=self.restore_defaults)
        m.add_command(label="Reload tracker", command=self.force_reload)
        m.add_separator()
        m.add_command(label="Close", command=self.close)
        m.tk_popup(e.x_root, e.y_root)

    def restore_defaults(self):
        have = {i["text"] for i in self.s["items"]}
        self.s["items"] += [{"text": t, "done": False} for t in DEFAULT_ITEMS if t not in have]
        save_state(self.s)
        self.render()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    lock = socket.socket()
    try:
        lock.bind(("127.0.0.1", 47613))  # single instance
    except OSError:
        sys.exit(0)
    Checklist().run()
