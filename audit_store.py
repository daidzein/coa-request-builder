"""SQLite history of weekly audits + the 'is this week's Wednesday audit done?' rule."""
import json
import os
import sqlite3
from datetime import timedelta

DB_PATH = os.getenv("AUDIT_DB", "data/audit.db")


def _conn():
    os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def init():
    with _conn() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS audits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kind TEXT NOT NULL,              -- AUDIT or SKIPPED
            iso_year INTEGER NOT NULL,
            iso_week INTEGER NOT NULL,
            plan_date TEXT,
            auditor TEXT,
            created_at TEXT NOT NULL,
            rows_total INTEGER,
            rows_match INTEGER,
            match_pct REAL,
            qty_agreement REAL,
            notes TEXT,
            detail_json TEXT)""")


def add_audit(now, kind, auditor, plan_date=None, summary=None, df=None, notes=""):
    y, w, _ = now.isocalendar()
    summary = summary or {}
    detail = json.dumps({"summary": summary,
                         "rows": json.loads(df.to_json(orient="records")) if df is not None else []})
    with _conn() as c:
        c.execute("""INSERT INTO audits (kind, iso_year, iso_week, plan_date, auditor, created_at,
                     rows_total, rows_match, match_pct, qty_agreement, notes, detail_json)
                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (kind, y, w, str(plan_date) if plan_date else None, auditor, now.isoformat(timespec="seconds"),
                   summary.get("rows_total"), summary.get("rows_match"), summary.get("match_pct"),
                   summary.get("qty_agreement"), notes, detail))


def list_audits():
    with _conn() as c:
        return [dict(r) for r in c.execute(
            "SELECT id, kind, iso_year, iso_week, plan_date, auditor, created_at, rows_total, rows_match, "
            "match_pct, qty_agreement, notes FROM audits ORDER BY id DESC")]


def get_detail(audit_id):
    with _conn() as c:
        r = c.execute("SELECT detail_json FROM audits WHERE id=?", (audit_id,)).fetchone()
    return json.loads(r["detail_json"]) if r and r["detail_json"] else None


def week_status(now, enforce=True):
    """State for the current ISO week. Audit is due Wednesday; Generate is locked from
    Wednesday until the week's audit is saved (or skipped with the supervisor PIN)."""
    y, w, _ = now.isocalendar()
    with _conn() as c:
        rec = c.execute("SELECT * FROM audits WHERE iso_year=? AND iso_week=? ORDER BY id DESC LIMIT 1",
                        (y, w)).fetchone()
    wd = now.weekday()      # Mon=0 ... Wed=2
    if rec:
        return dict(state="DONE" if rec["kind"] == "AUDIT" else "SKIPPED", locked=False,
                    iso_year=y, iso_week=w, record=dict(rec))
    if wd < 2:
        return dict(state="UPCOMING", locked=False, iso_year=y, iso_week=w, record=None)
    return dict(state="DUE" if wd == 2 else "OVERDUE", locked=bool(enforce),
                iso_year=y, iso_week=w, record=None)


def compliance(now, weeks=12):
    """One row per ISO week (newest first): Done / Skipped / Missed / Pending."""
    rows = list_audits()
    by_week = {}
    for r in rows:                              # newest first, so keep the first seen
        by_week.setdefault((r["iso_year"], r["iso_week"]), r)
    monday = (now - timedelta(days=now.weekday())).date()
    cur = now.isocalendar()[:2]
    out = []
    for i in range(weeks):
        d = monday - timedelta(weeks=i)
        key = d.isocalendar()[:2]
        r = by_week.get(key)
        if r:
            status = "Done" if r["kind"] == "AUDIT" else "Skipped"
        elif key == cur:
            status = "Pending" if now.weekday() <= 2 else "Overdue"
        else:
            status = "Missed"
        out.append({"Week": "%d-W%02d" % key, "Wednesday": (d + timedelta(days=2)).strftime("%d/%m/%Y"),
                    "Status": status,
                    "Auditor": r["auditor"] if r else "",
                    "Match %": r["match_pct"] if r else None})
    return out
