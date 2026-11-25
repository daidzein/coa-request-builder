"""Weekly Audit: compare the manually picked batches with the automated (FEFO) picks."""
import difflib
import io
import re
from collections import OrderedDict, defaultdict

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

import coa_engine as eng

QTY_TOL = 0.001

STATUS_MATCH = "MATCH"
STATUS_QTY = "QTY DIFFERENT"
STATUS_MISSING = "MISSING IN MANUAL"
STATUS_EXTRA = "EXTRA IN MANUAL"
STATUS_NO_MANUAL_BLOCK = "BLOCK MISSING IN MANUAL"
STATUS_NO_AUTO_BLOCK = "BLOCK NOT IN PLAN"

COLORS = {
    STATUS_MATCH: "E2F0D9",
    STATUS_QTY: "FFF2CC",
    STATUS_MISSING: "F8CBAD",
    STATUS_EXTRA: "F8CBAD",
    STATUS_NO_MANUAL_BLOCK: "F4B084",
    STATUS_NO_AUTO_BLOCK: "F4B084",
}


def norm(s):
    return re.sub(r"\s+", " ", str(s or "").strip().upper())


def fmt_date(d):
    return d.strftime("%d/%m/%Y") if d else ""


# ---------- read the manual file (same layout as the generated PERMINTAAN COA) ----------
def parse_manual(data):
    wb = load_workbook(io.BytesIO(data), data_only=True)
    rows = list(wb.worksheets[0].iter_rows(values_only=True))

    start = 3                                   # default: data starts on row 4
    for i, r in enumerate(rows[:10]):
        if any(isinstance(c, str) and c.strip().upper() == "QA FG" for c in r):
            start = i + 1
            break

    blocks, cur, last_prod = [], None, ""
    for r in rows[start:]:
        r = list(r) + [None] * 12
        cust = eng.to_text(r[0])
        prod = eng.to_text(r[1])
        lot = eng.to_code(r[2]).upper()
        qty = eng.to_float(r[8])
        if not (cust or prod or lot or qty):
            continue                            # blank separator row
        if cust:                                # customer is only filled on the first row of a block
            cur = dict(customer=cust, rows=[])
            blocks.append(cur)
            last_prod = ""
        if cur is None:
            continue
        if prod:
            last_prod = prod
        if not lot and not qty:
            continue
        cur["rows"].append(dict(product=last_prod, lot=lot, sled=eng.to_date(r[3]),
                                qty=qty, unit=eng.to_text(r[9]).upper()))
    return blocks


# ---------- pair up blocks ----------
def match_blocks(auto, manual):
    """Same customer matched in order of appearance (so the same customer on two trucks stays
    two blocks). Leftovers get a fuzzy name match to absorb typos."""
    pending = defaultdict(list)
    for j, m in enumerate(manual):
        pending[norm(m["customer"])].append(j)

    pairs, unmatched_auto = [], []
    for i, a in enumerate(auto):
        lst = pending.get(norm(a["customer"]))
        if lst:
            pairs.append((i, lst.pop(0)))
        else:
            unmatched_auto.append(i)

    used = {j for _, j in pairs}
    leftover = [j for j in range(len(manual)) if j not in used]
    for i in list(unmatched_auto):
        names = [norm(manual[j]["customer"]) for j in leftover]
        hit = difflib.get_close_matches(norm(auto[i]["customer"]), names, n=1, cutoff=0.85)
        if hit:
            j = leftover[names.index(hit[0])]
            pairs.append((i, j))
            leftover.remove(j)
            unmatched_auto.remove(i)
    pairs.sort()
    return pairs, unmatched_auto, leftover


# ---------- compare ----------
def compare(auto_blocks, manual_blocks, snapshot, name2kode):
    """auto_blocks: coa_engine.blocks_list(...). Returns (summary dict, DataFrame)."""

    def stock_lot(prod, lot):
        for l in snapshot.get(name2kode.get(norm(prod)), []):
            if lot in (l["lot"].upper(), l["batch"].upper()):
                return l
        return None

    pairs, ua, um = match_blocks(auto_blocks, manual_blocks)
    out = []

    def agg(rows):
        d = OrderedDict()
        for r in rows:
            k = (norm(r["product"]), str(r["lot"]).upper())
            if k in d:
                d[k]["qty"] += r["qty"]
            else:
                d[k] = dict(product=r["product"], qty=r["qty"], sled=r.get("sled"), unit=r.get("unit", ""))
        return d

    for i, j in pairs:
        A, M = auto_blocks[i], manual_blocks[j]
        a, m = agg(A["rows"]), agg(M["rows"])
        earliest_auto = {}
        for (p, _), v in a.items():
            if v["sled"] and (p not in earliest_auto or v["sled"] < earliest_auto[p]):
                earliest_auto[p] = v["sled"]

        for k in list(a) + [k for k in m if k not in a]:
            prod, lot = k
            av, mv = a.get(k), m.get(k)
            note = ""
            if av and mv:
                status = STATUS_MATCH if abs(av["qty"] - mv["qty"]) <= QTY_TOL else STATUS_QTY
            elif av:
                status = STATUS_MISSING
            else:
                status = STATUS_EXTRA
                sl = stock_lot(mv["product"], lot)
                if lot != "CEK STOK":
                    if name2kode.get(prod) is None:
                        note = "Product name not found in the plan"
                    elif sl is None:
                        note = "Lot not in unrestricted / non-expired stock"
                    elif sl["sled"] and earliest_auto.get(prod) and sl["sled"] > earliest_auto[prod]:
                        note = "Later SLED than the FEFO pick - check customer shelf-life requirement"
            out.append({
                "Block": i + 1, "Customer": A["customer"], "Truck": A["mobil"],
                "Product": (av or mv)["product"], "Lot": lot,
                "Auto Qty": round(av["qty"], 3) if av else None,
                "Manual Qty": round(mv["qty"], 3) if mv else None,
                "Unit": (av or mv)["unit"],
                "Auto SLED": fmt_date(av["sled"]) if av else "",
                "Manual SLED": fmt_date(mv["sled"]) if mv else "",
                "Status": status, "Note": note,
            })

    for i in ua:
        A = auto_blocks[i]
        for r in A["rows"]:
            out.append({"Block": i + 1, "Customer": A["customer"], "Truck": A["mobil"],
                        "Product": r["product"], "Lot": str(r["lot"]).upper(),
                        "Auto Qty": round(r["qty"], 3), "Manual Qty": None, "Unit": r["unit"],
                        "Auto SLED": fmt_date(r["sled"]), "Manual SLED": "",
                        "Status": STATUS_NO_MANUAL_BLOCK, "Note": "Customer block not found in the manual sheet"})
    for j in um:
        M = manual_blocks[j]
        for r in M["rows"]:
            out.append({"Block": None, "Customer": M["customer"], "Truck": "",
                        "Product": r["product"], "Lot": r["lot"],
                        "Auto Qty": None, "Manual Qty": round(r["qty"], 3), "Unit": r["unit"],
                        "Auto SLED": "", "Manual SLED": fmt_date(r["sled"]),
                        "Status": STATUS_NO_AUTO_BLOCK, "Note": "Customer block not found in the plan"})

    df = pd.DataFrame(out)
    total = len(df)
    counts = df["Status"].value_counts().to_dict() if total else {}
    n_match = int(counts.get(STATUS_MATCH, 0))
    mn = float(df[["Auto Qty", "Manual Qty"]].fillna(0).min(axis=1).sum()) if total else 0.0
    mx = float(df[["Auto Qty", "Manual Qty"]].fillna(0).max(axis=1).sum()) if total else 0.0
    summary = dict(
        rows_total=total, rows_match=n_match,
        match_pct=round(n_match / total * 100, 2) if total else 0.0,
        qty_agreement=round(mn / mx * 100, 2) if mx else 0.0,
        blocks_auto=len(auto_blocks), blocks_manual=len(manual_blocks),
        blocks_matched=len(pairs),
        by_status={k: int(v) for k, v in counts.items()},
    )
    return summary, df


# ---------- downloadable report ----------
def build_report(summary, df, meta=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    items = list((meta or {}).items()) + [
        ("Rows compared", summary["rows_total"]),
        ("Rows matching", summary["rows_match"]),
        ("Row match %", summary["match_pct"]),
        ("Qty agreement %", summary["qty_agreement"]),
        ("Blocks in plan / manual / paired",
         "%s / %s / %s" % (summary["blocks_auto"], summary["blocks_manual"], summary["blocks_matched"])),
    ] + [("Status: " + k, v) for k, v in summary["by_status"].items()]
    for i, (k, v) in enumerate(items, 1):
        ws.cell(i, 1, k).font = Font(bold=True)
        ws.cell(i, 2, v)
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 30

    wd = wb.create_sheet("Detail")
    cols = list(df.columns)
    for c, h in enumerate(cols, 1):
        cell = wd.cell(1, c, h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for r, rec in enumerate(df.itertuples(index=False), 2):
        fill = PatternFill("solid", fgColor=COLORS.get(rec[cols.index("Status")], "FFFFFF"))
        for c, v in enumerate(rec, 1):
            if v is not None and isinstance(v, float) and pd.isna(v):
                v = None
            wd.cell(r, c, v).fill = fill
    for col, w in zip("ABCDEFGHIJKL", (7, 34, 14, 36, 20, 11, 11, 7, 12, 12, 24, 60)):
        wd.column_dimensions[col].width = w
    wd.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
