"""COA request engine: Rencana Kirim (plan) + SAP EXPORT (stock) -> PERMINTAAN COA.

Same logic as the original standalone script, but every function takes uploaded
file contents instead of fixed file names, so a web app can call it.
"""
import copy
import io
import re
from datetime import date, datetime, timedelta

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

EPOCH = date(1899, 12, 30)

MONTHS_ID = ["januari", "februari", "maret", "april", "mei", "juni", "juli",
             "agustus", "september", "oktober", "november", "desember"]


# ---------------- helpers ----------------
def to_date(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return EPOCH + timedelta(days=int(float(v)))
    except (ValueError, TypeError):
        try:
            return pd.to_datetime(v).date()
        except Exception:
            return None


def to_float(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def to_code(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def to_text(v):
    return str(v).strip() if v is not None else ""


# ---------------- 0. plan sheet helpers ----------------
def list_plan_sheets(plan_file):
    wb = load_workbook(plan_file, read_only=True)
    return wb.sheetnames


def guess_sheet_index(sheets, d):
    """Pick the sheet named like '27 Oktober' for date d, else the last sheet."""
    want = "%d %s" % (d.day, MONTHS_ID[d.month - 1])
    for i, s in enumerate(sheets):
        if s.strip().lower() == want:
            return i
    return len(sheets) - 1


# ---------------- 1. read shipping plan ----------------
def read_plan(plan_file, plan_sheet, plan_date_override=None):
    wb = load_workbook(plan_file, data_only=True)
    name = next(s for s in wb.sheetnames if s.strip() == plan_sheet.strip())
    rows = list(wb[name].iter_rows(values_only=True))

    plan_date = None
    for r in rows[:10]:
        for c in r:
            if isinstance(c, str):
                m = re.search(r"TANGGAL\s*(\d+)\s*-\s*(\d+)\s*-\s*(\d{4})", c)
                if m:
                    d, mo, y = map(int, m.groups())
                    plan_date = date(y, mo, d)
    if plan_date_override:
        plan_date = plan_date_override
    if plan_date is None:
        raise ValueError("Could not find 'TANGGAL dd - mm - yyyy' in the first 10 rows of sheet '%s'." % name)

    hdr_i = next((i for i, r in enumerate(rows)
                  if any(isinstance(c, str) and c.strip().upper() == "NAMA CUSTOMER" for c in r)), None)
    if hdr_i is None:
        raise ValueError("Could not find the 'NAMA CUSTOMER' header row in sheet '%s'." % name)
    hdr = [str(c).strip().upper() if c is not None else "" for c in rows[hdr_i]]

    def col(pred):
        return next(i for i, h in enumerate(hdr) if pred(h))

    c_mobil = next((i for i, h in enumerate(hdr) if h.startswith("JENIS")), None)
    c_cust = col(lambda h: h == "NAMA CUSTOMER")
    c_kode = col(lambda h: h == "KODE BRG")
    c_name = col(lambda h: h == "NAMA BARANG")
    c_qty = col(lambda h: h.startswith("QTY"))
    c_st = col(lambda h: h == "ST")
    c_so = col(lambda h: h == "NOMOR SO")
    c_do = col(lambda h: h == "NOMOR DO")

    lines = []
    last_cust, last_mobil, group = "", "", 0
    for r in rows[hdr_i + 1:]:
        r = list(r) + [None] * 30

        # Track the truck block on EVERY row (even skipped ones) so blocks stay aligned with the plan.
        cust_cell = to_text(r[c_cust])
        mobil_cell = to_text(r[c_mobil]) if c_mobil is not None else ""
        if (mobil_cell and mobil_cell != last_mobil) or (cust_cell and cust_cell != last_cust):
            group += 1
        if cust_cell:
            last_cust = cust_cell
        if mobil_cell:
            last_mobil = mobil_cell

        kode = to_code(r[c_kode])
        so = to_code(r[c_so])
        # customer lines only: 10-digit SO. Skips "--->", TUKAR GULING, and internal transfers
        if not kode.isdigit() or not re.fullmatch(r"\d{10}", so):
            continue
        tail = " ".join(str(c) for c in r[c_do:] if c is not None)
        if "CANCEL" in tail.upper():
            continue
        qty = to_float(r[c_qty])
        if qty <= 0:
            continue
        lines.append(dict(group=group, mobil=last_mobil, customer=last_cust, kode=kode,
                          product=to_text(r[c_name]), qty=qty,
                          unit=to_text(r[c_st]).upper(), so=so))
    return plan_date, lines


# ---------------- 2. read SAP stock ----------------
def read_stock(export_file, plan_date, exclude_expired=True, trading_slocs=("2002",)):
    df = pd.read_excel(export_file, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    needed = ["material", "unrestricted", "sled/bbd", "batch", "storage location",
              "date of manufacture", "sled percentage", "base unit of measure"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError("EXPORT file is missing columns: %s" % ", ".join(missing))
    df = df[df["material"].fillna("").str.strip().str.fullmatch(r"\d+")]   # drops total rows
    df["unrestricted"] = pd.to_numeric(df["unrestricted"], errors="coerce").fillna(0)
    df = df[df["unrestricted"] > 0]       # unrestricted only: no transit, QI, blocked, returns

    trading_slocs = set(trading_slocs)
    pools = {}
    for _, r in df.iterrows():
        mat = r["material"].strip()
        sled = to_date(r["sled/bbd"])
        if exclude_expired and sled and sled <= plan_date:
            continue

        sp = r.get("special stock number")
        m = re.match(r"\s*(\d{10})", sp) if isinstance(sp, str) else None
        special = m.group(1) if m else ""          # "1200019764/ 200" -> SO 1200019764

        batch = str(r["batch"]).strip()
        sloc = str(r["storage location"]).strip()
        supplier = r.get("supplier batch")
        supplier = supplier.strip() if isinstance(supplier, str) else ""
        is_trading = sloc in trading_slocs or mat.startswith("41")
        label = supplier if (is_trading and supplier) else batch   # lot printed on the COA request

        lots = pools.setdefault(mat, {})
        key = (batch, special)
        if key in lots:
            lots[key]["qty"] += float(r["unrestricted"])
        else:
            lots[key] = dict(batch=batch, lot=label, special=special,
                             qty=float(r["unrestricted"]),
                             sled=sled, dom=to_date(r["date of manufacture"]),
                             export_pct=to_float(r["sled percentage"]),
                             unit=str(r["base unit of measure"]).strip().upper())
    return {m: list(v.values()) for m, v in pools.items()}


# ---------------- 3. allocate lots (FEFO) ----------------
def allocate(line, pools):
    lots = pools.get(line["kode"], [])
    if not lots:
        return [], line["qty"], "TIDAK ADA STOK UNRESTRICTED"
    stock_unit = lots[0]["unit"]
    if line["unit"] and line["unit"] != stock_unit:
        return [], None, "UNIT BEDA (rencana %s / stok %s)" % (line["unit"], stock_unit)

    need = line["qty"]
    cands = [l for l in lots
             if l["qty"] > 0.0005 and (l["special"] == "" or l["special"] == line["so"])]
    # stock reserved to this SO first, then FEFO
    cands.sort(key=lambda l: (0 if l["special"] else 1, l["sled"] or date.max, l["batch"]))
    out = []
    for l in cands:
        if need <= 0.0005:
            break
        take = min(l["qty"], need)
        l["qty"] -= take
        need -= take
        out.append((l, take))
    return out, need, ""


def shelf_pct(lot, ref):
    if lot["sled"] and lot["dom"] and lot["sled"] > lot["dom"]:
        total = (lot["sled"] - lot["dom"]).days
        return round((lot["sled"] - ref).days / total * 100, 2)
    return lot["export_pct"] or None


# ---------------- 4. build blocks + workbook ----------------
def build_blocks(plan_date, lines, pools, ref_offset_days=1):
    """One block per JENIS MOBIL / customer entry, in plan order (never merged per customer)."""
    ref = plan_date - timedelta(days=ref_offset_days)
    blocks = {}
    for ln in lines:
        picks, short, note = allocate(ln, pools)
        blk = blocks.setdefault(ln["group"], dict(customer=ln["customer"], mobil=ln["mobil"], rows={}))
        rows = blk["rows"]
        for lot, take in picks:
            k = (ln["product"], lot["lot"])
            if k in rows:
                rows[k]["qty"] += take
            else:
                rows[k] = dict(sled=lot["sled"], pct=shelf_pct(lot, ref),
                               qty=take, remark="", unit=lot["unit"])
        if note or (short and short > 0.0005):
            msg = note or "STOK KURANG %.1f %s" % (short, ln["unit"] or "")
            rows[(ln["product"], "CEK STOK")] = dict(
                sled=None, pct=None, qty=ln["qty"] if note else short,
                remark=msg, unit=ln["unit"])
    return blocks


def blocks_list(blocks):
    """Non-empty blocks as a plain ordered list (what the audit compares against)."""
    out = []
    for blk in blocks.values():
        if not blk["rows"]:
            continue
        out.append(dict(
            customer=blk["customer"], mobil=blk["mobil"],
            rows=[dict(product=p, lot=l, qty=d["qty"], sled=d["sled"], unit=d["unit"], remark=d["remark"])
                  for (p, l), d in blk["rows"].items()]))
    return out


def flatten(blocks):
    """Flat table (every cell filled) for on-screen preview."""
    recs = []
    for i, b in enumerate(blocks_list(blocks), 1):
        for r in b["rows"]:
            recs.append({"Block": i, "Customer": b["customer"], "Truck": b["mobil"],
                         "Product": r["product"], "Lot": r["lot"],
                         "Expired": r["sled"].strftime("%d/%m/%Y") if r["sled"] else "",
                         "Qty": round(r["qty"], 3), "Unit": r["unit"], "Remarks": r["remark"]})
    return pd.DataFrame(recs)


def write_workbook(blocks, plan_date, show_mobil=True):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = plan_date.strftime("%d/%m/%Y")
    heads = ["CUSTOMER", "PRODUK", "LOT", "EXPIRED DATE", "SELED PERCENTAGE",
             "SERAH TERIMA COA", "", "REMARKS", "QTY", "SAT"]
    if show_mobil:
        heads.append("JENIS MOBIL")
    for i, h in enumerate(heads, 1):
        ws.cell(2, i, h)
    ws.merge_cells("F2:G2")
    ws["F3"], ws["G3"] = "QA FG", "LOG"
    for c in ws[2] + ws[3]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")

    r = 4
    for blk in blocks.values():
        rows = blk["rows"]
        if not rows:
            continue
        first_row, last_p = True, None
        for (prod, lot), d in rows.items():
            ws.cell(r, 1, blk["customer"] if first_row else None)
            ws.cell(r, 2, prod if prod != last_p else None)
            ws.cell(r, 3, lot)
            if d["sled"]:
                ws.cell(r, 4, d["sled"]).number_format = "dd/mm/yyyy"
            if d["pct"] is not None:
                ws.cell(r, 5, d["pct"]).number_format = "0.00"
            ws.cell(r, 8, d["remark"] or None)
            ws.cell(r, 9, round(d["qty"], 3) if d["qty"] else None)
            ws.cell(r, 10, d["unit"] or None)
            if show_mobil and first_row:
                ws.cell(r, 11, blk["mobil"] or None)
            first_row, last_p = False, prod
            r += 1
        r += 1   # blank row between blocks

    widths = (34, 40, 22, 14, 18, 10, 10, 34, 12, 8, 18)
    for c, w in zip("ABCDEFGHIJK", widths):
        ws.column_dimensions[c].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------- 5. one-call pipeline ----------------
def generate(plan_bytes, sheet, export_bytes, ref_offset_days=1, exclude_expired=True,
             trading_slocs=("2002",), show_mobil=True, plan_date_override=None):
    plan_date, lines = read_plan(io.BytesIO(plan_bytes), sheet, plan_date_override)
    pools = read_stock(io.BytesIO(export_bytes), plan_date, exclude_expired, trading_slocs)
    snapshot = copy.deepcopy(pools)           # stock before allocation (used by the audit)
    blocks = build_blocks(plan_date, lines, pools, ref_offset_days)
    shown = [b for b in blocks.values() if b["rows"]]
    flagged = sum(1 for b in shown for k in b["rows"] if k[1] == "CEK STOK")
    return dict(
        plan_date=plan_date, lines=lines, blocks=blocks, snapshot=snapshot,
        name2kode={re.sub(r"\s+", " ", l["product"].upper()): l["kode"] for l in lines},
        xlsx=write_workbook(blocks, plan_date, show_mobil),
        stats=dict(blocks=len(shown), lines=len(lines), flagged=flagged),
    )
