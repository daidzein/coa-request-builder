"""Smoke test with synthetic data: python test_smoke.py"""
import copy
import io
import os
import tempfile
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
from openpyxl import Workbook

os.environ["AUDIT_DB"] = os.path.join(tempfile.mkdtemp(), "audit.db")
import audit as aud
import audit_store as store
import coa_engine as eng


def make_plan():
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
    ws.append([None, None, 1003, "PRODUK Z", 10, "KG", "1200000003", "8000003"])   # no stock -> CEK STOK
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_export():
    cols = ["Material", "Unrestricted", "SLED/BBD", "Special Stock Number", "Batch", "Storage Location",
            "Supplier Batch", "Date of Manufacture", "SLED Percentage", "Base Unit of Measure"]
    rows = [
        ["1001", "500", "2026-10-20", "", "B0", "1000", "", "2026-01-01", "0", "KG"],   # expired
        ["1001", "120", "2027-03-01", "", "B1", "1000", "", "2026-03-01", "0", "KG"],
        ["1001", "200", "2027-06-01", "", "B2", "1000", "", "2026-06-01", "0", "KG"],
        ["1002", "60", "2027-05-01", "", "B3", "1000", "", "2026-05-01", "0", "KG"],
        ["Total", "880", "", "", "", "", "", "", "", ""],
    ]
    buf = io.BytesIO()
    pd.DataFrame(rows, columns=cols).to_excel(buf, index=False)
    return buf.getvalue()


plan, export = make_plan(), make_export()
res = eng.generate(plan, "27 Oktober", export)
print(eng.flatten(res["blocks"]).to_string())
auto = eng.blocks_list(res["blocks"])
assert res["stats"] == dict(blocks=3, lines=5, flagged=1), res["stats"]
assert [b["customer"] for b in auto] == ["PT ALPHA", "PT BETA", "PT ALPHA"]     # same customer, two blocks
lots_a = {(r["product"], r["lot"]): r["qty"] for r in auto[0]["rows"]}
assert lots_a[("PRODUK X", "B1")] == 100                                          # FEFO, expired B0 skipped
lots_b = {(r["product"], r["lot"]): r["qty"] for r in auto[1]["rows"]}
assert lots_b == {("PRODUK X", "B1"): 20, ("PRODUK X", "B2"): 130}

# manual sheet identical to auto -> 100% match
manual_same = aud.parse_manual(eng.write_workbook(res["blocks"], res["plan_date"]))
s, df = aud.compare(auto, manual_same, res["snapshot"], res["name2kode"])
assert s["match_pct"] == 100.0 and s["blocks_matched"] == 3, s

# manual deviates on block B (all from B2) and block C (typo'd customer + an invalid lot)
blocks2 = copy.deepcopy(res["blocks"])
b = [x for x in blocks2.values() if x["customer"] == "PT BETA"][0]
b["rows"] = {("PRODUK X", "B2"): dict(sled=None, pct=None, qty=150, remark="", unit="KG")}
c = [x for x in blocks2.values() if x["customer"] == "PT ALPHA"][1]
c["customer"] = "PT ALPHA."
c["rows"][("PRODUK X", "B9")] = dict(sled=None, pct=None, qty=5, remark="", unit="KG")
manual = aud.parse_manual(eng.write_workbook(blocks2, res["plan_date"]))
s2, df2 = aud.compare(auto, manual, res["snapshot"], res["name2kode"])
print(df2[["Customer", "Product", "Lot", "Auto Qty", "Manual Qty", "Status", "Note"]].to_string())
print(s2)
st = df2["Status"].tolist()
assert aud.STATUS_QTY in st and aud.STATUS_MISSING in st and aud.STATUS_EXTRA in st
assert s2["blocks_matched"] == 3                                                  # fuzzy matched "PT ALPHA."
assert any("not in unrestricted" in n for n in df2["Note"])
assert len(aud.build_report(s2, df2, {"Auditor": "x"})) > 1000

# Wednesday rule
tz = ZoneInfo("Asia/Jakarta")
store.init()
mon, wed, thu = (datetime(2026, 10, d, 9, 0, tzinfo=tz) for d in (26, 28, 29))
assert store.week_status(mon)["state"] == "UPCOMING" and not store.week_status(mon)["locked"]
assert store.week_status(wed)["state"] == "DUE" and store.week_status(wed)["locked"]
assert store.week_status(thu)["state"] == "OVERDUE" and store.week_status(thu)["locked"]
assert not store.week_status(wed, enforce=False)["locked"]
store.add_audit(wed, "AUDIT", "tester", res["plan_date"], s2, df2, "n")
assert store.week_status(thu)["state"] == "DONE" and not store.week_status(thu)["locked"]
assert store.get_detail(store.list_audits()[0]["id"])["summary"]["rows_total"] == s2["rows_total"]
print(store.compliance(thu, 3))
print("ALL OK")
