"""COA Request Builder - web UI (Streamlit)."""
import hmac
import io
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import audit as aud
import audit_store as store
import coa_engine as eng

TZ = ZoneInfo(os.getenv("APP_TZ", "Asia/Jakarta"))
ENFORCE = os.getenv("ENFORCE_WEDNESDAY_AUDIT", "1") == "1"   # lock Generate from Wed until audit done
OVERRIDE_PIN = os.getenv("AUDIT_OVERRIDE_PIN", "")            # supervisor PIN to skip a week (logged)
APP_PASSWORD = os.getenv("APP_PASSWORD", "")                  # simple shared password (optional)

st.set_page_config(page_title="COA Request Builder", page_icon="📦", layout="wide")
store.init()
now = datetime.now(TZ)

# ---------------- login ----------------
if APP_PASSWORD and not st.session_state.get("authed"):
    st.title("COA Request Builder")
    pw = st.text_input("Password", type="password")
    if pw:
        if hmac.compare_digest(pw, APP_PASSWORD):
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Wrong password")
    st.stop()

status = store.week_status(now, ENFORCE)


# ---------------- shared widgets ----------------
def file_inputs(prefix, with_manual=False):
    c1, c2 = st.columns(2)
    plan = c1.file_uploader("Rencana Kirim (.xlsx)", type=["xlsx"], key=prefix + "plan")
    exp = c2.file_uploader("EXPORT SAP stock (.xlsx)", type=["xlsx"], key=prefix + "exp")
    manual = None
    if with_manual:
        manual = st.file_uploader("Manual PERMINTAAN COA (.xlsx) - the batches picked by hand, same layout",
                                  type=["xlsx"], key=prefix + "manual")
    sheet = None
    if plan:
        try:
            sheets = eng.list_plan_sheets(io.BytesIO(plan.getvalue()))
            sheet = st.selectbox("Plan sheet", sheets, index=eng.guess_sheet_index(sheets, now.date()),
                                 key=prefix + "sheet")
        except Exception as e:
            st.error("Cannot read the plan file: %s" % e)

    use_date = st.checkbox("Override plan date (otherwise the date in the sheet is used)", False,
                           key=prefix + "usedate")
    pdate = st.date_input("Plan date", now.date(), format="DD/MM/YYYY",
                          key=prefix + "date", disabled=not use_date)

    with st.expander("Settings"):
        a, b = st.columns(2)
        ref = a.number_input("Shelf-life % reference = plan date minus (days)", 0, 10, 1, key=prefix + "ref")
        excl = a.checkbox("Exclude expired lots (SLED on/before plan date)", True, key=prefix + "excl")
        slocs = b.text_input("Trading storage locations (comma separated)", "2002", key=prefix + "slocs")
        mobil = b.checkbox("Write JENIS MOBIL in column K", True, key=prefix + "mobil")

    opts = dict(ref_offset_days=int(ref), exclude_expired=excl, show_mobil=mobil,
                plan_date_override=pdate if use_date else None,
                trading_slocs=tuple(s.strip() for s in slocs.split(",") if s.strip()))
    return plan, exp, manual, sheet, opts


def run_engine(plan, exp, sheet, opts):
    return eng.generate(plan.getvalue(), sheet, exp.getvalue(), **opts)


# ---------------- sidebar + banner ----------------
st.sidebar.title("📦 COA Request Builder")
modes = ["Generate COA request", "Weekly Audit", "Audit history"]
default_mode = 1 if status["locked"] else 0
mode = st.sidebar.radio("Mode", modes, index=default_mode)
st.sidebar.caption("Now: %s (Jakarta)" % now.strftime("%a %d/%m/%Y %H:%M"))

s = status["state"]
if s == "DONE":
    st.success("Weekly audit for week %d is done." % status["iso_week"])
elif s == "SKIPPED":
    st.warning("Week %d audit was skipped by a supervisor (logged)." % status["iso_week"])
elif s == "UPCOMING":
    st.info("Weekly audit is due on Wednesday. Prepare your manual picks.")
elif s == "DUE":
    st.warning("Weekly audit is due TODAY (Wednesday)." +
               (" Generate is locked until it is saved." if status["locked"] else ""))
else:
    st.error("Weekly audit for week %d is OVERDUE." % status["iso_week"] +
             (" Generate is locked until it is saved." if status["locked"] else ""))

# ======================================================================
if mode == "Generate COA request":
    st.header("Generate COA request")
    if status["locked"]:
        st.error("Locked: complete this week's Weekly Audit first.")
        st.stop()

    plan, exp, _, sheet, opts = file_inputs("gen_")
    if st.button("Generate", type="primary", disabled=not (plan and exp and sheet)):
        try:
            st.session_state["gen"] = run_engine(plan, exp, sheet, opts)
        except Exception as e:
            st.session_state.pop("gen", None)
            st.error("Failed: %s" % e)

    res = st.session_state.get("gen")
    if res:
        st_ = res["stats"]
        m1, m2, m3 = st.columns(3)
        m1.metric("Blocks", st_["blocks"])
        m2.metric("Plan lines", st_["lines"])
        m3.metric("Lines flagged CEK STOK", st_["flagged"])
        if st_["flagged"]:
            st.warning("Some lines have no / not enough unrestricted stock. Look for 'CEK STOK' in the file.")
        st.dataframe(eng.flatten(res["blocks"]), use_container_width=True, hide_index=True)
        st.download_button(
            "Download PERMINTAAN COA", res["xlsx"],
            file_name="PERMINTAAN COA %s.xlsx" % res["plan_date"].strftime("%d %B %Y").upper(),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ======================================================================
elif mode == "Weekly Audit":
    st.header("Weekly Audit - manual picks vs automated picks")
    st.caption("Pick batches by hand for a plan, then upload your sheet here. Use the same plan and the same "
               "SAP EXPORT snapshot that you used when picking manually.")

    plan, exp, manual, sheet, opts = file_inputs("aud_", with_manual=True)
    auditor = st.text_input("Auditor name")
    ready = bool(plan and exp and manual and sheet and auditor.strip())

    if st.button("Compare", type="primary", disabled=not ready):
        try:
            res = run_engine(plan, exp, sheet, opts)
            manual_blocks = aud.parse_manual(manual.getvalue())
            if not manual_blocks:
                raise ValueError("No data found in the manual file (expected the PERMINTAAN COA layout).")
            summary, df = aud.compare(eng.blocks_list(res["blocks"]), manual_blocks,
                                      res["snapshot"], res["name2kode"])
            st.session_state["audit"] = dict(summary=summary, df=df, plan_date=res["plan_date"],
                                             auditor=auditor.strip())
        except Exception as e:
            st.session_state.pop("audit", None)
            st.error("Failed: %s" % e)
    elif not ready:
        st.caption("Upload all three files and enter your name to enable Compare.")

    a = st.session_state.get("audit")
    if a:
        sm, df = a["summary"], a["df"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Row match", "%.1f%%" % sm["match_pct"])
        c2.metric("Qty agreement", "%.1f%%" % sm["qty_agreement"])
        c3.metric("Rows compared", sm["rows_total"])
        c4.metric("Blocks paired", "%d / %d" % (sm["blocks_matched"], sm["blocks_auto"]))
        st.write({k: v for k, v in sm["by_status"].items()})

        only_diff = st.checkbox("Show differences only", True)
        view = df[df["Status"] != aud.STATUS_MATCH] if only_diff else df

        def colour(row):
            return ["background-color: #%s" % aud.COLORS.get(row["Status"], "FFFFFF")] * len(row)

        st.dataframe(view.style.apply(colour, axis=1), use_container_width=True, hide_index=True)

        notes = st.text_area("Notes / reason for differences (optional)")
        meta = {"Plan date": a["plan_date"].strftime("%d/%m/%Y"), "Auditor": a["auditor"],
                "Audit time": now.strftime("%d/%m/%Y %H:%M")}
        st.download_button("Download audit report", aud.build_report(sm, df, meta),
                           file_name="AUDIT %s.xlsx" % a["plan_date"].strftime("%Y-%m-%d"),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if st.button("Save this audit (counts for this week)"):
            store.add_audit(now, "AUDIT", a["auditor"], a["plan_date"], sm, df, notes)
            st.session_state.pop("audit", None)
            st.rerun()

    if status["state"] in ("DUE", "OVERDUE") and OVERRIDE_PIN:
        with st.expander("Cannot audit this week? (supervisor only)"):
            pin = st.text_input("Supervisor PIN", type="password", key="pin")
            who = st.text_input("Supervisor name", key="who")
            why = st.text_area("Reason for skipping", key="why")
            if st.button("Skip this week's audit"):
                if hmac.compare_digest(pin, OVERRIDE_PIN) and who.strip() and why.strip():
                    store.add_audit(now, "SKIPPED", who.strip(), notes=why.strip())
                    st.rerun()
                else:
                    st.error("Needs the correct PIN, a name and a reason.")

# ======================================================================
else:
    st.header("Audit history")
    comp = pd.DataFrame(store.compliance(now, 12))
    st.subheader("Last 12 weeks")
    st.dataframe(comp, use_container_width=True, hide_index=True)
    done = (comp["Status"] == "Done").sum()
    st.caption("Completed %d of the last %d weeks." % (done, len(comp)))

    st.subheader("All audits")
    hist = store.list_audits()
    if not hist:
        st.info("No audits saved yet.")
    else:
        h = pd.DataFrame(hist)
        st.dataframe(h.drop(columns=["notes"]), use_container_width=True, hide_index=True)
        pick = st.selectbox("Open an audit", [r["id"] for r in hist],
                            format_func=lambda i: next("#%d  %s  %s  (%s)" % (r["id"], r["plan_date"] or "-", r["auditor"], r["kind"])
                                                       for r in hist if r["id"] == i))
        det = store.get_detail(pick)
        if det and det["rows"]:
            ddf = pd.DataFrame(det["rows"])
            st.dataframe(ddf, use_container_width=True, hide_index=True)
            st.download_button("Download this report", aud.build_report(det["summary"], ddf),
                               file_name="AUDIT #%d.xlsx" % pick,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        note = next(r["notes"] for r in hist if r["id"] == pick)
        if note:
            st.write("**Notes:** " + note)
