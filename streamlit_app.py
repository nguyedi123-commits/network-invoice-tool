import hmac

import streamlit as st

from invoice_run import process_files

st.set_page_config(page_title="Network Invoice Tool", page_icon=":material/receipt_long:", layout="wide")

st.title("Network invoice tool")

# ---- password gate: the password lives in the app's Secrets (app_password = "..."), never in the code
if not st.session_state.get("signed_in"):
    try:
        expected = st.secrets.get("app_password", "")
    except Exception:  # no secrets set up at all
        expected = ""
    if not expected:
        st.error("The app password has not been set up yet (Settings > Secrets).", title="App locked")
        st.stop()
    with st.form("sign_in"):
        pw = st.text_input("Password", type="password", key="pw")
        ok = st.form_submit_button("Open", icon=":material/lock_open:", type="primary")
    if ok and hmac.compare_digest(pw.encode(), str(expected).encode()):
        st.session_state["signed_in"] = True
        st.rerun()
    if ok:
        st.error("Wrong password.")
    st.stop()

st.caption("Upload the network invoice files (NBCHARGES / NBREMIND). You get one zip with the Excel extract, "
           "the remit files and Network Lines.csv for Snowflake. Nothing is kept on the server.")

with st.form("run"):
    month = st.text_input("Month", placeholder="e.g. Oct", key="month",
                          help="Used in the file names, e.g. \"Oct Network Invoice.xlsx\".")
    uploads = st.file_uploader("Invoice files", accept_multiple_files=True, key="files",
                               help="All the .csv files from the networks for this run. Remit files are skipped.")
    go = st.form_submit_button("Create extract and remits", icon=":material/play_arrow:", type="primary")

if go:
    if not uploads:
        st.warning("Add at least one invoice file.")
    else:
        with st.spinner("Reading the files..."):
            st.session_state["result"] = process_files(month, [(f.name, f.getvalue()) for f in uploads])

res = st.session_state.get("result")
if res:
    rows = res["results"]
    with st.container(horizontal=True):
        st.metric("Files processed", len(rows), border=True)
        st.metric("Files skipped", len(res["skipped"]), border=True)
        st.metric("Total ex GST", f"${sum(r['Ex GST'] for r in rows):,.2f}", border=True)
        st.metric("Remits pay (inc GST)", f"${sum(r['Remit pays'] for r in rows):,.2f}", border=True)
        st.metric("Network lines", f"{res['lines']:,}", border=True)

    st.download_button(f"Download {res['zip_name']}", res["zip_bytes"], file_name=res["zip_name"],
                       mime="application/zip", icon=":material/download:", type="primary")

    if res["skipped"]:
        st.warning(f"{len(res['skipped'])} file(s) not processed", title="Check these files")
        st.dataframe(res["skipped"], hide_index=True, alt="Files not processed and why")

    if rows:
        money = st.column_config.NumberColumn(format="dollar")
        st.subheader("Files processed")
        st.dataframe(rows, hide_index=True, alt="One row per invoice file processed",
                     column_config={"Ex GST": money, "GST": money, "Inc GST": money,
                                    "GSL credit": money, "Remit pays": money})
