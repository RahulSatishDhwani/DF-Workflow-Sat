"""
SAT Assessment Report Converter
--------------------------------
Run locally with:
    pip install -r requirements.txt
    streamlit run app.py

Workflow:
 1. Upload a raw "Assessment Capture Report" (.xls or .xlsx).
 2. Set / load Program config (Program name, State, Month, IM names, Trainer names).
 3. Click "Convert" -> builds an editable review table.
 4. Edit any Accepted / Rejected / Reviewer / NGO / Category values that need
    correcting.
 5. Click "Generate Final Report" -> download the workflow-format .xlsx.
"""

import streamlit as st
import pandas as pd

from report_engine import (
    ProgramConfig,
    load_raw,
    build_review_table,
    aggregate_workflow,
    write_workflow_excel,
    reviewer_summary,
    list_saved_programs,
    save_program,
    load_program,
)

st.set_page_config(page_title="SAT Assessment Report Converter", layout="wide")
st.title("SAT Assessment Report Converter")

# --------------------------------------------------------------------
# Session state init
# --------------------------------------------------------------------
if "config" not in st.session_state:
    st.session_state.config = ProgramConfig()
if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "review_df" not in st.session_state:
    st.session_state.review_df = None

# --------------------------------------------------------------------
# Sidebar: Program / IM / Trainer configuration (the part that varies
# across programs & states)
# --------------------------------------------------------------------
with st.sidebar:
    st.header("Program Config")

    saved = list_saved_programs()
    choice = st.selectbox("Load saved program", ["-- new --"] + saved)
    if choice != "-- new --" and st.button("Load"):
        st.session_state.config = load_program(choice)
        st.rerun()

    cfg = st.session_state.config

    cfg.program_label = st.text_input("Program label (e.g. Saksham, Niranthara)", cfg.program_label)
    cfg.state_code = st.text_input("State / batch code (e.g. KA, TN, JH, OD, NNE_B2)", cfg.state_code)
    cfg.month_label = st.text_input("Month label (e.g. Jul, Aug)", cfg.month_label)
    cfg.assessment_name = st.text_input("Assessment Name (title shown in output)", cfg.assessment_name)
    cfg.report_date = st.text_input("Report Date (free text, e.g. '6th July, 2026')", cfg.report_date)

    im_text = st.text_area("IM names (one per line)", "\n".join(cfg.im_names), height=100)
    cfg.im_names = [n.strip() for n in im_text.split("\n") if n.strip()]

    trainer_text = st.text_area("Trainer names (one per line)", "\n".join(cfg.trainer_names), height=100)
    cfg.trainer_names = [n.strip() for n in trainer_text.split("\n") if n.strip()]

    cat_order_text = st.text_area(
        "Module/Category order (one per line, optional — leave blank to auto-detect)",
        "\n".join(cfg.category_order),
        height=100,
    )
    cfg.category_order = [c.strip() for c in cat_order_text.split("\n") if c.strip()]

    st.session_state.config = cfg

    if st.button("Save this program config"):
        path = save_program(cfg)
        st.success(f"Saved as {path.name}")

# --------------------------------------------------------------------
# Step 1: Upload + Convert
# --------------------------------------------------------------------
st.subheader("1. Upload raw Assessment Capture Report")
uploaded = st.file_uploader("Raw file (.xls or .xlsx)", type=["xls", "xlsx"])

if uploaded is not None and st.button("Convert"):
    with st.spinner("Reading raw file..."):
        raw_df = load_raw(uploaded.getvalue(), uploaded.name)
    st.session_state.raw_df = raw_df
    st.session_state.review_df = build_review_table(raw_df, st.session_state.config)
    st.success(f"Loaded {len(raw_df)} rows across {raw_df['NGO Name'].nunique()} NGOs.")

# --------------------------------------------------------------------
# Step 2: Editable review table
# --------------------------------------------------------------------
if st.session_state.review_df is not None:
    st.subheader("2. Review & edit")
    st.caption(
        "Accepted/Rejected/Reviewer flags are auto-filled best guesses. "
        "Correct any row here before generating the final report — "
        "e.g. fix a misspelled NGO name, reassign a reviewer, or flip an "
        "accept/reject call after checking the evidence."
    )

    filter_choice = st.radio(
        "Show", ["All rows", "Only rows with feedback (reviewed)", "Only rows missing a file"], horizontal=True
    )
    df = st.session_state.review_df
    if filter_choice == "Only rows with feedback (reviewed)":
        view = df[df["Feedback"].notna() & (df["Feedback"].astype(str).str.strip() != "")]
    elif filter_choice == "Only rows missing a file":
        view = df[df["File Name"].isna()]
    else:
        view = df

    edited = st.data_editor(
        view,
        num_rows="fixed",
        use_container_width=True,
        height=500,
        column_config={
            "Reviewed by IM": st.column_config.CheckboxColumn(),
            "Reviewed by Trainer": st.column_config.CheckboxColumn(),
            "Accepted": st.column_config.NumberColumn(min_value=0, max_value=1, step=1),
            "Rejected": st.column_config.NumberColumn(min_value=0, max_value=1, step=1),
        },
        key="editor",
    )

    if st.button("Apply edits back to full table"):
        st.session_state.review_df.update(edited)
        st.success("Edits applied.")

    # --------------------------------------------------------------------
    # Step 3: Generate final workflow-format report
    # --------------------------------------------------------------------
    st.subheader("3. Generate final report")
    if st.button("Generate Final Report"):
        ngo_list, categories, params_per_category, table = aggregate_workflow(
            st.session_state.review_df, st.session_state.config
        )
        excel_bytes = write_workflow_excel(
            ngo_list, categories, params_per_category, table, st.session_state.config,
            reviewer_summary_rows=reviewer_summary(st.session_state.review_df, st.session_state.config),
        )
        st.download_button(
            "Download final .xlsx",
            data=excel_bytes,
            file_name=f"{st.session_state.config.sheet_name}_Workflow_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success("Report generated.")
