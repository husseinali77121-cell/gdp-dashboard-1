import streamlit as st
import pandas as pd
import tempfile
import os
import sys
import io
import validation_bk310_biochemistry as vb

st.set_page_config(page_title="Biobase BK-310 Validation", layout="wide")
st.title("🧪 Biobase BK‑310 Biochemistry Analyzer - Validation Suite")

# ---------- Sidebar ----------
module = st.sidebar.selectbox(
    "Select Validation Module",
    ["Quality Control (QC)", "Precision", "Accuracy", "Linearity", "Carryover"]
)

st.sidebar.markdown("---")
st.sidebar.info(
    "1. Download the empty template.\n"
    "2. Fill in your BK‑310 results.\n"
    "3. Upload the filled file.\n"
    "4. View results & download the Excel report."
)

# ---------- Module map ----------
module_map = {
    "Quality Control (QC)": {
        "generate": vb.generate_qc_template,
        "process": vb.process_qc,
        "prefix": "qc"
    },
    "Precision": {
        "generate": vb.generate_precision_template,
        "process": vb.process_precision,
        "prefix": "precision"
    },
    "Accuracy": {
        "generate": vb.generate_accuracy_template,
        "process": vb.process_accuracy,
        "prefix": "accuracy"
    },
    "Linearity": {
        "generate": vb.generate_linearity_template,
        "process": vb.process_linearity,
        "prefix": "linearity"
    },
    "Carryover": {
        "generate": vb.generate_carryover_template,
        "process": vb.process_carryover,
        "prefix": "carryover"
    },
}

# ---------- Temp directory ----------
if "temp_dir" not in st.session_state:
    st.session_state.temp_dir = tempfile.TemporaryDirectory()
temp_dir = st.session_state.temp_dir.name

# ---------- Main interface ----------
st.header(module)
mod = module_map[module]

# 1. Generate template
st.subheader("1. Download Empty Template")
# For precision we need to pass replicates optionally, but we can use a default of 10.
if st.button("Generate Template"):
    # Redirect output folders to temp
    old_out = vb.OUTPUT_DIR
    old_tpl = vb.TEMPLATE_DIR
    vb.OUTPUT_DIR = temp_dir
    vb.TEMPLATE_DIR = temp_dir

    # Precision needs the replicates argument; others ignore it
    if module == "Precision":
        mod["generate"](num_replicates=10)  # you can make this a number input if desired
    else:
        mod["generate"]()

    vb.OUTPUT_DIR = old_out
    vb.TEMPLATE_DIR = old_tpl

    template_file = os.path.join(temp_dir, f"{mod['prefix']}_template.xlsx")
    if os.path.exists(template_file):
        with open(template_file, "rb") as f:
            st.download_button("📥 Download Template", f,
                               file_name=os.path.basename(template_file),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Template generation failed. Check the console for errors.")

st.markdown("---")

# 2. Upload filled template
st.subheader("2. Upload Your Filled Template")
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Save uploaded file
    upload_path = os.path.join(temp_dir, uploaded_file.name)
    with open(upload_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success("File uploaded ✓")

    # 3. Process and show results
    st.subheader("3. Validation Results")
    try:
        # Redirect stdout to capture prints (optional)
        old_stdout = sys.stdout
        sys.stdout = io.StringIO()

        # Redirect output folder so reports land in temp
        old_out = vb.OUTPUT_DIR
        vb.OUTPUT_DIR = temp_dir

        # Process the file (most functions return a DataFrame)
        result_df = mod["process"](upload_path)

        # Restore stdout and output folder
        sys.stdout = old_stdout
        vb.OUTPUT_DIR = old_out

        # Show summary DataFrame
        if result_df is not None and not result_df.empty:
            st.dataframe(result_df, use_container_width=True)
        else:
            st.warning("No summary table returned, but report may have been generated.")

        # Find the Excel report file
        report_name_map = {
            "Quality Control (QC)": "QC_Report.xlsx",
            "Precision": "Precision_Report.xlsx",
            "Accuracy": "Accuracy_Report.xlsx",
            "Linearity": "Linearity_Report.xlsx",
            "Carryover": "Carryover_Report.xlsx",
        }
        report_file = os.path.join(temp_dir, report_name_map[module])

        if os.path.exists(report_file):
            st.success("✅ Report generated successfully!")
            with open(report_file, "rb") as f:
                st.download_button("📥 Download Full Excel Report", f,
                                   file_name=report_name_map[module],
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.info("📊 The Excel file contains embedded charts. Open it to view them.")
        else:
            st.error("Report file not found. Please check the console output above.")

    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
        sys.stdout = old_stdout
