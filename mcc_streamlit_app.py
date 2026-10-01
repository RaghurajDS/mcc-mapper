"""
Streamlit front-end for mcc_universal_mapper_fixed.py

Keep both files in the same folder, then run:
    streamlit run mcc_streamlit_app.py
"""
import io
from collections import OrderedDict

import pdfplumber
import streamlit as st

import mcc_universal_mapper_fixed as m

st.set_page_config(page_title="MCC Round Mapper", layout="centered")
st.title("MCC NEET-UG Round Mapper")
st.caption("Upload any MCC round PDF (Round 1, 2, 3, ...). The round is detected automatically.")

pdf_file = st.file_uploader("MCC allotment PDF", type=["pdf"])
format_file = st.file_uploader(
    "FORMAT.xlsx (optional - leave empty to use the path set in the script)", type=["xlsx"]
)
skip_pages = st.number_input("Leading cover/legend pages to skip", min_value=0, value=2, step=1)

if st.button("Convert and map", type="primary", disabled=pdf_file is None):
    # 1. Lookup workbook
    fmt_src = io.BytesIO(format_file.getvalue()) if format_file else m.FORMAT_XLSX_PATH
    try:
        lookup = m.load_format_lookup(fmt_src)
    except Exception as e:
        st.error(f"Could not load FORMAT.xlsx: {e}")
        st.stop()
    st.write(f"{len(lookup)} reference institutes loaded.")

    # 2. Parse the PDF page by page with a progress bar
    records, layout = [], None
    bar = st.progress(0.0, text="Starting...")
    with pdfplumber.open(io.BytesIO(pdf_file.getvalue())) as pdf:
        latest_label = m.detect_latest_round_label(pdf)
        total = len(pdf.pages)
        for idx in range(int(skip_pages), total):
            page = pdf.pages[idx]
            try:
                tables = page.extract_tables()
            except Exception:
                tables = []
            for t in tables:
                for row in t:
                    if row is None or len(row) < 8 or not m.is_row_start(row[0]):
                        continue
                    ncols = len(row)
                    if layout is None:
                        layout = "simple" if ncols == 8 else "dual"
                    if ncols == 8:
                        inst = row[3]
                        other = " | ".join(m.clean_cell(x) for x in row[4:] if x not in (None, ""))
                    else:
                        inst = row[-6]
                        other = " | ".join(m.clean_cell(x) for x in row[-5:] if x not in (None, ""))
                    records.append({"id": m.clean_cell(row[0]), "institute": m.clean_cell(inst), "other": other})
            page.flush_cache()
            if (idx + 1) % 10 == 0 or idx + 1 == total:
                bar.progress((idx + 1) / total, text=f"Parsing page {idx + 1}/{total}")
    layout = layout or "simple"
    bar.progress(1.0, text="Parsing done. Mapping...")

    # 3. Map and build the workbook
    out_rows, stats = m.build_output_rows(records, lookup, layout)
    if layout == "simple":
        inst_label = "Allotted Institute"
        other_label = "Course / Category / Remarks (reference only)"
        sheet_title = "Allotment Data"
    else:
        tag = latest_label or "Latest Round"
        inst_label = f"[{tag}] Allotted Institute"
        other_label = f"[{tag}] Course / Category / Remarks (reference only)"
        sheet_title = f"{tag} Allotment Data"
    headers = ["Rank", inst_label, other_label, "Code", "State", "Institute Type", "Institute Name", "Match Status"]

    unmatched = OrderedDict()
    si, ii = headers.index("Match Status"), headers.index(inst_label)
    for r in out_rows:
        if r[si] == "No Match":
            unmatched[r[ii]] = unmatched.get(r[ii], 0) + 1

    buf = io.BytesIO()
    m.write_output_workbook(headers, out_rows, stats, buf, sheet_title, unmatched)

    # 4. Results
    rate = f"{stats['matched'] / stats['applicable'] * 100:.1f}%" if stats["applicable"] else "N/A"
    c1, c2, c3 = st.columns(3)
    c1.metric("Rows extracted", f"{stats['total']:,}")
    c2.metric("Matched", f"{stats['matched']:,} / {stats['applicable']:,}")
    c3.metric("Match rate", rate)
    st.write(f"Layout: **{layout}**" + (f" (latest = {latest_label})" if layout == "dual" else ""))
    if unmatched:
        st.warning(f"{len(unmatched)} unique institutes did not match.")
        st.dataframe(
            [{"Institute": k, "Rows": v} for k, v in unmatched.items()], use_container_width=True
        )

    out_name = pdf_file.name.rsplit(".", 1)[0] + "_mapped.xlsx"
    st.download_button(
        "Download mapped Excel",
        data=buf.getvalue(),
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
