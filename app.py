"""
Streamlit application to compare commercial opportunities between two quarters.
Auteur : Sofyan BENANIBA
Launch: streamlit run app.py
"""
import base64
import io
import pandas as pd
import streamlit as st

from comparison import compare_dataframes, KEY_COL, normalize_key, values_differ
from excel_export import generate_excel, generate_excel_by_ou
from word_export import generate_word, generate_word_by_ou

# ──────────────────────────────────────────────
# Page configuration
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Opportunity Comparator",
    page_icon="📈",
    layout="wide",
)

# ──────────────────────────────────────────────
# Custom CSS
# ──────────────────────────────────────────────
st.markdown("""
<style>
    .block-container {
        padding-top: 1rem;
    }
    .kpi-card {
        background: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        text-align: center;
        border-left: 5px solid;
        margin-bottom: 10px;
    }
    .kpi-new   { border-color: #FFD700; background: #FFFDE7; }
    .kpi-del   { border-color: #EF5350; background: #FFEBEE; }
    .kpi-mod   { border-color: #FF9800; background: #FFF3E0; }
    .kpi-num   { font-size: 2.5rem; font-weight: bold; margin: 0; }
    .kpi-label { font-size: 0.9rem; color: #555; margin: 0; }
    .status-badge-new  { background:#FFD700; color:#333; padding:2px 8px; border-radius:10px; font-size:0.8rem; }
    .status-badge-del  { background:#EF5350; color:#fff; padding:2px 8px; border-radius:10px; font-size:0.8rem; }
    .status-badge-mod  { background:#FF9800; color:#fff; padding:2px 8px; border-radius:10px; font-size:0.8rem; }
    .status-badge-unch { background:#90CAF9; color:#333; padding:2px 8px; border-radius:10px; font-size:0.8rem; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# Header (logo + title)
# ──────────────────────────────────────────────
_logo_b64 = base64.b64encode(open("Thales Logo.png", "rb").read()).decode()
st.markdown(
    f'<img src="data:image/png;base64,{_logo_b64}" width="420">',
    unsafe_allow_html=True,
)

st.markdown(
    '<h1 style="color:#242B75; font-size:25pt;">📈 Commercial Opportunity Comparator</h1>',
    unsafe_allow_html=True,
)
st.markdown("Compare your Excel files between two quarters and export the report (MOIR, SER, Order In Take, Extract T360...).")

# ──────────────────────────────────────────────
# File upload
# ──────────────────────────────────────────────
col_upload_l, col_upload_r = st.columns(2)

with col_upload_l:
    st.markdown(
        '<h3 style="color:#5DBFD5;">📁 Previous Quarter File (N-1)</h3>',
        unsafe_allow_html=True,
    )
    file_prev = st.file_uploader("Select the N-1 file", type=["xlsx", "xls"], key="prev")

with col_upload_r:
    st.markdown(
        '<h3 style="color:#5DBFD5;">📁 Current Quarter File</h3>',
        unsafe_allow_html=True,
    )
    file_curr = st.file_uploader("Select the current file", type=["xlsx", "xls"], key="curr")


def load_excel(file) -> pd.DataFrame | None:
    """Load an Excel file and return a DataFrame, or None on error."""
    try:
        df = pd.read_excel(file, engine="openpyxl")
        if df.empty:
            st.error(f"The file **{file.name}** is empty.")
            return None
        return df
    except Exception as e:
        st.error(f"Cannot read the file **{file.name}**: {e}")
        return None


# ──────────────────────────────────────────────
# Main processing
# ──────────────────────────────────────────────
if file_prev and file_curr:
    df_prev = load_excel(file_prev)
    df_curr = load_excel(file_curr)

    if df_prev is not None and df_curr is not None:

        # Clean column names
        df_prev.columns = df_prev.columns.str.strip()
        df_curr.columns = df_curr.columns.str.strip()

        # ── KIP column propagation ──
        # If KIP exists in N-1 but not in current, copy it over by matching Opportunity Name.
        KIP_COL = "KIP"
        if KIP_COL in df_prev.columns and KIP_COL not in df_curr.columns:
            prev_kip_map = {
                str(name).strip().lower(): val
                for name, val in zip(df_prev[KEY_COL], df_prev[KIP_COL])
            }
            df_curr[KIP_COL] = df_curr[KEY_COL].apply(
                lambda name: prev_kip_map.get(str(name).strip().lower(), "NEW so TO DEFINE")
            )
            st.info(
                f"Column **'{KIP_COL}'** was missing from the current file — "
                "it has been automatically filled from the N-1 file. "
                "Unmatched opportunities are set to **\"NEW so TO DEFINE\"**."
            )

        # Column validation
        cols_prev = set(df_prev.columns)
        cols_curr = set(df_curr.columns)

        if cols_prev != cols_curr:
            missing_in_curr = cols_prev - cols_curr
            missing_in_prev = cols_curr - cols_prev
            msg = "The two files do not have the same columns.\n\n"
            if missing_in_curr:
                msg += f"- Columns in N-1 but missing in current: **{', '.join(missing_in_curr)}**\n"
            if missing_in_prev:
                msg += f"- Columns in current but missing in N-1: **{', '.join(missing_in_prev)}**\n"
            st.error(msg)
            st.stop()

        if KEY_COL not in df_prev.columns:
            st.error(f"Key column **'{KEY_COL}'** not found in the files.")
            st.stop()

        # Detect GBU/BL column
        gbu_col = None
        for col in df_curr.columns:
            if "gbu" in col.lower() or "bl" in col.lower():
                gbu_col = col
                break

        # Detect Operating Unit / Legal Entity column
        ou_col_default = None
        for col in df_curr.columns:
            if any(k in col.lower() for k in ("operating", "legal", "entity", "ou/le", "ou ", "le ")):
                ou_col_default = col
                break

        # ── Comparison ──
        with st.spinner("Running comparison..."):
            result = compare_dataframes(df_prev, df_curr)

        n_new  = len(result["new"])
        n_del  = len(result["deleted"])
        n_mod  = len(result["modified"])
        n_unch = len(result["unchanged"])

        # ─────────────────────────────────────
        # KPIs
        # ─────────────────────────────────────
        st.markdown("---")
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)

        with kpi1:
            st.markdown(f"""
            <div class="kpi-card kpi-new">
                <p class="kpi-num">{n_new}</p>
                <p class="kpi-label">✨ New opportunities</p>
            </div>""", unsafe_allow_html=True)

        with kpi2:
            st.markdown(f"""
            <div class="kpi-card kpi-del">
                <p class="kpi-num">{n_del}</p>
                <p class="kpi-label">🗑️ Removed opportunities</p>
            </div>""", unsafe_allow_html=True)

        with kpi3:
            st.markdown(f"""
            <div class="kpi-card kpi-mod">
                <p class="kpi-num">{n_mod}</p>
                <p class="kpi-label">✏️ Modified opportunities</p>
            </div>""", unsafe_allow_html=True)

        with kpi4:
            st.markdown(f"""
            <div class="kpi-card" style="border-color:#42A5F5;background:#E3F2FD;">
                <p class="kpi-num">{n_unch}</p>
                <p class="kpi-label">✅ Unchanged</p>
            </div>""", unsafe_allow_html=True)

        # ─────────────────────────────────────
        # Filters
        # ─────────────────────────────────────
        st.markdown("---")
        st.markdown(
            '<h3 style="color:#5DBFD5;">🔍 Detail Table</h3>',
            unsafe_allow_html=True,
        )

        search_query = st.text_input("🔎 Search by Opportunity Name", placeholder="Type to filter...")

        filter_col1, filter_col2 = st.columns([1, 2])

        with filter_col1:
            status_filter = st.selectbox(
                "Filter by status",
                ["All", "New", "Modified", "Removed", "Unchanged"],
            )

        # Collect available GBU/BL values
        gbu_options = ["All"]
        if gbu_col:
            all_gbu = set()
            for df in [result["new"], result["deleted"], result["modified"], result["unchanged"]]:
                if not df.empty and gbu_col in df.columns:
                    all_gbu.update(
                        str(v).strip() if str(v).strip() not in ("", "nan", "NaT", "None", "<NA>")
                        else "(No GBU/BL)"
                        for v in df[gbu_col]
                    )
            gbu_options += sorted(all_gbu)

        with filter_col2:
            gbu_filter = st.selectbox("Filter by GBU/BL", gbu_options)

        # ─────────────────────────────────────
        # Build display table
        # ─────────────────────────────────────
        status_map = {
            "All":       ["new", "modified", "deleted", "unchanged"],
            "New":       ["new"],
            "Modified":  ["modified"],
            "Removed":   ["deleted"],
            "Unchanged": ["unchanged"],
        }
        active_statuses = status_map[status_filter]

        frames_display = []
        for status in active_statuses:
            df = result[status]
            if not df.empty:
                frames_display.append(df)

        if not frames_display:
            st.info("No data to display for this filter.")
        else:
            df_display = pd.concat(frames_display, ignore_index=True)

            # Search filter (Opportunity Name)
            if search_query.strip():
                df_display = df_display[
                    df_display[KEY_COL].astype(str).str.contains(search_query.strip(), case=False, na=False)
                ]

            # GBU filter
            if gbu_col and gbu_filter != "All":
                df_display = df_display[
                    df_display[gbu_col].apply(
                        lambda v: "(No GBU/BL)" if (v is None or str(v).strip() in ("", "nan", "NaT", "None", "<NA>"))
                        else str(v).strip()
                    ) == gbu_filter
                ]

            if df_display.empty:
                st.info("No data for this filter.")
            else:
                columns = result["columns"]
                diff_map = result["diff_map"]
                prev_indexed = result["prev_indexed"]

                # ── HTML table rendering ──
                def build_html_table(df: pd.DataFrame) -> str:
                    th_style = (
                        "background:#244061;color:white;padding:6px 10px;"
                        "text-align:left;border:1px solid #ccc;white-space:nowrap;"
                        "font-size:12px;"
                    )
                    headers = "<tr>" + "".join(
                        f"<th style='{th_style}'>{col}</th>" for col in ["Status"] + columns
                    ) + "</tr>"

                    rows_html = ""
                    for _, row in df.iterrows():
                        status = row.get("__status__", "unchanged")
                        norm_key = str(row.get(KEY_COL, "")).strip().lower()
                        changed_cols = diff_map.get(norm_key, set())

                        # Status badge
                        if status == "new":
                            badge = "<span class='status-badge-new'>New</span>"
                        elif status == "deleted":
                            badge = "<span class='status-badge-del'>Removed</span>"
                        elif status == "modified":
                            badge = "<span class='status-badge-mod'>Modified</span>"
                        else:
                            badge = "<span class='status-badge-unch'>Unchanged</span>"

                        row_style = "background:#FFFDE7;" if status == "new" else ""
                        cells_html = f"<td style='border:1px solid #ccc;padding:4px 8px;{row_style}'>{badge}</td>"

                        for col in columns:
                            val = row.get(col, "")
                            val_str = "" if pd.isna(val) else str(val)
                            cell_style = "border:1px solid #ccc;padding:4px 8px;font-size:12px;"

                            if status == "deleted":
                                cell_style += "text-decoration:line-through;background:#FFEBEE;color:#999;"
                                cells_html += f"<td style='{cell_style}'>{val_str}</td>"

                            elif status == "new":
                                cell_style += "background:#FFFDE7;"
                                cells_html += f"<td style='{cell_style}'>{val_str}</td>"

                            elif status == "modified" and col in changed_cols:
                                old_val = prev_indexed.loc[norm_key, col] if col in prev_indexed.columns else ""
                                old_str = "" if pd.isna(old_val) else str(old_val)
                                cell_style += "background:#FFF3E0;"
                                inner = ""
                                if old_str:
                                    inner += f"<span style='text-decoration:line-through;color:#999;font-size:11px;'>{old_str}</span> → "
                                inner += f"<strong>{val_str}</strong>"
                                cells_html += f"<td style='{cell_style}'>{inner}</td>"

                            else:
                                cells_html += f"<td style='{cell_style}'>{val_str}</td>"

                        rows_html += f"<tr>{cells_html}</tr>"

                    return f"""
                    <div style="overflow-x:auto;max-height:600px;overflow-y:auto;">
                    <table style="border-collapse:collapse;width:100%;font-size:12px;">
                        <thead style="position:sticky;top:0;z-index:1;">{headers}</thead>
                        <tbody>{rows_html}</tbody>
                    </table>
                    </div>
                    """

                st.markdown(build_html_table(df_display), unsafe_allow_html=True)
                st.caption(f"Showing {len(df_display)} row(s).")

        # ─────────────────────────────────────
        # Exports
        # ─────────────────────────────────────
        st.markdown("---")
        st.markdown(
            '<h3 style="color:#5DBFD5;">📥 Download Report</h3>',
            unsafe_allow_html=True,
        )

        # ── Export by GBU/BL ──
        st.markdown("**By GBU/BL**")
        exp_col1, exp_col2 = st.columns(2)

        with exp_col1:
            if st.button("🟢 Generate Excel by GBU/BL (.xlsx)"):
                with st.spinner("Generating Excel..."):
                    xlsx_bytes = generate_excel(result, gbu_col)
                st.download_button(
                    label="⬇️ Download Excel (GBU/BL)",
                    data=xlsx_bytes,
                    file_name="opportunity_report_gbu.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        with exp_col2:
            if st.button("🔵 Generate Word by GBU/BL (.docx)"):
                with st.spinner("Generating Word..."):
                    docx_bytes = generate_word(result, gbu_col)
                st.download_button(
                    label="⬇️ Download Word (GBU/BL)",
                    data=docx_bytes,
                    file_name="opportunity_report_gbu.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        # ── Export by Operating Unit / Legal Entity ──
        st.markdown("**By Operating Unit / Legal Entity**")

        ou_col = ou_col_default

        exp_col3, exp_col4 = st.columns(2)

        with exp_col3:
            if st.button("🟢 Generate Excel by OU/LE (.xlsx)"):
                with st.spinner("Generating Excel..."):
                    xlsx_bytes = generate_excel_by_ou(result, ou_col, gbu_col)
                st.download_button(
                    label="⬇️ Download Excel (OU/LE)",
                    data=xlsx_bytes,
                    file_name="opportunity_report_ou.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        with exp_col4:
            if st.button("🔵 Generate Word by OU/LE (.docx)"):
                with st.spinner("Generating Word..."):
                    docx_bytes = generate_word_by_ou(result, ou_col, gbu_col)
                st.download_button(
                    label="⬇️ Download Word (OU/LE)",
                    data=docx_bytes,
                    file_name="opportunity_report_ou.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

else:
    st.info("Please upload both Excel files to start the comparison.")

    with st.expander("📋 Expected file format"):
        st.markdown("""
        Both files must have **exactly the same columns** in both quarters.

        The **`Opportunity Name`** column is used as the unique identifier.

        Example structure:
        | Opportunity Name | GBU/BL | Status | Value | ... |
        |---|---|---|---|---|
        | Project Alpha | BL1 | In Progress | 50,000 | ... |
        | Project Beta | BL2 | Won | 120,000 | ... |
        """)

st.markdown("---")
st.markdown(
    '<p style="text-align:right; color:#888; font-size:0.78rem; line-height:1.5; margin-bottom:2px;">'
    '🔒 This application is hosted locally on your machine. No data leaves your workstation, '
    'meeting Thales C3 security level requirements.'
    '</p>',
    unsafe_allow_html=True,
)
st.markdown(
    '<p style="text-align:right; color:#666; font-size:0.85rem;">'
    'For any questions or enhancement requests, please contact sofyan.benaniba@thalesgroup.com'
    '</p>',
    unsafe_allow_html=True,
)
