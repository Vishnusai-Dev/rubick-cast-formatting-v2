"""
Rubick CAST Formatting - Streamlit Application

Two modes:
  1. Input Generator - builds the Types + Values input workbook from a Style file
     (required) and an optional SKU / size chart file.
  2. Export Format   - reindexes one uploaded file to the bundled Rubick OS template.
"""

from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from src.config import get_settings
from src.core.export_processor import process_export
from src.core.import_processor import process_import
from src.utils.logging import get_logger

logger = get_logger(__name__)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# --------------------------------------------------------------------------- #
# Shared helpers
# --------------------------------------------------------------------------- #
def setup_page():
    settings = get_settings()
    st.set_page_config(
        page_title=settings.app_name,
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed",
    )


def render_logo():
    logo_path = Path(__file__).parent / "logo.svg"
    if logo_path.exists():
        st.image(str(logo_path), width=120)


def header_with_back(title: str, subtitle: str):
    col1, col2 = st.columns([1, 6], vertical_alignment="center")
    with col1:
        if st.button("← Back", width="stretch"):
            st.session_state.current_mode = None
            st.rerun()
    with col2:
        st.markdown(
            f'<div class="page-title"><h1>{title}</h1><p>{subtitle}</p></div>',
            unsafe_allow_html=True,
        )
    st.write("")


def step(n: int, label: str, hint: str = ""):
    hint_html = f'<span class="step-hint">{hint}</span>' if hint else ""
    st.markdown(
        f'<div class="step"><span class="step-n">{n}</span>'
        f'<span class="step-label">{label}</span>{hint_html}</div>',
        unsafe_allow_html=True,
    )


def is_valid_filename(name: str) -> bool:
    return bool(name) and name.lower().endswith((".xlsx", ".xls"))


def to_xlsx_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    return buf.getvalue()


def read_headers(data: bytes) -> list[str]:
    """Read the header row of the first sheet only."""
    return [str(c) for c in pd.read_excel(BytesIO(data), nrows=0).columns]


# --------------------------------------------------------------------------- #
# Example files shown in the "how should the input be" guides
# (illustrative placeholder values - not real client data)
# --------------------------------------------------------------------------- #
EXAMPLE_STYLE = pd.DataFrame(
    {
        "styleId": ["1001", "1002"],
        "productDisplayName": ["Example Product A", "Example Product B"],
        "Front Image": ["https://example.com/a_front.jpg", "https://example.com/b_front.jpg"],
    }
)

EXAMPLE_SKU = pd.DataFrame(
    {
        "styleId": ["1001", "1001", "1002", "1002"],
        "Brand Size": ["S", "M", "S", "M"],
        "Chest (Inches)": ["36", "38", "36", "38"],
        "Length (Inches)": ["26", "27", "26", "27"],
    }
)

EXAMPLE_EXPORT_INPUT = pd.DataFrame(
    {
        "styleId": ["1001", "1002"],
        "productDisplayName": ["Example Product A", "Example Product B"],
        "Front Image": ["https://example.com/a_front.jpg", "https://example.com/b_front.jpg"],
        "AI Generated Image Flag": ["Yes", "No"],
    }
)


@st.cache_data
def example_style_bytes() -> bytes:
    return to_xlsx_bytes({"Sheet1": EXAMPLE_STYLE})


@st.cache_data
def example_sku_bytes() -> bytes:
    return to_xlsx_bytes({"Sheet1": EXAMPLE_SKU})


@st.cache_data
def example_export_input_bytes() -> bytes:
    return to_xlsx_bytes({"Sheet1": EXAMPLE_EXPORT_INPUT})



CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"], .stApp { font-family: 'Inter', -apple-system, 'Segoe UI', sans-serif; }
.block-container { max-width: 1240px; padding-top: 2.2rem; padding-bottom: 3rem; }
header[data-testid="stHeader"] { background: transparent; }
#MainMenu, footer { visibility: hidden; }

/* Hero (landing) */
.hero { text-align: center; padding: 0.5rem 0 1.6rem 0; }
.hero h1 {
  font-size: 2.5rem; font-weight: 700; letter-spacing: -0.02em; margin: 0.4rem 0 0.2rem 0;
  background: linear-gradient(90deg, #4f46e5, #0ea5e9);
  -webkit-background-clip: text; background-clip: text; color: transparent;
}
.hero p { opacity: 0.7; margin: 0; font-size: 1.02rem; }
.hero .ver { display: inline-block; margin-top: 0.7rem; padding: 2px 10px; border-radius: 999px;
  font-size: 0.75rem; border: 1px solid rgba(99,102,241,.35); color: #6366f1; }

/* Page title */
.page-title h1 { font-size: 1.7rem; font-weight: 700; letter-spacing: -0.01em; margin: 0; padding: 0; }
.page-title p { opacity: 0.7; margin: 0.15rem 0 0 0; font-size: 0.95rem; }

/* Steps */
.step { display: flex; align-items: center; gap: 0.6rem; margin: 1.1rem 0 0.5rem 0; }
.step-n { display: inline-flex; align-items: center; justify-content: center; width: 26px; height: 26px;
  border-radius: 50%; background: #4f46e5; color: #fff; font-size: 0.8rem; font-weight: 600; flex: none; }
.step-label { font-weight: 600; font-size: 1.05rem; }
.step-hint { opacity: 0.55; font-size: 0.85rem; margin-left: 0.2rem; }

/* Mode cards (landing) */
[class*="st-key-mode_"] {
  border: 1px solid rgba(128,128,128,.25); border-radius: 16px; padding: 1.3rem 1.4rem;
  transition: all .2s ease; background: rgba(128,128,128,.04);
}
[class*="st-key-mode_"]:hover { border-color: #6366f1; box-shadow: 0 6px 22px rgba(79,70,229,.14); transform: translateY(-2px); }
.st-key-mode_input { border-left: 5px solid #6366f1; }
.st-key-mode_export { border-left: 5px solid #10b981; }
.mode-card h3 { margin: 0 0 .25rem 0; font-size: 1.25rem; font-weight: 650; }
.mode-card p { opacity: .72; margin: 0 0 .5rem 0; font-size: .92rem; }
.mode-card ul { margin: 0 0 .6rem 0; padding-left: 1.1rem; opacity: .8; font-size: .88rem; }

/* Read me panel (right side, all pages) */
.st-key-readme {
  border: 1px solid rgba(99,102,241,.28); background: rgba(99,102,241,.06);
  border-radius: 16px; padding: 1.2rem 1.3rem;
}
.readme-title { font-weight: 700; font-size: 1.05rem; margin-bottom: .6rem; }
.st-key-readme h4 { font-size: .95rem; margin: 1rem 0 .3rem 0; font-weight: 650; }
.st-key-readme p, .st-key-readme li { font-size: .88rem; }

/* Controls */
div.stButton > button, div.stDownloadButton > button { border-radius: 10px; font-weight: 600; }
[data-testid="stFileUploaderDropzone"] { border-radius: 12px; border: 1.5px dashed rgba(99,102,241,.5); }
[data-testid="stMetric"] { background: rgba(128,128,128,.07); border-radius: 12px; padding: .6rem .8rem; }
[data-testid="stMetricLabel"] p { font-size: .78rem; opacity: .7; }
[data-testid="stMetricValue"] { font-size: 1.5rem; }
</style>
"""


def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)


def readme_panel():
    """Right-hand 'Read me' container, used on every page."""
    return st.container(key="readme")


# --------------------------------------------------------------------------- #
# Landing page
# --------------------------------------------------------------------------- #
def render_landing_page():
    settings = get_settings()

    st.markdown(
        f'<div class="hero"><h1>{settings.app_name}</h1>'
        f"<p>Generate input files and format outputs for the Rubick OS template.</p>"
        f'<span class="ver">v{settings.app_version}</span></div>',
        unsafe_allow_html=True,
    )

    main, side = st.columns([3, 2], gap="large")

    with main:
        with st.container(key="mode_input"):
            st.markdown(
                """
                <div class="mode-card">
                    <h3>🧩 Input Generator</h3>
                    <p>Build the input file (Types + Values) from your Style file and, optionally, a SKU size chart.</p>
                    <ul>
                        <li>Style file required, SKU optional</li>
                        <li>Types sheet generated automatically</li>
                    </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button("Open Input Generator", type="primary", width="stretch", key="input_gen"):
                st.session_state.current_mode = "input_generator"
                st.rerun()

        st.write("")

        with st.container(key="mode_export"):
            st.markdown(
                """
                <div class="mode-card">
                    <h3>📤 Export Format</h3>
                    <p>Reindex your data to the Rubick OS template in one step.</p>
                    <ul>
                        <li>Upload one data file</li>
                        <li>Template applied automatically</li>
                    </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button("Open Export Format", type="primary", width="stretch", key="export"):
                st.session_state.current_mode = "export"
                st.rerun()

    with side:
        with readme_panel():
            st.markdown('<div class="readme-title">📘 Read me</div>', unsafe_allow_html=True)
            st.markdown(
                """
                #### Typical flow
                1. **Input Generator**: combine the Style file (and SKU file, if you have one) into one input workbook.
                2. **Export Format**: upload your data file and download it reindexed to the Rubick OS template.

                #### Good to know
                - Files must be `.xlsx` or `.xls`, up to 50 MB.
                - Each page has its own read me on this side with the exact input format and example files.
                - Nothing is stored; files are processed in memory for your session.
                """
            )


# --------------------------------------------------------------------------- #
# Input Generator
# --------------------------------------------------------------------------- #
def render_input_readme():
    with readme_panel():
        st.markdown('<div class="readme-title">📘 Read me: how should the input be?</div>', unsafe_allow_html=True)
        st.markdown(
            """
            Two files, **one tab per category**. Tab names must be identical in both files.
            Tabs listed in *Exclude Sheets* (default `masterdata`) are skipped.

            | File | Required | Holds |
            |---|---|---|
            | **Style** (`style.xlsx`) | Yes | One row per style, product attributes |
            | **SKU** (`sku.xlsx`) | No | One row per style + size, size chart measurements |
            """
        )
        st.markdown("#### Style file")
        st.dataframe(EXAMPLE_STYLE, hide_index=True, width="stretch")
        st.markdown("#### SKU file (optional)")
        st.dataframe(EXAMPLE_SKU, hide_index=True, width="stretch")
        st.markdown(
            """
            #### Rules
            - Header row in row 1 of every tab.
            - Style identifier column named `styleId`, `style_id` or `SKU`.
            - SKU file: size columns start at **Brand Size** and run to the right. Rows are combined per style into comma-separated values (e.g. `S,M`).
            - No SKU file: no size columns are added.
            - Characters other than letters, numbers, `( ) _ - #` and spaces are removed from headers.

            #### Output
            One workbook with a `Types` sheet and a `Values` sheet.
            """
        )
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇️ Example style",
                data=example_style_bytes(),
                file_name="example_style.xlsx",
                mime=XLSX_MIME,
                width="stretch",
            )
        with c2:
            st.download_button(
                "⬇️ Example SKU",
                data=example_sku_bytes(),
                file_name="example_sku.xlsx",
                mime=XLSX_MIME,
                width="stretch",
            )
        st.caption("Example files use placeholder values only.")


def render_input_generator_page():
    settings = get_settings()
    header_with_back(
        "🧩 Input Generator",
        "Build the input file from your Style file and, optionally, a SKU size chart.",
    )

    main, side = st.columns([3, 2], gap="large")
    with side:
        render_input_readme()

    with main:
        step(1, "Upload files")
        col1, col2 = st.columns(2)
        with col1:
            product_details_file = st.file_uploader(
                "Product Details File (style.xlsx)",
                type=["xlsx", "xls"],
                key="input_product",
            )
        with col2:
            size_chart_file = st.file_uploader(
                "Size Chart File (sku.xlsx), optional",
                type=["xlsx", "xls"],
                key="input_size",
                help="Leave empty to generate the input from the Style file alone.",
            )

        step(2, "Options")
        output_filename = st.text_input(
            "Output Filename",
            value=settings.input_generator_output,
            help="Must end with .xlsx or .xls",
        )
        if output_filename and not is_valid_filename(output_filename):
            st.error("Output filename must end with .xlsx or .xls")
            output_filename = settings.input_generator_output

        exclude_sheets = st.text_input(
            "Exclude Sheets (comma-separated)",
            value=settings.default_exclude_sheets,
            help="Enter sheet names to exclude, separated by commas",
        )
        exclude_list = []
        if exclude_sheets.strip():
            exclude_list = [s.strip() for s in exclude_sheets.split(",") if s.strip()]
            invalid = [s for s in exclude_list if any(c in s for c in r"[]/?*:;{}")]
            if invalid:
                st.error(f"Invalid sheet names (cannot contain special characters): {invalid}")
                st.stop()

        max_size = settings.max_file_size_mb * 1024 * 1024
        if product_details_file and product_details_file.size > max_size:
            st.error(f"Style file exceeds maximum size ({settings.max_file_size_mb}MB)")
            product_details_file = None
        if size_chart_file and size_chart_file.size > max_size:
            st.error(f"SKU file exceeds maximum size ({settings.max_file_size_mb}MB)")
            size_chart_file = None

        step(3, "Generate")
        if not product_details_file:
            st.info("Upload the Style file to continue.")
        else:
            if not size_chart_file:
                st.info("No SKU file uploaded: size columns will not be added.")
            if st.button("Generate Input", type="primary"):
                result = process_import(
                    size_chart_data=size_chart_file.getvalue() if size_chart_file else None,
                    size_chart_filename=size_chart_file.name if size_chart_file else None,
                    product_details_data=product_details_file.getvalue(),
                    product_details_filename=product_details_file.name,
                    output_filename=output_filename,
                    exclude_sheets=exclude_list,
                )

                if result.success:
                    st.success("Input generated successfully!")
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Rows", result.rows_processed)
                    m2.metric("Columns", result.columns_count)
                    m3.metric("Sheets", result.sheets_processed)
                    st.download_button(
                        label="⬇️ Download Generated Input",
                        data=result.data,
                        file_name=result.filename,
                        mime=XLSX_MIME,
                    )
                else:
                    st.error(f"Error: {result.error_message}")


# --------------------------------------------------------------------------- #
# Export Format
# --------------------------------------------------------------------------- #
def render_export_readme():
    with readme_panel():
        st.markdown('<div class="readme-title">📘 Read me: how should the input be?</div>', unsafe_allow_html=True)
        st.markdown(
            """
            **One file:** your data file. It is reindexed to the Rubick OS template automatically.

            #### Rules
            - Header row in row 1; only the **first sheet** is read.
            - Headers must match the template **exactly** (spelling, spaces, capitalisation).
            - Matched columns follow template order.
            - Template columns missing from your file are added blank.
            - Columns not in the template are dropped, unless *Preserve unknown columns* is ticked (they are then added after the template columns).

            #### Example data file
            Header names must follow the template; this sample is illustrative.
            """
        )
        st.dataframe(EXAMPLE_EXPORT_INPUT, hide_index=True, width="stretch")
        st.download_button(
            "⬇️ Example data file",
            data=example_export_input_bytes(),
            file_name="example_export_input.xlsx",
            mime=XLSX_MIME,
            width="stretch",
        )


def render_column_match(input_bytes: bytes, template_bytes: bytes, preserve_unknown: bool):
    """Show how the uploaded data file lines up with the template headers."""
    try:
        in_cols = read_headers(input_bytes)
        tpl_cols = read_headers(template_bytes)
    except Exception as e:
        st.warning(f"Could not preview column match: {e}")
        return

    in_set, tpl_set = set(in_cols), set(tpl_cols)
    matched = [c for c in tpl_cols if c in in_set]
    missing = [c for c in tpl_cols if c not in in_set]
    extra = [c for c in in_cols if c not in tpl_set]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Template columns", len(tpl_cols))
    m2.metric("Matched", len(matched))
    m3.metric("Added blank", len(missing))
    m4.metric("Kept" if preserve_unknown else "Dropped", len(extra))

    if not matched:
        st.error("No input column matches the template headers. Check the header row of your data file.")

    with st.expander("Column match details"):
        st.markdown(f"**Added blank (in template, not in your file):** {', '.join(missing) or 'None'}")
        st.markdown(
            f"**{'Kept after template columns' if preserve_unknown else 'Dropped'} "
            f"(in your file, not in template):** {', '.join(extra) or 'None'}"
        )


def render_export_page():
    settings = get_settings()
    header_with_back(
        "📤 Export Format",
        "Upload your data file and download it reindexed to the Rubick OS template.",
    )

    template_path = Path(__file__).parent / settings.export_template
    if not template_path.is_file():
        st.error(f"Rubick OS template not found at {settings.export_template}.")
        st.stop()
    template_bytes = template_path.read_bytes()

    main, side = st.columns([3, 2], gap="large")
    with side:
        render_export_readme()

    with main:
        step(1, "Upload your data file")
        input_file = st.file_uploader(
            "Input Excel File",
            type=["xlsx", "xls"],
            key="export_input",
        )

        step(2, "Options")
        preserve_unknown = st.checkbox(
            "Preserve unknown columns",
            value=False,
            help="Keep columns from the input that are not in the template",
        )

        max_size = settings.max_file_size_mb * 1024 * 1024
        if input_file and input_file.size > max_size:
            st.error(f"Input file exceeds maximum size ({settings.max_file_size_mb}MB)")
            input_file = None

        step(3, "Check and format")
        if not input_file:
            st.info("Upload your data file to continue.")
        else:
            render_column_match(input_file.getvalue(), template_bytes, preserve_unknown)

            output_filename = st.text_input(
                "Output Filename",
                value=settings.export_output,
                help="Must end with .xlsx or .xls",
            )
            if output_filename and not is_valid_filename(output_filename):
                st.error("Output filename must end with .xlsx or .xls")
                output_filename = settings.export_output

            if st.button("Format Excel File", type="primary"):
                result = process_export(
                    input_file_data=input_file.getvalue(),
                    input_filename=input_file.name,
                    template_file_data=template_bytes,
                    template_filename=template_path.name,
                    output_filename=output_filename,
                    preserve_unknown_columns=preserve_unknown,
                )

                if result.success:
                    st.success("Format completed successfully!")
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Rows", result.rows_processed)
                    m2.metric("Input cols", result.columns_in_input)
                    m3.metric("Output cols", result.columns_in_output)
                    m4.metric("Added", result.columns_added)
                    st.download_button(
                        label="⬇️ Download Formatted File",
                        data=result.data,
                        file_name=result.filename,
                        mime=XLSX_MIME,
                    )
                else:
                    st.error(f"Error: {result.error_message}")


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def main():
    setup_page()
    inject_css()

    if "current_mode" not in st.session_state:
        st.session_state.current_mode = None

    mode = st.session_state.current_mode
    if mode is None:
        render_landing_page()
    elif mode == "input_generator":
        render_input_generator_page()
    elif mode == "export":
        render_export_page()


if __name__ == "__main__":
    main()
