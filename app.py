"""
Rubick CAST Formatting - Streamlit Application

Two modes:
  1. Input Generator - builds the Types + Values input workbook from a Style file
     (required) and an optional SKU / size chart file.
  2. Export Format   - reindexes the input to the uploaded Rubick OS template.
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
        layout="centered",
        initial_sidebar_state="collapsed",
    )


def render_logo():
    logo_path = Path(__file__).parent / "logo.svg"
    if logo_path.exists():
        st.image(str(logo_path), width=120)


def header_with_back(title: str, subtitle: str):
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.current_mode = None
            st.rerun()
    with col2:
        st.title(title)
        st.markdown(subtitle)
    st.markdown("---")


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


# --------------------------------------------------------------------------- #
# Landing page
# --------------------------------------------------------------------------- #
def render_landing_page():
    settings = get_settings()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        render_logo()
        st.title(settings.app_name)
        st.caption(f"Version {settings.app_version}")
        st.markdown("---")

    st.markdown("### Choose a Processing Mode")

    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown(
            """
            <div class="mode-card input-card">
                <h2>🧩 Input Generator</h2>
                <p>Build the input file (Types + Values) from your Style file and, optionally, a SKU size chart.</p>
                <ul>
                    <li>Style file required, SKU optional</li>
                    <li>Auto-generate Types sheet</li>
                    <li>Download the generated input</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Start Input Generator", use_container_width=True, type="primary", key="input_gen"):
            st.session_state.current_mode = "input_generator"
            st.rerun()

    with col2:
        st.markdown(
            """
            <div class="mode-card export-card">
                <h2>📤 Export Format</h2>
                <p>Reindex your data to the column structure of the Rubick OS template.</p>
                <ul>
                    <li>Upload your data file</li>
                    <li>Upload the Rubick OS template</li>
                    <li>Download formatted output</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Start Export", use_container_width=True, type="primary", key="export"):
            st.session_state.current_mode = "export"
            st.rerun()

    st.markdown(
        """
        <style>
        .mode-card {
            background-color: #f8f9fa;
            border-radius: 10px;
            padding: 20px;
            text-align: center;
            height: 100%;
            border: 2px solid transparent;
            transition: all 0.3s ease;
        }
        .mode-card:hover { border-color: #4e8cff; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
        .mode-card h2 { margin-bottom: 10px; color: #1f2937; }
        .mode-card p { color: #6b7280; font-size: 0.9rem; margin-bottom: 15px; }
        .mode-card ul { text-align: left; color: #4b5563; font-size: 0.85rem; padding-left: 20px; }
        .input-card { border-left: 4px solid #6366f1; }
        .export-card { border-left: 4px solid #10b981; }
        </style>
        """,
        unsafe_allow_html=True,
    )


# --------------------------------------------------------------------------- #
# Input Generator
# --------------------------------------------------------------------------- #
def render_input_guide():
    with st.expander("📘 How should the input be?", expanded=True):
        st.markdown(
            """
            **Two files, one tab per category.** Tab names must be identical in both files.
            Tabs named in *Exclude Sheets* (default `masterdata`) are skipped.

            | File | Required | What it holds |
            |---|---|---|
            | **Style file** (`style.xlsx`) | Yes | One row per style with product attributes |
            | **SKU file** (`sku.xlsx`) | No | One row per style + size with size chart measurements |
            """
        )

        st.markdown("**Style file**: one row per style")
        st.dataframe(EXAMPLE_STYLE, hide_index=True, use_container_width=True)

        st.markdown("**SKU file (optional)**: one row per style + size")
        st.dataframe(EXAMPLE_SKU, hide_index=True, use_container_width=True)

        st.markdown(
            """
            **Rules**
            - Header row in row 1 of every tab.
            - The style identifier column must be named `styleId`, `style_id` or `SKU`.
            - SKU file: size columns start at the **Brand Size** column and run to the right.
              Rows are combined per style into comma-separated values (e.g. `S,M`).
            - Without a SKU file, no size columns are added; the output holds Style file columns only.
            - Characters other than letters, numbers, `( ) _ - #` and spaces are removed from headers.

            **Output:** one workbook with a `Types` sheet and a `Values` sheet.
            """
        )

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇️ Example style.xlsx",
                data=example_style_bytes(),
                file_name="example_style.xlsx",
                mime=XLSX_MIME,
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "⬇️ Example sku.xlsx",
                data=example_sku_bytes(),
                file_name="example_sku.xlsx",
                mime=XLSX_MIME,
                use_container_width=True,
            )
        st.caption("Example files use placeholder values only.")


def render_input_generator_page():
    settings = get_settings()
    header_with_back(
        "🧩 Input Generator",
        "Build the input file from your Style file and, optionally, a SKU size chart.",
    )

    render_input_guide()

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

    if product_details_file and not size_chart_file:
        st.info("No SKU file uploaded: size columns will not be added.")

    if product_details_file:
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
                with st.expander("Processing Details", expanded=True):
                    st.json(
                        {
                            "Rows Processed": result.rows_processed,
                            "Columns": result.columns_count,
                            "Sheets Processed": result.sheets_processed,
                            "Size Chart Used": size_chart_file is not None,
                        }
                    )
                st.download_button(
                    label="Download Generated Input",
                    data=result.data,
                    file_name=result.filename,
                    mime=XLSX_MIME,
                )
            else:
                st.error(f"Error: {result.error_message}")


# --------------------------------------------------------------------------- #
# Export Format
# --------------------------------------------------------------------------- #
def render_export_guide():
    with st.expander("📘 How should the input be?", expanded=True):
        st.markdown(
            """
            **Two files:** your data file and the **Rubick OS template**.

            - Data file: header row in row 1; only the **first sheet** is read.
            - Column headers must match the template headers **exactly** (spelling, spaces, capitalisation).
            - Matched columns are placed in template order. Template columns missing from your file are added blank.
            - Columns not in the template are dropped, unless *Preserve unknown columns* is ticked
              (they are then added after the template columns).
            - Example data file (header names must follow your template; the sample below is illustrative):
            """
        )
        st.dataframe(EXAMPLE_EXPORT_INPUT, hide_index=True, use_container_width=True)
        st.download_button(
            "⬇️ Example data file",
            data=example_export_input_bytes(),
            file_name="example_export_input.xlsx",
            mime=XLSX_MIME,
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
        "Reindex your data to the column structure of the Rubick OS template.",
    )

    render_export_guide()

    col1, col2 = st.columns(2)
    with col1:
        input_file = st.file_uploader(
            "Input Excel File",
            type=["xlsx", "xls"],
            key="export_input",
        )
    with col2:
        template_file = st.file_uploader(
            "Rubick OS Template",
            type=["xlsx", "xls"],
            key="export_template",
            help="Upload the Rubick OS template. Its column headers define the output format.",
        )

    preserve_unknown = st.checkbox(
        "Preserve unknown columns",
        value=False,
        help="Keep columns from the input that are not in the template",
    )

    max_size = settings.max_file_size_mb * 1024 * 1024
    if input_file and input_file.size > max_size:
        st.error(f"Input file exceeds maximum size ({settings.max_file_size_mb}MB)")
        input_file = None
    if template_file and template_file.size > max_size:
        st.error(f"Template file exceeds maximum size ({settings.max_file_size_mb}MB)")
        template_file = None

    if input_file and template_file:
        render_column_match(input_file.getvalue(), template_file.getvalue(), preserve_unknown)

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
                template_file_data=template_file.getvalue(),
                template_filename=template_file.name,
                output_filename=output_filename,
                preserve_unknown_columns=preserve_unknown,
            )

            if result.success:
                st.success("Format completed successfully!")
                with st.expander("Processing Details", expanded=True):
                    st.json(
                        {
                            "Rows Processed": result.rows_processed,
                            "Input Columns": result.columns_in_input,
                            "Output Columns": result.columns_in_output,
                            "Columns Added": result.columns_added,
                        }
                    )
                st.download_button(
                    label="Download Formatted File",
                    data=result.data,
                    file_name=result.filename,
                    mime=XLSX_MIME,
                )
            else:
                st.error(f"Error: {result.error_message}")
    elif input_file or template_file:
        st.info("Upload both the input file and the Rubick OS template to continue.")


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def main():
    setup_page()

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
