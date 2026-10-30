"""
Input generator module: builds the Types + Values input workbook from Style (required) and SKU / size chart (optional) files.

Aggregates size chart rows per styleId, merges all tabs into one dataset,
and outputs Excel with Types and Values sheets in CAST format.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


@dataclass
class ImportResult:
    """Result of an import merging operation."""
    success: bool
    output_path: Path | None = None
    rows_processed: int = 0
    columns_count: int = 0
    sheets_processed: int = 0
    error_message: str | None = None


class ImportValidationError(Exception):
    """Custom exception for import validation errors."""
    pass


INVALID_HEADER_CHARS = re.compile(r"[^A-Za-z0-9()_\-# ]")
DUPLICATE_HEADER_SUFFIX = re.compile(r"\.\d+$")
EMPTY_TEXT_VALUES = {"", "nan", "none"}


def normalize_col(col: str) -> str:
    """Normalize column name for canonical mapping."""
    return re.sub(r"\s+", "", str(col).lower())


def sanitize_output_header(col: str) -> str:
    """Strip unsupported characters from output headers."""
    sanitized = INVALID_HEADER_CHARS.sub("", str(col))
    return sanitized if sanitized.strip() else "Column"


def base_output_header(col: str) -> str:
    """Return the display name for a header after removing duplicate suffixes."""
    base_name = DUPLICATE_HEADER_SUFFIX.sub("", str(col))
    return sanitize_output_header(base_name)


def normalized_output_header(col: str) -> str:
    """Return the canonical key used for grouping duplicate output columns."""
    return normalize_col(base_output_header(col)) or "column"


def has_meaningful_value(value) -> bool:
    """Return True when a cell contains usable data."""
    if pd.isna(value):
        return False
    if isinstance(value, str):
        return value.strip().lower() not in EMPTY_TEXT_VALUES
    return True


def count_meaningful_values(series: pd.Series) -> int:
    """Count non-empty values in a series."""
    return int(series.map(has_meaningful_value).sum())


def consolidate_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sanitize headers and merge duplicates into a single column.

    The most-populated column is used as the base. Sparser duplicate columns
    overwrite matching rows when they contain data.
    """
    if len(df.columns) == 0:
        return df.copy()

    value_counts = [count_meaningful_values(df.iloc[:, idx]) for idx in range(len(df.columns))]

    groups: dict[str, list[int]] = {}
    for idx, col in enumerate(df.columns):
        groups.setdefault(normalized_output_header(col), []).append(idx)

    merged_columns = []
    merged_names = []

    for _, indices in sorted(groups.items(), key=lambda item: min(item[1])):
        sorted_indices = sorted(indices, key=lambda idx: (-value_counts[idx], idx))
        merged = df.iloc[:, sorted_indices[0]].copy()
        merged_name = base_output_header(df.columns[sorted_indices[0]])

        for idx in sorted_indices[1:]:
            series = df.iloc[:, idx]
            overwrite_mask = series.map(has_meaningful_value)
            merged.loc[overwrite_mask] = series.loc[overwrite_mask]

        merged_columns.append(merged)
        merged_names.append(merged_name)

    consolidated = pd.concat(merged_columns, axis=1)
    consolidated.columns = merged_names
    return consolidated


def find_style_column(columns, sheet_name: str = "sheet") -> tuple[str | None, str | None]:
    """
    Find the style ID column in a list of columns.

    Returns:
        Tuple of (found_column, suggestion) or (None, None) if not found
    """
    patterns = [
        (re.compile(r"^style[_\s\-]?id$", re.I), "style_id"),
        (re.compile(r"^sku$", re.I), "SKU"),
        (re.compile(r"styleid", re.I), "styleId"),
    ]
    for c in columns:
        for p, suggestion in patterns:
            if p.search(str(c)):
                return c, None
    # Return first column as suggestion if no style column found
    if len(columns) > 0:
        return None, str(columns[0])
    return None, None


def find_brand_size_column(columns) -> str | None:
    """Find the brand size column (starts brand size columns)."""
    for c in columns:
        s = str(c).lower()
        if "brand" in s and "size" in s:
            return c
    for c in columns:
        if "size" in str(c).lower():
            return c
    return None


def validate_excel_file(file_path: Path, file_name: str) -> None:
    """
    Validate that an Excel file is readable and has valid structure.

    Args:
        file_path: Path to the Excel file
        file_name: Name of the file for error messages

    Raises:
        ImportValidationError: If file is invalid or empty
    """
    try:
        xl = pd.ExcelFile(file_path, engine="openpyxl")

        if len(xl.sheet_names) == 0:
            raise ImportValidationError(f"{file_name}: Excel file has no sheets")

        # Check first sheet has data
        df = xl.parse(xl.sheet_names[0], dtype=str)
        if df.empty and len(xl.sheet_names) == 1:
            raise ImportValidationError(f"{file_name}: First sheet is empty")
        if len(df.columns) == 0:
            raise ImportValidationError(f"{file_name}: Sheet has no columns")

    except ImportValidationError:
        raise
    except pd.errors.EmptyDataError:
        raise ImportValidationError(f"{file_name}: Excel file is empty")
    except pd.errors.ParserError as e:
        raise ImportValidationError(f"{file_name}: Could not parse Excel file: {e}")
    except Exception as e:
        if isinstance(e, ImportValidationError):
            raise
        raise ImportValidationError(f"{file_name}: Error reading file: {e}")


def aggregate_list(series) -> str | None:
    """Aggregate multiple values into a comma-separated string."""
    vals = (
        series.dropna()
        .astype(str)
        .str.strip()
        .replace({"": None, "nan": None, "none": None})
        .dropna()
        .unique()
    )
    return ",".join(vals) if len(vals) else None


DETAILS_ANGLE_PATTERN = re.compile(r"detail[s]?[\s_\-]*angle", re.I)
IMAGE_PATTERN = re.compile(r"image|img|cdn", re.I)
URL_PATTERN = re.compile(r"url", re.I)


def infer_column_type(col_name: str) -> str:
    """
    Infer the Types-sheet data type for a column based on its name.

    - "Details angle" columns -> image
    - Columns with image / img / cdn in the name -> image
    - Other columns with url in the name -> url
    - Everything else -> string
    """
    name = str(col_name)
    if DETAILS_ANGLE_PATTERN.search(name):
        return "image"
    if IMAGE_PATTERN.search(name):
        return "image"
    if URL_PATTERN.search(name):
        return "url"
    return "string"


def merge_sizechart_productdetails(
    size_chart_path: Path | None,
    product_details_path: Path,
    output_path: Path,
    exclude_sheets: list[str] | None = None
) -> ImportResult:
    """
    Generate the input workbook (Types + Values sheets) from Style and optional SKU files.

    This function:
    1. Loads the product details (Style) file and, if provided, the size chart (SKU) file
    2. Aggregates size chart rows per styleId (only when a size chart is provided)
    3. Merges all sheets into one dataset
    4. Creates Types sheet with column metadata
    5. Creates Values sheet with merged data

    Args:
        size_chart_path: Path to the size chart Excel file, or None to skip size aggregation
        product_details_path: Path to the product details Excel file
        output_path: Path for the generated Excel file
        exclude_sheets: List of sheet names to exclude

    Returns:
        ImportResult with operation details
    """
    try:
        exclude_sheets = exclude_sheets or ["masterdata"]

        # Validate input files
        validate_excel_file(product_details_path, "Product details file")
        if size_chart_path is not None:
            validate_excel_file(size_chart_path, "Size chart file")

        prod_xl = pd.ExcelFile(product_details_path, engine="openpyxl")
        size_xl = (
            pd.ExcelFile(size_chart_path, engine="openpyxl")
            if size_chart_path is not None
            else None
        )

        prod_sheets = [s for s in prod_xl.sheet_names if s not in exclude_sheets]

        if size_xl is not None:
            # Size chart drives the sheet list (existing behaviour)
            sheets = [s for s in size_xl.sheet_names if s not in exclude_sheets]
            if not sheets:
                return ImportResult(
                    success=False,
                    error_message="No valid sheets found in size chart after exclusions"
                )
            size_sheet_set = set(size_xl.sheet_names)
            prod_sheet_set = set(prod_xl.sheet_names)
        else:
            sheets = prod_sheets
            if not sheets:
                return ImportResult(
                    success=False,
                    error_message="No valid sheets found in product details file after exclusions"
                )
            size_sheet_set = set()
            prod_sheet_set = set(prod_xl.sheet_names)

        canonical_cols = {}
        ordered_columns = []
        final_dfs = []
        any_valid_data = False

        for sheet in sheets:
            size_df = (
                size_xl.parse(sheet, dtype=str)
                if size_xl is not None and sheet in size_sheet_set
                else pd.DataFrame()
            )
            prod_df = (
                prod_xl.parse(sheet, dtype=str)
                if sheet in prod_sheet_set
                else pd.DataFrame()
            )

            size_df.columns = [str(c) for c in size_df.columns]
            prod_df.columns = [str(c) for c in prod_df.columns]

            # The reference file for the style ID column is the size chart when present,
            # otherwise the product details sheet.
            ref_df = size_df if size_xl is not None else prod_df
            ref_name = "Size chart" if size_xl is not None else "Product details"
            style_ref, _ = find_style_column(ref_df.columns, sheet)

            if style_ref is None:
                # Only error on first sheet - subsequent sheets might be intentionally empty
                if sheet == sheets[0]:
                    columns_list = ", ".join(ref_df.columns[:5].tolist())
                    if len(ref_df.columns) > 5:
                        columns_list += ", ..."
                    return ImportResult(
                        success=False,
                        error_message=f"Could not find style ID column in {ref_name} sheet '{sheet}'. "
                                      f"Expected column like 'style_id', 'SKU', or 'styleId'. "
                                      f"Found: [{columns_list}]. "
                                      f"Rename your identifier column to match one of these patterns."
                    )
                continue

            any_valid_data = True

            if size_xl is not None:
                style_size = style_ref
                if prod_df.empty:
                    prod_df = pd.DataFrame({style_size: size_df[style_size].dropna().unique()})
                    style_prod = style_size
                else:
                    style_prod = find_style_column(prod_df.columns)[0]
                    if style_prod is None:
                        style_prod = style_size
                        prod_df[style_prod] = prod_df.iloc[:, 0]
            else:
                style_size = None
                style_prod = style_ref

            # Register product columns
            for col in prod_df.columns:
                key = normalize_col(col)
                if key not in canonical_cols:
                    canonical_cols[key] = col
                    ordered_columns.append(col)

            if not prod_df.empty:
                prod_df.rename(columns=lambda c: canonical_cols[normalize_col(c)], inplace=True)

            # Size aggregation (only when a size chart was provided)
            brand_col = find_brand_size_column(size_df.columns) if style_size else None

            if brand_col:
                size_cols = size_df.columns[size_df.columns.get_loc(brand_col):]

                melted = size_df.melt(
                    id_vars=[style_size],
                    value_vars=size_cols,
                    var_name="col",
                    value_name="val",
                )

                aggregated = (
                    melted.groupby([style_size, "col"])["val"]
                    .apply(aggregate_list)
                    .reset_index()
                )

                pivot = aggregated.pivot(
                    index=style_size, columns="col", values="val"
                ).reset_index()

                pivot.rename(columns={style_size: style_prod}, inplace=True)
                prod_df = prod_df.merge(pivot, on=style_prod, how="outer")

                for col in pivot.columns:
                    if col == style_prod:
                        continue
                    key = normalize_col(col)
                    if key not in canonical_cols:
                        canonical_cols[key] = col
                        ordered_columns.append(col)

            final_dfs.append(prod_df)

        if not any_valid_data:
            return ImportResult(
                success=False,
                error_message="No valid data found in any sheet. Check that files contain style ID columns."
            )

        # Final values dataframe
        final_df = pd.concat(final_dfs, ignore_index=True)

        if final_df.empty:
            return ImportResult(
                success=False,
                error_message="No data produced after merging. Check input file formats."
            )

        for col in ordered_columns:
            if col not in final_df.columns:
                final_df[col] = None

        final_df = final_df[ordered_columns]
        final_df = consolidate_duplicate_columns(final_df)
        final_df = final_df.where(pd.notnull(final_df), "")
        ordered_columns = final_df.columns.tolist()

        # Build Types sheet
        types_columns = ["Column1", "Column2"] + ordered_columns
        types_df = pd.DataFrame(columns=types_columns)

        for col in ordered_columns:
            types_df.loc[0, col] = col
        for col in ordered_columns:
            types_df.loc[1, col] = "mandatory"
        for col in ordered_columns:
            types_df.loc[2, col] = infer_column_type(col)

        # Write Excel output
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            types_df.to_excel(writer, sheet_name="Types", index=False)
            final_df.to_excel(writer, sheet_name="Values", index=False)

        return ImportResult(
            success=True,
            output_path=output_path,
            rows_processed=len(final_df),
            columns_count=len(final_df.columns),
            sheets_processed=len(sheets)
        )

    except FileNotFoundError as e:
        return ImportResult(
            success=False,
            error_message=f"File not found: {e}"
        )
    except ImportValidationError as e:
        return ImportResult(
            success=False,
            error_message=str(e)
        )
    except Exception as e:
        return ImportResult(
            success=False,
            error_message=f"Error processing files: {e}"
        )
