"""
Export formatter module for reindexing Excel files to match template columns.

Handles formatting input data to match a target template structure,
preserving column order and adding missing columns as empty.
"""

import pandas as pd
from pathlib import Path
from dataclasses import dataclass
from typing import Optional


@dataclass
class ExportResult:
    """Result of an export formatting operation."""
    success: bool
    output_path: Optional[Path] = None
    rows_processed: int = 0
    columns_in_input: int = 0
    columns_in_output: int = 0
    columns_added: int = 0
    error_message: Optional[str] = None


class ExcelValidationError(Exception):
    """Custom exception for Excel validation errors."""
    pass


def validate_excel_file(file_path: Path, file_name: str) -> None:
    """
    Validate that an Excel file is readable and not empty.

    Args:
        file_path: Path to the Excel file
        file_name: Name of the file for error messages

    Raises:
        ExcelValidationError: If file is invalid or empty
    """
    try:
        # Try reading the file
        df = pd.read_excel(file_path)

        if df is None:
            raise ExcelValidationError(f"{file_name}: Could not read Excel file (file may be corrupted)")

        if df.empty:
            raise ExcelValidationError(f"{file_name}: Excel file contains no data")

        # Check for valid columns
        if len(df.columns) == 0:
            raise ExcelValidationError(f"{file_name}: Excel file has no columns")

    except pd.errors.EmptyDataError:
        raise ExcelValidationError(f"{file_name}: Excel file is empty or contains no parseable data")
    except pd.errors.ParserError as e:
        raise ExcelValidationError(f"{file_name}: Could not parse Excel file: {e}")
    except Exception as e:
        if isinstance(e, ExcelValidationError):
            raise
        raise ExcelValidationError(f"{file_name}: Error reading file: {e}")


def format_excel_file(
    input_file: Path,
    template_file: Path,
    output_file: Path,
    preserve_unknown_columns: bool = False
) -> ExportResult:
    """
    Format an Excel file to match a template's column structure.

    This function:
    1. Loads the template to get the target column structure
    2. Loads the input data
    3. Reindexes input columns to match template (reordering, dropping extras)
    4. Adds missing template columns (filled with empty values)
    5. Saves the formatted output

    Args:
        input_file: Path to the input Excel file
        template_file: Path to the template Excel file
        output_file: Path for the formatted output
        preserve_unknown_columns: If True, keep columns from input not in template

    Returns:
        ExportResult with operation details
    """
    try:
        # Validate input files
        validate_excel_file(input_file, "Input file")

        # Load only the template headers - a header-only Rubick OS template is valid
        try:
            df_template = pd.read_excel(template_file, nrows=0)
        except Exception as e:
            raise ExcelValidationError(f"Template file: Error reading file: {e}")
        target_columns = df_template.columns.tolist()

        if len(target_columns) == 0:
            return ExportResult(
                success=False,
                error_message="Template file has no columns to match against"
            )

        # Load the input data that needs formatting
        df_input = pd.read_excel(input_file)
        input_columns = df_input.columns.tolist()

        if len(input_columns) == 0:
            return ExportResult(
                success=False,
                error_message="Input file has no columns to process"
            )

        # Determine output columns based on configuration
        if preserve_unknown_columns:
            # Combine template columns with extra input columns (after template columns)
            output_columns = list(target_columns) + [
                col for col in input_columns if col not in target_columns
            ]
        else:
            output_columns = target_columns

        # Reindex the input dataframe to match target structure
        df_output = df_input.reindex(columns=output_columns)

        # Save the result
        df_output.to_excel(output_file, index=False)

        return ExportResult(
            success=True,
            output_path=output_file,
            rows_processed=len(df_output),
            columns_in_input=len(input_columns),
            columns_in_output=len(output_columns),
            columns_added=len([c for c in output_columns if c not in input_columns])
        )

    except FileNotFoundError as e:
        return ExportResult(
            success=False,
            error_message=f"File not found: {e}"
        )
    except ExcelValidationError as e:
        return ExportResult(
            success=False,
            error_message=str(e)
        )
    except Exception as e:
        return ExportResult(
            success=False,
            error_message=f"Error processing file: {e}"
        )


def normalize_column_name(col: str) -> str:
    """Normalize a column name for matching (lowercase, stripped)."""
    return col.strip().lower()


def find_column_in_list(
    target_columns: list[str],
    column_name: str
) -> Optional[str]:
    """
    Find a column in a list using normalized matching.

    Args:
        target_columns: List of column names to search
        column_name: Column name to find

    Returns:
        Matching column name or None
    """
    normalized_target = {normalize_column_name(c): c for c in target_columns}
    normalized_name = normalize_column_name(column_name)
    return normalized_target.get(normalized_name)
