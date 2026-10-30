"""
Export processor module for Streamlit integration.

Provides a clean interface for the web UI to process export operations
without dealing with file handling details.
"""

import tempfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Optional

import pandas as pd

from src.utils.logging import get_logger
from .export_formatter import format_excel_file, ExportResult

logger = get_logger(__name__)


@dataclass
class ExportProcessorResult:
    """Result of an export processing operation for Streamlit."""
    success: bool
    data: Optional[BytesIO] = None
    filename: Optional[str] = None
    rows_processed: int = 0
    columns_in_input: int = 0
    columns_in_output: int = 0
    columns_added: int = 0
    error_message: Optional[str] = None


# Maximum file size in bytes (50MB)
MAX_FILE_SIZE = 50 * 1024 * 1024


def validate_file_size(data: bytes, filename: str, max_size: int = MAX_FILE_SIZE) -> Optional[str]:
    """
    Validate file size is within limits.

    Args:
        data: File data in bytes
        filename: Filename for error messages
        max_size: Maximum allowed size in bytes

    Returns:
        Error message if invalid, None if valid
    """
    if len(data) > max_size:
        size_mb = len(data) / (1024 * 1024)
        max_mb = max_size / (1024 * 1024)
        return f"{filename} is {size_mb:.1f}MB, exceeds maximum size of {max_mb:.0f}MB"
    return None


def process_export(
    input_file_data: bytes,
    input_filename: str,
    template_file_data: bytes,
    template_filename: str,
    output_filename: str = "Formatted_Output.xlsx",
    preserve_unknown_columns: bool = False
) -> ExportProcessorResult:
    """
    Process an export formatting operation for Streamlit.

    This function handles the complete workflow:
    1. Receives raw file data from uploaded files
    2. Creates temporary files for processing
    3. Runs the export formatter
    4. Returns the result as a BytesIO object for download

    Args:
        input_file_data: Raw bytes of the input Excel file
        input_filename: Original filename of the input file
        template_file_data: Raw bytes of the template Excel file
        template_filename: Original filename of the template file
        output_filename: Name for the output file (default: "Formatted_Output.xlsx")
        preserve_unknown_columns: If True, keep columns from input not in template

    Returns:
        ExportProcessorResult with operation details and downloadable data
    """
    logger.info("Starting export processing", extra_data={
        "input_filename": input_filename,
        "template_filename": template_filename,
        "output_filename": output_filename,
        "preserve_unknown_columns": preserve_unknown_columns
    })

    # Validate file sizes
    error = validate_file_size(input_file_data, input_filename)
    if error:
        logger.warning("Input file size validation failed", extra_data={"filename": input_filename, "error": error})
        return ExportProcessorResult(success=False, error_message=error)

    error = validate_file_size(template_file_data, template_filename)
    if error:
        logger.warning("Template file size validation failed", extra_data={"filename": template_filename, "error": error})
        return ExportProcessorResult(success=False, error_message=error)

    # Validate filename extensions
    if not input_filename.lower().endswith(('.xlsx', '.xls')):
        logger.warning("Invalid input file extension", extra_data={"filename": input_filename})
        return ExportProcessorResult(
            success=False,
            error_message=f"Input file must be an Excel file (.xlsx or .xls)"
        )

    if not template_filename.lower().endswith(('.xlsx', '.xls')):
        logger.warning("Invalid template file extension", extra_data={"filename": template_filename})
        return ExportProcessorResult(
            success=False,
            error_message=f"Template file must be an Excel file (.xlsx or .xls)"
        )

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # Create temp file paths
            input_path = tmpdir / input_filename
            template_path = tmpdir / template_filename
            output_path = tmpdir / output_filename

            # Write uploaded data to temp files
            with open(input_path, "wb") as f:
                f.write(input_file_data)
            with open(template_path, "wb") as f:
                f.write(template_file_data)

            # Process the files
            result = format_excel_file(
                input_file=input_path,
                template_file=template_path,
                output_file=output_path,
                preserve_unknown_columns=preserve_unknown_columns
            )

            if result.success:
                logger.info("Export processing completed successfully", extra_data={
                    "rows_processed": result.rows_processed,
                    "columns_in_input": result.columns_in_input,
                    "columns_in_output": result.columns_in_output,
                    "columns_added": result.columns_added
                })
                # Read the output into BytesIO for download
                with open(output_path, "rb") as f:
                    output_data = BytesIO(f.read())

                return ExportProcessorResult(
                    success=True,
                    data=output_data,
                    filename=output_filename,
                    rows_processed=result.rows_processed,
                    columns_in_input=result.columns_in_input,
                    columns_in_output=result.columns_in_output,
                    columns_added=result.columns_added
                )
            else:
                logger.warning("Export formatting failed", extra_data={"error": result.error_message})
                return ExportProcessorResult(
                    success=False,
                    error_message=result.error_message
                )

    except Exception as e:
        logger.error("Unexpected error during export processing", extra_data={"error": str(e)})
        return ExportProcessorResult(
            success=False,
            error_message=f"Processing error: {e}"
        )
