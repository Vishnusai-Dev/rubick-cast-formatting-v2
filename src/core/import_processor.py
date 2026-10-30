"""
Import processor module for Streamlit integration.

Provides a clean interface for the web UI to process import operations
without dealing with file handling details.
"""

import tempfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Optional

from src.utils.logging import get_logger
from .import_formatter import merge_sizechart_productdetails, ImportResult

logger = get_logger(__name__)


@dataclass
class ImportProcessorResult:
    """Result of an import processing operation for Streamlit."""
    success: bool
    data: Optional[BytesIO] = None
    filename: Optional[str] = None
    rows_processed: int = 0
    columns_count: int = 0
    sheets_processed: int = 0
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


def process_import(
    size_chart_data: Optional[bytes],
    size_chart_filename: Optional[str],
    product_details_data: bytes,
    product_details_filename: str,
    output_filename: str = "Input_Generated.xlsx",
    exclude_sheets: Optional[list[str]] = None
) -> ImportProcessorResult:
    """
    Process an import merging operation for Streamlit.

    This function handles the complete workflow:
    1. Receives raw file data from uploaded files
    2. Creates temporary files for processing
    3. Runs the import formatter (merges Style + optional SKU files)
    4. Returns the result as a BytesIO object for download

    Args:
        size_chart_data: Raw bytes of the size chart (SKU) Excel file
        size_chart_filename: Original filename of the size chart file
        product_details_data: Raw bytes of the product details (Style) Excel file
        product_details_filename: Original filename of the product details file
        output_filename: Name for the output file
        exclude_sheets: List of sheet names to exclude from processing

    Returns:
        ImportProcessorResult with operation details and downloadable data
    """
    logger.info("Starting import processing", extra_data={
        "size_chart_filename": size_chart_filename,
        "product_details_filename": product_details_filename,
        "output_filename": output_filename,
        "exclude_sheets": exclude_sheets
    })

    has_size_chart = size_chart_data is not None and size_chart_filename is not None

    # Validate file sizes
    if has_size_chart:
        error = validate_file_size(size_chart_data, size_chart_filename)
        if error:
            logger.warning("Size chart file size validation failed", extra_data={"filename": size_chart_filename, "error": error})
            return ImportProcessorResult(success=False, error_message=error)

    error = validate_file_size(product_details_data, product_details_filename)
    if error:
        logger.warning("Product details file size validation failed", extra_data={"filename": product_details_filename, "error": error})
        return ImportProcessorResult(success=False, error_message=error)

    # Validate filename extensions
    if has_size_chart and not size_chart_filename.lower().endswith(('.xlsx', '.xls')):
        logger.warning("Invalid size chart file extension", extra_data={"filename": size_chart_filename})
        return ImportProcessorResult(
            success=False,
            error_message="Size chart file must be an Excel file (.xlsx or .xls)"
        )

    if not product_details_filename.lower().endswith(('.xlsx', '.xls')):
        logger.warning("Invalid product details file extension", extra_data={"filename": product_details_filename})
        return ImportProcessorResult(
            success=False,
            error_message=f"Product details file must be an Excel file (.xlsx or .xls)"
        )

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # Create temp file paths
            size_chart_path = (tmpdir / size_chart_filename) if has_size_chart else None
            product_details_path = tmpdir / product_details_filename
            output_path = tmpdir / output_filename

            # Write uploaded data to temp files
            if has_size_chart:
                with open(size_chart_path, "wb") as f:
                    f.write(size_chart_data)
            with open(product_details_path, "wb") as f:
                f.write(product_details_data)

            # Process the files
            result = merge_sizechart_productdetails(
                size_chart_path=size_chart_path,
                product_details_path=product_details_path,
                output_path=output_path,
                exclude_sheets=exclude_sheets
            )

            if result.success:
                logger.info("Import processing completed successfully", extra_data={
                    "rows_processed": result.rows_processed,
                    "columns_count": result.columns_count,
                    "sheets_processed": result.sheets_processed
                })
                # Read the output into BytesIO for download
                with open(output_path, "rb") as f:
                    output_data = BytesIO(f.read())

                return ImportProcessorResult(
                    success=True,
                    data=output_data,
                    filename=output_filename,
                    rows_processed=result.rows_processed,
                    columns_count=result.columns_count,
                    sheets_processed=result.sheets_processed
                )
            else:
                logger.warning("Import formatting failed", extra_data={"error": result.error_message})
                return ImportProcessorResult(
                    success=False,
                    error_message=result.error_message
                )

    except Exception as e:
        logger.error("Unexpected error during import processing", extra_data={"error": str(e)})
        return ImportProcessorResult(
            success=False,
            error_message=f"Processing error: {e}"
        )
