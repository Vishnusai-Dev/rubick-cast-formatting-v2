import re

import pandas as pd

from src.core.import_formatter import (
    consolidate_duplicate_columns,
    merge_sizechart_productdetails,
)

ALLOWED_HEADER_PATTERN = re.compile(r"^[A-Za-z0-9()_\-# ]+$")


def test_consolidate_duplicate_columns_overwrites_dense_column_with_sparse_values():
    df = pd.DataFrame(
        [
            ["1", None, "Machine wash"],
            ["2", "5", None],
            ["3", None, "Dry clean"],
        ],
        columns=["Net Quantity", "Net Quantity.1", "Material/Care?"],
    )

    result = consolidate_duplicate_columns(df)

    assert list(result.columns) == ["Net Quantity", "MaterialCare"]
    assert result["Net Quantity"].tolist() == ["1", "5", "3"]
    material = result["MaterialCare"].tolist()
    assert material[0] == "Machine wash"
    assert pd.isna(material[1])
    assert material[2] == "Dry clean"


def test_merge_sizechart_productdetails_sanitizes_headers_and_merges_duplicates(tmp_path):
    size_chart_path = tmp_path / "sku.xlsx"
    product_details_path = tmp_path / "style.xlsx"
    output_path = tmp_path / "output.xlsx"

    size_df = pd.DataFrame(
        {
            "styleId": ["A", "B", "C"],
            "Brand Size": ["S", "M", "L"],
            "Chest (Inches)": ["34", "36", "38"],
        }
    )
    product_df = pd.DataFrame(
        {
            "styleId": ["A", "B", "C"],
            "Net Quantity": ["1", "2", "3"],
            "Net Quantity.1": [None, "5", None],
            "Material/Care?": ["Machine wash", None, "Dry clean"],
        }
    )

    with pd.ExcelWriter(size_chart_path, engine="openpyxl") as writer:
        size_df.to_excel(writer, sheet_name="Apparel", index=False)

    with pd.ExcelWriter(product_details_path, engine="openpyxl") as writer:
        product_df.to_excel(writer, sheet_name="Apparel", index=False)

    result = merge_sizechart_productdetails(
        size_chart_path=size_chart_path,
        product_details_path=product_details_path,
        output_path=output_path,
    )

    assert result.success is True

    values_df = pd.read_excel(output_path, sheet_name="Values", dtype=str)
    types_df = pd.read_excel(output_path, sheet_name="Types", dtype=str)

    assert values_df.columns.tolist().count("Net Quantity") == 1
    assert values_df.loc[values_df["styleId"] == "B", "Net Quantity"].iloc[0] == "5"
    assert "MaterialCare" in values_df.columns

    for header in values_df.columns:
        assert ALLOWED_HEADER_PATTERN.fullmatch(header), header

    for header in types_df.columns:
        assert ALLOWED_HEADER_PATTERN.fullmatch(header), header


def test_merge_without_size_chart_uses_style_file_only(tmp_path):
    product_details_path = tmp_path / "style.xlsx"
    output_path = tmp_path / "output.xlsx"

    pd.DataFrame(
        {"styleId": ["A", "B"], "productDisplayName": ["Alpha", "Beta"]}
    ).to_excel(product_details_path, index=False)

    result = merge_sizechart_productdetails(
        size_chart_path=None,
        product_details_path=product_details_path,
        output_path=output_path,
    )

    assert result.success
    assert result.rows_processed == 2
    values = pd.read_excel(output_path, sheet_name="Values", dtype=str)
    assert list(values.columns) == ["styleId", "productDisplayName"]
    types = pd.read_excel(output_path, sheet_name="Types", dtype=str)
    assert types.loc[1, "styleId"] == "mandatory"


def test_merge_without_size_chart_reports_missing_style_column(tmp_path):
    product_details_path = tmp_path / "style.xlsx"
    pd.DataFrame({"name": ["Alpha"]}).to_excel(product_details_path, index=False)

    result = merge_sizechart_productdetails(
        size_chart_path=None,
        product_details_path=product_details_path,
        output_path=tmp_path / "out.xlsx",
    )

    assert not result.success
    assert "style ID column" in result.error_message
