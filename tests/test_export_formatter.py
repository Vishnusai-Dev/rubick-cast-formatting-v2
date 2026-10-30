import pandas as pd

from src.core.export_formatter import format_excel_file


def test_header_only_template_is_accepted(tmp_path):
    template = tmp_path / "template.xlsx"
    data = tmp_path / "data.xlsx"
    out = tmp_path / "out.xlsx"

    pd.DataFrame(columns=["styleId", "Brand", "Colour"]).to_excel(template, index=False)
    pd.DataFrame({"Colour": ["Red"], "styleId": ["1"], "Extra": ["x"]}).to_excel(data, index=False)

    result = format_excel_file(data, template, out)

    assert result.success
    df = pd.read_excel(out, dtype=str)
    assert list(df.columns) == ["styleId", "Brand", "Colour"]
    assert df.loc[0, "styleId"] == "1"
    assert result.columns_added == 1
