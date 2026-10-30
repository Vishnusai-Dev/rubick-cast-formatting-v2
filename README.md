# Rubick CAST Formatting (v0.2)

Streamlit tool with two modes:

- **Input Generator**: builds the input workbook (`Types` + `Values`) from a Style file (required) and a SKU / size chart file (optional). Includes an on-page guide and example files showing the expected input.
- **Export Format**: reindexes your data file to the column structure of an uploaded Rubick OS template (single template option; no default).

## Run

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

## Test

```bash
pip install pytest
python -m pytest -q
```

## Deploy (Streamlit Community Cloud)

Push this folder to a GitHub repo, set `app.py` as the main file. `requirements.txt` is picked up automatically.
