# SN Info → Result Chart Converter

A Streamlit web app that converts `SN_Info_list.xlsx` (or CSV/ZIP) into a
fully-populated `Result_Chart.xlsx` with Temperature Distribution and CH Distribution sheets.

## How It Works

1. User uploads `SN_Info_list.xlsx`, a `.csv`, or a `.zip` of CSV files
2. App writes the data into the **Raw Data** sheet of `Result_Chart_template.xlsx`
3. User downloads the resulting `Result_Chart.xlsx`
4. On opening in Excel, the **Temperature Distribution** and **CH Distribution** sheets
   auto-recalculate from the live SUMIFS / AGGREGATE formulas

## Files

| File | Purpose |
|------|---------|
| `app.py` | Streamlit application |
| `Result_Chart_template.xlsx` | Template with formula sheets — **must be in the repo** |
| `requirements.txt` | Python dependencies |

## Local Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

1. Push this folder to a **GitHub repository** (public or private)
   - ⚠️ Make sure `Result_Chart_template.xlsx` is committed to the repo
2. Go to https://share.streamlit.io → **New app**
3. Select your repo, branch, and set **Main file path** = `app.py`
4. Click **Deploy**

### Supported Upload Formats

| Format | Notes |
|--------|-------|
| `.xlsx` | Standard Excel — Sheet1 is used |
| `.csv`  | Single CSV file |
| `.zip`  | ZIP containing one or more CSV files (all are merged) |
