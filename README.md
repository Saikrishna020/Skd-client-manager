# Health Open Cases Report

Streamlit app: upload a "Health open cases" export and download two reports:

1. **Client Wise** - Total cases per Client, broken down by Sub Product.
2. **Manager Pending Closure** - for cases where the FO has completed
   fieldwork (`Status == "FO Completed"`), per Manager:
   - `Total` - count of such cases
   - `More than 7 days` - `SKD TAT-D > 7`
   - `Cashless/Spot intimation` - Sub Product is "Cashless" or "Spot Intimation"
   - `More than 1 day after closure by CAT` - `CAT Completed Date` is 2 or
     more days before today (today = the date the app is run)

## Run locally

```
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Render

1. Push this repo to GitHub.
2. On Render, create a new **Web Service** from the repo (or use the
   included `render.yaml` via "New > Blueprint").
3. Build command: `pip install -r requirements.txt`
4. Start command: `streamlit run app.py --server.port=$PORT --server.address=0.0.0.0`

## Other scripts in this repo

- `fo_tat_report.py` - FO-wise TAT compliance report from a "Cat Closed" export.
- `manager_tat_report.py` - Manager-wise TAT compliance report from a
  "Health Managers closed" export.
- `health_open_cases_report.py` - CLI version of the report the Streamlit
  app above is built on; also runnable directly:
  `python health_open_cases_report.py "Health open cases.xlsx"`
