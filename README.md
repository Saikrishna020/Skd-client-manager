# SKD TAT Reports

Streamlit app: upload an export, pick a report from the dropdown, and
download the result(s) as Excel. Three reports are covered:

## 1. Health Open Cases

Upload a "Health open cases" export. Produces two reports:

- **Client Wise** - Total cases per Client, broken down by Sub Product,
  plus `FO Closed` (`Status == "FO Completed"`) and `Pending`
  (`Status == "Pending"`) counts.
- **Manager Pending Closure** - for cases where the FO has completed
  fieldwork (`Status == "FO Completed"`), per Manager:
  - `Total` - count of such cases
  - `More than 5 days` - `SKD TAT-D > 5`
  - `Cashless/Spot intimation` - Sub Product is "Cashless" or "Spot Intimation"
  - `More than 1 day after closure by CAT` - `CAT Completed Date` is 2 or
    more days before today (today = the date the app is run)
  - `% more than 1 day after closure by CAT` and `% More than 5 days` -
    the above counts as a percentage of Total (rounded half up)

## 2. FO TAT

Upload a "Cat Closed" export. Produces one report: cases, within-TAT,
and Discrepant counts per FO and Sub Product category, against a
standard TAT per category (Cashless 1 day, FULL INVESTIGATION 7 days,
TP 21 days, etc. - see `STANDARD_TAT` in `fo_tat_report.py`).

`Manager Days` (when present in the file) is recomputed as
`Manager Completed Date - FO Completed Date` rather than trusting the
source file's own formula, which evaluates to `#VALUE!` or a wrong
number whenever its date columns are stored as text. Rows where the
recomputed value is negative are excluded.

## 3. Manager TAT

Upload a "Health Managers closed" export. Produces one report:
cases, within-TAT, Discrepant, and "Manager closed days >= 2" counts
per Manager and Sub Product category.

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

## Files

- `app.py` - the Streamlit app; a thin UI over the three report modules below.
- `health_open_cases_report.py` - Health Open Cases report logic; also
  runnable directly: `python health_open_cases_report.py "Health open cases.xlsx"`
- `fo_tat_report.py` - FO TAT report logic; also runnable directly:
  `python fo_tat_report.py "Cat Closed.xlsx"`
- `manager_tat_report.py` - Manager TAT report logic; also runnable
  directly: `python manager_tat_report.py "Health Managers closed.xlsx"`
