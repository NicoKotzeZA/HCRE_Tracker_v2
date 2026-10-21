# Dynamic Project Budget & Cash Flow Tracker

Works for **any project**: configurable income & cost categories (headers), sub-lines, budgets, actuals and forecasts.

## Features

- **Dashboard** – KPIs, charts, facility utilisation, cash snapshot
- **Income structure** – add/edit/delete category headers and sub-lines; set budgets
- **Cost structure** – same for costs
- **Monthly values** – editable grid for actuals & forecasts
- **Export / Import** – full project JSON (round-trip) + CSV exports
- **Sidebar** – project name, currency, facility limit, interest rate, months (add/remove, Actual vs Forecast)

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Place `Cash Flow Forecast.xlsx` in the same folder as `app.py` if you want the HCRE template button to work.

## Deploy to Streamlit Community Cloud

1. Create a **public** GitHub repo
2. Upload these files at the **root**:
   - `app.py`
   - `requirements.txt`
   - `README.md`
   - `Cash Flow Forecast.xlsx` (optional – only needed for HCRE template)
3. Go to https://share.streamlit.io → New app → Main file path: `app.py` → Deploy

## Typical workflow

1. Set project name / facility / months in the sidebar  
2. Build **Income** and **Cost** structure (headers + sub-lines + budgets)  
3. Enter monthly actuals & forecasts on the Monthly values tab  
4. Review the Dashboard  
5. Export JSON to save your work; re-import later
