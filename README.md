# SJCL - IEX Landed Cost per Unit

A Streamlit app for calculating landed cost from IEX Scheduling Request (SCH)
and Daily Obligation Summary Report (DOR) spreadsheets.

## Run locally

```text
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Deploy with Streamlit Community Cloud

Create a new app from this repository, select the `main` branch, and set
`app.py` as the app file.

The app processes uploaded reports in memory and does not save them as files.
Community Cloud apps are publicly accessible by default. Settings entered in
the hosted app are private to each browser session and reset when that session
ends. Do not commit `settings.json`, uploaded reports, or credentials.
