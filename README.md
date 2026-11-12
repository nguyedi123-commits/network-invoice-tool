# Network invoice tool

Turns NEM network invoice files (NBCHARGES / NBREMIND) into:
- `<Month> Network Invoice.xlsx`: the Excel extract
- `Remit/`: one NBREMITT remit file per invoice file
- `Network Lines.csv`: every charge line, to load into Snowflake for the reconciliation

All three come back in one zip. Nothing is stored on the server.

## Use it
1. Open the app link.
2. Type the month (used in file names), drop all the invoice files, then click **Create extract and remits**.
3. Check the table ("Check" column, skipped files), then download the zip.

## Snowflake (reconciliation)
- `snowflake/network_lines.sql`: run step 1 (table) and step 3 (view) once.
- After each run, load that run's `Network Lines.csv` into `NETWORK_LINES_RAW` (Snowsight > Load data, first line = header).
- Query the `NETWORK_LINES` view. It removes duplicates (a file loaded twice, V2 replacing V1, the same invoice in a later file) and leaves out interest. It's the same result as the old tool's "Network Lines" tab.

## Files
- `streamlit_app.py`: the page
- `invoice_run.py`: one run (files in, zip out)
- `nem_core.py`: reading invoices, Excel rows, remits (same logic as the old .exe)
- `gsl_codes.txt`: GSL charge codes. Add new codes here and commit; the app picks them up.

## Deploy (Streamlit Community Cloud)
1. share.streamlit.io > **Create app** > from GitHub: repo `network-invoice-tool`, branch `main`, file `streamlit_app.py`.
2. Advanced settings: Python 3.12. Deploy.
3. App > **Share**: keep it private and add the email addresses allowed to open it.

## Change the code
- Run locally: `pip install -r requirements.txt`, then `streamlit run streamlit_app.py`.
- Push to `main`: the app updates by itself.
