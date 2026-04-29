# S&I Reimbursement API

Python/Flask backend for the S&I Travel Reimbursement app.
Fills the official company Excel template with trip data and returns
a perfectly formatted .xlsx file.

## Files
- `app.py` — Flask API
- `Reimbursement_Request_Form.xlsx` — blank company template
- `requirements.txt` — Python dependencies
- `Procfile` — Render start command

## Deploy on Render
1. Connect this repo to Render as a new Web Service
2. Build command: `pip install -r requirements.txt`
3. Start command: `gunicorn app:app`
