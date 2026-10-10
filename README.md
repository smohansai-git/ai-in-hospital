# AI IN HOSPITAL

AI IN HOSPITAL is a Flask-based hospital resource optimization system that uses Excel workbooks as the only persistent data store. It includes patient admission, isolation-aware bed assignment, discharge workflows, doctor and nurse assignment, medicine tracking, billing, and activity logging.

## Features

- Secure login and registration using hashed passwords
- Role-based access for Administrator, Doctor, Nurse, and Receptionist
- Excel-based persistence with workbook setup on first launch
- Smart bed assignment logic that prioritizes communicable disease isolation
- Dashboard with real metrics and Chart.js visualizations
- Patient discharge and billing workflow
- Medicine inventory monitoring and stock alerts
- Activity log capture for key actions

## Demo credentials

- User ID: admin
- Password: admin123

## Run locally

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Then open http://127.0.0.1:5000

## Deployment notes

This app is set up for Render with Gunicorn. Render's default ephemeral filesystem may lose Excel changes after restarts or redeployments. If a persistent disk is unavailable on the selected plan, the app will still run but workbook edits will not survive redeployments.

## Data storage

All application data is kept in the `excel_data` folder as `.xlsx` files. These files are intentionally excluded from the repository by `.gitignore`.
