from __future__ import annotations

import os
import re
import threading
from datetime import date, datetime
from pathlib import Path

from flask import Flask, flash, jsonify, redirect, render_template, request, session, url_for
from openpyxl import Workbook, load_workbook
from werkzeug.security import check_password_hash, generate_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "ai-hospital-demo-secret")

DATA_DIR = Path(__file__).resolve().parent / "excel_data"
FILE_HEADERS = {
    "users": [
        "User ID",
        "Full Name",
        "Email",
        "Phone",
        "Role",
        "Password Hash",
        "Registration Date",
        "Account Status",
    ],
    "patients": [
        "Patient ID",
        "Patient Name",
        "Age",
        "Gender",
        "Contact Number",
        "Emergency Contact",
        "Disease Name",
        "Disease Category",
        "Severity",
        "Symptoms",
        "Admission Date",
        "Assigned Doctor",
        "Assigned Nurse",
        "Assigned Bed",
        "Treatment Status",
        "Discharge Status",
        "Estimated Cost",
        "Actual Bill",
        "Discharge Date",
    ],
    "beds": [
        "Bed ID",
        "Ward",
        "Bed Type",
        "Isolation Capability",
        "Equipment Available",
        "Current Patient ID",
        "Bed Status",
        "Daily Charge",
    ],
    "doctors": [
        "Doctor ID",
        "Doctor Name",
        "Specialization",
        "Availability Status",
        "Current Patient Count",
    ],
    "medicines": [
        "Medicine ID",
        "Medicine Name",
        "Available Quantity",
        "Minimum Stock",
        "Unit Price",
        "Expiry Date",
    ],
    "billing": [
        "Bill ID",
        "Patient ID",
        "Bed Charges",
        "Medicine Charges",
        "Doctor Charges",
        "Other Charges",
        "Total Amount",
        "Payment Status",
        "Generated Date",
    ],
    "activity_log": [
        "Timestamp",
        "User ID",
        "Action",
        "Record Type",
        "Record ID",
        "Description",
    ],
}

ROLE_OPTIONS = ["Administrator", "Doctor", "Nurse", "Receptionist"]
DISEASE_CATEGORIES = ["Communicable", "Non-communicable"]
NURSE_POOL = [
    "Nurse Aisha Rahman",
    "Nurse Daniel Okafor",
    "Nurse Priya Nair",
    "Nurse Samuel Mensah",
]

DATA_LOCK = threading.RLock()


def today_iso() -> str:
    return date.today().isoformat()


def ensure_workbook(file_name: str, headers: list[str]):
    DATA_DIR.mkdir(exist_ok=True, parents=True)
    workbook_path = DATA_DIR / f"{file_name}.xlsx"
    if not workbook_path.exists() or workbook_path.stat().st_size == 0:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        sheet.append(headers)
        workbook.save(workbook_path)
        return

    try:
        workbook = load_workbook(workbook_path, read_only=False)
    except Exception:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        sheet.append(headers)
        workbook.save(workbook_path)
        return

    sheet = workbook.active
    actual_headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
    if actual_headers != headers:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        sheet.append(headers)
        workbook.save(workbook_path)


def read_rows(file_name: str):
    ensure_workbook(file_name, FILE_HEADERS[file_name])
    workbook_path = DATA_DIR / f"{file_name}.xlsx"
    workbook = load_workbook(workbook_path, read_only=False)
    sheet = workbook.active
    rows = []
    headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
    if not headers:
        sheet.append(FILE_HEADERS[file_name])
        workbook.save(workbook_path)
        headers = FILE_HEADERS[file_name]

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not any(value is not None and str(value).strip() for value in row):
            continue
        row_data = {headers[index]: row[index] if index < len(row) else "" for index in range(len(headers))}
        rows.append(row_data)
    return rows


def write_rows(file_name: str, rows: list[dict]):
    with DATA_LOCK:
        ensure_workbook(file_name, FILE_HEADERS[file_name])
        workbook_path = DATA_DIR / f"{file_name}.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        sheet.append(FILE_HEADERS[file_name])
        for row in rows:
            values = []
            for header in FILE_HEADERS[file_name]:
                value = row.get(header, "")
                values.append(value)
            sheet.append(values)
        workbook.save(workbook_path)


def log_activity(user_id: str, action: str, record_type: str, record_id: str, description: str):
    rows = read_rows("activity_log")
    rows.append(
        {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "User ID": user_id,
            "Action": action,
            "Record Type": record_type,
            "Record ID": record_id,
            "Description": description,
        }
    )
    write_rows("activity_log", rows)


def seed_demo_data():
    users = read_rows("users")
    if not users:
        password_hash = generate_password_hash("admin123")
        write_rows(
            "users",
            [
                {
                    "User ID": "admin",
                    "Full Name": "System Administrator",
                    "Email": "admin@aihospital.local",
                    "Phone": "+0000000000",
                    "Role": "Administrator",
                    "Password Hash": password_hash,
                    "Registration Date": today_iso(),
                    "Account Status": "Active",
                }
            ],
        )
        log_activity("admin", "Created demo administrator", "user", "admin", "Generated default admin demo account")

    doctors = read_rows("doctors")
    if not doctors:
        doctor_rows = [
            {"Doctor ID": "DOC-101", "Doctor Name": "Dr. Rachel Kim", "Specialization": "Internal Medicine", "Availability Status": "Available", "Current Patient Count": 1},
            {"Doctor ID": "DOC-102", "Doctor Name": "Dr. Omar Hassan", "Specialization": "Pulmonology", "Availability Status": "Available", "Current Patient Count": 0},
            {"Doctor ID": "DOC-103", "Doctor Name": "Dr. Sneh Patel", "Specialization": "Orthopedics", "Availability Status": "Available", "Current Patient Count": 0},
            {"Doctor ID": "DOC-104", "Doctor Name": "Dr. Lydia Wu", "Specialization": "Infectious Diseases", "Availability Status": "Available", "Current Patient Count": 0},
        ]
        write_rows("doctors", doctor_rows)

    beds = read_rows("beds")
    if not beds:
        beds_rows = [
            {"Bed ID": "B-101", "Ward": "Ward A", "Bed Type": "General", "Isolation Capability": "No", "Equipment Available": "Monitor, IV Stand", "Current Patient ID": "", "Bed Status": "Available", "Daily Charge": 2500},
            {"Bed ID": "B-102", "Ward": "Ward A", "Bed Type": "General", "Isolation Capability": "No", "Equipment Available": "Monitor, IV Stand", "Current Patient ID": "", "Bed Status": "Available", "Daily Charge": 2500},
            {"Bed ID": "B-103", "Ward": "Ward A", "Bed Type": "General", "Isolation Capability": "No", "Equipment Available": "Monitor, IV Stand", "Current Patient ID": "", "Bed Status": "Available", "Daily Charge": 2500},
            {"Bed ID": "B-201", "Ward": "Ward B", "Bed Type": "ICU", "Isolation Capability": "No", "Equipment Available": "Ventilator, ECG", "Current Patient ID": "", "Bed Status": "Available", "Daily Charge": 4800},
            {"Bed ID": "B-202", "Ward": "Ward B", "Bed Type": "ICU", "Isolation Capability": "No", "Equipment Available": "Ventilator, ECG", "Current Patient ID": "", "Bed Status": "Available", "Daily Charge": 4800},
            {"Bed ID": "B-301", "Ward": "Isolation", "Bed Type": "Isolation", "Isolation Capability": "Yes", "Equipment Available": "HEPA Filter, PPE Kit", "Current Patient ID": "", "Bed Status": "Available", "Daily Charge": 5400},
            {"Bed ID": "B-302", "Ward": "Isolation", "Bed Type": "Isolation", "Isolation Capability": "Yes", "Equipment Available": "HEPA Filter, PPE Kit", "Current Patient ID": "", "Bed Status": "Available", "Daily Charge": 5400},
            {"Bed ID": "B-303", "Ward": "Isolation", "Bed Type": "Isolation", "Isolation Capability": "Yes", "Equipment Available": "HEPA Filter, PPE Kit", "Current Patient ID": "", "Bed Status": "Available", "Daily Charge": 5400},
            {"Bed ID": "B-401", "Ward": "Private", "Bed Type": "Private", "Isolation Capability": "No", "Equipment Available": "Monitor, Recliner", "Current Patient ID": "", "Bed Status": "Available", "Daily Charge": 6200},
            {"Bed ID": "B-402", "Ward": "Private", "Bed Type": "Private", "Isolation Capability": "No", "Equipment Available": "Monitor, Recliner", "Current Patient ID": "", "Bed Status": "Available", "Daily Charge": 6200},
        ]
        write_rows("beds", beds_rows)

    medicines = read_rows("medicines")
    if not medicines:
        medicine_rows = [
            {"Medicine ID": "MED-001", "Medicine Name": "Paracetamol", "Available Quantity": 110, "Minimum Stock": 40, "Unit Price": 15, "Expiry Date": "2027-12-31"},
            {"Medicine ID": "MED-002", "Medicine Name": "Amoxicillin", "Available Quantity": 60, "Minimum Stock": 30, "Unit Price": 28, "Expiry Date": "2026-09-15"},
            {"Medicine ID": "MED-003", "Medicine Name": "Insulin", "Available Quantity": 30, "Minimum Stock": 20, "Unit Price": 90, "Expiry Date": "2027-08-20"},
            {"Medicine ID": "MED-004", "Medicine Name": "Azithromycin", "Available Quantity": 24, "Minimum Stock": 25, "Unit Price": 48, "Expiry Date": "2026-11-30"},
            {"Medicine ID": "MED-005", "Medicine Name": "Salbutamol Inhaler", "Available Quantity": 18, "Minimum Stock": 20, "Unit Price": 55, "Expiry Date": "2027-06-15"},
        ]
        write_rows("medicines", medicine_rows)

    if not read_rows("billing"):
        write_rows("billing", [])

    if not read_rows("activity_log"):
        write_rows(
            "activity_log",
            [
                {
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "User ID": "admin",
                    "Action": "System startup",
                    "Record Type": "application",
                    "Record ID": "startup",
                    "Description": "Initial workbook setup completed.",
                }
            ],
        )


def get_user_by_id(user_id: str):
    for user in read_rows("users"):
        if str(user.get("User ID", "")).strip().lower() == str(user_id).strip().lower():
            return user
    return None


def get_current_user():
    user_id = session.get("user_id")
    if not user_id:
        return None
    user = get_user_by_id(user_id)
    if user and str(user.get("Account Status", "")).lower() == "active":
        return user
    session.clear()
    return None


def get_dashboard_metrics():
    patients = read_rows("patients")
    beds = read_rows("beds")
    doctors = read_rows("doctors")
    medicines = read_rows("medicines")

    total_patients = len(patients)
    admitted_patients = sum(1 for row in patients if str(row.get("Treatment Status", "")).strip().lower() == "admitted")
    available_beds = sum(1 for row in beds if str(row.get("Bed Status", "")).strip().lower() == "available")
    occupied_beds = sum(1 for row in beds if str(row.get("Bed Status", "")).strip().lower() == "occupied")
    isolation_available = sum(1 for row in beds if str(row.get("Isolation Capability", "")).strip().lower() == "yes" and str(row.get("Bed Status", "")).strip().lower() == "available")
    pending_discharges = sum(1 for row in patients if str(row.get("Discharge Status", "")).strip().lower() in {"pending", "requested"})
    available_doctors = sum(1 for row in doctors if str(row.get("Availability Status", "")).strip().lower() == "available")
    low_stock_medicines = sum(1 for row in medicines if float(row.get("Available Quantity", 0) or 0) < float(row.get("Minimum Stock", 0) or 0))

    revenue_value = 0.0
    for row in patients:
        value = row.get("Estimated Cost", "0")
        if value not in (None, ""):
            try:
                revenue_value += float(value)
            except (TypeError, ValueError):
                pass

    disease_counts = {}
    for row in patients:
        disease_name = str(row.get("Disease Name", "Unknown")).strip() or "Unknown"
        disease_counts[disease_name] = disease_counts.get(disease_name, 0) + 1

    return {
        "total_patients": total_patients,
        "currently_admitted": admitted_patients,
        "available_beds": available_beds,
        "occupied_beds": occupied_beds,
        "isolation_beds_available": isolation_available,
        "pending_discharges": pending_discharges,
        "available_doctors": available_doctors,
        "low_stock_medicines": low_stock_medicines,
        "todays_estimated_revenue": round(revenue_value, 2),
        "disease_counts": disease_counts,
    }


def find_available_bed(category: str, severity: str):
    beds = read_rows("beds")
    severity_weight = {"Low": 1, "Moderate": 2, "High": 3, "Critical": 4}.get(str(severity).strip().title(), 1)
    preferred = []
    for bed in beds:
        bed_status = str(bed.get("Bed Status", "")).strip().lower()
        isolation_flag = str(bed.get("Isolation Capability", "")).strip().lower()
        if bed_status != "available":
            continue
        if category == "Communicable" and isolation_flag == "yes":
            preferred.append(bed)
        elif category == "Communicable" and severity_weight >= 3:
            preferred.append(bed)
        else:
            preferred.append(bed)
    if not preferred:
        return None
    preferred.sort(key=lambda item: (0 if str(item.get("Isolation Capability", "")).strip().lower() == "yes" else 1, float(item.get("Daily Charge", 0) or 0)))
    return preferred[0]


def get_available_doctor():
    doctors = read_rows("doctors")
    for doctor in sorted(doctors, key=lambda row: int(row.get("Current Patient Count", 0) or 0)):
        if str(doctor.get("Availability Status", "")).strip().lower() == "available":
            return doctor
    return None


def assign_patient_to_doctor(doctor_name: str | None = None):
    doctors = read_rows("doctors")
    if doctor_name:
        for doctor in doctors:
            if str(doctor.get("Doctor Name", "")).strip() == doctor_name:
                return doctor
    available_doctor = get_available_doctor()
    if not available_doctor:
        return None
    for doctor in doctors:
        if str(doctor.get("Doctor ID", "")).strip() == str(available_doctor.get("Doctor ID", "")).strip():
            doctor["Current Patient Count"] = int(doctor.get("Current Patient Count", 0) or 0) + 1
            doctor["Availability Status"] = "Available"
            break
    write_rows("doctors", doctors)
    return available_doctor


def release_doctor(doctor_name: str):
    doctors = read_rows("doctors")
    for doctor in doctors:
        if str(doctor.get("Doctor Name", "")).strip() == doctor_name:
            count = int(doctor.get("Current Patient Count", 0) or 0)
            doctor["Current Patient Count"] = max(0, count - 1)
            break
    write_rows("doctors", doctors)


def get_next_nurse():
    nurse_list = list(NURSE_POOL)
    current = session.get("nurse_index")
    if current is None:
        session["nurse_index"] = 0
        current = 0
    nurse = nurse_list[current % len(nurse_list)]
    session["nurse_index"] = (current + 1) % len(nurse_list)
    return nurse


def create_bill_for_patient(patient_id: str):
    patients = read_rows("patients")
    beds = read_rows("beds")
    medicines = read_rows("medicines")
    patient = next((row for row in patients if str(row.get("Patient ID", "")).strip() == str(patient_id).strip()), None)
    if patient is None:
        return None

    bed = next((row for row in beds if str(row.get("Bed ID", "")).strip() == str(patient.get("Assigned Bed", "")).strip()), None)
    days = 1
    bed_charges = float(bed.get("Daily Charge", 0) or 0) * days if bed else 0
    medicine_charges = 0.0
    for medicine in medicines:
        if medicine.get("Medicine Name", "").lower() in {"paracetamol", "amoxicillin", "azithromycin", "salbutamol inhaler", "insulin"}:
            medicine_charges += float(medicine.get("Unit Price", 0) or 0) * 1.5
    doctor_charges = 1800 if patient.get("Disease Category", "") == "Communicable" else 1200
    other_charges = 450 if patient.get("Disease Category", "") == "Communicable" else 250
    total = bed_charges + medicine_charges + doctor_charges + other_charges

    bill_id = f"BILL-{patient_id}"
    billing_rows = read_rows("billing")
    existing_index = None
    for index, item in enumerate(billing_rows):
        if str(item.get("Bill ID", "")).strip() == bill_id:
            existing_index = index
            break

    bill_record = {
        "Bill ID": bill_id,
        "Patient ID": patient_id,
        "Bed Charges": round(bed_charges, 2),
        "Medicine Charges": round(medicine_charges, 2),
        "Doctor Charges": round(doctor_charges, 2),
        "Other Charges": round(other_charges, 2),
        "Total Amount": round(total, 2),
        "Payment Status": "Pending",
        "Generated Date": today_iso(),
    }

    if existing_index is not None:
        billing_rows[existing_index] = bill_record
    else:
        billing_rows.append(bill_record)
    write_rows("billing", billing_rows)

    patient["Estimated Cost"] = round(total, 2)
    patient["Actual Bill"] = round(total, 2)
    for idx, row in enumerate(patients):
        if str(row.get("Patient ID", "")).strip() == str(patient_id).strip():
            patients[idx] = patient
            break
    write_rows("patients", patients)
    return bill_record


@app.before_request
def require_login():
    if request.endpoint in {"login", "register", "static"}:
        return None
    if not session.get("user_id"):
        return redirect(url_for("login"))


@app.context_processor
def inject_current_user():
    return {"current_user": get_current_user()}


@app.route("/")
def index():
    if session.get("user_id"):
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user_id = request.form.get("user_id", "").strip()
        password = request.form.get("password", "")
        if not user_id or not password:
            flash("User ID and password are required.", "error")
            return render_template("login.html")

        user = get_user_by_id(user_id)
        if not user or not check_password_hash(str(user.get("Password Hash", "")), password):
            flash("Invalid credentials. Please try again.", "error")
            return render_template("login.html")

        if str(user.get("Account Status", "")).lower() != "active":
            flash("This account is inactive. Please contact the administrator.", "error")
            return render_template("login.html")

        session["user_id"] = user["User ID"]
        flash("Login successful.", "success")
        return redirect(url_for("dashboard"))

    return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        user_id = request.form.get("user_id", "").strip()
        email = request.form.get("email", "").strip()
        phone = request.form.get("phone", "").strip()
        role = request.form.get("role", "").strip()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")

        if not all([full_name, user_id, email, phone, role, password, confirm_password]):
            flash("All fields are required.", "error")
            return render_template("register.html")

        if password != confirm_password:
            flash("Passwords do not match.", "error")
            return render_template("register.html")

        if len(password) < 6:
            flash("Password must be at least 6 characters long.", "error")
            return render_template("register.html")

        if not re.match(r"^[A-Za-z0-9_@.-]+$", user_id):
            flash("User ID can only contain letters, numbers, underscores, dots, hyphens, and @.", "error")
            return render_template("register.html")

        if get_user_by_id(user_id):
            flash("User ID already exists. Please choose another one.", "error")
            return render_template("register.html")

        if role not in ROLE_OPTIONS:
            flash("Choose a valid role.", "error")
            return render_template("register.html")

        users = read_rows("users")
        users.append(
            {
                "User ID": user_id,
                "Full Name": full_name,
                "Email": email,
                "Phone": phone,
                "Role": role,
                "Password Hash": generate_password_hash(password),
                "Registration Date": today_iso(),
                "Account Status": "Active",
            }
        )
        write_rows("users", users)
        log_activity(user_id, "Registered user account", "user", user_id, f"Created account for {full_name} with role {role}.")
        flash("Registration successful. You can now log in.", "success")
        return redirect(url_for("login"))

    return render_template("register.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Logged out successfully.", "success")
    return redirect(url_for("login"))


@app.route("/dashboard")
def dashboard():
    metrics = get_dashboard_metrics()
    patients = read_rows("patients")
    disease_labels = list(metrics["disease_counts"].keys())
    disease_values = list(metrics["disease_counts"].values())
    beds = read_rows("beds")
    bed_status_labels = ["Available", "Occupied"]
    bed_status_values = [sum(1 for row in beds if str(row.get("Bed Status", "")).strip().lower() == "available"), sum(1 for row in beds if str(row.get("Bed Status", "")).strip().lower() == "occupied")]
    summary_rows = [
        {"label": "Total Patients", "value": metrics["total_patients"], "icon": "fa-user-injured", "color": "primary"},
        {"label": "Currently Admitted", "value": metrics["currently_admitted"], "icon": "fa-bed", "color": "warning"},
        {"label": "Available Beds", "value": metrics["available_beds"], "icon": "fa-hospital", "color": "success"},
        {"label": "Occupied Beds", "value": metrics["occupied_beds"], "icon": "fa-procedures", "color": "info"},
        {"label": "Isolation Beds Available", "value": metrics["isolation_beds_available"], "icon": "fa-shield-virus", "color": "danger"},
        {"label": "Pending Discharges", "value": metrics["pending_discharges"], "icon": "fa-sign-out-alt", "color": "secondary"},
        {"label": "Available Doctors", "value": metrics["available_doctors"], "icon": "fa-user-md", "color": "dark"},
        {"label": "Low-Stock Medicines", "value": metrics["low_stock_medicines"], "icon": "fa-pills", "color": "danger"},
        {"label": "Today's Estimated Revenue", "value": f"${metrics['todays_estimated_revenue']:.2f}", "icon": "fa-dollar-sign", "color": "success"},
    ]
    return render_template("dashboard.html", metrics=metrics, cards=summary_rows, patients=patients[:5], disease_labels=disease_labels, disease_values=disease_values, bed_status_labels=bed_status_labels, bed_status_values=bed_status_values, active_page="Overview")


@app.route("/patients")
def patients():
    rows = read_rows("patients")
    return render_template("patients.html", patients=rows, active_page="Patient Management")


@app.route("/patients/admit", methods=["GET", "POST"])
def admit_patient():
    if request.method == "POST":
        patient_id = request.form.get("patient_id", "").strip()
        patient_name = request.form.get("patient_name", "").strip()
        age = request.form.get("age", "").strip()
        gender = request.form.get("gender", "").strip()
        contact_number = request.form.get("contact_number", "").strip()
        emergency_contact = request.form.get("emergency_contact", "").strip()
        disease_name = request.form.get("disease_name", "").strip()
        disease_category = request.form.get("disease_category", "").strip()
        severity = request.form.get("severity", "").strip()
        symptoms = request.form.get("symptoms", "").strip()

        if not all([patient_id, patient_name, age, gender, contact_number, disease_name, disease_category, severity]):
            flash("Please complete all required patient fields.", "error")
            return render_template("admit_patient.html", active_page="Patient Admission")

        patient_rows = read_rows("patients")
        if any(str(row.get("Patient ID", "")).strip() == patient_id for row in patient_rows):
            flash("A patient with this ID already exists.", "error")
            return render_template("admit_patient.html", active_page="Patient Admission")

        if disease_category not in DISEASE_CATEGORIES:
            flash("Please choose a valid disease category.", "error")
            return render_template("admit_patient.html", active_page="Patient Admission")

        bed = find_available_bed(disease_category, severity)
        if not bed:
            flash("No suitable bed is currently available for this patient. Please try again later.", "error")
            return render_template("admit_patient.html", active_page="Patient Admission")

        doctor = assign_patient_to_doctor()
        nurse = get_next_nurse()
        estimated_cost = float(bed.get("Daily Charge", 0) or 0) * 1.5
        if disease_category == "Communicable":
            estimated_cost += 1800
        else:
            estimated_cost += 900

        patient_rows.append(
            {
                "Patient ID": patient_id,
                "Patient Name": patient_name,
                "Age": age,
                "Gender": gender,
                "Contact Number": contact_number,
                "Emergency Contact": emergency_contact,
                "Disease Name": disease_name,
                "Disease Category": disease_category,
                "Severity": severity,
                "Symptoms": symptoms,
                "Admission Date": today_iso(),
                "Assigned Doctor": doctor["Doctor Name"] if doctor else "Unassigned",
                "Assigned Nurse": nurse,
                "Assigned Bed": bed["Bed ID"],
                "Treatment Status": "Admitted",
                "Discharge Status": "Pending",
                "Estimated Cost": round(estimated_cost, 2),
                "Actual Bill": "",
                "Discharge Date": "",
            }
        )
        write_rows("patients", patient_rows)

        beds_rows = read_rows("beds")
        for index, bed_row in enumerate(beds_rows):
            if str(bed_row.get("Bed ID", "")).strip() == str(bed["Bed ID"]).strip():
                bed_row["Current Patient ID"] = patient_id
                bed_row["Bed Status"] = "Occupied"
                beds_rows[index] = bed_row
                break
        write_rows("beds", beds_rows)

        log_activity(
            session.get("user_id", "system"),
            "Admitted patient",
            "patient",
            patient_id,
            f"Admitted {patient_name} to bed {bed['Bed ID']} with {disease_category} disease category.",
        )
        flash("Patient admitted successfully and a suitable bed was assigned.", "success")
        return redirect(url_for("patients"))

    return render_template("admit_patient.html", active_page="Patient Admission")


@app.route("/beds")
def beds():
    rows = read_rows("beds")
    return render_template("beds.html", beds=rows, active_page="Intelligent Bed Assignment")


@app.route("/discharge", methods=["GET", "POST"])
def discharge():
    patients_rows = read_rows("patients")
    if request.method == "POST":
        patient_id = request.form.get("patient_id", "").strip()
        patient = next((row for row in patients_rows if str(row.get("Patient ID", "")).strip() == patient_id), None)
        if not patient:
            flash("Patient record not found.", "error")
            return redirect(url_for("discharge"))

        bed_id = patient.get("Assigned Bed", "")
        patient["Treatment Status"] = "Discharged"
        patient["Discharge Status"] = "Completed"
        patient["Discharge Date"] = today_iso()
        patient["Estimated Cost"] = patient.get("Estimated Cost", 0) or 0

        for index, row in enumerate(patients_rows):
            if str(row.get("Patient ID", "")).strip() == patient_id:
                patients_rows[index] = patient
                break
        write_rows("patients", patients_rows)

        bed_rows = read_rows("beds")
        for index, row in enumerate(bed_rows):
            if str(row.get("Bed ID", "")).strip() == str(bed_id).strip():
                row["Current Patient ID"] = ""
                row["Bed Status"] = "Available"
                bed_rows[index] = row
                break
        write_rows("beds", bed_rows)

        if patient.get("Assigned Doctor"):
            release_doctor(patient.get("Assigned Doctor"))

        bill = create_bill_for_patient(patient_id)
        log_activity(
            session.get("user_id", "system"),
            "Discharged patient",
            "patient",
            patient_id,
            f"Discharged {patient.get('Patient Name')} and created billing record {bill['Bill ID'] if bill else 'N/A'}.",
        )
        flash("Patient discharge recorded successfully.", "success")
        return redirect(url_for("discharge"))

    active_patients = [row for row in patients_rows if str(row.get("Treatment Status", "")).strip().lower() != "discharged"]
    return render_template("discharge.html", patients=active_patients, active_page="Patient Discharge")


@app.route("/doctors")
def doctors():
    doctor_rows = read_rows("doctors")
    return render_template("doctors.html", doctors=doctor_rows, active_page="Doctors and Nurses")


@app.route("/medicines")
def medicines():
    rows = read_rows("medicines")
    return render_template("medicines.html", medicines=rows, active_page="Medicine Inventory")


@app.route("/billing")
def billing():
    billing_rows = read_rows("billing")
    return render_template("billing.html", billing=billing_rows, active_page="Billing and Payments")


@app.route("/reports")
def reports():
    metrics = get_dashboard_metrics()
    patients = read_rows("patients")
    return render_template("reports.html", metrics=metrics, patients=patients, active_page="Reports and Analytics")


@app.route("/activity")
def activity():
    rows = read_rows("activity_log")
    rows = sorted(rows, key=lambda r: str(r.get("Timestamp", "")).strip(), reverse=True)
    return render_template("activity_log.html", activities=rows, active_page="Activity Logs")


@app.route("/settings")
def settings():
    user = get_current_user()
    return render_template("settings.html", user=user, active_page="Settings")


@app.route("/api/optimization")
def optimization_api():
    beds = read_rows("beds")
    patients = read_rows("patients")
    recommendations = []
    for patient in patients:
        if str(patient.get("Treatment Status", "")).strip().lower() == "admitted":
            continue
        category = patient.get("Disease Category", "Non-communicable")
        severity = patient.get("Severity", "Moderate")
        bed = find_available_bed(category, severity)
        recommendations.append({
            "patient_id": patient.get("Patient ID"),
            "patient_name": patient.get("Patient Name"),
            "category": category,
            "severity": severity,
            "recommended_bed": bed.get("Bed ID") if bed else "None available",
        })
    return jsonify({"recommendations": recommendations, "available_beds": sum(1 for row in beds if str(row.get("Bed Status", "")).strip().lower() == "available")})


seed_demo_data()


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
