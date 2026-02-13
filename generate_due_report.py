import os
import csv
import zipfile
import smtplib
import pyodbc
from datetime import datetime, timedelta
from email.message import EmailMessage

import config


def connect_sql_server():
    conn_str = (
        f"DRIVER={{{config.DB_DRIVER}}};"
        f"SERVER={config.DB_SERVER};"
        f"DATABASE={config.DB_NAME};"
        f"UID={config.DB_USERNAME};"
        f"PWD={config.DB_PASSWORD};"
        "TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str)


def fetch_loans_due_next_7_days(conn):
    query = f"""
        SELECT member_number, member_name, due_date, loan_amount
        FROM {config.TABLE_NAME}
        WHERE due_date >= CAST(GETDATE() AS DATE)
          AND due_date <= DATEADD(DAY, 7, CAST(GETDATE() AS DATE))
        ORDER BY due_date ASC
    """
    cursor = conn.cursor()
    cursor.execute(query)
    rows = cursor.fetchall()
    return rows


def create_report_folder(base_dir, today_str):
    reports_dir = os.path.join(base_dir, "reports")
    os.makedirs(reports_dir, exist_ok=True)

    folder_name = f"LoanDueReport_{today_str}"
    folder_path = os.path.join(reports_dir, folder_name)
    os.makedirs(folder_path, exist_ok=True)

    return folder_name, folder_path


def generate_csv(folder_path, today_str, loans):
    csv_filename = f"Loans_Due_Next_7_Days_{today_str}.csv"
    csv_path = os.path.join(folder_path, csv_filename)

    total_amount = 0

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Member Number", "Member Name", "Due Date", "Loan Amount"])

        for row in loans:
            member_number = row[0]
            member_name = row[1]
            due_date = row[2]
            loan_amount = float(row[3])

            # Convert SQL Server date to string
            if hasattr(due_date, "strftime"):
                due_date = due_date.strftime("%Y-%m-%d")

            total_amount += loan_amount

            writer.writerow([member_number, member_name, due_date, f"{loan_amount:,.2f}"])

    return csv_path, len(loans), total_amount


def zip_folder(reports_dir, folder_name, folder_path):
    zip_path = os.path.join(reports_dir, f"{folder_name}.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                full_path = os.path.join(root, file)
                relative_path = os.path.relpath(full_path, folder_path)
                zipf.write(full_path, arcname=relative_path)

    return zip_path


def send_email(zip_path, today_str, total_loans, total_amount):
    msg = EmailMessage()
    msg["Subject"] = f"Loan Due Report (Next 7 Days) - {today_str}"
    msg["From"] = f"{config.FROM_NAME} <{config.FROM_EMAIL}>"
    msg["To"] = f"{config.TO_NAME} <{config.TO_EMAIL}>"

    body = f"""Hello Credit & Loans Team,

Attached is the weekly report for loans due within the next 7 days.

Report Date: {today_str}
Total Loans Due: {total_loans}
Total Loan Amount: {total_amount:,.2f}

Regards,
Loan Report Bot
"""
    msg.set_content(body)

    # Attach ZIP
    with open(zip_path, "rb") as f:
        zip_data = f.read()

    msg.add_attachment(
        zip_data,
        maintype="application",
        subtype="zip",
        filename=os.path.basename(zip_path)
    )

    # SMTP send
    if config.SMTP_USE_TLS:
        with smtplib.SMTP(config.SMTP_HOST, config.SMTP_PORT) as server:
            server.starttls()
            server.login(config.SMTP_USERNAME, config.SMTP_PASSWORD)
            server.send_message(msg)
    else:
        with smtplib.SMTP_SSL(config.SMTP_HOST, config.SMTP_PORT) as server:
            server.login(config.SMTP_USERNAME, config.SMTP_PASSWORD)
            server.send_message(msg)


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    today_str = datetime.now().strftime("%Y-%m-%d")

    print("Connecting to SQL Server...")
    conn = connect_sql_server()

    print("Fetching loans due in next 7 days...")
    loans = fetch_loans_due_next_7_days(conn)
    conn.close()

    folder_name, folder_path = create_report_folder(base_dir, today_str)

    print("Generating CSV report...")
    csv_path, total_loans, total_amount = generate_csv(folder_path, today_str, loans)

    reports_dir = os.path.join(base_dir, "reports")

    print("Zipping report folder...")
    zip_path = zip_folder(reports_dir, folder_name, folder_path)

    print("Sending email via SMTP...")
    send_email(zip_path, today_str, total_loans, total_amount)

    print("✅ DONE: Report generated + zipped + emailed successfully!")
    print(f"CSV: {csv_path}")
    print(f"ZIP: {zip_path}")


if __name__ == "__main__":
    main()
