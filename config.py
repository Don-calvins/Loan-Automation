# config.py

# SQL SERVER DATABASE SETTINGS

DB_SERVER = "YOUR_SERVER_NAME"      # Example: 192.168.1.50 OR SERVER\\SQLEXPRESS
DB_NAME = "YOUR_DATABASE_NAME"
DB_USERNAME = "YOUR_READONLY_USERNAME"
DB_PASSWORD = "YOUR_PASSWORD"

# If you don't know the driver, use this common one:
DB_DRIVER = "ODBC Driver 17 for SQL Server"
# You can also try:
# DB_DRIVER = "ODBC Driver 18 for SQL Server"

TABLE_NAME = "dbo.Loans"  # Change to your real table

# -----------------------------
# SMTP SETTINGS (COMPANY SMTP)
# -----------------------------
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587  # 587 (TLS) or 465 (SSL)
SMTP_USERNAME = "info@maishaborasacco.com"
SMTP_PASSWORD = "YOUR_EMAIL_PASSWORD"
SMTP_USE_TLS = True  # True for 587 TLS, False for SSL 465

FROM_EMAIL = "info@maishaborasacco.com"
FROM_NAME = "Loan Due Report"

TO_EMAIL = "creditloans@company.com"
TO_NAME = "Credit & Loans Department"

