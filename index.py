import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import zipfile
import os

# Step 1: Load and filter data
def extract_due_loans(data_path, output_file='due_loans.csv'):
    # Assume data is in a CSV with columns like 'loan_id', 'due_date' (YYYY-MM-DD), 'amount', etc.
    df = pd.read_csv(data_path)
    df['due_date'] = pd.to_datetime(df['due_date'])
    
    # Filter for loans due in the next 7 days
    today = datetime.now()
    one_week_later = today + timedelta(days=7)
    due_loans = df[(df['due_date'] >= today) & (df['due_date'] <= one_week_later)]
    
    # Save to CSV
    due_loans.to_csv(output_file, index=False)
    return output_file

# Step 2: Create ZIP folder
def create_zip(file_to_zip, zip_name='loan_data.zip'):
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        zipf.write(file_to_zip, os.path.basename(file_to_zip))
    return zip_name

# Step 3: Send email with attachment
def send_email(zip_file, recipient_email, sender_email, sender_password, smtp_server='smtp.gmail.com', smtp_port=587):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = 'Weekly Loan Due Report'
    
    # Attach ZIP
    with open(zip_file, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(zip_file)}')
        msg.attach(part)
    
    # Send via SMTP
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)
    server.sendmail(sender_email, recipient_email, msg.as_string())
    server.quit()

# Main execution
if __name__ == '__main__':
    data_path = 'path/to/your/financial_data.csv'  # Replace with your data source
    recipient = 'credit.loans@company.com'  # Replace with actual email
    sender = 'your.email@example.com'  # Replace with your email
    password = 'your_app_password'  # Use app password for Gmail, etc.
    
    # Extract and package
    csv_file = extract_due_loans(data_path)
    zip_file = create_zip(csv_file)
    
    # Send email
    send_email(zip_file, recipient, sender, password)
    
    # Cleanup
    os.remove(csv_file)
    os.remove(zip_file)
