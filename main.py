import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Load the recipient data from CSV
data = pd.read_csv('recipents.csv')

# Email settings
smtp_server = 'smtp.gmail.com'
smtp_port = 587
email_user = '#'  # Your Gmail address
email_password = '#'  # Your Gmail "app password"

# Set up the SMTP server
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(email_user, email_password)

# Function to send email
def send_email(to_email, subject, body, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = to_email
    msg['Subject'] = subject
    
    msg.attach(MIMEText(body, 'plain'))
    
    if attachment_path and os.path.isfile(attachment_path):
        print(f"Attaching file: {attachment_path}")
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
            msg.attach(part)
    else:
        print(f"Attachment file not found: {attachment_path}")
    
    text = msg.as_string()
    server.sendmail(email_user, to_email, text)
    print(f"Email sent to {to_email}")

# Iterate through the data and send emails
for index, row in data.iterrows():
    name = row['Name']
    to_email = row['Email']
    certificate_file = row['Certificate']
    
    # Source path of the certificates
    certificate_path = f'certificate/{certificate_file}' if pd.notna(certificate_file) else None
    
    # Subject of the Mail
    subject = "Certificate of Participation"
    
    # Body of the mail
    body = f"Good Evening {name},\n\nThis is your certificate of participation in CME Sleep Physiology.\n\nThank you."
     
    send_email(to_email, subject, body, certificate_path)

# Close the SMTP server
server.quit()
