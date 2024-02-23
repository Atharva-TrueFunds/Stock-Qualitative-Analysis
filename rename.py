import ssl
import smtplib
from email.message import EmailMessage

email_sender = "atharvachoudhari.truefunds@gmail.com"
email_password = "urnp mvgv rnei vgrr"
email_receiver = "atharvachoudhari06@gmail.com"

subject = "Subject of your email"
body = "Body of your email"

# Create an EmailMessage object
em = EmailMessage()
em['From'] = email_sender
em['To'] = email_receiver
em['Subject'] = subject
em.set_content(body)

# Attach files to the email
files = [
    "C:/Users/HP/Staging/currentValue_Analysis.xlsx",
    "C:/Users/HP/Staging/52WeekHigh_Low.xlsx",
    "C:/Users/HP/Staging/price_Change_day_to_day.xlsx",
    "C:/Users/HP/Staging/price_Change_day_to_each_day.xlsx"
]  # Add the paths to your files

for file in files:
    with open(file, "rb") as f:
        file_data = f.read()
        file_name = f.name.split('/')[-1]  # Extracting the file name from the path
    em.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

# Connect to Gmail's SMTP server securely
context = ssl.create_default_context()
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(email_sender, email_password)
    smtp.send_message(em)
