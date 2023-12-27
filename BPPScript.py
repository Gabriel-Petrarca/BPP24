import smtplib

# Define the email server and port
SMTP_SERVER = 'smtp.gmail.com'  # Example for Gmail
SMTP_PORT = 587  # Example for Gmail's TLS port

# Your email credentials
FROM = 'petrigabo07@gmail.com'
PASSWORD = 'qvdb qqbo srob zdsj'  # You should use an 'App Password' or enable 'Less secure apps' for this to work

TO = ["dlperry008@gmail.com"]  # Must be a list

SUBJECT = "Hello!"
TEXT = "This python email is so cool."

# Prepare the email message
message = f"""\
From: {FROM}
To: {", ".join(TO)}
Subject: {SUBJECT}

{TEXT}
"""

# Try to establish a connection to the SMTP server and send the email
try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()  # Upgrade the connection to a secure TLS connection
    server.login(FROM, PASSWORD)
    server.sendmail(FROM, TO, message)
    server.quit()
    print("Email sent successfully")
except smtplib.SMTPException as e:
    print(f"An error occurred: {e}")
