import csv
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Define your Titan email hosting details
titan_server = 'smtp.titan.email'
titan_port = 587
titan_username = 'contact@improdata.com'
titan_password = '12345N@b'

# Define email bodies for different job titles
email_bodies = {
    "Data Analyst": "Hey I want to apply as data analyst",
    "Business Intelligence": "Hi I want to apply for business intelligence post ",
    "Data Scientist": "Hi I want to apply as Data Scientist ",
    # Add other job titles and bodies as needed
}

# Email validation function
def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email)

def send_email(receiver_email, subject, body):
    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = titan_username
    message['To'] = receiver_email
    message['Subject'] = subject

    # Add body to the email
    message.attach(MIMEText(body, 'plain'))

    # Create SMTP session
    with smtplib.SMTP(titan_server, titan_port) as server:
        server.starttls()
        server.login(titan_username, titan_password)
        text = message.as_string()
        server.sendmail(titan_username, receiver_email, text)

# Read CSV and send emails
with open("C:\\Users\\Electro Store\\Downloads\\recipients.csv", newline='', encoding='ISO-8859-1') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        email = row['email']
        job_title = row['Job Title']
        if is_valid_email(email):
            subject = f"Application for {job_title}"
            body = email_bodies.get(job_title, "Default body if job title not found")
            send_email(email, subject, body)
        else:
            print(f"Invalid email address: {email}")
