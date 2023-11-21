import csv
import smtplib
import re
import imaplib
import time
import mimetypes
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from docx import Document  # Make sure to install python-docx

# Titan email hosting details
titan_smtp_server = 'smtp.titan.email'
titan_imap_server = 'imap.titan.email'
titan_port = 587
titan_username = 'contact@improdata.com'
titan_password = '12345N@b'

# Specify the absolute path for the "sent_emails.csv" file
sent_emails_file_path = r'F:\sent emails\sent_emails.csv'

# Dictionary mapping keywords to their respective file paths, now with .docs instead of .txt
email_body_files = {
    "analyst": "F:\\proposals\\Analyst.docx",
    "business intelligence": "F:\\proposals\\Business Intelligence.docx",
    "data scientist": "F:\\proposals\\Data_Scientist.docx",
    "power bi": "F:\\proposals\\Power_BI.docx",
    # Add other keywords and file paths as needed
}

# Function to read .docs file content
def read_docs_file(file_path):
    try:
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return ""

# Function to select email body based on job title, now using read_docs_file function
def select_email_body(job_title):
    job_title_lower = job_title.lower()
    for keyword, file_path in email_body_files.items():
        if keyword in job_title_lower:
            content = read_docs_file(file_path)
            if content:
                formatted_content = content.replace('\n\n', '</p><p>').replace('\n', '<br>')
                return f"<p>{formatted_content}</p>"
            else:
                return "<p>Default body if job title not found</p>"
    return "<p>Default body if job title not found</p>"

# Email validation function
def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email)

# Function to attach files
def attach_file(message, filepath, custom_filename=None):
    ctype, encoding = mimetypes.guess_type(filepath)
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    with open(filepath, 'rb') as fp:
        attachment = MIMEBase(maintype, subtype)
        attachment.set_payload(fp.read())
        encoders.encode_base64(attachment)
        disposition = f'attachment; filename="{custom_filename or filepath.split("/")[-1]}"'
        attachment.add_header('Content-Disposition', disposition)
    message.attach(attachment)

# Function to save email to the "Sent" folder using IMAP and update the CSV file
def save_sent_email(imap_server, imap_user, imap_password, email_message, sent_emails_file):
    with imaplib.IMAP4_SSL(imap_server) as imap:
        imap.login(imap_user, imap_password)
        imap.append('Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), email_message.as_bytes())
        imap.logout()

    # Update the CSV file to mark the email as sent
    with open(sent_emails_file, mode='a', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow([email_message['To'], 'Sent'])

# Function to send email
def send_email(receiver_email, subject, body, attachments=[], logo_path=None, company_name=None):
    # Create a MIMEMultipart message with 'mixed' type for mixed content (attachments and inline)
    message = MIMEMultipart('mixed')
    message['From'] = f"Improdata <{titan_username}>"
    message['To'] = receiver_email
    message['Subject'] = subject

    # Create an 'alternative' part for plain text and HTML
    alt_part = MIMEMultipart('alternative')
    message.attach(alt_part)

    # Create a 'related' part for HTML content and inline images
    html_part = MIMEMultipart('related')
    alt_part.attach(html_part)

    # HTML body with resized and positioned sticker (logo), and the name/title
    if company_name:
        html_body_content = f"""
        <html>
            <body>
                Hi {company_name} Team,<br>
                {body}
                <div style="text-align: left; margin-top: 50px;">
                    <p><b>Jahanzeb Khan</b><br>Co-Founder</p>
                    <img src="cid:logo" style="width: 100px; height: auto;">
                </div>
            </body>
        </html>
        """
    else:
        html_body_content = f"""
        <html>
            <body>
                {body}
                <div style="text-align: left; margin-top: 50px;">
                    <p><b>Jahanzeb Khan</b><br>Co-Founder</p>
                    <img src="cid:logo" style="width: 100px; height: auto;">
                </div>
            </body>
        </html>
        """

    html_body_part = MIMEText(html_body_content, 'html')
    html_part.attach(html_body_part)

    # Attach logo as an inline image for the signature
    if logo_path:
        with open(logo_path, 'rb') as fp:
            logo_attachment = MIMEImage(fp.read(), _subtype="jpeg")
            logo_attachment.add_header('Content-ID', '<logo>')
            logo_attachment.add_header('Content-Disposition', 'inline', filename="logo")
        html_part.attach(logo_attachment)

    # Attach other files
    for attachment in attachments:
        attach_file(message, attachment, "Portfolio.pdf" if "portfolio" in attachment.lower() else None)

    # Create SMTP session and send the email
    with smtplib.SMTP(titan_smtp_server, titan_port) as server:
        server.starttls()
        server.login(titan_username, titan_password)
        server.send_message(message)

    # Save sent email to the "Sent" folder using IMAP
    save_sent_email(titan_imap_server, titan_username, titan_password, message, sent_emails_file_path)

# Read CSV and send emails
with open("C:\\Users\\Electro Store\\Downloads\\recipients.csv", newline='', encoding='ISO-8859-1') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        email = row['email']
        job_title = row['Job Title']
        company_name = row['Company']  # Extract the company name from CSV
        if is_valid_email(email):
            subject = f"Application for {job_title}"
            body = select_email_body(job_title)
            portfolio_path = 'C:\\Users\\Electro Store\\Downloads\\Improdata - Data, AI, ML & CV Projects Portfolio.pdf'
            attachments = [portfolio_path]
            logo_path = 'C:\\Users\\Electro Store\\Downloads\\WhatsApp Image 2023-11-16 at 6.31.42 PM.jpeg'
            send_email(email, subject, body, attachments, logo_path, company_name)
        else:
            print(f"Invalid email address: {email}")

            

