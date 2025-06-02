import subprocess
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time
import os
import sys

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        # base_path = root folder (one level up from current script folder)
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

# Paths
output_folder = get_resource_path('assets')
rrd_template_path = get_resource_path('assets/pendingiqos.rrd')
pending_rpt_path = get_resource_path('assets/Pending IQOS.rpt')
pending_pdf_path = os.path.join(output_folder, "Pending IQOS.pdf")

def escape_windows_path(path):
    # Convert forward slashes to backslashes, then double each backslash for .rrd file formatting
    return path.replace('/', '\\\\').replace('\\', '\\\\')

def patch_rrd_file():
    """Reads the .rrd template, replaces absolute paths with current user paths (escaped),
    writes a temp .rrd file, and returns its path."""
    with open(rrd_template_path, 'r') as f:
        rrd_contents = f.read()

    # Prepare escaped paths for the .rrd file
    escaped_rpt_path = escape_windows_path(pending_rpt_path)
    escaped_pdf_path = escape_windows_path(pending_pdf_path)

    # Replace hardcoded absolute paths with the dynamic ones
    rrd_contents = rrd_contents.replace(
        r'C:\PythonProjects\pythonProject\IbcoStockControl\Pending IQOS.rpt',
        escaped_rpt_path
    ).replace(
        r'C:\PythonProjects\pythonProject\IbcoStockControl\Pending IQOS.pdf',
        escaped_pdf_path
    )

    # Write patched .rrd to a temporary file in the assets folder
    temp_rrd_path = os.path.join(output_folder, 'pendingiqos_temp.rrd')
    with open(temp_rrd_path, 'w') as f:
        f.write(rrd_contents)

    return temp_rrd_path

# List of recipients
recipients = [
        "anwar@ibco.co.uk",
        "bilal@ibco.co.uk",
        "kabir@ibco.co.uk",
        "tareque@ibco.co.uk",
        "kazi@ibco.co.uk",
        "JahedulAmin@ibco.co.uk",
        "nazmul@ibco.co.uk",
        "redwyan@ibco.co.uk",
        "jewel@ibco.co.uk",
        "nasim@ibco.co.uk",
        "zabir@ibco.co.uk",
        "mijan@ibco.co.uk",
        "logistics@ibco.co.uk",
        "daniel@ibco.co.uk",

]

# SMTP Server Details
smtp_server = "remote.ibco.co.uk"
smtp_port = 587
sender_email = "daniel@ibco.co.uk"
username = r"IBCO\daniel"
password = "WWebp01nt;"

def send_email(subject, body, attachments):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipients)  # Multiple recipients
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    for attachment_path in attachments:
        if os.path.exists(attachment_path):
            attachment_name = os.path.basename(attachment_path)
            part = MIMEBase('application', 'octet-stream')
            with open(attachment_path, "rb") as attachment:
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={attachment_name}")
            msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(username, password)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {str(e)}")

def generate_reports():
    try:
        # Patch the .rrd file with correct paths and get temp file path
        temp_rrd_path = patch_rrd_file()

        report_command = [
            "C:/Program Files (x86)/SaberLogic/Logicity/Logicity Desktop.exe",
            "--quiet",
            temp_rrd_path
        ]
        subprocess.run(report_command, check=True)
        print("Reports generated successfully.")
        time.sleep(5)  # Ensure files are written
    except subprocess.CalledProcessError as e:
        print(f"Error generating reports: {str(e)}")

def clear_output_folder():
    try:
        file_path = os.path.join(output_folder, "Pending IQOS.pdf")
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"Removed file: {file_path}")
        # Optionally remove the temp .rrd file
        temp_rrd_path = os.path.join(output_folder, 'pendingiqos_temp.rrd')
        if os.path.exists(temp_rrd_path):
            os.remove(temp_rrd_path)
            print(f"Removed temp rrd file: {temp_rrd_path}")
    except Exception as e:
        print(f"Error clearing output folder: {str(e)}")

# Run the process
generate_reports()

attachments = [pending_pdf_path]

email_body = """

Please find the current pending Iqos report attached.

Thanks and Regards,
Daniel Jackson,
Stock Control
________________________________________________________________________________________________________

Head Office: IBCO Building, Hulme Hall Lane/Lord North Street, Manchester M40 8AD, UK
Tel: +44 (0)  161 202 8200  Fax: +44 (0)  161 202 8201 Email: bilal@ibco.co.uk
Visit us online at: http:/ www.ibco.co.uk 

Confidentiality Notice: The information in this email is confidential or privileged and subject to copyright. Access to this email by anyone other than the addressee is unauthorized. If you are not the intended recipient, any disclosure, copying, distribution or any action taken or omitted to be taken on reliance on it, is prohibited and may be unlawful. No liability or responsibility is accepted for viruses - it is your responsibility to scan attachments (if any). 

"""

send_email("Pending Iqos Transfers", email_body, attachments)

clear_output_folder()
