import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json
import shutil
import os
from datetime import datetime


def zip_folder(folder_path, zip_filename):
    # Create a zip file from the specified folder
    shutil.make_archive(zip_filename, 'zip', folder_path)
    return zip_filename + '.zip'


import smtplib
from email.mime.text import MIMEText
import base64
import json


def send_email_with_image(image_paths):
    # Load credentials
    with open("D:\\Item_replenishment_report_automation\\Credentials.json", "r+") as crednt:
        data = json.load(crednt)
        password = data["password"]

    sender_email = "Bcs.notifications@building-controls.com"
    sender_password = password
    receiver_emails = [
            "mithul.murugaadev@building-controls.com" #,
            #"brian.ackerman@building-controls.com",
            #"adam.martinez@building-controls.com",
            #"harriette.henderson@building-controls.com",
            #"jason.bail@building-controls.com"

        ]
    subject = 'Discrepancy stats - Automated Reports'

    # Read and encode image
    with open(image_paths[0], "rb") as img_file:
        encoded_image_1 = base64.b64encode(img_file.read()).decode("utf-8")

    # Read and encode image
    with open(image_paths[1], "rb") as img_file:
        encoded_image_2 = base64.b64encode(img_file.read()).decode("utf-8")

    # Determine the MIME image type
    img_ext_1 = image_paths[0].split('.')[-1].lower()
    mime_type_1 = f"image/{'jpeg' if img_ext_1 in ['jpg', 'jpeg'] else img_ext_1}"

    img_ext_2 = image_paths[1].split('.')[-1].lower()
    mime_type_2 = f"image/{'jpeg' if img_ext_2 in ['jpg', 'jpeg'] else img_ext_2}"

    # Prepare the inline HTML
    html = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 0;
                background-color: #f4f4f4;
            }}

            .newsletter-container {{
                max-width: 700px;
                margin: 30px auto;
                background-color: #ffffff;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 0 10px rgba(0,0,0,0.1);
            }}

            h1 {{
                font-size: 26px;
                color: #2b2b2b;
                margin-bottom: 20px;
            }}

            p {{
                font-size: 16px;
                line-height: 1.6;
                color: #444444;
                margin-bottom: 20px;
            }}

            img {{
                max-width: 100%;
                height: auto;
                border-radius: 8px;
                margin: 20px 0;
            }}

            .footer {{
                margin-top: 40px;
                text-align: center;
                font-size: 14px;
                color: #aaaaaa;
            }}

            .footer a {{
                color: #007bff;
                text-decoration: none;
            }}

            .footer a:hover {{
                text-decoration: underline;
            }}
        </style>
    </head>
    <body>
        <div class="newsletter-container">
            <h1>Discrepancy reports status - dashboard</h1>
            <p>Hi Team,</p>
            <p>A visual snapshot of the updated dashboard is attached below. This was generated and shared via the automation process.
            
            </p>
            
            <h3>Daily insights:
            
            </h3>
            <img src="data:{mime_type_1};base64,{encoded_image_1}" alt="Report Image" />
            <br>
            <h3>Trend analysis:
            
            </h3>
            <img src="data:{mime_type_2};base64,{encoded_image_2}" alt="Report Image" />

            <p>Regards,</p>
            <p>Mithul</p>

            <p style="font-size: 13px; color: #999999; font-style: italic;">
                This snapshot is auto-generated and accurately reflects the status of updated underlying data. Please note that the accompanying automated AI-generated analysis is still in testing and may not always be fully accurate.
            </p>


            <div class="footer">
                <p>&copy; 2025 Building Controls. All rights reserved.</p>
                <p><a href="mailto:bcs.notifications@building-controls.com">Contact Us</a></p>
            </div>
        </div>
    </body>
    </html>
    """

    # Email setup
    message = MIMEText(html, 'html')
    message['Subject'] = subject
    message['From'] = sender_email
    message['To'] = ', '.join(receiver_emails)

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_emails, message.as_string())
        server.quit()
        print("✅ Email sent successfully with inline image and styled layout.")
        return True

    except Exception as e:
        raise ValueError(f'❌ Failed to send email: {e}')


def sender(image_url):
    current_time = datetime.now()
    day = current_time.day
    month =  current_time.strftime("%b")
    year = current_time.year

    result = send_email_with_image(image_url)

    return result