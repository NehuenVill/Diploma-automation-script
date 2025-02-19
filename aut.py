import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from dotenv import load_dotenv
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase
import pickle

load_dotenv()

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# Define the function
def generate_and_send_diplomas():
    # Load the Excel file
    data = pd.read_excel("form.xlsx")

    # Ensure the output folder exists
    os.makedirs("diplomas", exist_ok=True)

    # Iterate through each row in the Excel file
    for index, row in data.iterrows():
        name = row['Nombre y apellido']
        name = " ".join([name.split(" ")[0].capitalize(),name.split(" ")[1].capitalize()])
        email = row['Email address']
        dni = str(row['DNI (sin puntos)'])

        string_dni = ""

        for i in range(len(dni)):
            string_dni += dni[i]

            if len(dni) == 8:

                if i == 1 or i==4:

                    string_dni += "."

            elif len(dni) ==7:

                if i == 0 or i==3:

                    string_dni += "."



        # Open the diploma template
        with Image.open("diploma_1.png") as img:
            draw = ImageDraw.Draw(img)

            # Define text position and font
            font = ImageFont.truetype("arial.ttf", 55)

            # Center the text on the image
            position_name = (785, 515)
            position_dni = (540, 585)

            # Add the text to the image
            draw.text(position_name, name, fill="black", font=font,align="left")
            draw.text(position_dni, string_dni, fill="black", font=font,align="left")

            # Save the personalized diploma
            diploma_filename = os.path.join("diplomas", f"Diploma_campamento_para_{name}.png")
            img.save(diploma_filename)

        # Send the diploma via email

        send_email_with_attachment(email, "Campamento Shaolin 2025", "",diploma_filename)
    
def authenticate_gmail():
    """Authenticate the user and return the Gmail API service."""
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        
        # Save credentials for next use
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    service = build("gmail", "v1", credentials=creds)
    return service

# Send an email
def send_email_with_attachment(recipient, subject, message, attachment_path):
    """Send an email with an attachment using Gmail API."""
    service = authenticate_gmail()

    # Create email message
    msg = MIMEMultipart()
    msg["to"] = recipient
    msg["subject"] = subject

    # Attach email body
    msg.attach(MIMEText(message, "plain"))

    # Attach file
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={os.path.basename(attachment_path)}",
    )
    msg.attach(part)

    # Encode email in base64
    raw_msg = base64.urlsafe_b64encode(msg.as_bytes()).decode()

    send_message = {"raw": raw_msg}

    # Send email
    service.users().messages().send(userId="me", body=send_message).execute()
    print(f"✅ Email sent to {recipient} with attachment: {os.path.basename(attachment_path)}")



if __name__ == "__main__":
    generate_and_send_diplomas()