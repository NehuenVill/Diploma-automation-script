import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image, ImageDraw, ImageFont

# Define the function
def generate_and_send_diplomas(smtp_server, smtp_port, sender_email, sender_password):
    # Load the Excel file
    data = pd.read_excel("form.xlsx")

    # Ensure the output folder exists
    os.makedirs("diplomas", exist_ok=True)

    # Iterate through each row in the Excel file
    for index, row in data.iterrows():
        name = row['Nombre y apellido']
        email = row['Email address']
        dni = row['DNI (sin puntos)']

        # Open the diploma template
        with Image.open("diploma_1.png") as img:
            draw = ImageDraw.Draw(img)

            # Define text position and font
            font = ImageFont.truetype("arial.ttf", 70)

            # Center the text on the image
            text_width, text_height = draw.textsize(name, font=font)
            image_width, image_height = img.size
            position = (((image_width - text_width) // 2)+100, ((image_height - text_height) // 2)-155)

            # Add the text to the image
            draw.text(position, name, fill="black", font=font,align="left")

            # Save the personalized diploma
            diploma_filename = os.path.join("diplomas", f"Diploma_campamento_para_{name}.png")
            img.save(diploma_filename)

"""        # Send the diploma via email
        try:
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = email
            msg['Subject'] = "Your Diploma"

            # Attach the diploma
            with open(diploma_filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={os.path.basename(diploma_filename)}",
                )
                msg.attach(part)

            # Connect to the SMTP server and send the email
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, email, msg.as_string())

            print(f"Email sent to {email}")

        except Exception as e:
            print(f"Failed to send email to {email}: {e}")
"""
if __name__ == "__main__":
    generate_and_send_diplomas(
        smtp_server="smtp.gmail.com", 
        smtp_port=587, 
        sender_email="your_email@example.com",
        sender_password="your_password"
    )