import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image, ImageDraw, ImageFont

# Define the function
def generate_and_send_diplomas(excel_path, template_path, font_path, output_folder, smtp_server, smtp_port, sender_email, sender_password):
    # Load the Excel file
    data = pd.read_excel(excel_path)

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Iterate through each row in the Excel file
    for index, row in data.iterrows():
        name = row['Name']
        surname = row['Surname']
        email = row['Email']

        # Open the diploma template
        with Image.open(template_path) as img:
            draw = ImageDraw.Draw(img)

            # Define text position and font
            font = ImageFont.truetype(font_path, 60)
            text = f"{name} {surname}"

            # Center the text on the image
            text_width, text_height = draw.textsize(text, font=font)
            image_width, image_height = img.size
            position = ((image_width - text_width) // 2, (image_height - text_height) // 2)

            # Add the text to the image
            draw.text(position, text, fill="black", font=font)

            # Save the personalized diploma
            diploma_filename = os.path.join(output_folder, f"diploma_{index + 1}.png")
            img.save(diploma_filename)

        # Send the diploma via email
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

# Example usage
if __name__ == "__main__":
    generate_and_send_diplomas(
        excel_path="data.xlsx",  # Path to your Excel file
        template_path="diploma_template.png",  # Path to your diploma template image
        font_path="arial.ttf",  # Path to your .ttf font file
        output_folder="diplomas",  # Output folder for the generated diplomas
        smtp_server="smtp.gmail.com",  # Your SMTP server (e.g., Gmail)
        smtp_port=587,  # SMTP port (587 for TLS)
        sender_email="your_email@example.com",  # Your email address
        sender_password="your_password"  # Your email password
    )