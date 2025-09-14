import os
import base64
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# ----------------------------
# CONFIGURATION
# ----------------------------
TEMPLATE_PATH = "COP_template.pdf"
OUTPUT_DIR = "outputs"

SUBJECT = "Your Certificate of Participation"
BODY = """
Congratulations on Successfully Participating in the Git & GitHub Workshop!

Dear {name},

Congratulations! 🎊
You have successfully participated in the Git & GitHub Workshop held on 13th September 2025.

We truly appreciate your enthusiasm and active involvement throughout the session. As a token of your successful participation, we are pleased to award you with a Certificate of Participation. 🏅

This certificate recognizes your effort in learning the fundamentals of Git & GitHub, tools that are essential for version control, collaboration, and professional development in today’s tech-driven world.

Keep up the great spirit of learning and collaboration – this is just the beginning of your coding journey! 🚀

Best wishes,
Team CSI
"""

# Toggle preview mode (True = only generate 1 certificate, no emails)
PREVIEW_MODE = False
PREVIEW_NAME = "BUDDA NAGA SAMBA S V DURGA SAI" 

# Certificate text placement
MAX_FONT_SIZE = 46   # starting point
MIN_FONT_SIZE = 20   # don’t go smaller than this
MAX_TEXT_WIDTH = 600  # maximum width allowed for the name box (adjust!)

TEXT_POSITION_Y = 305   # vertical placement of name
TEXT_CENTER_X = 420     # horizontal center point

# Gmail OAuth2
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


# ----------------------------
# FUNCTION: Generate Certificate
# ----------------------------
def generate_certificate(name, template_path, output_dir):
    overlay_path = "overlay.pdf"

    # Dynamically adjust font size to fit bounding box
    font_size = MAX_FONT_SIZE
    while font_size >= MIN_FONT_SIZE:
        text_width = pdfmetrics.stringWidth(str(name), "Helvetica-Bold", font_size)
        if text_width <= MAX_TEXT_WIDTH:
            break
        font_size -= 1

    # Get template page size
    template_pdf = PdfReader(open(template_path, "rb"))
    page = template_pdf.pages[0]
    page_width = float(page.mediabox.width)
    page_height = float(page.mediabox.height)

    # Create overlay with same size
    c = canvas.Canvas(overlay_path, pagesize=(page_width, page_height))
    c.setFont("Helvetica-Bold", font_size)
    c.drawCentredString(TEXT_CENTER_X, TEXT_POSITION_Y, name)
    c.save()

    # Merge overlay with template
    overlay_pdf = PdfReader(open(overlay_path, "rb"))
    writer = PdfWriter()
    page.merge_page(overlay_pdf.pages[0])
    writer.add_page(page)

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    output_path = os.path.join(output_dir, f"{name}.pdf")
    with open(output_path, "wb") as f:
        writer.write(f)

    return output_path


# ----------------------------
# FUNCTION: Gmail Authentication
# ----------------------------
def gmail_authenticate():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)


# ----------------------------
# FUNCTION: Send Email with Attachment
# ----------------------------
def send_email_with_attachment(service, sender, to, subject, body, attachment_path):
    message = MIMEMultipart()
    message["To"] = to
    message["From"] = sender
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    # Attach PDF
    with open(attachment_path, "rb") as f:
        mime_base = MIMEBase("application", "pdf")
        mime_base.set_payload(f.read())
        encoders.encode_base64(mime_base)
        mime_base.add_header(
            "Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}"
        )
        message.attach(mime_base)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    send_message = {"raw": raw_message}

    return service.users().messages().send(userId="me", body=send_message).execute()


# ----------------------------
# MAIN SCRIPT
# ----------------------------
if __name__ == "__main__":
    if PREVIEW_MODE:
        print("🔍 Running in PREVIEW MODE...")
        cert_path = generate_certificate(PREVIEW_NAME, TEMPLATE_PATH, OUTPUT_DIR)
        print(f"✅ Preview certificate generated: {cert_path}")
    else:
        # Authenticate Gmail once
        service = gmail_authenticate()
        sender_email = "csibodyevents@gmail.com"  # replace with your Gmail

        # Load participants
        df = pd.read_csv("attended_list_final.csv")  # or pd.read_excel("participants.xlsx")

        for _, row in df.iterrows():
            name = row["NAME"]
            email = row["EMAIL"]

            # Generate certificate
            cert_path = generate_certificate(name, TEMPLATE_PATH, OUTPUT_DIR)

            # Personalize body
            personalized_body = BODY.format(name=name)

            # Send email
            send_email_with_attachment(
                service,
                sender=sender_email,
                to=email,
                subject=SUBJECT,
                body=personalized_body,
                attachment_path=cert_path,
            )

            print(f"✅ Sent to {name} ({email})")
