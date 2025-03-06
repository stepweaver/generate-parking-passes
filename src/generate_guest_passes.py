import pandas as pd
import os
from datetime import datetime
import pdfkit
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pickle
import base64
import json
from dotenv import load_dotenv
from tqdm import tqdm

load_dotenv()  # Load environment variables from .env file

# --- OAuth 2.0 Configuration ---
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.settings.basic',
    'https://www.googleapis.com/auth/gmail.settings.sharing'  # Required for delegation
]

# Setup paths
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CREDENTIALS_DIR = os.path.join(PROJECT_ROOT, 'credentials')
os.makedirs(CREDENTIALS_DIR, exist_ok=True)
TOKEN_FILE = os.path.join(CREDENTIALS_DIR, 'token.pickle')
ASSETS_DIR = os.path.join(PROJECT_ROOT, "assets")
TEMPLATES_DIR = os.path.join(PROJECT_ROOT, "templates")

def authenticate_gmail():
    """Authenticate with Gmail API and return credentials"""
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, 'rb') as token:
                creds = pickle.load(token)
        except Exception as e:
            print(f"Error loading token file: {e}")
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Error refreshing token: {e}")
                creds = None
        
        if not creds:
            # Create a temporary credentials file
            credentials_dict = {
                "installed": {
                    "client_id": os.getenv("GMAIL_CLIENT_ID"),
                    "project_id": os.getenv("GMAIL_PROJECT_ID"),
                    "auth_uri": os.getenv("GMAIL_AUTH_URI"),
                    "token_uri": os.getenv("GMAIL_TOKEN_URI"),
                    "auth_provider_x509_cert_url": os.getenv("GMAIL_AUTH_PROVIDER_CERT_URL"),
                    "client_secret": os.getenv("GMAIL_CLIENT_SECRET"),
                    "redirect_uris": [os.getenv("GMAIL_REDIRECT_URI")]
                }
            }
            temp_creds_file = "temp_credentials.json"
            
            with open(temp_creds_file, 'w') as f:
                json.dump(credentials_dict, f)
            
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    temp_creds_file, SCOPES)
                creds = flow.run_local_server(port=0)
                
                # Save the credentials for the next run
                with open(TOKEN_FILE, 'wb') as token:
                    pickle.dump(creds, token)
            except Exception as e:
                print(f"Error in authorization flow: {e}")
                return None
            finally:
                # Clean up temporary file
                if os.path.exists(temp_creds_file):
                    os.remove(temp_creds_file)

    return creds

def generate_email(to_email, subject, body, pdf_path=None):
    """Send an email with optional PDF attachment"""
    creds = authenticate_gmail()
    if not creds:
        print("Failed to authenticate with Gmail")
        return False
        
    try:
        service = build('gmail', 'v1', credentials=creds)
    except Exception as e:
        print(f"Error building Gmail service: {e}")
        return False

    msg = MIMEMultipart()
    delegate_email = os.getenv('GMAIL_DELEGATE_EMAIL', 'idcard@nd.edu')
    msg['From'] = delegate_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))

    if pdf_path and os.path.exists(pdf_path):
        try:
            with open(pdf_path, "rb") as attachment:
                part = MIMEApplication(attachment.read(), Name=os.path.basename(pdf_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_path)}"'
                msg.attach(part)
        except Exception as e:
            print(f"Error attaching PDF: {e}")

    try:
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        raw_message = {'raw': raw}
        
        # Use delegation if configured
        if delegate_email and delegate_email != 'idcard@nd.edu':
            message = service.users().messages().send(
                userId='me', 
                body=raw_message,
                fields='',
                headers={'X-Goog-User-Delegation': delegate_email}
            ).execute()
        else:
            message = service.users().messages().send(
                userId='me', 
                body=raw_message
            ).execute()
            
        return message
    except Exception as e:
        error_details = getattr(e, 'details', str(e))
        print(f"Error sending email: {e}")
        print(f"Error details: {error_details}")
        
        if "OAuth" in str(e) or "client" in str(e) or "credential" in str(e):
            if os.path.exists(TOKEN_FILE):
                print("Deleting token file to force reauthentication")
                os.remove(TOKEN_FILE)
        return False

def generate_diamond_pass_pdf(data, output_path="diamondPass.pdf"):
    """Generate a PDF parking pass from HTML template"""
    template_path = os.path.join(TEMPLATES_DIR, "diamondPass.html")
    nd_logo_path = os.path.join(ASSETS_DIR, "NotreDameFightingIrish.png")
    footer_logo_path = os.path.join(ASSETS_DIR, "A91waj2z0_18kacb_mug.png")
    
    try:
        with open(template_path, "r") as f:
            html_template = f.read()
    except FileNotFoundError:
        print(f"Error: Template file not found at {template_path}")
        return None

    # Replace template variables
    replacements = {
        "{{academic_year_start}}": str(data.get('ACADEMIC_YEAR_START', '')),
        "{{academic_year_end}}": str(data.get('ACADEMIC_YEAR_END', '')),
        "{{pass_type}}": str(data.get('PASS_TYPE', 'UNIVERSITY OF NOTRE DAME')),
        "{{parking_type}}": str(data.get('PARKING_TYPE', 'GUEST PARKING PASS')),
        "{{valid_until}}": str(data.get('VALID_UNTIL', '')),
        "{{lot_name}}": str(data.get('LOT', 'C LOT')),
        "{{add_lot}}": str(data.get('ADD LOT', '')),
        "{{pass_number}}": str(data.get('PASS_NUMBER', ''))
    }
    
    html_content = html_template
    for key, value in replacements.items():
        html_content = html_content.replace(key, value)

    # Embed images as base64
    try:
        # Load and encode images
        images = {
            'src="NotreDameFightingIrish.png"': nd_logo_path,
            'src="A91waj2z0_18kacb_mug.png"': footer_logo_path
        }
        
        for img_src, img_path in images.items():
            with open(img_path, "rb") as img_file:
                img_base64 = base64.b64encode(img_file.read()).decode('utf-8')
                html_content = html_content.replace(
                    img_src,
                    f'src="data:image/png;base64,{img_base64}"'
                )
    except FileNotFoundError as e:
        print(f"\nError: Required image files are missing! {e}")
        return None
    
    try:
        config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')
        
        options = {
            'enable-local-file-access': None,
            'quiet': '',
            'page-size': 'Letter',
            'margin-top': '0mm',
            'margin-right': '0mm',
            'margin-bottom': '0mm',
            'margin-left': '0mm',
            'encoding': 'UTF-8',
            'no-outline': None,
            'dpi': '300',
            'image-quality': '100',
            'enable-smart-shrinking': None,
            'zoom': '1.0',
            'javascript-delay': '1000'
        }
        
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        
        pdfkit.from_string(
            html_content, 
            output_path, 
            configuration=config,
            options=options
        )
        
        return output_path if os.path.exists(output_path) else None
            
    except Exception as e:
        print(f"Error generating PDF: {e}")
        return None

def parse_date(date_str):
    """Parse different date formats and return pandas Timestamp"""
    try:
        # If it's already a pandas Timestamp
        if isinstance(date_str, pd.Timestamp):
            return date_str
            
        # Try pandas default parsing first
        try:
            return pd.to_datetime(date_str)
        except:
            # Handle JavaScript date format by extracting the date portion
            date_str = date_str.split('GMT')[0].strip()
            if 'T' in date_str:
                # For ISO format
                return pd.to_datetime(date_str)
            elif len(date_str.split(' ')) >= 4:
                # For format like "Thu Jan 30 2025 08:00:00"
                month = date_str.split(' ')[1]
                day = date_str.split(' ')[2]
                year = date_str.split(' ')[3]
                clean_date = f"{month} {day} {year}"
                return pd.to_datetime(clean_date)
            else:
                # Try one more time with cleaned string
                return pd.to_datetime(date_str)
    except Exception as e:
        print(f"Warning: Could not parse date '{date_str}'. Error: {e}")
        return None

def format_date_range(start_date, end_date):
    """Format date range for VALID_UNTIL field"""
    start = parse_date(start_date)
    end = parse_date(end_date)
    
    if start is None or end is None:
        return "Invalid Date"
    
    start_str = start.strftime('%m/%d/%y')
    end_str = end.strftime('%m/%d/%y')
    
    if start_str == end_str:
        return start_str
    return f"{start_str} - {end_str}"

def format_email_date(date):
    """Format date in long form for emails"""
    return date.strftime('%B %d, %Y')  # e.g. "February 11, 2025"

def format_email_date_range(start_date, end_date):
    """Format date range for emails"""
    if start_date.date() == end_date.date():
        return format_email_date(start_date)
    return f"{format_email_date(start_date)} - {format_email_date(end_date)}"

def generate_parkmobile_email_body(row, start_date, end_date):
    """Generate email body for ParkMobile access code"""
    
    # Get ParkMobile image
    try:
        image_path = os.path.join(ASSETS_DIR, "image.png")
        with open(image_path, "rb") as img_file:
            image_base64 = base64.b64encode(img_file.read()).decode('utf-8')
    except Exception as e:
        print(f"Warning: Could not load ParkMobile image: {e}")
        image_base64 = None

    # Get the current date
    email_generated_date = datetime.now().strftime('%B %d, %Y')

    return f"""<html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                line-height: 1.8;
                color: #333;
                max-width: 600px;
                margin: 0 auto;
                padding: 20px;
                font-size: 16px;
                background-color: #ffffff;
            }}
            .email-container {{ width: 100%; max-width: 600px; margin: 0 auto; }}
            .info-box {{
                margin: 20px 0;
                padding: 20px;
                background-color: #f8f9fa;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(12, 35, 64, 0.1);
            }}
            .access-code {{
                font-size: 1.6em;
                color: #0c2340;
                font-weight: bold;
                background-color: #e9ecef;
                padding: 15px;
                border-radius: 8px;
                display: inline-block;
                width: auto;
                min-width: 100px;
                max-width: 80%;
                margin: 15px 0;
                text-align: left;
                border: 2px dashed #0c2340;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            .warning-box {{
                margin: 20px 0;
                padding: 20px;
                background-color: #fff3cd;
                border-left: 6px solid #ffc107;
                color: #856404;
                font-size: 1.1em;
                border-radius: 8px;
            }}
            .charge-notice {{
                margin: 20px 0;
                padding: 15px;
                background-color: #0c2340;
                color: white;
                border-radius: 8px;
            }}
            .signature {{
                margin-top: 30px;
                padding-top: 20px;
                border-top: 1px solid #dee2e6;
                color: #666;
            }}
            img {{
                max-width: 100%;
                margin: 20px 0;
                border: 1px solid #ddd;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            ol li, ul li {{ margin-bottom: 12px; padding-left: 5px; }}
            ol, ul {{ padding-left: 30px; }}
            h3 {{
                color: #0c2340;
                margin-top: 25px;
                margin-bottom: 15px;
                font-size: 1.3em;
            }}
            .button {{
                display: inline-block;
                background-color: #0c2340;
                color: white;
                padding: 15px 30px;
                text-decoration: none;
                font-weight: bold;
                border-radius: 8px;
                margin: 20px 0;
                font-size: 1.2em;
                text-align: center;
            }}
            a {{
                color: #0c2340;
                text-decoration: underline;
                font-weight: bold;
            }}
        </style>
    </head>
    <body>
        <p>Greetings <span style="font-weight: bold; font-size: 1.1em;">{row['FIRST_NAME']},</span></p>

        <div class="info-box">
            <strong>Event:</strong> {row.get('EVENT', 'Event Name Not Provided')}<br>
            <strong>Event Date:</strong> {format_email_date_range(start_date, end_date)}<br>
            <strong>ParkMobile Access Code:</strong><br>
            <div class="access-code">{row['PARKMOBILE']}</div>
        </div>

        <p>Your ParkMobile Access Code has been assigned! Please share this information with your guests. The Access Code may be used to reserve parking for the event date. This can be done prior to arriving at Notre Dame.</p>

        <div class="warning-box">
            <strong>Note:</strong> Please allow 1-2 business days from {email_generated_date} for the access code to become active in the ParkMobile app.
        </div>
        {f'<img src="data:image/png;base64,{image_base64}" alt="ParkMobile Interface">' if image_base64 else ''}
        <div class="info-box">
            <h3 style="color: #0c2340; margin-top: 0;">How to Use ParkMobile:</h3>
            <p style="font-weight: bold; margin-bottom: 20px; font-size: 1.1em;">Follow these steps to reserve your parking spot:</p>
            <ol style="padding-left: 20px;">
                <li><span style="font-weight: bold;">FIRST:</span> Download the ParkMobile app on your phone* <strong>OR</strong> visit <a href="https://parkmobile.io" style="color: #0c2340; font-weight: bold; font-size: 1.1em;">ParkMobile.io</a></li>
                <li><span style="font-weight: bold;">NEXT:</span> Your specific event will appear in a blue bar at the top of the screen</li>
                <li><span style="font-weight: bold;">THEN:</span> Click "Filters & Access Codes" (located just below the blue bar)</li>
                <li><span style="font-weight: bold;">ENTER THIS CODE:</span>
                    <div class="access-code">{row['PARKMOBILE']}</div>
                </li>
                <li><span style="font-weight: bold;">CLICK:</span> "Apply" to unlock complimentary parking in available lots</li>
                <li><span style="font-weight: bold;">SELECT:</span> Your preferred parking lot from the list</li>
                <li><span style="font-weight: bold;">CLICK:</span> The green "Reserve" button</li>
                <li><span style="font-weight: bold;">ENTER:</span> Your email address and vehicle license plate number</li>
                <li><span style="font-weight: bold;">COMPLETE:</span> Follow the remaining prompts to finish your reservation</li>
            </ol>
            <p style="background-color: #e9ecef; padding: 10px; border-radius: 8px; margin-top: 20px;"><strong>Tip:</strong> <em>You can either continue as a guest or create an account for future use.</em></p>
        </div>

        <div class="warning-box" style="border: 3px solid #ffc107; text-align: center;">
            <h3 style="color: #856404; margin-top: 0; font-size: 1.3em;">⚠️ IMPORTANT FOR ANDROID USERS ⚠️</h3>
            <ul style="margin: 10px 0 0 0; padding-left: 20px; text-align: left; list-style-type: none;">
                <li style="margin-bottom: 10px;">📱 Android users MUST use the <a href="https://parkmobile.io" style="color: #856404; font-weight: bold; text-decoration: underline;">ParkMobile.io website</a> (not the app)</li>
                <li style="margin-bottom: 10px;">❌ Access codes are NOT supported in the Android app</li>
                <li style="margin-bottom: 10px;">✅ Once reserved, your parking will appear in your ParkMobile app account</li>
            </ul>
        </div>

        <div class="info-box" style="border: 2px solid #0c2340;">
            <h3 style="color: #0c2340; margin-top: 0; text-align: center; font-size: 1.4em;">📅 On the Day of Parking:</h3>
            <ul style="padding-left: 20px;">
                <li style="margin-bottom: 15px;">✅ <span style="font-weight: bold; font-size: 1.1em;">No physical parking pass needed</span></li>
                <li style="margin-bottom: 15px;">✅ <span style="font-weight: bold; font-size: 1.1em;">NDPD Parking Enforcement will verify your parking using your license plate</span></li>
                <li style="margin-bottom: 15px;">⚠️ <span style="font-weight: bold; font-size: 1.1em; color: #856404;">IMPORTANT: Make sure the license plate number is entered correctly</span></li>
            </ul>
        </div>

        <div class="charge-notice">
            <p style="margin: 0;"><strong>After the event, your departmental FOAPAL will be charged $5.50 for each use of the access code.</strong></p>
        </div>

        <p>We recommend testing the link and code yourself before sharing with guests, so you can assist if they have questions.</p>

        <div class="signature">
            <p style="font-size: 1.1em; margin-bottom: 5px;">Thank you,</p>
            <p style="font-size: 1.2em; font-weight: bold; color: #0c2340; margin-top: 0;">NDPD Parking Services Team</p>
            <hr style="border: 1px solid #dee2e6; margin: 15px 0;">
            <p style="color: #666; font-size: 0.9em;">Pass Number: {row['PASS #']}</p>
        </div>
    </body>
    </html>"""

def generate_diamond_email_body(row, start_date, end_date):
    """Generate email body for Diamond Pass emails"""
    return f"""<html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                line-height: 1.8;
                color: #333;
                max-width: 600px;
                margin: 0 auto;
                padding: 20px;
                font-size: 16px;
                background-color: #ffffff;
            }}
            .email-container {{ width: 100%; max-width: 600px; margin: 0 auto; }}
            .date-box {{
                margin: 20px 0;
                padding: 20px;
                background-color: #f8f9fa;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(12, 35, 64, 0.1);
                font-size: 1.1em;
                border: 1px solid #dee2e6;
            }}
            .important-notice {{
                margin: 20px 0;
                padding: 15px;
                background-color: #0c2340;
                color: white;
                border-radius: 8px;
            }}
            .contact-info {{
                margin: 20px 0;
                padding: 15px;
                background-color: #e9ecef;
                border-radius: 8px;
            }}
            .signature {{
                margin-top: 30px;
                padding-top: 20px;
                border-top: 1px solid #dee2e6;
                color: #666;
            }}
            .button {{
                display: inline-block;
                background-color: #0c2340;
                color: white;
                padding: 15px 30px;
                text-decoration: none;
                font-weight: bold;
                border-radius: 8px;
                margin: 20px 0;
                font-size: 1.2em;
                text-align: center;
            }}
            a {{ color: #0c2340; text-decoration: underline; font-weight: bold; }}
        </style>
    </head>
    <body>
        <p>Greetings <span style="font-weight: bold; font-size: 1.1em;">{row['FIRST_NAME']},</span></p>

        <div class="date-box">
            <h3 style="color: #0c2340; margin-top: 0; text-align: left;">📝 Guest Parking Pass Information</h3>
            <p style="font-size: 1.1em;">A Guest Parking Pass PDF has been <span style="font-weight: bold; text-decoration: underline;">attached to this email</span> for use by your guest(s) on:</p>
            <div style="font-size: 1.5em; color: #0c2340; font-weight: bold; background-color: #e9ecef; padding: 15px; border-radius: 8px; text-align: left; margin: 15px 0; border: 2px dashed #0c2340; display: inline-block; width: auto;">
                {format_email_date_range(start_date, end_date)}
            </div>
        </div>

        <div style="font-size: 1.1em; line-height: 1.8; margin: 25px 0; padding: 15px; background-color: #f8f9fa; border-radius: 8px;">
            <p>📄 This PDF version of the Guest Parking Pass:</p>
            <ul style="padding-left: 30px;">
                <li>Should be <span style="font-weight: bold;">emailed to your guest(s)</span> before their visit</li>
                <li>Must be <span style="font-weight: bold;">printed out</span> by your guest</li>
                <li>Needs to be <span style="font-weight: bold;">placed on their vehicle's dashboard</span> while parked</li>
                <li>Is <span style="font-weight: bold; color: #856404;">only valid for the date(s) shown on the pass</span></li>
            </ul>
        </div>

        <div class="important-notice">
            <p style="margin: 0; font-size: 1.1em;"><strong>⚠️ Important:</strong> The FOAPAL number provided will be charged for 
            the number of guest passes requested after the usage date.</p>
        </div>

        <div class="contact-info" style="text-align: left;">
            <h3 style="color: #0c2340;">Need Help?</h3>
            <p style="font-size: 1.1em;">Contact our office at:</p>
            <a href="tel:574-631-5053" class="button">📞 Call: 574-631-5053</a><br>
            <a href="mailto:parking@nd.edu" class="button">✉️ Email: parking@nd.edu</a>
        </div>

        <div class="signature">
            <p style="font-size: 1.1em; margin-bottom: 5px;">Thank you,</p>
            <p style="font-size: 1.2em; font-weight: bold; color: #0c2340; margin-top: 0;">NDPD Parking Services Office</p>
            <hr style="border: 1px solid #dee2e6; margin: 15px 0;">
            <p style="color: #666; font-size: 0.9em;">Pass Number: {row['PASS #']}</p>
        </div>
    </body>
    </html>"""

def main():
    """Main function to process the master file and generate passes"""
    # Base directory for the master file
    directory_path = r"G:\Shared drives\Card Office\Department Guest Parking Passes"
    csv_path = os.path.join(directory_path, "master_file.csv")
    
    # Define the output directory for Diamond Pass PDFs
    diamond_pass_pdf_dir = os.path.join(directory_path, "Diamond Passes")
    os.makedirs(diamond_pass_pdf_dir, exist_ok=True)

    try:
        df = pd.read_csv(csv_path, on_bad_lines='skip')
        df['VEHICLE_COUNT'] = pd.to_numeric(df['VEHICLE_COUNT'], errors='coerce').fillna(0).astype(int)
    except Exception as e:
        print(f"Failed to read CSV: {e}")
        return

    diamond_passes = 0
    emails_sent = 0
    errors = []
    
    # Process each row in the CSV
    for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="Processing rows"):
        if not row['GENERATE']:
            continue
            
        try:
            start_date = parse_date(row['START'])
            end_date = parse_date(row['END'])
            
            if start_date is None or end_date is None:
                errors.append(f"Pass {row['PASS #']}: Invalid dates - START: {row['START']}, END: {row['END']}")
                continue

            # Create data for diamond pass
            data = {
                'ACADEMIC_YEAR_START': str(datetime.now().year),
                'ACADEMIC_YEAR_END': str(datetime.now().year + 1),
                'PASS_TYPE': 'UNIVERSITY OF NOTRE DAME',
                'PARKING_TYPE': 'GUEST PARKING PASS',
                'VALID_UNTIL': format_date_range(row['START'], row['END']),
                'LOT': 'C LOT',
                'ADD LOT': f"OR {row['ADD LOT']}" if pd.notna(row.get('ADD LOT', '')) else '',
                'PASS_NUMBER': str(row['PASS #'])
            }

            # Determine if this is a diamond pass or parkmobile pass
            if row['VEHICLE_COUNT'] <= 10:
                # Generate diamond pass
                filename = f"diamondPass_{row['DEPARTMENT']}_{row['PASS #']}.pdf"
                filename = "".join(c for c in filename if c.isalnum() or c in ('_', '-', '.'))
                
                pdf_path = generate_diamond_pass_pdf(data, os.path.join(diamond_pass_pdf_dir, filename))
                if pdf_path:
                    email_body = generate_diamond_email_body(row, start_date, end_date)
                    if generate_email(row['EMAIL'], "Diamond Parking Pass", email_body, pdf_path):
                        diamond_passes += 1
                        emails_sent += 1
                    else:
                        errors.append(f"Pass {row['PASS #']}: Failed to send Diamond Pass email to {row['EMAIL']}")
                else:
                    errors.append(f"Pass {row['PASS #']}: Failed to generate PDF")
            else:
                # Generate parkmobile pass
                email_body = generate_parkmobile_email_body(row, start_date, end_date)
                if generate_email(row['EMAIL'], "ParkMobile Access Code", email_body):
                    emails_sent += 1
                else:
                    errors.append(f"Pass {row['PASS #']}: Failed to send ParkMobile email to {row['EMAIL']}")
                
        except Exception as e:
            errors.append(f"Pass {row['PASS #']}: Unexpected error - {str(e)}")
            
    # Print summary
    print(f"Diamond Passes generated: {diamond_passes}")
    print(f"Total emails sent: {emails_sent}")
    if errors:
        print("\nErrors encountered:")
        for error in errors:
            print(f"- {error}")

if __name__ == "__main__":
    main()