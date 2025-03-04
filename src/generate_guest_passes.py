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
from tqdm import tqdm  # Add this import at the top with other imports

load_dotenv()  # Load environment variables from .env file

# --- OAuth 2.0 Configuration ---
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.settings.basic',
    'https://www.googleapis.com/auth/gmail.settings.sharing'  # Required for delegation
]  # Gmail scopes for sending and managing settings

# Get the project root directory (one level up from src)
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# Create a credentials directory if it doesn't exist
CREDENTIALS_DIR = os.path.join(PROJECT_ROOT, 'credentials')
os.makedirs(CREDENTIALS_DIR, exist_ok=True)
# Store token.pickle in the credentials directory
TOKEN_FILE = os.path.join(CREDENTIALS_DIR, 'token.pickle')

def get_credentials_dict():
    return {
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

def authenticate_gmail():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Create a temporary credentials file
            credentials_dict = get_credentials_dict()
            temp_creds_file = "temp_credentials.json"
            
            with open(temp_creds_file, 'w') as f:
                json.dump(credentials_dict, f)
            
            flow = InstalledAppFlow.from_client_secrets_file(
                temp_creds_file, SCOPES)
            creds = flow.run_local_server(port=0)
            
            # Clean up temporary file
            os.remove(temp_creds_file)

        # Save the credentials for the next run
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)

    return creds


def generate_email(to_email, subject, body, pdf_path=None):
    creds = authenticate_gmail()
    service = build('gmail', 'v1', credentials=creds)

    msg = MIMEMultipart()
    delegate_email = os.getenv('GMAIL_DELEGATE_EMAIL', 'parking@nd.edu')
    msg['From'] = delegate_email  # Use delegate email as the sender
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))

    if pdf_path:
        try:
            with open(pdf_path, "rb") as attachment:
                part = MIMEApplication(attachment.read(), Name=os.path.basename(pdf_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_path)}"'
                msg.attach(part)
        except FileNotFoundError:
            print(f"Warning: PDF file not found at {pdf_path}")

    try:
        # If using delegation, set the 'me' value to the delegate's email
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        raw_message = {'raw': raw}
        
        # Use delegation if configured
        if delegate_email and delegate_email != 'parking@nd.edu':
            message = service.users().messages().send(
                userId='me', 
                body=raw_message,
                # Add delegation header
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
        print(f"Error sending email: {e}")
        return False


def generate_diamond_pass_pdf(data, output_path="diamondPass.pdf"):
    # Get the directory where the script is located and go up one level
    script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    # Fix path joining using os.path.join properly from parent directory
    template_path = os.path.join(script_dir, "templates", "diamondPass.html")
    nd_logo_path = os.path.join(script_dir, "assets", "NotreDameFightingIrish.png")
    footer_logo_path = os.path.join(script_dir, "assets", "A91waj2z0_18kacb_mug.png")
    
    try:
        with open(template_path, "r") as f:
            html_template = f.read()
    except FileNotFoundError:
        print(f"Error: Template file not found at {template_path}")
        print("Please ensure the diamondPass.html template exists in the templates directory")
        return None

    # Replace template variables
    html_content = html_template.replace("{{academic_year_start}}", str(data.get('ACADEMIC_YEAR_START', '')))
    html_content = html_content.replace("{{academic_year_end}}", str(data.get('ACADEMIC_YEAR_END', '')))
    html_content = html_content.replace("{{pass_type}}", str(data.get('PASS_TYPE', 'UNIVERSITY OF NOTRE DAME')))
    html_content = html_content.replace("{{parking_type}}", str(data.get('PARKING_TYPE', 'GUEST PARKING PASS')))
    html_content = html_content.replace("{{valid_until}}", str(data.get('VALID_UNTIL', '')))
    html_content = html_content.replace("{{lot_name}}", str(data.get('LOT', 'C LOT')))
    html_content = html_content.replace("{{add_lot}}", str(data.get('ADD LOT', '')))
    html_content = html_content.replace("{{pass_number}}", str(data.get('PASS_NUMBER', '')))

    # Embed images as base64
    try:
        with open(nd_logo_path, "rb") as img_file:
            nd_logo_base64 = base64.b64encode(img_file.read()).decode('utf-8')
        with open(footer_logo_path, "rb") as img_file:
            footer_logo_base64 = base64.b64encode(img_file.read()).decode('utf-8')

        # Update the image replacement to match the exact src attributes in the HTML
        html_content = html_content.replace(
            'src="NotreDameFightingIrish.png"',
            f'src="data:image/png;base64,{nd_logo_base64}"'
        )
        html_content = html_content.replace(
            'src="A91waj2z0_18kacb_mug.png"',
            f'src="data:image/png;base64,{footer_logo_base64}"'
        )

    except FileNotFoundError as e:
        print(f"\nError: Required image files are missing!")
        print("Please ensure the following files exist in the assets directory:")
        print("- NotreDameFightingIrish.png")
        print("- A91waj2z0_18kacb_mug.png")
        print(f"Looking in: {os.path.join(script_dir, 'assets')}")
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

def format_date_range(start_date, end_date):
    """Format date range for VALID_UNTIL field"""
    def parse_date(date_str):
        try:
            # First try pandas default parsing
            return pd.to_datetime(date_str)
        except:
            # If that fails, try to handle JavaScript date format
            try:
                # Remove timezone info from the string
                date_str = date_str.split('(')[0].strip()
                return pd.to_datetime(date_str)
            except:
                print(f"Warning: Could not parse date: {date_str}")
                return None

    start = parse_date(start_date)
    end = parse_date(end_date)
    
    if start is None or end is None:
        return "Invalid Date"
    
    start_str = start.strftime('%m/%d/%y')
    end_str = end.strftime('%m/%d/%y')
    
    if start_str == end_str:
        return start_str
    return f"{start_str} - {end_str}"

def format_filename_date(date_str):
    try:
        # If it's already a pandas Timestamp
        if isinstance(date_str, pd.Timestamp):
            date = date_str
        else:
            # Handle JavaScript date format by extracting the date portion
            # Example: "Thu Jan 30 2025 08:00:00 GMT-0500 (Eastern Standard Time)"
            date_parts = date_str.split(' ')
            if len(date_parts) >= 4:
                # Extract month, day, year
                month = date_parts[1]
                day = date_parts[2]
                year = date_parts[3]
                # Reconstruct in a format pandas can parse
                clean_date = f"{month} {day} {year}"
                date = pd.to_datetime(clean_date)
            else:
                # Fallback to direct parsing if it's in a different format
                date = pd.to_datetime(date_str)
    except Exception as e:
        print(f"Warning: Could not parse date '{date_str}'. Using current date.")
        date = pd.Timestamp.now()
    
    return date.strftime('%Y%m%d')  # Format as YYYYMMDD

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
    
    # Fix path joining for ParkMobile image
    try:
        script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        image_path = os.path.join(script_dir, "assets", "image.png")
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
                line-height: 1.6;
                color: #333;
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
            }}
            .info-box {{
                margin: 20px 0;
                padding: 20px;
                background-color: #f8f9fa;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(12, 35, 64, 0.1);
            }}
            .access-code {{
                font-size: 1.4em;
                color: #0c2340;
                font-weight: bold;
                background-color: #e9ecef;
                padding: 10px;
                border-radius: 4px;
                display: inline-block;
                margin: 10px 0;
            }}
            .warning-box {{
                margin: 20px 0;
                padding: 15px;
                background-color: #fff3cd;
                border-left: 4px solid #ffc107;
                color: #856404;
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
        </style>
    </head>
    <body>
        <p>Greetings {row['FIRST_NAME']},</p>

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
            <ol style="padding-left: 20px;">
                <li>Download the ParkMobile app on your phone* <strong>OR</strong> visit <a href="https://parkmobile.io" style="color: #0c2340; font-weight: bold;">ParkMobile.io</a></li>
                <li>Your specific event will appear in a blue bar at the top of the screen</li>
                <li>Click "Filters & Access Codes" (located just below the blue bar)</li>
                <li>Enter your access code: <div class="access-code">{row['PARKMOBILE']}</div></li>
                <li>Click "Apply" to unlock complimentary parking in available lots</li>
                <li>Select your preferred parking lot from the list on the left</li>
                <li>Click the green "Reserve" button</li>
                <li>Enter your email address and vehicle license plate number</li>
                <li>Complete the reservation process following the prompts</li>
            </ol>
            <p><em>Note: You can either continue as a guest or create an account for future use.</em></p>
        </div>

        <div class="warning-box">
            <strong>*IMPORTANT FOR ANDROID USERS:</strong>
            <ul style="margin: 10px 0 0 0; padding-left: 20px;">
                <li>Android users MUST use the ParkMobile.io website (not the app) to make reservations</li>
                <li>Access codes are not supported in the Android version of the app</li>
                <li>Once reserved, your parking will appear in your ParkMobile app account</li>
            </ul>
        </div>

        <div class="info-box">
            <h3 style="color: #0c2340; margin-top: 0;">On the Day of Parking:</h3>
            <ul style="padding-left: 20px;">
                <li>No physical parking pass is needed</li>
                <li>NDPD Parking Enforcement will verify your parking using your license plate</li>
                <li><strong>Important:</strong> Make sure the license plate number is entered correctly</li>
            </ul>
        </div>

        <div class="charge-notice">
            <p style="margin: 0;"><strong>After the event, your departmental FOAPAL will be charged $5.50 for each use of the access code.</strong></p>
        </div>

        <p>We recommend testing the link and code yourself before sharing with guests, so you can assist if they have questions.</p>

        <div class="signature">
            <p>Thank you,<br>
            <em>NDPD Parking Services Team</em></p>
            <p style="color: #666; font-size: 0.9em;">Pass Number: {row['PASS #']}</p>
        </div>
    </body>
    </html>"""

def parse_js_date(date_str):
    """Parse JavaScript date format and return datetime object"""
    try:
        # Remove timezone and parenthetical info
        date_str = date_str.split('GMT')[0].strip()
        # Parse with pandas
        return pd.to_datetime(date_str)
    except Exception as e:
        print(f"Warning: Could not parse date '{date_str}'. Error: {e}")
        return None

def main():
    # Base directory for the master file
    directory_path = r"G:\Shared drives\Card Office\Department Guest Parking Passes"
    csv_path = os.path.join(directory_path, "master_file.csv")
    
    # Define the output directory for Diamond Pass PDFs
    diamond_pass_pdf_dir = os.path.join(directory_path, "Diamond Passes")
    os.makedirs(diamond_pass_pdf_dir, exist_ok=True)  # Create directory if it doesn't exist

    try:
        df = pd.read_csv(csv_path, on_bad_lines='skip')
        df['VEHICLE_COUNT'] = pd.to_numeric(df['VEHICLE_COUNT'], errors='coerce').fillna(0).astype(int)
    except Exception as e:
        print(f"Failed to read CSV: {e}")
        return

    diamond_passes = 0
    emails_sent = 0
    errors = []  # Changed to list to store error messages
    
    # Wrap the DataFrame iteration with tqdm for a progress bar
    for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="Processing rows"):
        if not row['GENERATE']:
            continue
            
        try:
            start_date = parse_js_date(row['START'])
            end_date = parse_js_date(row['END'])
            
            if start_date is None or end_date is None:
                errors.append(f"Pass {row['PASS #']}: Invalid dates - START: {row['START']}, END: {row['END']}")
                continue

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

            if row['VEHICLE_COUNT'] <= 10:
                filename = f"diamondPass_{row['DEPARTMENT']}_{row['PASS #']}.pdf"
                filename = "".join(c for c in filename if c.isalnum() or c in ('_', '-', '.'))
                
                # Use the Diamond Passes subdirectory for outputting PDFs
                pdf_path = generate_diamond_pass_pdf(data, os.path.join(diamond_pass_pdf_dir, filename))
                if pdf_path:
                    email_body = f"""<html>
                    <head>
                        <style>
                            body {{
                                font-family: Arial, sans-serif;
                                line-height: 1.6;
                                color: #333;
                                max-width: 800px;
                                margin: 0 auto;
                                padding: 20px;
                            }}
                            .date-box {{
                                margin: 20px 0;
                                padding: 20px;
                                background-color: #f8f9fa;
                                border-radius: 8px;
                                box-shadow: 0 2px 4px rgba(12, 35, 64, 0.1);
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
                        </style>
                    </head>
                    <body>
                        <p>Dear {row['FIRST_NAME']},</p>

                        <div class="date-box">
                            A Guest Parking Pass .pdf has been attached for use by your guest(s) on:<br>
                            <span style="font-size: 1.4em; color: #0c2340; font-weight: bold; display: block; margin: 10px 0;">
                                {format_email_date_range(start_date, end_date)}
                            </span>
                        </div>

                        <p>This PDF version of the Guest Parking Pass can be emailed to your guest(s) in advance of their visit, 
                        and should be printed and placed on their vehicle's dash while parked on campus. It is valid only for the 
                        date(s) indicated on the pass.</p>

                        <div class="important-notice">
                            <p style="margin: 0;"><strong>Important:</strong> The FOAPAL number provided will be charged for 
                            the number of guest passes requested after the usage date.</p>
                        </div>

                        <div class="contact-info">
                            <p>If you have any questions, please contact our office at:<br>
                            📞 <a href="tel:574-631-5053">574-631-5053</a><br>
                            ✉️ <a href="mailto:parking@nd.edu">parking@nd.edu</a></p>
                        </div>

                        <div class="signature">
                            <p>Thank you,<br>
                            <em>NDPD Parking Services Office</em></p>
                            <p style="color: #666; font-size: 0.9em;">Pass Number: {row['PASS #']}</p>
                        </div>
                    </body>
                    </html>"""
                    if generate_email(row['EMAIL'], "Diamond Parking Pass", email_body, pdf_path):
                        diamond_passes += 1
                        emails_sent += 1
                    else:
                        errors.append(f"Pass {row['PASS #']}: Failed to send Diamond Pass email to {row['EMAIL']}")
                else:
                    errors.append(f"Pass {row['PASS #']}: Failed to generate PDF")
            else:
                email_body = generate_parkmobile_email_body(row, start_date, end_date)
                if generate_email(row['EMAIL'], "ParkMobile Access Code", email_body):
                    emails_sent += 1
                else:
                    errors.append(f"Pass {row['PASS #']}: Failed to send ParkMobile email to {row['EMAIL']}")
                
        except Exception as e:
            errors.append(f"Pass {row['PASS #']}: Unexpected error - {str(e)}")
            
    print(f"Diamond Passes generated: {diamond_passes}")
    print(f"Total emails sent: {emails_sent}")
    if errors:
        print("\nErrors encountered:")
        for error in errors:
            print(f"- {error}")

if __name__ == "__main__":
    main()