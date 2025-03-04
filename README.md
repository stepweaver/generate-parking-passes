# Parking Guest Pass Generator

A Python application that automates the generation and distribution of parking passes. This tool creates PDF parking passes from a data source and sends them via email using Gmail API.

## Features

- Generates Diamond Pass PDFs from input data
- Sends automated emails with parking pass attachments
- Integrates with Gmail API for email sending
- Formats dates and information for clarity
- Supports batch processing of multiple passes

## Requirements

- Python 3.6+
- Required Python packages (see requirements.txt):
  - pandas - For data manipulation
  - pdfkit - For PDF generation from HTML templates
  - google-api-python-client - For Gmail API integration
  - google-auth, google-auth-oauthlib - For Google authentication
  - python-dotenv - For environment variable management

## Installation

1. Clone this repository:

   ```
   git clone <repository-url>
   cd Parking
   ```

2. Install the required packages:

   ```
   pip install -r requirements.txt
   ```

3. Install wkhtmltopdf (required by pdfkit):

   - Windows: Download from https://wkhtmltopdf.org/downloads.html
   - macOS: `brew install wkhtmltopdf`
   - Linux: `sudo apt-get install wkhtmltopdf`

4. Configure Gmail API credentials:
   - Create a project in Google Cloud Console
   - Enable the Gmail API
   - Create OAuth 2.0 credentials
   - Fill in the .env file with your credentials (see .env.example)

## Configuration

Create a `.env` file in the root directory with the following variables:

```
# Google Cloud project credentials
GMAIL_CLIENT_ID=your-client-id.apps.googleusercontent.com
GMAIL_PROJECT_ID=your-project-id
GMAIL_AUTH_URI=https://accounts.google.com/o/oauth2/auth
GMAIL_TOKEN_URI=https://oauth2.googleapis.com/token
GMAIL_AUTH_PROVIDER_CERT_URL=https://www.googleapis.com/oauth2/v1/certs
GMAIL_CLIENT_SECRET=your-client-secret
GMAIL_REDIRECT_URI=http://localhost
# Delegate email setting (if applicable)
GMAIL_DELEGATE_EMAIL=parking@example.com
```

## Usage

Run the main script from the command line:

```
python src/generate_guest_passes.py
```

The application will:

1. Authenticate with Gmail using your configured credentials
2. Process input data
3. Generate PDF parking passes using the templates
4. Send emails with the passes attached

## Project Structure

```
├── assets/              # Static assets for the project
├── credentials/         # Stores OAuth token files (auto-generated)
├── src/                 # Source code
│   └── generate_guest_passes.py  # Main application script
├── templates/           # HTML templates for PDF generation
│   └── diamondPass.html # Template for Diamond passes
├── .env                 # Environment variables (add your own)
├── .gitignore           # Git ignore file
└── requirements.txt     # Python dependencies
```

## Authentication

The application uses OAuth 2.0 to authenticate with Gmail. On first run, it will:

1. Open a browser window to authorize access
2. Store the authentication token in `credentials/token.pickle`
3. Use this token for future runs until it expires

## License

[Specify your license here]

## Author

[Your Name/Organization]
