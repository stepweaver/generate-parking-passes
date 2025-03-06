# Setup Instructions for New Email Sender

This guide explains how to configure the ND Parking Guest Pass System to use a different Gmail account and the parking@nd.edu delegate email.

## Prerequisites

1. You must have already configured your Gmail account to allow sending mail from parking@nd.edu
2. You need a Google Cloud project with Gmail API enabled
3. You need Python installed on your system

## Configuration Steps

### 1. Update the Environment Variables

Edit the `.env` file in the root directory of the project and update the following variables:

```
# Google Cloud project credentials
GMAIL_CLIENT_ID=your_client_id_here
GMAIL_PROJECT_ID=your_project_id_here
GMAIL_AUTH_URI=https://accounts.google.com/o/oauth2/auth
GMAIL_TOKEN_URI=https://oauth2.googleapis.com/token
GMAIL_AUTH_PROVIDER_CERT_URL=https://www.googleapis.com/oauth2/v1/certs
GMAIL_CLIENT_SECRET=your_client_secret_here
GMAIL_REDIRECT_URI=http://localhost
# Delegate email setting
GMAIL_DELEGATE_EMAIL=parking@nd.edu
```

Replace:

- `your_client_id_here` with your Google Cloud project client ID
- `your_project_id_here` with your Google Cloud project ID
- `your_client_secret_here` with your Google Cloud project client secret

### 2. Creating Google Cloud Project (if needed)

If you need to set up a new Google Cloud project:

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable the Gmail API for your project
4. Create OAuth 2.0 credentials
   - Application type: Desktop application
   - Redirect URI: http://localhost
5. Copy the Client ID and Client Secret to the `.env` file

### 3. Reset Authentication

Delete the existing token file to force re-authentication:

1. Navigate to the `credentials` folder
2. Delete the `token.pickle` file (if it exists)

### 4. First Run

Run the script for the first time:

1. This will prompt you to authenticate with your Google account
2. Follow the instructions in the browser to authorize the application
3. Once authorized, the script will create a new token file

## Verifying the Setup

To verify that the setup is working correctly:

1. Run the script with a test entry
2. Check that the email is sent from parking@nd.edu
3. Verify that the recipient receives the email with the correct parking pass

## Troubleshooting

- **Authentication Error**: Make sure your Google Cloud project has the correct API permissions
- **Delegate Email Error**: Verify that your Gmail account has permission to send as parking@nd.edu
- **Token File Issues**: If you encounter persistent authentication problems, delete the token file and try again

For further assistance, contact the original developer or IT support.
