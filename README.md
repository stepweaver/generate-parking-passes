# ND Parking Guest Pass System

This system automates the generation and delivery of parking passes for Notre Dame University events.

## Workflow Overview

1. Access the Google Sheet
2. Export responses to CSV
3. Run the Python script to generate and email parking passes

## Detailed Instructions

### 1. Access the Department Event Guest Parking Request Sheet

Navigate to the Department Event Guest Parking Request (Responses) sheet at:
https://docs.google.com/spreadsheets/d/1kP_Z0ceHyxOVUvSmeTJhjqWRIR9Rm9_rpcYafaYLLSQ/edit?gid=2117242300#gid=2117242300

### 2. Prepare and Export Data

1. Review and make any necessary updates to the responses
2. Enter ParkMobile and Additional lots as necessary
3. Check the box in the **Generate** column for each entry you want to export (generates pass #)
4. From the menu, select '**Export Guest Passes!**' and click '**Do it!**'

This will:

- Export the selected responses to `master_file.csv`
- Archive the processed responses on a separate tab

### 3. Generate and Send Parking Passes

1. Click the **Generate Parking Passes** shortcut on the desktop
2. The system will:
   - Process the CSV data
   - Generate parking passes as PDF files
   - Send emails with the passes to guests
   - Display progress and completion status

## Technical Information

### Dependencies

- pandas: Data processing
- pdfkit: PDF generation
- Google API libraries: Gmail integration
- python-dotenv: Environment variable management

### Project Structure

- `src/`: Contains the main Python script
- `templates/`: HTML templates for parking passes
- `assets/`: Images and other resources
- `credentials/`: OAuth credentials
- `.env`: Configuration settings

## Troubleshooting

- If authentication fails, check that the Gmail credentials are valid
- Ensure the `.env` file contains the correct configuration
- Verify that the generated PDFs look correct before sending

For technical support, please contact the system administrator.
