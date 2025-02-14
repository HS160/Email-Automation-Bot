# Email Automation Tool

A Python script for sending personalized emails to multiple recipients using a CSV file for contact information and a Word document as an email template.

## Features

- Automatic CSV encoding detection
- Flexible CSV column name matching
- Support for Word document templates with personalization
- Gmail SMTP integration
- Error handling and reporting
- Progress tracking for email sending

## Prerequisites

- Python 3.x
- Required Python packages:
  - python-docx
  - chardet

## Installation

1. Clone this repository
2. Install required packages:
```bash
pip install python-docx chardet
```

## CSV File Format

The script accepts CSV files with the following column headers (case-insensitive):

- Name/Full Name/First Name/Firstname
- Email/Email Address/E-mail/Mail

Example:
```csv
Name,Email
Harsh Solanki,harsh@example.com
Dev,dev@example.com
```

## Word Template Format

Create a Word document (.docx) with your email content. Use `[Name]` as a placeholder where you want the recipient's name to appear.

## Usage

1. Run the script:
```bash
python email_sender.py
```

2. Follow the prompts to provide:
   - Path to your CSV file
   - Path to your Word template
   - Your Gmail address
   - Your Gmail app password

## Gmail App Password Setup

1. Go to your Google Account settings
2. Navigate to Security > 2-Step Verification
3. Scroll to the bottom and select "App passwords"
4. Generate a new app password for this script

## Security Notes

- Never share your Gmail app password
- Store recipient data securely
- Review all emails before sending

## Error Handling

The script includes error handling for:
- Invalid CSV formats
- Incorrect file paths
- Email sending failures
- Authentication errors

## Author

HS160

