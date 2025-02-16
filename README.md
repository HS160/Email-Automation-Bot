# Email-Automation-Bot
### Created by HS160

A Python-based desktop application for sending personalized bulk emails using CSV data and Word document templates.

## Features

- User-friendly graphical interface for easy operation
- CSV file support for recipient data management
- Word document template support with mail merge capabilities
- Progress tracking with visual progress bar
- Detailed status logging
- Multi-threading to prevent GUI freezing
- Automatic CSV encoding detection
- Smart column detection for name and email fields

## Prerequisites

Before running the application, ensure you have Python installed and the following dependencies:

```bash
pip install python-docx chardet
```

## Setup

1. Clone or download this repository
2. Install the required dependencies
3. For Gmail users, you'll need to set up an App Password:
   - Go to your Google Account settings
   - Navigate to Security > 2-Step Verification
   - At the bottom, select "App passwords"
   - Generate a new app password for "Mail"
   - Use this password in the application instead of your regular Gmail password

## Usage

1. Launch the application by running:
   ```bash
   python email_sender.py
   ```

2. Configure the following:
   - Select your CSV file containing recipient information
   - Choose your Word template document
   - Enter your email address
   - Enter your app password

3. CSV File Format:
   - Must include columns for name and email
   - Accepted name column headers: "name", "full name", "firstname", "first name"
   - Accepted email column headers: "email", "email address", "e-mail", "mail"

4. Word Template Format:
   - Use `[Name]` as a placeholder where you want the recipient's name to appear
   - The template content will be sent as plain text in the email body

5. Click "Send Emails" to begin the sending process

## Security Notes

- The application uses Gmail's SMTP server with TLS encryption
- Email credentials are not stored and must be entered each time
- App passwords are recommended over regular passwords for enhanced security

## Error Handling

The application includes robust error handling for:
- Malformed CSV files
- Invalid email credentials
- Network connectivity issues
- Individual email sending failures

## Limitations

- Currently supports Gmail accounts only
- Sends plain text emails only (no HTML or attachments)
- One template placeholder (`[Name]`) supported

## Contributing

Feel free to fork this repository and submit pull requests for any improvements.

## License

This project is open source and available under the MIT License.

