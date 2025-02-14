import csv
import smtplib
from email.message import EmailMessage
import docx
import sys
from getpass import getpass
import chardet

def detect_encoding(file_path):
    """Detect the encoding of a file."""
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
        return result['encoding']

def read_csv(csv_path):
    """Read CSV file with automatic encoding detection and flexible column names."""
    try:
        # Detect the file encoding
        encoding = detect_encoding(csv_path)
        
        # First pass: check headers
        with open(csv_path, 'r', encoding=encoding) as file:
            # Try different CSV dialects
            sample = file.read(1024)
            sniffer = csv.Sniffer()
            dialect = sniffer.sniff(sample)
            file.seek(0)
            
            reader = csv.reader(file, dialect=dialect)
            headers = next(reader)
            
            # Case-insensitive header matching
            name_col = None
            email_col = None
            
            for i, header in enumerate(headers):
                header_clean = header.strip().lower()
                if header_clean in ['name', 'full name', 'firstname', 'first name']:
                    name_col = i
                elif header_clean in ['email', 'email address', 'e-mail', 'mail']:
                    email_col = i
            
            if name_col is None or email_col is None:
                raise ValueError("CSV must contain columns for name and email. "
                               "Accepted headers:\nName: 'name', 'full name', 'firstname', 'first name'\n"
                               "Email: 'email', 'email address', 'e-mail', 'mail'")
            
            # Read all recipients
            recipients = []
            for row in reader:
                if row:  # Skip empty rows
                    try:
                        recipients.append({
                            'Name': row[name_col].strip(),
                            'Email': row[email_col].strip()
                        })
                    except IndexError:
                        print(f"Warning: Skipping malformed row: {row}")
                        continue
                    
            return recipients

    except Exception as e:
        print(f"Error reading CSV file: {str(e)}")
        sys.exit(1)

def read_template(template_path):
    """Read and return the content of the Word template."""
    try:
        doc = docx.Document(template_path)
        return '\n'.join([paragraph.text for paragraph in doc.paragraphs])
    except Exception as e:
        print(f"Error reading template: {str(e)}")
        sys.exit(1)

def send_emails(csv_path, template_path, email, password):
    """Send emails to recipients from CSV using the template."""
    try:
        # Read recipients from CSV
        recipients = read_csv(csv_path)
        
        # Read template
        template_content = read_template(template_path)
        
        # Connect to SMTP server
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            try:
                server.login(email, password)
            except smtplib.SMTPAuthenticationError:
                print("Authentication failed. Please check your email and app password.")
                return
            
            # Send emails to all recipients
            for recipient in recipients:
                try:
                    # Create email message
                    msg = EmailMessage()
                    msg['From'] = email
                    msg['To'] = recipient['Email']
                    msg['Subject'] = "Important Message"
                    
                    # Replace [Name] with recipient's name
                    content = template_content.replace('[Name]', recipient['Name'])
                    msg.set_content(content)
                    
                    # Send email
                    server.send_message(msg)
                    print(f"Successfully sent email to {recipient['Name']} ({recipient['Email']})")
                    
                except Exception as e:
                    print(f"Error sending email to {recipient['Email']}: {str(e)}")
                    continue
                    
        print("\nEmail sending process completed!")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        sys.exit(1)

def main():
    # Get file paths
    csv_path = input("Enter path to CSV file: ").strip()
    template_path = input("Enter path to Word template file: ").strip()
    
    # Get email credentials
    email = input("Enter your email address: ").strip()
    password = getpass("Enter your app password: ")
    
    # Confirm before sending
    print("\nReady to send emails with the following settings:")
    print(f"CSV file: {csv_path}")
    print(f"Template file: {template_path}")
    print(f"From email: {email}")
    
    confirm = input("\nProceed with sending emails? (y/n): ").strip().lower()
    if confirm == 'y':
        send_emails(csv_path, template_path, email, password)
    else:
        print("Operation cancelled")

if __name__ == "__main__":
    main()