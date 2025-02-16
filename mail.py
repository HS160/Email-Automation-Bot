import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import smtplib
from email.message import EmailMessage
import docx
import sys
from getpass import getpass
import chardet
import threading

class EmailSenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Email Sender")
        self.root.geometry("600x700")
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # File Selection Section
        ttk.Label(main_frame, text="File Selection", font=('Helvetica', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10, sticky=tk.W)
        
        # CSV File Selection
        ttk.Label(main_frame, text="CSV File:").grid(row=1, column=0, sticky=tk.W)
        self.csv_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.csv_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_csv).grid(row=1, column=2)
        
        # Template File Selection
        ttk.Label(main_frame, text="Word Template:").grid(row=2, column=0, sticky=tk.W)
        self.template_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.template_path, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_template).grid(row=2, column=2)
        
        # Email Configuration Section
        ttk.Label(main_frame, text="Email Configuration", font=('Helvetica', 12, 'bold')).grid(row=3, column=0, columnspan=2, pady=(20,10), sticky=tk.W)
        
        # Email Address
        ttk.Label(main_frame, text="Email Address:").grid(row=4, column=0, sticky=tk.W)
        self.email = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.email, width=50).grid(row=4, column=1, padx=5)
        
        # Password
        ttk.Label(main_frame, text="App Password:").grid(row=5, column=0, sticky=tk.W)
        self.password = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.password, show="*", width=50).grid(row=5, column=1, padx=5)
        
        # Progress Section
        ttk.Label(main_frame, text="Progress", font=('Helvetica', 12, 'bold')).grid(row=6, column=0, columnspan=2, pady=(20,10), sticky=tk.W)
        
        # Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, length=400, mode='determinate', variable=self.progress_var)
        self.progress_bar.grid(row=7, column=0, columnspan=3, pady=10)
        
        # Status Text
        self.status_text = tk.Text(main_frame, height=15, width=60)
        self.status_text.grid(row=8, column=0, columnspan=3, pady=10)
        
        # Send Button
        self.send_button = ttk.Button(main_frame, text="Send Emails", command=self.start_sending)
        self.send_button.grid(row=9, column=0, columnspan=3, pady=10)
        
        # Configure grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

    def browse_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if filename:
            self.csv_path.set(filename)

    def browse_template(self):
        filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if filename:
            self.template_path.set(filename)

    def update_status(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()

    def detect_encoding(self, file_path):
        with open(file_path, 'rb') as file:
            raw_data = file.read()
            result = chardet.detect(raw_data)
            return result['encoding']

    def read_csv(self, csv_path):
        try:
            encoding = self.detect_encoding(csv_path)
            
            with open(csv_path, 'r', encoding=encoding) as file:
                sample = file.read(1024)
                sniffer = csv.Sniffer()
                dialect = sniffer.sniff(sample)
                file.seek(0)
                
                reader = csv.reader(file, dialect=dialect)
                headers = next(reader)
                
                name_col = email_col = None
                for i, header in enumerate(headers):
                    header_clean = header.strip().lower()
                    if header_clean in ['name', 'full name', 'firstname', 'first name']:
                        name_col = i
                    elif header_clean in ['email', 'email address', 'e-mail', 'mail']:
                        email_col = i
                
                if name_col is None or email_col is None:
                    raise ValueError("Required columns not found in CSV")
                
                recipients = []
                for row in reader:
                    if row:
                        try:
                            recipients.append({
                                'Name': row[name_col].strip(),
                                'Email': row[email_col].strip()
                            })
                        except IndexError:
                            self.update_status(f"Warning: Skipping malformed row: {row}")
                            continue
                
                return recipients

        except Exception as e:
            self.update_status(f"Error reading CSV file: {str(e)}")
            return None

    def read_template(self, template_path):
        try:
            doc = docx.Document(template_path)
            return '\n'.join([paragraph.text for paragraph in doc.paragraphs])
        except Exception as e:
            self.update_status(f"Error reading template: {str(e)}")
            return None

    def send_emails(self):
        try:
            # Disable send button
            self.send_button.state(['disabled'])
            
            # Read recipients and template
            recipients = self.read_csv(self.csv_path.get())
            if not recipients:
                raise ValueError("No recipients found in CSV")
                
            template_content = self.read_template(self.template_path.get())
            if not template_content:
                raise ValueError("Could not read template content")
            
            # Configure progress bar
            total_recipients = len(recipients)
            self.progress_var.set(0)
            
            # Connect to SMTP server
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                try:
                    server.login(self.email.get(), self.password.get())
                except smtplib.SMTPAuthenticationError:
                    self.update_status("Authentication failed. Please check your email and app password.")
                    return
                
                # Send emails
                for i, recipient in enumerate(recipients):
                    try:
                        msg = EmailMessage()
                        msg['From'] = self.email.get()
                        msg['To'] = recipient['Email']
                        msg['Subject'] = "Important Message"
                        
                        content = template_content.replace('[Name]', recipient['Name'])
                        msg.set_content(content)
                        
                        server.send_message(msg)
                        self.update_status(f"Successfully sent email to {recipient['Name']} ({recipient['Email']})")
                        
                        # Update progress
                        progress = ((i + 1) / total_recipients) * 100
                        self.progress_var.set(progress)
                        
                    except Exception as e:
                        self.update_status(f"Error sending email to {recipient['Email']}: {str(e)}")
                        continue
            
            self.update_status("\nEmail sending process completed!")
            
        except Exception as e:
            self.update_status(f"An error occurred: {str(e)}")
        finally:
            # Re-enable send button
            self.send_button.state(['!disabled'])

    def start_sending(self):
        # Validate inputs
        if not all([self.csv_path.get(), self.template_path.get(), self.email.get(), self.password.get()]):
            messagebox.showerror("Error", "Please fill in all fields")
            return
            
        # Confirm before sending
        if messagebox.askyesno("Confirm", "Are you sure you want to send the emails?"):
            # Clear status text
            self.status_text.delete(1.0, tk.END)
            
            # Start sending in a separate thread
            thread = threading.Thread(target=self.send_emails)
            thread.start()

def main():
    root = tk.Tk()
    app = EmailSenderGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
