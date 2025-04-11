import imaplib
import email
from email.header import decode_header
import pandas as pd
import os
import datetime
import re 
import uuid
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, messagebox
import configparser
import sys
from dateutil import parser as date_parser


def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))
    
def load_config():
    config = configparser.ConfigParser()
    config_path = os.path.join(get_base_dir(), 'email_config.ini')

    if os.path.exists(config_path):
        config.read(config_path)
    else:
        config['Credentials'] = {
            'email': '',
            'password': ''
        }
        config['Search'] = {
            'subject_keyword': 'Bukti Pembayaran Transaksi PT. KAI Persero',
            'unread_only': 'True'
        }
        config['Output'] = {
            'excel_file': 'email_attachment_report.xlsx',
            'attachments_dir': os.path.join(get_base_dir(), 'email_attachments')
        }
        with open(config_path, 'w') as f:
            config.write(f)
    return config, config_path

def save_config(config, config_path):
    with open(config_path, 'w') as f:
        config.write(f)

def create_attachments_dir(base_dir):
    attachments_dir = os.path.join(base_dir, 'email_attachments')
    os.makedirs(attachments_dir, exist_ok=True)
    return attachments_dir

def get_month_folder(attachments_dir, email_date):
    try:
        parsed_date = None
        if isinstance(email_date, str):
            try:
                # Use dateutil parser for more flexible date parsing
                parsed_date = date_parser.parse(email_date.split('(')[0].strip())
            except Exception:
                print(f"Couldn't parse date:{email_date}, using current date")
                parsed_date = datetime.datetime.now()
        else:
            # If email_date is already a datetime object
            parsed_date = email_date
            
        month_folder_name = parsed_date.strftime("%B %Y")
        month_folder_path = os.path.join(attachments_dir, month_folder_name)
        os.makedirs(month_folder_path, exist_ok=True)
        return month_folder_path
    except Exception as e:
        print(f"Error creating month folder: {e}")
        return attachments_dir

def save_attachment(part, attachments_dir, email_date=None):
    saved_attachments = []
    filename = part.get_filename()
    if not filename: 
        return saved_attachments
    try: 
        filename = decode_header(filename)[0][0]
        if isinstance(filename, bytes):
            filename = filename.decode('utf-8', errors='ignore')
    except Exception:
        filename = f"attachment_{uuid.uuid4()}"
    filename = re.sub(r'[^\w\-_\.]','_', filename)
    try:
        if email_date:
            save_dir = get_month_folder(attachments_dir, email_date)
        else:
            save_dir = attachments_dir
        
        filepath = os.path.join(save_dir, filename)
        counter = 1
        base, ext = os.path.splitext(filepath)
        while os.path.exists(filepath):
            filepath = f"{base}_{counter}{ext}"
            counter += 1
        
        with open(filepath, 'wb') as f:
            f.write(part.get_payload(decode=True))

        saved_attachments.append(filepath)
        return saved_attachments
    except Exception as e:
        print(f"Error saving attachment {filename}: {e}")
        return saved_attachments

def append_to_excel(new_data, output_file):
    try:
        if os.path.exists(output_file):
            existing_df = pd.read_excel(output_file)
            combined_df = pd.concat([existing_df, new_data], ignore_index=True)
            combined_df.drop_duplicates(subset=['Subject', 'Sender', 'Date'], keep='first', inplace=True)
            combined_df.to_excel(output_file, index=False)
            return f"Appended {len(new_data)} new emails to {output_file}"
        else: 
            new_data.to_excel(output_file, index=False)
            return f"Created new file {output_file} with {len(new_data)} emails"      
    except Exception as e:
        return f"Error appending to Excel: {e}"

def clean_subject(subject):
    if subject: 
        decoded_subject = []
        for part, encoding in decode_header(subject):
            if isinstance(part, bytes):
                part = part.decode(encoding or 'utf-8', errors='ignore')
            decoded_subject.append(part)
        return ' '.join(decoded_subject)
    return ''

def extract_email_content(email_message):
    email_content = ""
    
    if isinstance(email_message, str):
        try:
            email_message = email.message_from_string(email_message)
        except Exception as e:
            print(f"Could not parse email message: {e}")
            return ""
    
    if email_message.is_multipart():
        for part in email_message.walk():
            content_type = part.get_content_type()
            
            if content_type == 'text/plain':
                try:
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset('utf-8')
                    email_content += payload.decode(charset, errors='ignore') + "\n\n"
                except Exception as e:
                    print(f"Error decoding plain text part: {e}")
            
            elif content_type == 'text/html':
                try:
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset('utf-8')
                    html_content = payload.decode(charset, errors='ignore')
                    
                    soup = BeautifulSoup(html_content, 'html.parser')
                    
                    for script in soup(["script", "style"]):
                        script.decompose()
                    
                    text = soup.get_text(separator=' ', strip=True)
                    
                    text = re.sub(r'\s+', ' ', text).strip()
                    
                    email_content += text + "\n\n"
                except Exception as e:
                    print(f"Error processing HTML content: {e}")
    
    else:
        content_type = email_message.get_content_type()
        try:
            payload = email_message.get_payload(decode=True)
            charset = email_message.get_content_charset('utf-8')
            
            if content_type == 'text/plain':
                email_content = payload.decode(charset, errors='ignore')
            elif content_type == 'text/html':
                soup = BeautifulSoup(payload.decode(charset, errors='ignore'), 'html.parser')
                
                for script in soup(["script", "style"]):
                    script.decompose()
                
                text = soup.get_text(separator=' ', strip=True)
                
                email_content = re.sub(r'\s+', ' ', text).strip()
        except Exception as e:
            print(f"Error processing single-part email: {e}")
    
    return email_content.strip()
    
def search_emails(
        mail,
        attachments_dir,
        subject_keyword=None,
        start_date=None, 
        end_date=None,
        unread_only=False,
        status_callback=None
):
    search_criteria = []
    if subject_keyword:
        search_criteria.append(f'SUBJECT "{subject_keyword}"')
        
    if start_date:
        search_criteria.append(f'SINCE "{start_date.strftime("%d-%b-%Y")}"')
    if end_date:
        search_criteria.append(f'BEFORE "{end_date.strftime("%d-%b-%Y")}"')
    if unread_only:
        search_criteria.append('UNSEEN')
    
    try:
        search_string = ' '.join(search_criteria) if search_criteria else 'ALL'
        if status_callback:
            status_callback(f"Executing IMAP search with criteria: {search_string}")
        
        result, data = mail.search(None, search_string)
        
        if not data[0]:
            if status_callback:
                status_callback("No emails found matching the search criteria")
            return []
            
        email_count = len(data[0].split())
        if status_callback:
            status_callback(f"Found {email_count} emails matching search criteria")
        
        email_list = []
        for i, num in enumerate(data[0].split()):
            if status_callback:
                status_callback(f"Processing email {i+1}/{email_count}")
            result, email_data = mail.fetch(num, '(RFC822)')
            try:
                raw_email = email_data[0][1]
                email_message = email.message_from_bytes(raw_email)
                email_subject = clean_subject(email_message['Subject'])
                email_sender = email_message['From']
                email_date = email_message['Date']
                email_content = extract_email_content(email_message)
                
                attachment_paths = []
                if email_message.is_multipart():
                    for part in email_message.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                            
                        is_attachment = False
                        
                        if part.get('Content-Disposition') and part.get('Content-Disposition').startswith('attachment'):
                            is_attachment = True
                         
                        elif part.get_content_type() == 'application/octet-stream' or part.get_content_maintype() == 'application':
                            is_attachment = True
                            
                        elif part.get_filename():
                            is_attachment = True
                        
                        if is_attachment:
                            saved = save_attachment(part, attachments_dir, email_date)
                            if saved:
                                attachment_paths.extend(saved)
                                if status_callback:
                                    status_callback(f"  - Saved attachment: {os.path.basename(saved[0])}")
                
                email_list.append({
                    'Subject': email_subject,
                    'Sender': email_sender,
                    'Date': email_date, 
                    'Content': email_content,
                    'Attachments': '; '.join(attachment_paths) if attachment_paths else ''
                })
            except Exception as email_error:
                if status_callback:
                    status_callback(f"Error processing email {num}: {email_error}")
        
        if status_callback:
            status_callback(f"Successfully processed {len(email_list)} emails")
        return email_list

    except Exception as search_error:
        if status_callback:
            status_callback(f"Email search error: {search_error}")
        return []
    

class EmailProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title = 'Email Attachment Processor'
        self.root.geometry("700x600")
        self.config, self.config_path = load_config()
        
        # Create tabs
        self.tab_control = ttk.Notebook(root)
        self.setup_tab = ttk.Frame(self.tab_control)
        self.process_tab = ttk.Frame(self.tab_control)
        self.log_tab = ttk.Frame(self.tab_control)
        
        self.tab_control.add(self.setup_tab, text='Setup')
        self.tab_control.add(self.process_tab, text='Process Emails')
        self.tab_control.add(self.log_tab, text='Log')
        
        self.tab_control.pack(expand=1, fill="both")
        
        # Setup tab
        self.create_setup_tab()
        
        # Process tab
        self.create_process_tab()
        
        # Log tab
        self.create_log_tab()
        
        # Load config values
        self.load_config_values()

    def create_setup_tab(self):
        # Email credentials frame
        cred_frame = ttk.LabelFrame(self.setup_tab, text="Email Credentials")
        cred_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(cred_frame, text="Email:").grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
        self.email_var = tk.StringVar()
        ttk.Entry(cred_frame, textvariable=self.email_var, width=40).grid(column=1, row=0, padx=5, pady=5)
        
        ttk.Label(cred_frame, text="Password:").grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
        self.password_var = tk.StringVar()
        ttk.Entry(cred_frame, textvariable=self.password_var, show="*", width=40).grid(column=1, row=1, padx=5, pady=5)
        
        # Search criteria frame
        search_frame = ttk.LabelFrame(self.setup_tab, text="Search Criteria")
        search_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(search_frame, text="Subject Keyword:").grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
        self.subject_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.subject_var, width=40).grid(column=1, row=0, padx=5, pady=5)
        
        self.unread_var = tk.BooleanVar()
        ttk.Checkbutton(search_frame, text="Unread Only", variable=self.unread_var).grid(column=0, row=1, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # Output settings frame
        output_frame = ttk.LabelFrame(self.setup_tab, text="Output Settings")
        output_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(output_frame, text="Excel File:").grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
        self.excel_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.excel_var, width=40).grid(column=1, row=0, padx=5, pady=5)
        
        ttk.Label(output_frame, text="Attachments Directory:").grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
        self.dir_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.dir_var, width=40).grid(column=1, row=1, padx=5, pady=5)
        
        # Save button
        ttk.Button(self.setup_tab, text="Save Configuration", command=self.save_config_values).pack(pady=10)

    def create_process_tab(self):
        # Date range frame
        date_frame = ttk.LabelFrame(self.process_tab, text="Date Range")
        date_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(date_frame, text="Start Date (YYYY-MM-DD):").grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
        self.start_date_var = tk.StringVar()
        ttk.Entry(date_frame, textvariable=self.start_date_var, width=15).grid(column=1, row=0, padx=5, pady=5)
        
        ttk.Label(date_frame, text="End Date (YYYY-MM-DD):").grid(column=2, row=0, sticky=tk.W, padx=5, pady=5)
        self.end_date_var = tk.StringVar()
        ttk.Entry(date_frame, textvariable=self.end_date_var, width=15).grid(column=3, row=0, padx=5, pady=5)
        
        # Today button
        def set_today():
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            self.end_date_var.set(today)
        
        ttk.Button(date_frame, text="Set Today", command=set_today).grid(column=4, row=0, padx=5, pady=5)
        
        # Process button
        ttk.Button(self.process_tab, text="Process Emails", command=self.process_emails).pack(pady=10)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(self.process_tab, text="Progress")
        progress_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=5, pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        ttk.Label(progress_frame, textvariable=self.status_var).pack(padx=5, pady=5)
    
    def create_log_tab(self):
        log_frame = ttk.Frame(self.log_tab)
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=25)
        self.log_text.pack(side=tk.LEFT, fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        ttk.Button(self.log_tab, text="Clear Log", command=self.clear_log).pack(pady=5)
    
    def load_config_values(self):
        if 'Credentials' in self.config:
            self.email_var.set(self.config['Credentials'].get('email', ''))
            self.password_var.set(self.config['Credentials'].get('password', ''))
        
        if 'Search' in self.config:
            self.subject_var.set(self.config['Search'].get('subject_keyword', ''))
            self.unread_var.set(self.config['Search'].getboolean('unread_only', True))
        
        if 'Output' in self.config:
            self.excel_var.set(self.config['Output'].get('excel_file', 'email_attachment_report.xlsx'))
            self.dir_var.set(self.config['Output'].get('attachments_dir', ''))
    
    def save_config_values(self):
        if 'Credentials' not in self.config:
            self.config['Credentials'] = {}
        self.config['Credentials']['email'] = self.email_var.get()
        self.config['Credentials']['password'] = self.password_var.get()
        
        if 'Search' not in self.config:
            self.config['Search'] = {}
        self.config['Search']['subject_keyword'] = self.subject_var.get()
        self.config['Search']['unread_only'] = str(self.unread_var.get())
        
        if 'Output' not in self.config:
            self.config['Output'] = {}
        self.config['Output']['excel_file'] = self.excel_var.get()
        self.config['Output']['attachments_dir'] = self.dir_var.get()
        
        save_config(self.config, self.config_path)
        messagebox.showinfo("Configuration", "Configuration saved successfully!")
    
    def update_status(self, message):
        self.status_var.set(message)
        self.log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def process_emails(self):
        try:
            # Parse dates
            start_date = None
            end_date = None
            
            if self.start_date_var.get().strip():
                try:
                    start_date = datetime.datetime.strptime(self.start_date_var.get().strip(), "%Y-%m-%d")
                except ValueError:
                    messagebox.showerror("Date Error", "Invalid start date format. Use YYYY-MM-DD")
                    return
            
            if self.end_date_var.get().strip():
                try:
                    end_date = datetime.datetime.strptime(self.end_date_var.get().strip(), "%Y-%m-%d")
                    # Add one day to include emails from the end date
                    end_date += datetime.timedelta(days=1)
                except ValueError:
                    messagebox.showerror("Date Error", "Invalid end date format. Use YYYY-MM-DD")
                    return
            
            # Validate required fields
            if not self.email_var.get() or not self.password_var.get():
                messagebox.showerror("Input Error", "Email and password are required")
                return
            
            # Create attachments directory if needed
            attachments_dir = self.dir_var.get()
            if not attachments_dir:
                attachments_dir = create_attachments_dir(get_base_dir())
            else:
                os.makedirs(attachments_dir, exist_ok=True)
            
            self.update_status("Connecting to email server...")
            
            # Process in a separate thread to avoid freezing UI
            self.tab_control.select(self.log_tab)  # Switch to log tab
            
            # Connect to email
            try:
                with imaplib.IMAP4_SSL('imap.gmail.com') as mail:
                    self.update_status("Logging in...")
                    mail.login(self.email_var.get(), self.password_var.get())
                    self.update_status("Connected successfully")
                    
                    mail.select('inbox')
                    self.update_status("Searching for emails...")
                    
                    emails = search_emails(
                        mail, 
                        attachments_dir,
                        subject_keyword=self.subject_var.get(),
                        start_date=start_date,
                        end_date=end_date, 
                        unread_only=self.unread_var.get(),
                        status_callback=self.update_status
                    )
                    
                    if emails:
                        df = pd.DataFrame(emails)
                        output_file = self.excel_var.get()
                        if not output_file:
                            output_file = os.path.join(get_base_dir(), 'email_attachment_report.xlsx')
                        
                        result = append_to_excel(df, output_file)
                        self.update_status(result)
                        messagebox.showinfo("Process Complete", f"Successfully processed {len(emails)} emails")
                    else:
                        self.update_status("No emails were found or processed")
                        messagebox.showinfo("Process Complete", "No emails were found matching your criteria")
            
            except imaplib.IMAP4.error as login_error:
                error_msg = f"IMAP Login Error: {login_error}"
                self.update_status(error_msg)
                messagebox.showerror("Login Error", error_msg)
            
            except Exception as e:
                error_msg = f"Unexpected error: {e}"
                self.update_status(error_msg)
                messagebox.showerror("Error", error_msg)
        
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")

def run_cli():
    config, __ = load_config()
    print("Email Attachment Processor - CLI Mode")
    print("-------------------------------------")
    email = config['Credentials'].get('email', '')
    password = config['Credentials'].get('password', '')

    if not email or not password: 
        print('Email credentials not found in config. Please run the GUI version first to set up.')
        return
    attachments_dir = config['Output'].get('attachments_dir', '')
    if not attachments_dir:
        attachments_dir = create_attachments_dir(get_base_dir())
    
    try:
        print("Connecting to email server...")
        with imaplib.IMAP4_SSL('imap.gmail.com') as mail:
            print("Logging in...")
            mail.login(email, password)
            print("Connected successfully")
            mail.select('inbox')
            print("Searching for emails...")

            emails = search_emails(
                mail,
                attachments_dir,
                subject_keyword= config['Search'].get('subject_keyword', ''),
                unread_only=config['Search'].getboolean('unread_only', True),
                status_callback=print
            )

            if emails: 
                df = pd.DataFrame(emails)
                output_file = config['Output'].get('excel_file', 'email_attachment_report.xlsx')
                if not os.path.isabs(output_file):
                    output_file = os.path.join(get_base_dir(), output_file)
                result = append_to_excel(df, output_file)
                print(result)
            else:
                print('No emails were found or processed')
    except imaplib.IMAP4.error as login_error:
        print(f"IMAP Login Error:{login_error}")
    except Exception as e: 
        print(f"Unexpected error: {e}")

def main():
    if len(sys.argv) > 1 and sys.argv[1] == '--cli':
        run_cli()
    else:
        root = tk.Tk()
        app = EmailProcessorApp(root)
        root.mainloop()

if __name__ == '__main__':
    main()