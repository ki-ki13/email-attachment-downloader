import imaplib
import email
from email.header import decode_header
import pandas as pd
import os
from dotenv import load_dotenv
import datetime
import re 
from bs4 import BeautifulSoup

load_dotenv()

def append_to_excel(new_data, output_file):
    try:
        if os.path.exists(output_file):
            existing_df = pd.read_excel(output_file)
            combined_df = pd.concat([existing_df, new_data], ignore_index=True)
            combined_df.drop_duplicates(subset=['Subject', 'Sender', 'Date'], keep='first', inplace=True)
            combined_df.to_excel(output_file, index=False)
            print(f"Appended {len(new_data)} new emails to {output_file}")
        else: 
            new_data.to_excel(output_file, index=False)
            print(f"Created new file {output_file} with {len(new_data)} emails")      
    except Exception as e:
        print(f"Error appending to Excel: {e}")

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
        subject_keyword=None,
        start_date=None, 
        end_date=None,
        unread_only=False
):
    search_criteria = []
    if start_date : 
        search_criteria.append(f'SINCE "{start_date.strftime("%d-%b-%Y")}"')
    if end_date : 
        search_criteria.append(f'BEFORE "{end_date.strftime("%d-%b-%Y")}"')
    if unread_only: 
        search_criteria.append('UNSEEN')
    try:
        search_string = ' '.join(search_criteria) if search_criteria else 'ALL'
        result, data = mail.search(None, search_string)

        email_list = []
        for num in data[0].split():
            result, email_data = mail.fetch(num, '(RFC822)')
            try:
                raw_email = email_data[0][1]
                email_message = email.message_from_bytes(raw_email)
                email_subject = clean_subject(email_message['Subject'])
                if subject_keyword and subject_keyword.lower() not in email_subject.lower():
                    continue
                email_sender = email_message['From']
                email_date = email_message['Date']
                email_content = extract_email_content(email_message)
                email_list.append({
                    'Subject': email_subject,
                    'Sender': email_sender,
                    'Date': email_date, 
                    'Content': email_content
                })
            except Exception as email_error:
                print(f"Error processing email {num}: {email_error}")
        return email_list

    except Exception as search_error:
        print(f"Email search error: {search_error}")
        return []
    

def main():
    user_mail= os.getenv('email')
    user_pass = os.getenv('password')
    output_file = 'email_data3.xlsx'
    try:
        with imaplib.IMAP4_SSL('imap.gmail.com') as mail:
            mail.login(user_mail, user_pass)
            mail.select('inbox')
            emails = search_emails(
                mail, 
                subject_keyword='Your Grab E-Receipt',
                start_date= datetime.datetime(2025,3,20),
                end_date=datetime.datetime(2025,4,7), 
                unread_only=True
            )
            if emails: 
                df = pd.DataFrame(emails)
                append_to_excel(df, output_file)
    except imaplib.IMAP4.error as login_error:
        print(f"IMAP Login Error: {login_error}")
    except Exception as e: 
        print(f"Unexpected error: {e}")


if __name__ == '__main__':
    main()