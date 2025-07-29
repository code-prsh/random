import pandas as pd
import os
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
import time
import mimetypes

class EmailSystem:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.data = None
        self.resume_link = None
        self.template = """Subject: Application for [Position] in [Company Name]

Dear [Company Name] HR Team,

I hope this email finds you well. I am writing to express my interest in the [Position] position at [Company Name].

[Your custom message here]

You can find my resume here: [Resume Link]

I would welcome the opportunity to discuss how my skills and experience align with your needs.

Thank you for your time and consideration. I look forward to your response.

Best regards,
[Your Name]
[Your Contact Information]"""

    def load_data(self):
        """Load and clean data from Excel file"""
        try:
            # Read the Excel file
            xls = pd.ExcelFile(self.excel_file)
            print(f"Available sheets: {xls.sheet_names}")
            
            # Read the first sheet
            sheet_name = xls.sheet_names[0]
            print(f"\nReading sheet: {sheet_name}")
            
            # Read data with first row as header
            self.data = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=0)
            
            # Clean up column names and handle any leading/trailing spaces
            self.data.columns = [str(col).strip() for col in self.data.columns]
            
            # Convert all data to strings to handle any mixed types
            self.data = self.data.astype(str)
            
            # Remove any empty rows
            self.data = self.data.replace('nan', '').replace('None', '').replace('', pd.NA).dropna(how='all')
            
            print("\nFirst few rows of data:")
            print(self.data.head())
            
            print("\nAvailable columns (with first few values):")
            max_display = min(3, len(self.data))  # Show values from first 3 rows max
            for idx, col in enumerate(self.data.columns):
                sample_values = self.data[col].head(max_display).tolist()
                sample_text = ", ".join([str(v) for v in sample_values if pd.notna(v)])[:50]
                if len(sample_text) > 47:
                    sample_text = sample_text[:47] + "..."
                print(f"{idx + 1}. {col} (Sample: {sample_text})")
                
        except Exception as e:
            print(f"Error loading Excel file: {str(e)}")

    def set_template(self, template=None):
        """Set or update the email template"""
        if template:
            self.template = template
        print("\nCurrent email template:")
        print("-" * 50)
        print(self.template)
        print("-" * 50)

    def attach_file(self, msg, filepath):
        """Attach a file to the email"""
        if not os.path.isfile(filepath):
            print(f"Warning: File not found: {filepath}")
            return False
            
        # Guess the content type based on the file's extension
        ctype, encoding = mimetypes.guess_type(filepath)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
            
        maintype, subtype = ctype.split('/', 1)
        
        try:
            with open(filepath, 'rb') as fp:
                part = MIMEBase(maintype, subtype)
                part.set_payload(fp.read())
                encoders.encode_base64(part)
                
                # Add header with filename
                filename = os.path.basename(filepath)
                part.add_header('Content-Disposition', 'attachment', filename=filename)
                msg.attach(part)
            return True
        except Exception as e:
            print(f"Error attaching file {filepath}: {str(e)}")
            return False

    def send_emails(self, smtp_config, test_mode=True, batch_size=100, email_col=None, company_col=None, progress_callback=None):
        """
        Send emails to the companies in batches
        
        Args:
            smtp_config (dict): SMTP configuration
            test_mode (bool): If True, only show previews
            batch_size (int): Number of emails to send in each batch
            email_col (str): Name of the column containing email addresses
            company_col (str): Name of the column containing company names
            progress_callback (callable): Optional callback for progress updates (0-1)
        """
        if self.data is None:
            print("No data loaded. Please load data first.")
            return
            
        if test_mode:
            print("\n--- TEST MODE - No emails will be sent ---")
            
        # Ask for resume link if not set
        if not hasattr(self, 'resume_link') or not self.resume_link:
            self.resume_link = input("\nPlease enter your Google Drive resume link (or press Enter to skip): ").strip()
            if self.resume_link and self.resume_link.startswith(('http://', 'https://')):
                print("Resume link added.")
            elif self.resume_link:
                print("Warning: The link should start with http:// or https://")
                self.resume_link = None
            else:
                print("No resume link will be included.")
        
        # Display available columns and get user input for mapping
        print("\nAvailable columns in your Excel file:")
        for idx, col in enumerate(self.data.columns):
            print(f"{idx + 1}. {col}")
        
        # Auto-detect columns based on common patterns
        email_col = None
        company_col = None
        
        # First, try exact matches for known column names
        for col in self.data.columns:
            col_lower = str(col).lower()
            if col_lower in ['email', 'e-mail', 'email address']:
                email_col = col
            elif col_lower in ['org. name', 'org name', 'organization', 'company', 'company name']:
                company_col = col
        
        # If not found, try partial matches
        if email_col is None or company_col is None:
            for col in self.data.columns:
                col_lower = str(col).lower()
                if email_col is None and 'mail' in col_lower:
                    email_col = col
                if company_col is None and ('org' in col_lower or 'company' in col_lower):
                    company_col = col
        
        # If still not found, use first column for company name and look for email
        if company_col is None and len(self.data.columns) > 0:
            company_col = self.data.columns[0]
            
        if email_col is None and len(self.data.columns) > 1:
            # Look for any column that looks like an email
            for col in self.data.columns[1:]:
                if self.data[col].astype(str).str.contains('@').any():
                    email_col = col
                    break
        
        # If we still don't have both, ask the user
        if email_col is None or company_col is None:
            print("\nCould not automatically detect all required columns.")
            
        if email_col is None:
            print("\nPlease select the column containing email addresses:")
            for idx, col in enumerate(self.data.columns):
                print(f"{idx + 1}. {col}")
            email_col_idx = int(input("Enter the column number: ").strip()) - 1
            email_col = self.data.columns[email_col_idx]
        else:
            print(f"\nUsing column '{email_col}' for email addresses")
            
        if company_col is None:
            print("\nPlease select the column containing company names:")
            for idx, col in enumerate(self.data.columns):
                print(f"{idx + 1}. {col}")
        
        # Filter out rows without email addresses
        valid_emails = self.data.dropna(subset=[email_col]).copy()
        total_emails = len(valid_emails)
        
        if total_emails == 0:
            print("No valid email addresses found in the selected column.")
            return
            
        print(f"\nFound {total_emails} valid email addresses.")
        print(f"Will send emails in batches of {batch_size}.")
        
        # Calculate number of batches
        num_batches = (total_emails + batch_size - 1) // batch_size
        
        # User details should be provided in the template by Streamlit
        user_details = smtp_config.get('user_details', {})
        
        # Additional columns for personalization (handled by Streamlit UI)
        additional_cols = {}
        
        # If additional columns are provided in smtp_config, use them
        if 'additional_cols' in smtp_config:
            additional_cols = smtp_config['additional_cols']
            print(f"Using {len(additional_cols)} additional columns for personalization")
        
        # Ask for confirmation before starting
        if not test_mode:
            try:
                print(f"\n{'='*50}")
                print(f"Connecting to SMTP server {smtp_config['smtp_server']}:{smtp_config.get('smtp_port', 587)}...")
                server = smtplib.SMTP(smtp_config['smtp_server'], smtp_config.get('smtp_port', 587), timeout=30)
                server.ehlo()
                server.starttls()
                server.ehlo()
                print("Logging in to SMTP server...")
                server.login(smtp_config['smtp_username'], smtp_config['smtp_password'])
                print("Successfully connected to SMTP server")
                print("="*50 + "\n")
            except Exception as e:
                print(f"Error connecting to SMTP server: {str(e)}")
                if server:
                    try:
                        server.quit()
                    except:
                        pass
                return
        
        def update_progress(progress):
            if callable(progress_callback):
                try:
                    progress_callback(progress)
                except Exception as e:
                    print(f"Error in progress callback: {e}")
        
        # Process emails in batches
        for batch_num in range(num_batches):
            start_idx = batch_num * batch_size
            end_idx = min((batch_num + 1) * batch_size, total_emails)
            batch = valid_emails.iloc[start_idx:end_idx]
            
            print(f"\nProcessing batch {batch_num + 1}/{num_batches} ({len(batch)} emails)")
            
            # Update progress at start of batch
            if progress_callback:
                update_progress(start_idx / total_emails)
            
            # Process each email in the current batch
            for idx, (_, row) in enumerate(batch.iterrows(), 1):
                try:
                    # Update progress before each email
                    if progress_callback:
                        current_progress = (start_idx + idx - 1) / total_emails
                        update_progress(current_progress)
                    
                    company_email = str(row[email_col]).strip()
                    company_name = str(row[company_col]).strip()
                    
                    # Skip if email is not valid
                    if '@' not in company_email:
                        print(f"Skipping invalid email: {company_email}")
                        continue
                        
                    # Personalize the template with company information
                    email_content = self.template
                    
                    # Replace company information from Excel
                    email_content = email_content.replace('[Company Name]', company_name)
                    for col, placeholder in additional_cols.items():
                        if col in row and pd.notna(row[col]):
                            email_content = email_content.replace(f'[{placeholder}]', str(row[col]))
                        else:
                            email_content = email_content.replace(f'[{placeholder}]', '')
                    
                    # Replace user details and resume link
                    for key, value in user_details.items():
                        email_content = email_content.replace(f'[{key}]', value)
                    
                    # Add resume link if available
                    if hasattr(self, 'resume_link') and self.resume_link:
                        email_content = email_content.replace('[Resume Link]', self.resume_link)
                    else:
                        email_content = email_content.replace('You can find my resume here: [Resume Link]\n\n', '')
                    
                    if test_mode:
                        subject_line = email_content.split('\n')[0]
                        body = '\n'.join(email_content.split('\n')[1:])
                        print("\n" + "="*50)
                        print(f"To: {company_email}")
                        print(f"Subject: {subject_line}")
                        print("\n" + body)
                        print("="*50 + "\n")
                    else:
                        # Create the email
                        msg = MIMEMultipart()
                        msg['From'] = smtp_config['smtp_username']
                        msg['To'] = company_email
                        msg['Subject'] = email_content.split('\n')[0].replace('Subject: ', '')
                        
                        # Add the email body
                        msg.attach(MIMEText("\n".join(email_content.split('\n')[1:]), 'plain'))
                        
                        try:
                            # Send the email
                            if not test_mode and server:
                                server.send_message(msg)
                                print(f"Email sent to {company_email}")
                            else:
                                print(f"[TEST MODE] Would send email to {company_email}")
                                print(f"Subject: {msg['Subject']}")
                                print("-" * 50)
                        except Exception as e:
                            print(f"Error sending email to {company_email}: {str(e)}")
                            # If there's an error, try to reconnect
                            if not test_mode and "smtplib.SMTPServerDisconnected" in str(e):
                                print("Attempting to reconnect to SMTP server...")
                                try:
                                    server = smtplib.SMTP(smtp_config['smtp_server'], smtp_config.get('smtp_port', 587))
                                    server.starttls()
                                    server.login(smtp_config['smtp_username'], smtp_config['smtp_password'])
                                    print("Reconnected to SMTP server")
                                    # Retry sending the email
                                    server.send_message(msg)
                                    print(f"Email sent to {company_email} after reconnection")
                                except Exception as retry_error:
                                    print(f"Failed to resend to {company_email}: {str(retry_error)}")
                    
                    # Add delay to avoid being flagged as spam (only if not in test mode)
                    if not test_mode:
                        delay_between_emails = 2
                        if idx < len(batch):
                            time.sleep(delay_between_emails)
                            
                except Exception as e:
                    print(f"Error sending email to {row[email_col]}: {str(e)}")
                    continue
                    
                # Update progress after each email
                if progress_callback:
                    current_progress = (start_idx + idx) / total_emails
                    update_progress(current_progress)
        
            # Add a delay between batches
            if batch_num < num_batches - 1 and len(batch) > 0:
                delay_between_batches = 15
                print(f"Waiting {delay_between_batches} seconds before next batch...")
                # Update progress during batch delay
                if progress_callback:
                    progress = end_idx / total_emails
                    update_progress(progress)
                time.sleep(delay_between_batches)
        
        # Final progress update
        if progress_callback:
            update_progress(1.0)
        
        # Close SMTP connection at the very end
        if not test_mode and server:
            try:
                print("\nClosing SMTP connection...")
                server.quit()
                print("SMTP connection closed successfully")
            except Exception as e:
                print(f"Error closing SMTP connection: {str(e)}")
        
        print("\nEmail sending process completed!")

if __name__ == "__main__":
    # Initialize the email system
    excel_file = 'SampleData.xlsx'
    email_system = EmailSystem(excel_file)
    
    # Load and display data
    email_system.load_data()
    
    # Show current template
    email_system.set_template()
    
    # Ask if user wants to customize the template
    if input("\nDo you want to customize the email template? (y/n): ").lower() == 'y':
        print("\nCurrent template will be shown. Make your changes and save the file when done.")
        print("Note: Use [Placeholder] for variables that will be replaced.")
        input("Press Enter to continue...")
        
        # Open default text editor with the template
        import tempfile
        import os
        
        with tempfile.NamedTemporaryFile(mode='w+', suffix='.txt', delete=False) as temp:
            temp.write(email_system.template)
            temp_path = temp.name
        
        # Open the default text editor
        os.system(f'notepad.exe {temp_path}')
        
        # Wait for user to finish editing
        input("Press Enter after saving the template...")
        
        # Read the modified template
        with open(temp_path, 'r', encoding='utf-8') as f:
            new_template = f.read()
        
        # Clean up
        os.unlink(temp_path)
        
        # Update the template
        email_system.set_template(new_template)
    
    # Show sending summary
    print("\n" + "="*50)
    print("READY TO SEND EMAILS")
    print("="*50)
    print(f"Total emails to send: {len(email_system.data)}")
    print(f"Number of batches: {(len(email_system.data) + 100 - 1) // 100}")
    print(f"Emails per batch: 100")
    print("="*50)
    
    # Skip test mode - we'll send real emails directly
    test_mode = False
    
    # Set batch size to 2 for testing
    batch_size = 2
    
    if not test_mode:
        print("\nPlease provide SMTP configuration:")
        smtp_config = {
            'smtp_server':"smtp.gmail.com",
            'smtp_port': 587,
            'smtp_username': input("Your email address: "),
            'smtp_password': input("Your email password or app password: ")
        }
    else:
        smtp_config = {}
    
    # Start sending emails
    email_system.send_emails(smtp_config, test_mode=test_mode, batch_size=batch_size)
    
    print("\nProcess completed. Check the logs above for details.")
