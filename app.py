import streamlit as st
import pandas as pd
import time
import sys
import os
from io import StringIO
from email_system import EmailSystem

# Global flag for cancellation
if 'cancelled' not in st.session_state:
    st.session_state.cancelled = False
    
# Function to handle cancellation
def cancel_sending():
    st.session_state.cancelled = True

# Set page config
st.set_page_config(
    page_title="Cold Email Sender",
    page_icon="✉️",
    layout="wide"
)

# Custom CSS for better UI
st.markdown("""
<style>
    .main {
        max-width: 1200px;
        padding: 2rem;
    }
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    .success-msg {
        color: #4CAF50;
        font-weight: bold;
    }
    .error-msg {
        color: #f44336;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# App title and description
st.title("✉️ Cold Email Sender")
st.write("Upload your company dataset and send personalized emails in batches.")

# Initialize session state
if 'email_sent' not in st.session_state:
    st.session_state.email_sent = False
if 'smtp_connected' not in st.session_state:
    st.session_state.smtp_connected = False
if 'progress' not in st.session_state:
    st.session_state.progress = 0

# Sidebar for SMTP configuration
with st.sidebar:
    st.header("SMTP Configuration")
    smtp_server = st.text_input("SMTP Server", "smtp.gmail.com", key="smtp_server")
    smtp_port = st.number_input("SMTP Port", value=587, min_value=1, max_value=65535, key="smtp_port")
    smtp_username = st.text_input("Your Email", key="smtp_username")
    smtp_password = st.text_input("Password/App Password", type="password", key="smtp_password")
    
    st.session_state.smtp_config = {
        'smtp_server': smtp_server,
        'smtp_port': smtp_port,
        'smtp_username': smtp_username,
        'smtp_password': smtp_password
    }

# Main content area
tab1, tab2 = st.tabs(["📤 Send Emails", "📊 Dataset Preview"])

with tab1:
    # File upload
    uploaded_file = st.file_uploader("Upload Company Dataset (Excel/CSV)", type=['xlsx', 'xls', 'csv'], key="file_uploader")
    
    if uploaded_file is not None:
        try:
            # Read the uploaded file
            if uploaded_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(uploaded_file)
            else:
                df = pd.read_csv(uploaded_file)
                
            # Store in session state
            st.session_state.df = df
            st.session_state.file_uploaded = True
            
            # Show success message
            st.success(f"Successfully uploaded {uploaded_file.name} with {len(df)} rows")
            
            # Show column selector
            email_col = st.selectbox(
                "Select the column containing email addresses",
                df.columns,
                index=0,
                help="Select the column that contains the recipient email addresses",
                key="email_column_selector"
            )
            
            company_col = st.selectbox(
                "Select the column containing company names",
                df.columns,
                index=1 if len(df.columns) > 1 else 0,
                help="Select the column that contains the company names",
                key="company_column_selector"
            )
            
            # Additional columns for personalization
            st.subheader("Additional Personalization")
            additional_cols = {}
            cols = st.columns(2)
            for i, col in enumerate(df.columns):
                if col not in [email_col, company_col]:
                    with cols[i % 2]:
                        if st.checkbox(f"Use '{col}' in email", key=f"col_{col}"):
                            placeholder = st.text_input(
                                f"Placeholder for {col}",
                                value=col.replace(" ", "").title(),
                                key=f"ph_{col}",
                                help=f"Use [placeholder] in your email template"
                            )
                            additional_cols[col] = placeholder
            
            # Email template
            st.subheader("Email Template")
            default_template = """Subject: Application for [Position] - [Your Name]

Dear [Company Name] HR Team,

I hope this email finds you well. I am writing to express my interest in the [Position] position at [Company Name].

[Your custom message here]

You can find my resume here: [Resume Link]

I would welcome the opportunity to discuss how my skills and experience align with your needs.

Thank you for your time and consideration. I look forward to your response.

Best regards,
[Your Name]
[Your Contact Information]"""
            
            email_template = st.text_area("Email Template", value=default_template, height=300, key="email_template")
            
            # User details
            st.subheader("Your Information")
            col1, col2 = st.columns(2)
            with col1:
                your_name = st.text_input("Your Full Name", key="your_name")
                your_position = st.text_input("Position You're Applying For", key="your_position")
            with col2:
                your_email = st.text_input("Your Email", key="your_email")
                your_phone = st.text_input("Your Phone Number", key="your_phone")
            
            custom_message = st.text_area("Custom Message", key="custom_message")
            
            # Resume Link Section
            st.subheader("Resume Link")
            resume_link = st.text_input("Google Drive Link to Your Resume", 
                                     placeholder="https://drive.google.com/your-resume-link",
                                     help="Make sure the link is set to 'Anyone with the link can view'",
                                     key="resume_link")
            
            # Sending Settings Section
            with st.expander("⚙️ Advanced Sending Settings"):
                batch_size = st.number_input("Emails per batch", 
                                          min_value=1, 
                                          max_value=50, 
                                          value=2,
                                          key="batch_size_input")
                delay_between_emails = st.number_input("Delay between emails (seconds)", 
                                                    min_value=1, 
                                                    max_value=60, 
                                                    value=2,
                                                    key="email_delay_input")
                delay_between_batches = st.number_input("Delay between batches (seconds)", 
                                                     min_value=5, 
                                                     max_value=300, 
                                                     value=15,
                                                     key="batch_delay_input")
            
            # Column Selection
            st.markdown("---")
            st.subheader("Column Selection")
            
            # Auto-detect likely columns
            email_col = st.selectbox(
                "Select Email Column",
                options=df.columns,
                index=next((i for i, col in enumerate(df.columns) if 'email' in col.lower()), 0),
                key="email_column_selector_main"
            )
            
            company_col = st.selectbox(
                "Select Company Name Column",
                options=df.columns,
                index=next((i for i, col in enumerate(df.columns) if 'company' in col.lower() or 'org' in col.lower() or 'name' in col.lower()), 0),
                key="company_column_selector_main"
            )
            
            # Ready to Send Section
            st.markdown("---")
            st.subheader("Ready to Send")
            
            # Check all required fields are filled
            required_fields = {
                "SMTP Configuration": all([smtp_server, smtp_port, smtp_username, smtp_password]),
                "Your Information": all([your_name, your_position, your_email, your_phone]),
                "Resume Link": bool(resume_link.strip()),
                "Email Template": bool(email_template.strip())
            }
            
            # Show validation status
            st.write("### Validation Check")
            for field, is_valid in required_fields.items():
                status = "✅" if is_valid else "❌"
                st.write(f"{status} {field}")
            
            all_valid = all(required_fields.values())
            
            # Disable button if not all fields are valid
            if st.button("🚀 Send Emails", 
                       type="primary", 
                       disabled=not all_valid,
                       help="Fill in all required fields to enable"):
                if not all_valid:
                    st.error("Please fill in all required fields")
                    st.stop()
                else:
                    # Initialize email system
                    email_system = EmailSystem("temp.xlsx")
                    email_system.data = df
                    email_system.template = email_template
                    
                    # Set user details
                    user_details = {
                        'Your Name': your_name,
                        'Position': your_position,
                        'Your Email': your_email,
                        'Your Phone': your_phone,
                        'Your custom message here': custom_message,
                        'Your Contact Information': f"Email: {your_email}\nPhone: {your_phone}"
                    }
                    
                    # Set up progress bar
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Prepare email system
                    email_system = EmailSystem(uploaded_file)
                    email_system.load_data()
                    
                    # Set resume link
                    email_system.resume_link = resume_link
                    
                    # Set user details in the template
                    template = email_template
                    template = template.replace("[Your Name]", your_name)
                    template = template.replace("[Your Email]", your_email)
                    template = template.replace("[Your Phone]", your_phone)
                    template = template.replace("[Position]", your_position)
                    template = template.replace("[Your custom message here]", custom_message)
                    template = template.replace(
                        "[Your Contact Information]", 
                        f"Email: {your_email}\nPhone: {your_phone}"
                    )
                    template = template.replace("[Resume Link]", resume_link)
                    
                    email_system.template = template
                    
                    # Set up SMTP config with user details
                    smtp_config = {
                        'smtp_server': smtp_server,
                        'smtp_port': smtp_port,
                        'smtp_username': smtp_username,
                        'smtp_password': smtp_password,
                        'user_details': {
                            'Your Name': your_name,
                            'Position': your_position,
                            'Your Email': your_email,
                            'Your Phone': your_phone,
                            'Your custom message here': custom_message,
                            'Your Contact Information': f"Email: {your_email} | Phone: {your_phone}",
                            'Resume Link': resume_link
                        },
                        'additional_cols': {}
                    }
                    
                    # Set up progress tracking
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Create a placeholder for logs
                    log_placeholder = st.empty()
                    
                    # Redirect print to the UI
                    import sys
                    from io import StringIO
                    
                    class StdoutCatcher:
                        def __init__(self):
                            self.log = ""
                        
                        def write(self, message):
                            self.log += message
                            log_placeholder.text_area("Sending Logs", value=self.log, height=200)
                    
                    # Start capturing stdout
                    old_stdout = sys.stdout
                    sys.stdout = StdoutCatcher()
                    
                    # Reset cancellation flag at start
                    st.session_state.cancelled = False
                    
                    # Create containers for UI elements
                    progress_container = st.empty()
                    stop_container = st.empty()
                    
                    # Initialize progress bar
                    progress_bar = progress_container.progress(0)
                    
                    # Add stop button
                    stop_button = stop_container.button("🛑 Stop Sending", 
                                                    on_click=cancel_sending,
                                                    key="stop_button")
                    
                    # Status text
                    status_text = st.empty()
                    
                    try:
                        # Get total number of valid emails (non-null email addresses)
                        valid_emails = df[df[email_col].notna() & (df[email_col] != '')]
                        total_emails = len(valid_emails)
                        
                        # Use session state to track progress across reruns
                        if 'progress_state' not in st.session_state:
                            st.session_state.progress_state = {
                                'last_update': 0,
                                'last_progress': -1
                            }
                        
                        def update_progress(progress):
                            """Update progress bar and status text"""
                            try:
                                # Get current timestamp
                                current_time = time.time()
                                
                                # Throttle updates to at most once every 0.5 seconds
                                if current_time - st.session_state.progress_state['last_update'] < 0.5:
                                    return
                                
                                # Only update if progress has changed significantly (1% or more)
                                current_progress = int(progress * 100)
                                if current_progress == st.session_state.progress_state['last_progress']:
                                    return
                                
                                # Update state
                                st.session_state.progress_state.update({
                                    'last_update': current_time,
                                    'last_progress': current_progress
                                })
                                
                                # Ensure progress is between 0 and 1
                                safe_progress = max(0.0, min(float(progress), 1.0))
                                
                                # Update progress bar with error handling
                                try:
                                    progress_bar.progress(safe_progress)
                                except Exception as e:
                                    print(f"Progress bar error: {str(e)}")
                                
                                # Calculate current email being processed
                                current = min(int(round(safe_progress * total_emails)), total_emails)
                                
                                # Update status text with error handling
                                try:
                                    status_text.text(f"Sending email {current} of {total_emails}...")
                                except Exception as e:
                                    print(f"Status text error: {str(e)}")
                                
                                # Check for cancellation
                                if st.session_state.get('cancelled', False):
                                    raise Exception("Process cancelled by user")
                                    
                            except Exception as e:
                                # Only re-raise cancellation exceptions
                                if 'cancelled' in str(e).lower():
                                    raise e
                                # Log other errors but don't crash the app
                                print(f"Progress update error: {str(e)}")
                                if st.session_state.get('cancelled', False):
                                    raise Exception("Process cancelled by user")
                        
                        # Call the email sending function
                        with st.spinner("Sending emails..."):
                            # Send emails with progress tracking
                            try:
                                email_system.send_emails(
                                    smtp_config=smtp_config,
                                    test_mode=False,
                                    batch_size=batch_size,
                                    email_col=email_col,
                                    company_col=company_col,
                                    progress_callback=update_progress
                                )
                            except Exception as e:
                                if "cancelled" not in str(e).lower():
                                    raise e
                            
                            # If we get here, sending completed successfully
                            if not st.session_state.cancelled:
                                progress_bar.progress(1.0)
                                status_text.success("✅ All emails sent successfully!")
                                st.balloons()
                        
                        # Restore stdout
                        sys.stdout = old_stdout
                        
                        # Show success message
                        st.session_state.email_sent = True
                        st.session_state.progress = 100
                        progress_bar.progress(100)
                        status_text.success("✅ All emails sent successfully!")
                        st.balloons()
                        
                    except Exception as e:
                        # Restore stdout in case of error
                        sys.stdout = old_stdout
                        if "cancelled" in str(e).lower():
                            st.warning("⚠️ Email sending was cancelled. Some emails may have been sent.")
                        else:
                            st.error(f"❌ Error sending emails: {str(e)}")
                            st.exception(e)
        
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.exception(e)
    else:
        st.info("Please upload a dataset to get started.")

with tab2:
    if 'df' in st.session_state:
        st.subheader("Dataset Preview")
        st.dataframe(st.session_state.df.head())
        
        st.subheader("Column Information")
        st.json({
            "Total Rows": len(st.session_state.df),
            "Columns": list(st.session_state.df.columns),
            "Data Types": {col: str(dtype) for col, dtype in st.session_state.df.dtypes.items()}
        })
    else:
        st.info("Upload a dataset to see the preview.")

# Add some space at the bottom
st.markdown("---")
st.caption("© 2023 Cold Email Sender | Made with ❤️")
