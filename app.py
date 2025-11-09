import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import subprocess
import platform
import base64

# Page configuration
st.set_page_config(
    page_title="PO Line Comparison Tool",
    page_icon="📊",
    layout="wide"
)

st.title("📊 PO Line Comparison Tool")
st.markdown("Upload two Excel files to compare PO lines between previous week and current week")

# Sidebar for email configuration
with st.sidebar:
    st.header("📧 Email Configuration")
    
    # Email settings
    recipient_email = st.text_input("Recipient Email", value="galiaf@nvidia.com")
    
    # Email method selection
    st.subheader("Send Method")
    email_method = st.radio(
        "Choose how to send email:",
        options=["Outlook (Mac) - No password needed! 🎉", "Save as .eml file - No password needed! 📧", "SMTP - Requires password ⚙️"],
        help="Select your preferred email sending method"
    )
    
    # Show SMTP settings only if SMTP is selected
    if "SMTP" in email_method:
        with st.expander("⚙️ SMTP Settings", expanded=True):
            smtp_server = st.text_input("SMTP Server", value="smtp.office365.com", 
                                         help="For NVIDIA: smtp.office365.com")
            smtp_port = st.number_input("SMTP Port", value=587, min_value=1, max_value=65535)
            
            from_email = st.text_input("Your Email Address", 
                                       help="Your NVIDIA email address")
            email_password = st.text_input("Email Password / App Password", type="password",
                                          help="Your email password or app-specific password")
            
            st.info("💡 **For NVIDIA Office 365:**\n- Use your full email as username\n- May need an app password\n- SMTP: smtp.office365.com:587")
    else:
        # Default values for non-SMTP methods
        smtp_server = ""
        smtp_port = 587
        from_email = ""
        email_password = ""
        
        if "Outlook" in email_method:
            st.success("✅ **No password required!**\n\nWill create a draft in Outlook with the attachment. You can review and send.")
        elif ".eml" in email_method:
            st.success("✅ **No password required!**\n\nWill download a .eml file that you can open in any email client (Outlook, Mail, Gmail, etc.)")

# File uploaders
col1, col2 = st.columns(2)

with col1:
    st.subheader("Previous Week File")
    prev_week_file = st.file_uploader(
        "Upload previous week Excel file",
        type=['xlsx', 'xls'],
        key='prev_week'
    )

with col2:
    st.subheader("Current Week File")
    curr_week_file = st.file_uploader(
        "Upload current week Excel file",
        type=['xlsx', 'xls'],
        key='curr_week'
    )


def parse_excel_file(file):
    """Parse Excel file and return DataFrame with standardized column names"""
    try:
        df = pd.read_excel(file)
        
        # Map columns to standardized names
        column_mapping = {
            'Purch.doc.': 'PO_No',
            'Item': 'PO_Line',
            'Short text': 'PN',
            'Order': 'PWO',
            'Type': 'PO_Type',
            'ComDate': 'ComDate'
        }
        
        # Rename columns if they exist
        df.rename(columns=column_mapping, inplace=True)
        
        # Check for required columns
        required_cols = ['PO_No', 'PO_Line', 'PN', 'ComDate']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.warning(f"⚠️ Missing columns after mapping: {', '.join(missing_cols)}")
            st.info("Please ensure your Excel has: Purch.doc., Item, Short text, Order, Type, ComDate")
        
        # Ensure ComDate is datetime
        if 'ComDate' in df.columns:
            df['ComDate'] = pd.to_datetime(df['ComDate'], errors='coerce')
        
        # Create unique identifier for each PO line
        if 'PO_No' in df.columns and 'PO_Line' in df.columns:
            df['PO_LineID'] = df['PO_No'].astype(str) + '_' + df['PO_Line'].astype(str)
        
        return df
    except Exception as e:
        st.error(f"Error parsing file: {str(e)}")
        return None


def compare_po_lines(prev_df, curr_df):
    """Compare PO lines between two weeks and identify changes"""
    
    results = []
    
    # Get unique PO line IDs from both datasets
    prev_lines = set(prev_df['PO_LineID'].unique())
    curr_lines = set(curr_df['PO_LineID'].unique())
    
    # Lines that exist in both weeks
    common_lines = prev_lines.intersection(curr_lines)
    
    # Lines only in current week (new lines)
    new_lines = curr_lines - prev_lines
    
    # Check for pushed lines
    pushed_count = 0
    for line_id in common_lines:
        prev_line = prev_df[prev_df['PO_LineID'] == line_id].iloc[0]
        curr_line = curr_df[curr_df['PO_LineID'] == line_id].iloc[0]
        
        # Check if ComDate changed
        if pd.notna(prev_line['ComDate']) and pd.notna(curr_line['ComDate']):
            # Check if date changed at all (pushed OR pulled back)
            if curr_line['ComDate'] != prev_line['ComDate']:
                days_pushed = (curr_line['ComDate'] - prev_line['ComDate']).days
                
                result = {
                    'PO_No': curr_line['PO_No'],
                    'PO_Line': curr_line['PO_Line'],
                    'PN': curr_line['PN'] if 'PN' in curr_line else '',
                    'PWO': curr_line['PWO'] if 'PWO' in curr_line else '',
                    'PO_Type': curr_line['PO_Type'] if 'PO_Type' in curr_line else '',
                    'Prev_ComDate': prev_line['ComDate'],
                    'Curr_ComDate': curr_line['ComDate'],
                    'Days_Pushed': days_pushed,
                    'Status': 'Pushed' if days_pushed > 0 else 'Pulled Back',
                    'Alert': '🚨 ALERT' if days_pushed > 7 else ''
                }
                results.append(result)
                if days_pushed > 0:
                    pushed_count += 1
    
    # Check for split lines (same PO but multiple lines in current vs previous)
    split_count = 0
    if 'PN' in prev_df.columns and 'PN' in curr_df.columns:
        for po_no in prev_df['PO_No'].unique():
            prev_po_lines = prev_df[prev_df['PO_No'] == po_no]
            curr_po_lines = curr_df[curr_df['PO_No'] == po_no]
            
            if len(curr_po_lines) > len(prev_po_lines):
                # Check if it's a split (same PN but more lines)
                prev_pns = set(prev_po_lines['PN'].unique())
                curr_pns = set(curr_po_lines['PN'].unique())
                
                if prev_pns.intersection(curr_pns):  # Same part numbers exist
                    for _, curr_line in curr_po_lines.iterrows():
                        if curr_line['PO_LineID'] in new_lines:
                            result = {
                                'PO_No': curr_line['PO_No'],
                                'PO_Line': curr_line['PO_Line'],
                                'PN': curr_line['PN'] if 'PN' in curr_line else '',
                                'PWO': curr_line['PWO'] if 'PWO' in curr_line else '',
                                'PO_Type': curr_line['PO_Type'] if 'PO_Type' in curr_line else '',
                                'Prev_ComDate': None,
                                'Curr_ComDate': curr_line['ComDate'] if 'ComDate' in curr_line else None,
                                'Days_Pushed': 0,
                                'Status': 'Split',
                                'Alert': ''
                            }
                            results.append(result)
                            split_count += 1
    
    # Check for re-pushed lines (pushed multiple times)
    # Track if a line was pushed in previous comparisons (this would need historical data)
    # For now, we'll mark lines pushed >7 days as potentially re-pushed
    for result in results:
        if result['Status'] == 'Pushed' and result['Days_Pushed'] > 7:
            # This could indicate multiple pushes
            result['Status'] = 'Re-Pushed (>7 days)'
    
    return pd.DataFrame(results)


def style_dataframe(df):
    """Apply styling to the results dataframe"""
    def highlight_alerts(row):
        if row['Alert'] == '🚨 ALERT':
            return ['background-color: #ffcccc'] * len(row)
        elif row['Status'] == 'Split':
            return ['background-color: #ffffcc'] * len(row)
        else:
            return [''] * len(row)
    
    return df.style.apply(highlight_alerts, axis=1)


def send_email_with_attachment(to_email, subject, body, excel_data, filename, smtp_server, smtp_port, from_email, password):
    """Send email with Excel attachment using SMTP"""
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Add body as HTML
        html_body = body.replace('\n', '<br>')
        msg.attach(MIMEText(html_body, 'html'))
        
        # Add Excel attachment
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(excel_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={filename}')
        msg.attach(part)
        
        # Connect to SMTP server and send email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(from_email, password)
            server.send_message(msg)
        
        return True, "Email sent successfully!"
        
    except smtplib.SMTPAuthenticationError:
        return False, "Authentication failed. Please check your email and password."
    except smtplib.SMTPException as e:
        return False, f"SMTP error: {str(e)}"
    except Exception as e:
        return False, f"Error: {str(e)}"


def send_via_outlook_mac(to_email, subject, body, excel_data, filename):
    """Send email via Outlook on Mac using AppleScript"""
    try:
        import tempfile
        import os
        
        # Save Excel file temporarily
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, filename)
        with open(temp_file, 'wb') as f:
            f.write(excel_data)
        
        # Create AppleScript
        applescript = f'''
        tell application "Microsoft Outlook"
            set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{body}"}}
            make new recipient at newMessage with properties {{email address:{{address:"{to_email}"}}}}
            make new attachment at newMessage with properties {{file:POSIX file "{temp_file}"}}
            open newMessage
        end tell
        '''
        
        # Execute AppleScript
        subprocess.run(['osascript', '-e', applescript], check=True)
        
        return True, "Outlook draft created with attachment! Please review and send."
        
    except Exception as e:
        return False, f"Error: {str(e)}"


def save_as_eml_file(to_email, subject, body, excel_data, filename):
    """Save email with attachment as .eml file that can be opened in any email client"""
    try:
        # Create message
        msg = MIMEMultipart()
        msg['To'] = to_email
        msg['Subject'] = subject
        msg['From'] = 'PO Comparison Tool'
        
        # Add body
        msg.attach(MIMEText(body, 'plain'))
        
        # Add Excel attachment
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(excel_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={filename}')
        msg.attach(part)
        
        # Generate .eml content
        eml_content = msg.as_bytes()
        
        return True, eml_content
        
    except Exception as e:
        return False, str(e)


# Main comparison logic
if prev_week_file and curr_week_file:
    with st.spinner("Processing files..."):
        # Parse files
        prev_df = parse_excel_file(prev_week_file)
        curr_df = parse_excel_file(curr_week_file)
        
        if prev_df is not None and curr_df is not None:
            st.success("Files loaded successfully!")
            
            # Show file summaries
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Previous Week PO Lines", len(prev_df))
            with col2:
                st.metric("Current Week PO Lines", len(curr_df))
            
            # Compare PO lines
            results_df = compare_po_lines(prev_df, curr_df)
            
            if not results_df.empty:
                st.subheader("📋 Comparison Results")
                
                # Summary metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    pushed_count = len(results_df[results_df['Status'] == 'Pushed'])
                    st.metric("Pushed Lines", pushed_count)
                with col2:
                    split_count = len(results_df[results_df['Status'] == 'Split'])
                    st.metric("Split Lines", split_count)
                with col3:
                    repushed_count = len(results_df[results_df['Status'].str.contains('Re-Pushed', na=False)])
                    st.metric("Re-Pushed Lines", repushed_count)
                with col4:
                    alert_count = len(results_df[results_df['Alert'] != ''])
                    st.metric("🚨 Alerts (>7 days)", alert_count)
                
                # Filter options
                st.subheader("🔍 Filter Results")
                status_filter = st.multiselect(
                    "Filter by Status",
                    options=results_df['Status'].unique(),
                    default=results_df['Status'].unique()
                )
                
                show_alerts_only = st.checkbox("Show only alerts (>7 days)")
                
                # Apply filters
                filtered_df = results_df[results_df['Status'].isin(status_filter)]
                if show_alerts_only:
                    filtered_df = filtered_df[filtered_df['Alert'] != '']
                
                # Display results
                st.dataframe(
                    style_dataframe(filtered_df),
                    use_container_width=True,
                    height=500
                )
                
                # Download and Email results
                st.subheader("💾 Export Results")
                
                # Convert to Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Comparison Results')
                
                excel_data = output.getvalue()
                filename = f"po_comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="📥 Download Results as Excel",
                        data=excel_data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col2:
                    # Prepare email content
                    results_summary = {
                        'total': len(filtered_df),
                        'pushed': len(filtered_df[filtered_df['Status'].str.contains('Pushed', na=False)]),
                        'split': len(filtered_df[filtered_df['Status'] == 'Split']),
                        'alerts': len(filtered_df[filtered_df['Alert'] != ''])
                    }
                    
                    email_subject = f"PO Comparison Results - {datetime.now().strftime('%Y-%m-%d')}"
                    email_body = f"""Hi,

Please find attached the PO Line Comparison results:

SUMMARY:
- Total changes: {results_summary['total']}
- Pushed lines: {results_summary['pushed']}
- Split lines: {results_summary['split']}
- Alerts (>7 days): {results_summary['alerts']}

Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Best regards,
PO Line Comparison Tool
"""
                    
                    # Email button based on selected method
                    if "Outlook" in email_method:
                        # Outlook Mac method
                        if st.button("📧 Create Outlook Draft", key="send_email", use_container_width=True):
                            with st.spinner("Creating Outlook draft..."):
                                success, message = send_via_outlook_mac(
                                    to_email=recipient_email,
                                    subject=email_subject,
                                    body=email_body,
                                    excel_data=excel_data,
                                    filename=filename
                                )
                                
                                if success:
                                    st.success(f"✅ {message}")
                                else:
                                    st.error(f"❌ {message}")
                                    st.info("💡 Make sure Microsoft Outlook is installed on your Mac")
                    
                    elif ".eml" in email_method:
                        # Save as .eml file method
                        success, result = save_as_eml_file(
                            to_email=recipient_email,
                            subject=email_subject,
                            body=email_body,
                            excel_data=excel_data,
                            filename=filename
                        )
                        
                        if success:
                            eml_filename = f"po_email_{datetime.now().strftime('%Y%m%d_%H%M%S')}.eml"
                            st.download_button(
                                label="📧 Download .eml File (with attachment)",
                                data=result,
                                file_name=eml_filename,
                                mime="message/rfc822",
                                use_container_width=True
                            )
                            st.info("💡 Download the .eml file and double-click to open in Outlook, Mail, or any email client. The Excel file will be already attached!")
                        else:
                            st.error(f"❌ Error creating .eml file: {result}")
                    
                    else:
                        # SMTP method
                        if st.button("📧 Send Email via SMTP", key="send_email", use_container_width=True):
                            # Validate email settings
                            if not from_email or not email_password:
                                st.error("⚠️ Please enter your email and password in the SMTP Settings!")
                            else:
                                with st.spinner("Sending email..."):
                                    success, message = send_email_with_attachment(
                                        to_email=recipient_email,
                                        subject=email_subject,
                                        body=email_body,
                                        excel_data=excel_data,
                                        filename=filename,
                                        smtp_server=smtp_server,
                                        smtp_port=smtp_port,
                                        from_email=from_email,
                                        password=email_password
                                    )
                                    
                                    if success:
                                        st.success(f"✅ {message}")
                                        st.balloons()
                                    else:
                                        st.error(f"❌ {message}")
                    
                st.info("💡 Choose your preferred email method in the sidebar. The Excel file will be automatically attached!")
                
            else:
                st.info("No changes detected between the two weeks.")
else:
    st.info("👆 Please upload both Excel files to begin comparison")

# Footer
st.markdown("---")
st.markdown("### 📖 Instructions")
st.markdown("""
1. **Choose Email Method** (Sidebar): Select your preferred method:
   - **Outlook (Mac)** - Creates a draft in Outlook automatically (No password needed!)
   - **.eml File** - Download a file you can open in any email client (No password needed!)
   - **SMTP** - Direct email sending (Requires password)

2. **Upload Files**: Upload the previous week and current week Excel files

3. **Review Results**: The app will automatically identify:
   - **Pushed Lines**: PO lines with later commit dates
   - **Split Lines**: PO lines that were divided into multiple lines
   - **Re-Pushed Lines**: Lines pushed multiple times (>7 days indicates potential re-push)

4. **Alerts**: Lines pushed more than 7 days are highlighted with 🚨 ALERT

5. **Export Options**:
   - **Download Excel**: Save results as Excel file
   - **Email**: Send to galiaf@nvidia.com with Excel attachment (method depends on your selection in sidebar)
""")

