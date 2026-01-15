import os
import re
import mimetypes
import zipfile
from datetime import datetime
from email.message import EmailMessage
import pandas as pd
import json
try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False
    print("Warning: win32com not available. Install with: pip install pywin32")


def extract_variables(template_text):
    """Return a list of variables like ['FirstName', 'Company'] found in the template."""
    return re.findall(r"\[([^\]]+)\]", template_text)


def apply_text_formatting(text):
    """Apply text formatting like bold, italic, etc."""
    if not text:
        return text
    
    # Convert markdown-style formatting to HTML
    # **text** -> <b>text</b> (bold)
    text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', text)
    # *text* -> <i>text</i> (italic)
    text = re.sub(r'\*(.*?)\*', r'<i>\1</i>', text)
    # __text__ -> <u>text</u> (underline)
    text = re.sub(r'__(.*?)__', r'<u>\1</u>', text)
    
    return text


def convert_to_html_email(text):
    """Convert plain text to HTML format for email."""
    if not text:
        return text
    
    # Apply text formatting first
    text = apply_text_formatting(text)
    
    # Convert line breaks to HTML
    text = text.replace('\n', '<br>')
    
    # Wrap in basic HTML structure
    html_body = f"""<html>
<head>
<meta charset="UTF-8">
</head>
<body style="font-family: Arial, sans-serif; font-size: 11pt;">
{text}
</body>
</html>"""
    
    return html_body


def clean_email_encoding(text):
    """Clean up email encoding artifacts like quoted-printable characters."""
    if not text:
        return text
    
    # Remove quoted-printable encoding artifacts
    text = text.replace('=\n', '')  # Remove soft line breaks
    text = text.replace('=\r\n', '')  # Remove soft line breaks (Windows)
    text = text.replace('=\r', '')  # Remove soft line breaks (Mac)
    
    # Common quoted-printable encoded characters
    replacements = {
        '=20': ' ',    # space
        '=3D': '=',    # equals sign
        '=0A': '\n',   # newline
        '=0D': '\r',   # carriage return
        '=22': '"',    # double quote
        '=27': "'",    # single quote
        '=2C': ',',    # comma
        '=3B': ';',    # semicolon
        '=3A': ':',    # colon
        '=2E': '.',    # period
        '=2D': '-',    # hyphen
        '=5F': '_',    # underscore
        '=40': '@',    # at symbol
        '=24': '$',    # dollar sign
        '=25': '%',    # percent sign
        '=26': '&',    # ampersand
        '=2B': '+',    # plus sign
        '=3C': '<',    # less than
        '=3E': '>',    # greater than
        '=3F': '?',    # question mark
        '=21': '!',    # exclamation mark
        '=28': '(',    # left parenthesis
        '=29': ')',    # right parenthesis
        '=5B': '[',    # left bracket
        '=5D': ']',    # right bracket
        '=7B': '{',    # left brace
        '=7D': '}',    # right brace
        '=7C': '|',    # pipe
        '=5C': '\\',   # backslash
        '=2F': '/',    # forward slash
        '=7E': '~',    # tilde
    }
    
    for encoded, decoded in replacements.items():
        text = text.replace(encoded, decoded)
    
    # Remove any remaining stray = characters that might be encoding artifacts
    # But be careful not to remove legitimate = signs in content
    # Only remove = that appear at line breaks or followed by unusual characters
    text = re.sub(r'=(?=\s|$)', '', text)  # Remove = at end of lines or before whitespace
    
    return text


def fill_template(template_text, row, variables):
    """Fill template placeholders with row values, leaving placeholders if missing."""
    result = template_text
    for var in variables:
        if var in row and pd.notna(row[var]):
            value = str(row[var])
            # Clean encoding artifacts from the value
            value = clean_email_encoding(value)
            result = result.replace(f"[{var}]", value)
    
    # Clean the final result as well
    result = clean_email_encoding(result)
    return result


def parse_email_addresses(email_str):
    """
    Parse comma-separated email addresses and return a list.
    Ignores any addresses enclosed in square brackets, e.g., [test@example.com],
    which can be used as a manual circuit breaker to temporarily disable a recipient.
    """
    if not email_str or pd.isna(email_str):
        return []
    
    # Split by comma, strip whitespace, and filter out bracketed or empty emails
    emails = [
        email.strip() 
        for email in str(email_str).split(',') 
        if email.strip() and not (email.strip().startswith('[') and email.strip().endswith(']'))
    ]
    return emails


def create_email_message(row, template_text, variables, attachments_dir, attachment_columns, conditional_lines, is_html_template=False, attachment_mode="global", per_recipient_base=None, identifier_column=None):
    """Create EmailMessage object for a DataFrame row as a clean email (no draft prefixes).
    
    Args:
        row: DataFrame row with recipient data
        template_text: Email template text
        variables: List of template variables
        attachments_dir: Global attachments directory (used in global mode)
        attachment_columns: Columns containing individual attachment filenames
        conditional_lines: Dictionary of conditional content
        is_html_template: Whether the template is HTML
        attachment_mode: "global" or "per_recipient"
        per_recipient_base: Base folder for per-recipient attachments
        identifier_column: Column to use for per-recipient folder names
    """
    msg = EmailMessage()

    # Handle To addresses (clean, no [DRAFT] prefix)
    to_addresses = parse_email_addresses(row.get('Email') or row.get('email') or row.get('To'))
    if to_addresses:
        msg['To'] = ', '.join(to_addresses)
    
    # Handle CC addresses (clean, no [DRAFT] prefix)
    cc_addresses = parse_email_addresses(row.get('CC') or row.get('cc'))
    # Add default CC address - removing the hardcoded bracket for Streamlit version
    # default_cc = '['  # Always CC this address
    # if default_cc:
    #     cc_addresses.append(default_cc)
    if cc_addresses:
        msg['CC'] = ', '.join(cc_addresses)
    
    # Handle BCC addresses (clean, no [DRAFT] prefix)
    bcc_addresses = parse_email_addresses(row.get('BCC') or row.get('bcc'))
    if bcc_addresses:
        msg['BCC'] = ', '.join(bcc_addresses)
    
    msg['From'] = row.get('From') or 'sender@example.com'
    msg['Subject'] = row.get('Subject') or 'No Subject'  # Clean subject, no [DRAFT] prefix
    
    # Add headers to mark as editable draft
    # X-Unsent: 1 tells email clients (Outlook, Thunderbird) to open as draft
    msg['X-Unsent'] = '1'
    msg['X-Draft-Info'] = 'Generated email - Review before sending'

    # First, fill in the main template placeholders
    body = fill_template(template_text, row, variables)

    # Second, handle conditional placeholders
    conditional_placeholders = re.findall(r"\[Conditional:([^\]]+)\]", body)
    for placeholder_key in conditional_placeholders:
        placeholder_tag = f"[Conditional:{placeholder_key}]"
        replacement_text = "" # Default to empty string
        # Check if the flag is set to 1 in the Excel row
        if placeholder_key in conditional_lines and placeholder_key in row and row[placeholder_key] == 1:
            # If so, get the text from the JSON and fill its own placeholders
            raw_text = conditional_lines[placeholder_key]
            replacement_text = fill_template(raw_text, row, variables)
        
        body = body.replace(placeholder_tag, replacement_text)

    # Set clean email body without draft notices
    # Support both plain text and HTML
    if is_html_template:
        # Template is already HTML, use it directly
        # Create a plain text version from HTML
        plain_text = re.sub(r'<[^>]+>', '', body)  # Simple HTML tag removal
        plain_text = plain_text.replace('&nbsp;', ' ').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        plain_text = plain_text.replace('&quot;', '"').replace('&#39;', "'")  # Additional entity replacements
        # Set plain text first, then add HTML alternative
        msg.set_content(plain_text)
        msg.add_alternative(body, subtype='html')
    else:
        # Convert plain text to HTML
        html_body = convert_to_html_email(body)
        msg.set_content(body)  # Plain text version
        msg.add_alternative(html_body, subtype='html')  # HTML version

    # Determine which attachments directory to use
    actual_attachments_dir = attachments_dir
    attachment_source = "global"
    
    if attachment_mode == "per_recipient" and per_recipient_base and identifier_column:
        # Get the identifier value for this recipient
        identifier_value = str(row.get(identifier_column, "")).strip()
        if identifier_value:
            # Clean identifier for folder name (remove special characters)
            clean_identifier = re.sub(r'[^\w\s-]', '', identifier_value).strip()
            if clean_identifier:
                per_recipient_dir = os.path.join(per_recipient_base, clean_identifier)
                if os.path.exists(per_recipient_dir):
                    actual_attachments_dir = per_recipient_dir
                    attachment_source = f"per-recipient ({clean_identifier})"
                    print(f"  📁 Using per-recipient folder: {clean_identifier}")
                else:
                    print(f"  ⚠️ Per-recipient folder not found: {clean_identifier}, falling back to global attachments")
    
    # Attach files from the determined directory
    total_size_mb = 0
    attached_files = []
    
    if actual_attachments_dir and os.path.exists(actual_attachments_dir):
        for filename in os.listdir(actual_attachments_dir):
            file_path = os.path.join(actual_attachments_dir, filename)
            # Only attach files (not directories)
            if os.path.isfile(file_path):
                try:
                    file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
                    
                    # Warn about large files
                    if file_size_mb > 10:
                        print(f"  ⚠️ Large attachment: {filename} ({file_size_mb:.2f} MB)")
                    
                    with open(file_path, 'rb') as fh:
                        data = fh.read()
                    ctype, _ = mimetypes.guess_type(file_path)
                    if ctype:
                        maintype, subtype = ctype.split('/', 1)
                    else:
                        maintype, subtype = 'application', 'octet-stream'
                    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
                    attached_files.append(filename)
                    total_size_mb += file_size_mb
                    print(f"  ✅ Attached: {filename} ({file_size_mb:.2f} MB) from {attachment_source}")
                except Exception as e:
                    print(f"  ❌ Failed to attach {filename}: {e}")
    
    if total_size_mb > 25:
        print(f"  ⚠️ Warning: Total attachment size is {total_size_mb:.2f} MB - may exceed email limits")
    
    # Legacy: Also handle attachment columns if specified (for backward compatibility)
    for col in attachment_columns:
        filename = row.get(col)
        if filename and pd.notna(filename):
            # Extract just the filename if it's a full path
            filename_str = str(filename).strip().strip('"\'')  # Remove quotes and whitespace
            base_filename = os.path.basename(filename_str)  # Extract just the filename part
            file_path = os.path.join(attachments_dir, base_filename)
            try:
                with open(file_path, 'rb') as fh:
                    data = fh.read()
                ctype, _ = mimetypes.guess_type(file_path)
                if ctype:
                    maintype, subtype = ctype.split('/', 1)
                else:
                    maintype, subtype = 'application', 'octet-stream'
                msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=base_filename)
                attached_files.append(base_filename)
                print(f"  ✅ Attached from Excel: {base_filename}")
            except FileNotFoundError:
                print(f"  ❌ Attachment not found: {file_path} (original: {base_filename})")
            except Exception as e:
                print(f"  ❌ Failed to attach from Excel {base_filename}: {e}")
    
    # Return msg with attachment info for tracking
    return msg


def create_outlook_draft(row, template_text, variables, attachments_dir, attachment_columns, conditional_lines, is_html_template=False, attachment_mode="global", per_recipient_base=None, identifier_column=None, output_dir="generated_emails"):
    """Create an actual Outlook .msg file from the row data.
    
    Args:
        row: DataFrame row with recipient data
        template_text: Email template text
        variables: List of template variables
        attachments_dir: Global attachments directory (used in global mode)
        attachment_columns: Columns containing individual attachment filenames
        conditional_lines: Dictionary of conditional content
        is_html_template: Whether the template is HTML
        attachment_mode: "global" or "per_recipient"
        per_recipient_base: Base folder for per-recipient attachments
        identifier_column: Column to use for per-recipient folder names
        output_dir: Directory to save .msg files
    """
    if not OUTLOOK_AVAILABLE:
        raise ImportError("win32com.client not available. Install with: pip install pywin32")
    
    try:
        print("  📧 Connecting to Outlook...", end="", flush=True)
        outlook = win32com.client.Dispatch("Outlook.Application")
        print(" ✅")
        
        print("  📝 Creating email item...", end="", flush=True)
        mail = outlook.CreateItem(0)  # 0 = Mail item
        print(" ✅")
        
        # Handle To addresses (original addresses without [DRAFT] prefix for Outlook)
        to_addresses = parse_email_addresses(row.get('Email') or row.get('email') or row.get('To'))
        if to_addresses:
            mail.To = '; '.join(to_addresses)
        
        # Handle CC addresses
        cc_addresses = parse_email_addresses(row.get('CC') or row.get('cc'))
        # Remove default CC bracket for Streamlit version
        # default_cc = '['  # Always CC this address
        # if default_cc:
        #     cc_addresses.append(default_cc)
        if cc_addresses:
            mail.CC = '; '.join(cc_addresses)
        
        # Handle BCC addresses
        bcc_addresses = parse_email_addresses(row.get('BCC') or row.get('bcc'))
        if bcc_addresses:
            mail.BCC = '; '.join(bcc_addresses)
        
        # Set sender if specified
        sender = row.get('From')
        if sender:
            try:
                mail.SentOnBehalfOfName = sender
            except Exception:
                print(f"    ⚠️ Could not set sender to {sender}")
        
        # Set subject (clean, no [DRAFT] prefix since this is already a draft in Outlook)
        subject = row.get('Subject') or 'No Subject'
        mail.Subject = subject
        
        # Set email format to HTML (olFormatHTML = 2)
        mail.BodyFormat = 2
        
        print("  📄 Processing email body...", end="", flush=True)
        # Process email body
        body = fill_template(template_text, row, variables)
        
        # Handle conditional placeholders
        conditional_placeholders = re.findall(r"\[Conditional:([^\]]+)\]", body)
        for placeholder_key in conditional_placeholders:
            placeholder_tag = f"[Conditional:{placeholder_key}]"
            replacement_text = ""
            if placeholder_key in conditional_lines and placeholder_key in row and row[placeholder_key] == 1:
                raw_text = conditional_lines[placeholder_key]
                replacement_text = fill_template(raw_text, row, variables)
            body = body.replace(placeholder_tag, replacement_text)
        
        # Set HTML body (setting HTMLBody automatically preserves HTML format)
        if is_html_template:
            # Template is already HTML
            mail.HTMLBody = body
        else:
            # Convert plain text to HTML
            mail.HTMLBody = convert_to_html_email(body)
        print(" ✅")
        
        # Add attachments
        print("  📎 Adding attachments...", end="", flush=True)
        attachment_count = 0
        attachment_debug = []
        
        # Determine which attachments directory to use
        actual_attachments_dir = attachments_dir
        attachment_source = "global"
        
        if attachment_mode == "per_recipient" and per_recipient_base and identifier_column:
            # Get the identifier value for this recipient
            identifier_value = str(row.get(identifier_column, "")).strip()
            if identifier_value:
                # Clean identifier for folder name (remove special characters)
                clean_identifier = re.sub(r'[^\w\s-]', '', identifier_value).strip()
                if clean_identifier:
                    per_recipient_dir = os.path.join(per_recipient_base, clean_identifier)
                    if os.path.exists(per_recipient_dir):
                        actual_attachments_dir = per_recipient_dir
                        attachment_source = f"per-recipient ({clean_identifier})"
                        attachment_debug.append(f"Using per-recipient folder: {clean_identifier}")
                    else:
                        attachment_debug.append(f"Per-recipient folder not found: {clean_identifier}, using global")
        
        # Attach files from the determined directory
        total_size_mb = 0
        if actual_attachments_dir and os.path.exists(actual_attachments_dir):
            for filename in os.listdir(actual_attachments_dir):
                file_path = os.path.join(actual_attachments_dir, filename)
                # Only attach files (not directories)
                if os.path.isfile(file_path):
                    file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
                    attachment_debug.append(f"Trying to attach: {file_path} ({file_size_mb:.2f} MB) from {attachment_source}")
                    
                    # Warn about large files
                    if file_size_mb > 10:
                        attachment_debug.append(f"⚠️ Large attachment: {filename} ({file_size_mb:.2f} MB)")
                    
                    try:
                        mail.Attachments.Add(file_path)
                        attachment_count += 1
                        total_size_mb += file_size_mb
                        attachment_debug.append(f"✅ Successfully attached: {filename}")
                    except Exception as e:
                        attachment_debug.append(f"❌ Failed to attach {filename}: {e}")
        
        if total_size_mb > 25:
            attachment_debug.append(f"⚠️ Warning: Total attachment size is {total_size_mb:.2f} MB")
        
        # Legacy: Also handle attachment columns if specified (for backward compatibility)
        for col in attachment_columns:
            filename = row.get(col)
            if filename and pd.notna(filename):
                filename_str = str(filename).strip().strip('"\'')
                base_filename = os.path.basename(filename_str)
                file_path = os.path.join(attachments_dir, base_filename) if attachments_dir else base_filename
                attachment_debug.append(f"Trying to attach from Excel: {file_path}")
                try:
                    mail.Attachments.Add(file_path)
                    attachment_count += 1
                    attachment_debug.append(f"✅ Successfully attached from Excel: {base_filename}")
                except Exception as e:
                    attachment_debug.append(f"❌ Failed to attach from Excel {base_filename}: {e}")
        
        print(f" ✅ ({attachment_count} files)")
        
        # Generate filename for .msg file
        recipient_name = row.get('FirstName') or row.get('Email') or row.get('email') or 'recipient'
        subject = row.get('Subject') or 'No Subject'
        
        # Sanitize the filename components
        clean_recipient = sanitize_filename(str(recipient_name))
        clean_subject = sanitize_filename(str(subject))
        
        # Create the filename
        filename = f"{clean_recipient}_{clean_subject}.msg"
        
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        # Full path for the .msg file
        msg_path = os.path.join(output_dir, filename)
        
        # Save as .msg file (olMSG format = 3)
        print("  💾 Saving .msg file...", end="", flush=True)
        olMSG = 3
        mail.SaveAs(msg_path, olMSG)
        print(" ✅")
        
        print(f"  🎯 .msg file saved: {filename}")
        
        # Return debug info for troubleshooting
        return {
            'success': True,
            'attachment_count': attachment_count,
            'debug_info': attachment_debug,
            'msg_file': msg_path
        }
        
    except Exception as e:
        print(f" ❌ Error: {e}")
        return {
            'success': False,
            'error': str(e),
            'debug_info': attachment_debug if 'attachment_debug' in locals() else []
        }


def sanitize_filename(filename):
    """Convert a string to a safe filename by removing/replacing problematic characters."""
    # Replace problematic characters with underscores
    sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
    # Remove any remaining control characters
    sanitized = ''.join(char for char in sanitized if ord(char) >= 32)
    # Limit length and strip whitespace
    sanitized = sanitized.strip()[:100]
    return sanitized if sanitized else "email"


def create_zip_bundle(output_dir, zip_filename=None):
    """
    Create a ZIP file containing all generated email files (.eml and .msg).

    Args:
        output_dir: Directory containing the generated email files
        zip_filename: Optional custom filename for the ZIP (without extension)

    Returns:
        str: Path to the created ZIP file, or None if no files found
    """
    if not os.path.exists(output_dir):
        print(f"  Output directory not found: {output_dir}")
        return None

    # Find all email files
    email_files = []
    for filename in os.listdir(output_dir):
        if filename.endswith(('.eml', '.msg')):
            email_files.append(filename)

    if not email_files:
        print("  No email files found to bundle")
        return None

    # Generate ZIP filename with timestamp
    if zip_filename is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"email_drafts_{timestamp}"

    zip_path = os.path.join(output_dir, f"{zip_filename}.zip")

    print(f"  Creating ZIP bundle: {zip_filename}.zip")

    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for filename in email_files:
                file_path = os.path.join(output_dir, filename)
                # Add file to ZIP with just the filename (no directory path)
                zf.write(file_path, filename)
                print(f"    Added: {filename}")

        # Get ZIP file size
        zip_size_mb = os.path.getsize(zip_path) / (1024 * 1024)
        print(f"  ZIP bundle created: {zip_path} ({zip_size_mb:.2f} MB)")
        print(f"  Contains {len(email_files)} email file(s)")

        return zip_path

    except Exception as e:
        print(f"  Error creating ZIP bundle: {e}")
        return None


def main(template_path, excel_path, attachments_dir=None, output_dir="generated_emails",
         conditionals_path="conditional_lines.json", use_outlook=True, create_eml_backup=True, is_html_template=False,
         attachment_mode="global", per_recipient_base=None, identifier_column=None, create_zip=True):
    """
    Main function to generate emails from template and Excel data.

    Args:
        template_path: Path to email template file
        excel_path: Path to Excel file with recipient data
        attachments_dir: Path to directory containing attachments (optional)
        output_dir: Directory to save .eml files (optional)
        conditionals_path: Path to conditional lines JSON file (optional)
        use_outlook: Whether to create Outlook drafts (requires pywin32)
        create_eml_backup: Whether to create .eml backup files
        is_html_template: Whether the template is HTML
        attachment_mode: "global" or "per_recipient"
        per_recipient_base: Base folder for per-recipient attachments
        identifier_column: Column to use for per-recipient folder names
        create_zip: Whether to create a ZIP bundle of all generated emails

    Returns:
        dict: Result dictionary with success status, counts, zip_path, and any warnings/errors
    """
    
    # Load template
    print(f"📖 Loading template from: {template_path}")
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            template_text = f.read()
        print("✅ Template loaded")
        # Log if this is an HTML template
        if is_html_template:
            print("📝 Template format: HTML")
        else:
            print("📝 Template format: Plain text")
    except Exception as e:
        print(f"❌ Error loading template: {e}")
        return False
    
    # Extract variables from template
    variables = extract_variables(template_text)
    print(f"🔍 Found variables: {variables}")
    
    # Load Excel data
    print(f"📊 Loading Excel data from: {excel_path}")
    try:
        df = pd.read_excel(excel_path)
        print(f"✅ Loaded {len(df)} rows")
    except Exception as e:
        print(f"❌ Error loading Excel: {e}")
        return False
    
    # Load conditional lines if they exist
    conditional_lines = {}
    if conditionals_path and os.path.exists(conditionals_path):
        try:
            with open(conditionals_path, 'r', encoding='utf-8') as f:
                conditional_lines = json.load(f)
            print(f"🔀 Loaded {len(conditional_lines)} conditional lines")
        except Exception as e:
            print(f"⚠️ Warning: Could not load conditional lines: {e}")
    
    # Create output directory
    if create_eml_backup:
        os.makedirs(output_dir, exist_ok=True)
        print(f"📁 Output directory: {output_dir}")
    
    # Find attachment columns
    attachment_columns = [col for col in df.columns if col.lower().startswith('attachment')]
    if attachment_columns:
        print(f"📎 Found attachment columns: {attachment_columns}")
    
    # Check Outlook availability
    if use_outlook:
        if OUTLOOK_AVAILABLE:
            print("🔗 Outlook integration available")
        else:
            print("⚠️ Outlook integration not available - falling back to .eml files only")
            use_outlook = False
            create_eml_backup = True
    
    # Log attachment mode
    if attachment_mode == "per_recipient":
        print(f"📂 Attachment mode: Per-recipient")
        print(f"   Base folder: {per_recipient_base}")
        print(f"   Identifier column: {identifier_column}")
    else:
        print(f"📂 Attachment mode: Global")
        if attachments_dir:
            print(f"   Attachments folder: {attachments_dir}")
    
    # Process each row
    success_count = 0
    error_count = 0
    
    print(f"\n🚀 Processing {len(df)} emails...")
    
    for idx, row in df.iterrows():
        print(f"\n📧 Processing row {idx + 1}/{len(df)}")
        
        # Get recipient name for filename
        recipient_name = row.get('FirstName') or row.get('Email') or row.get('email') or f"row_{idx}"
        print(f"  👤 Recipient: {recipient_name}")
        
        try:
            # Create Outlook .msg file if requested
            outlook_success = True
            if use_outlook:
                result = create_outlook_draft(
                    row, template_text, variables, attachments_dir, attachment_columns, conditional_lines, 
                    is_html_template, attachment_mode, per_recipient_base, identifier_column, output_dir
                )
                outlook_success = result['success']
                if not outlook_success:
                    print(f"  ⚠️ Outlook .msg file creation failed: {result.get('error', 'Unknown error')}")
            
            # Create .eml backup if requested
            eml_success = True
            if create_eml_backup:
                # Create email message
                msg = create_email_message(
                    row, template_text, variables, attachments_dir, attachment_columns, conditional_lines, 
                    is_html_template, attachment_mode, per_recipient_base, identifier_column
                )
                
                # Generate filename
                subject = row.get('Subject') or 'No Subject'
                safe_subject = sanitize_filename(subject)
                filename = f"{sanitize_filename(recipient_name)}_{safe_subject}.eml"
                filepath = os.path.join(output_dir, filename)
                
                # Save .eml file
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(str(msg))
                print(f"  💾 Saved: {filename}")
            
            if outlook_success or eml_success:
                success_count += 1
            else:
                error_count += 1
                
        except Exception as e:
            print(f"  ❌ Error processing row {idx + 1}: {e}")
            error_count += 1
    
    # Summary
    print(f"\n📊 Summary:")
    print(f"✅ Successful: {success_count}")
    print(f"❌ Failed: {error_count}")
    print(f"📧 Total processed: {len(df)}")

    if use_outlook and success_count > 0:
        print(f"📧 Created {success_count} editable .msg files in: {output_dir}")

    if create_eml_backup and success_count > 0:
        print(f"📁 Draft .eml files saved to: {output_dir}")
        print(f"   (Files include X-Unsent header for draft mode in Classic Outlook)")

    # Create ZIP bundle if requested
    zip_path = None
    if create_zip and success_count > 0:
        print(f"\n📦 Creating ZIP bundle...")
        zip_path = create_zip_bundle(output_dir)
        if zip_path:
            print(f"✅ ZIP bundle ready for download")

    # Return detailed result dictionary
    return {
        'success': success_count > 0,
        'success_count': success_count,
        'error_count': error_count,
        'total_count': len(df),
        'msg_files_created': use_outlook and success_count > 0,
        'eml_files_created': create_eml_backup and success_count > 0,
        'output_dir': output_dir,
        'zip_path': zip_path
    }


if __name__ == "__main__":
    # Default configuration - modify as needed
    result = main(
        template_path="email_template.txt",
        excel_path="recipients.xlsx",
        attachments_dir="attachments",
        output_dir="generated_emails",
        conditionals_path="conditional_lines.json",
        use_outlook=True,
        create_eml_backup=True
    )
    
    if result:
        print("\n🎉 Email generation completed successfully!")
    else:
        print("\n💥 Email generation failed!")
