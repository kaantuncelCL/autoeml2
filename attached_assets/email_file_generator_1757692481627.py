import os
import re
import mimetypes
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


def create_email_message(row, template_text, variables, attachments_dir, attachment_columns, conditional_lines):
    """Create EmailMessage object for a DataFrame row as a clean email (no draft prefixes)."""
    msg = EmailMessage()

    # Handle To addresses (clean, no [DRAFT] prefix)
    to_addresses = parse_email_addresses(row.get('Email') or row.get('email') or row.get('To'))
    if to_addresses:
        msg['To'] = ', '.join(to_addresses)
    
    # Handle CC addresses (clean, no [DRAFT] prefix)
    cc_addresses = parse_email_addresses(row.get('CC') or row.get('cc'))
    # Add default CC address
    default_cc = '['  # Always CC this address
    if default_cc:
        cc_addresses.append(default_cc)
    if cc_addresses:
        msg['CC'] = ', '.join(cc_addresses)
    
    # Handle BCC addresses (clean, no [DRAFT] prefix)
    bcc_addresses = parse_email_addresses(row.get('BCC') or row.get('bcc'))
    if bcc_addresses:
        msg['BCC'] = ', '.join(bcc_addresses)
    
    msg['From'] = row.get('From') or 'sender@example.com'
    msg['Subject'] = row.get('Subject') or 'No Subject'  # Clean subject, no [DRAFT] prefix
    
    # Add some email headers for better compatibility
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
    html_body = convert_to_html_email(body)
    msg.set_content(body)  # Plain text version
    msg.add_alternative(html_body, subtype='html')  # HTML version

    # Attach ALL files from attachments directory (ignoring Excel columns)
    if attachments_dir and os.path.exists(attachments_dir):
        for filename in os.listdir(attachments_dir):
            file_path = os.path.join(attachments_dir, filename)
            # Only attach files (not directories)
            if os.path.isfile(file_path):
                try:
                    with open(file_path, 'rb') as fh:
                        data = fh.read()
                    ctype, _ = mimetypes.guess_type(file_path)
                    if ctype:
                        maintype, subtype = ctype.split('/', 1)
                    else:
                        maintype, subtype = 'application', 'octet-stream'
                    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
                    print(f"  ✅ Attached: {filename}")
                except Exception as e:
                    print(f"  ❌ Failed to attach {filename}: {e}")
    
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
                print(f"  ✅ Attached from Excel: {base_filename}")
            except FileNotFoundError:
                print(f"  ❌ Attachment not found: {file_path} (original: {filename})")
    return msg


def create_outlook_draft(row, template_text, variables, attachments_dir, attachment_columns, conditional_lines):
    """Create an actual Outlook draft from the row data."""
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
        # Add default CC address
        default_cc = '['  # Always CC this address
        if default_cc:
            cc_addresses.append(default_cc)
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
        
        # Set clean email body without draft notices
        # Convert to HTML for better formatting support
        html_body = convert_to_html_email(body)
        mail.HTMLBody = html_body
        print(" ✅")
        
        # Add attachments
        print("  📎 Adding attachments...", end="", flush=True)
        attachment_count = 0
        attachment_debug = []
        
        # Attach ALL files from attachments directory (ignoring Excel columns)
        if attachments_dir and os.path.exists(attachments_dir):
            for filename in os.listdir(attachments_dir):
                file_path = os.path.join(attachments_dir, filename)
                # Only attach files (not directories)
                if os.path.isfile(file_path):
                    attachment_debug.append(f"Trying to attach: {file_path}")
                    try:
                        mail.Attachments.Add(file_path)
                        attachment_count += 1
                        attachment_debug.append(f"✅ Successfully attached: {filename}")
                    except Exception as e:
                        attachment_debug.append(f"❌ Error attaching {file_path}: {e}")
        
        # Legacy: Also handle attachment columns if specified (for backward compatibility)
        for col in attachment_columns:
            filename = row.get(col)
            if filename and pd.notna(filename):
                filename_str = str(filename).strip().strip('"\'')
                base_filename = os.path.basename(filename_str)
                file_path = os.path.join(attachments_dir, base_filename)
                attachment_debug.append(f"Trying to attach from Excel: {file_path}")
                
                try:
                    if os.path.exists(file_path):
                        mail.Attachments.Add(file_path)
                        attachment_count += 1
                        attachment_debug.append(f"✅ Successfully attached from Excel: {base_filename}")
                    else:
                        attachment_debug.append(f"❌ File not found: {file_path}")
                except Exception as e:
                    attachment_debug.append(f"❌ Error attaching {file_path}: {e}")
        
        if attachment_debug:
            print(f"\n    " + "\n    ".join(attachment_debug))
        
        print(f" ✅ ({attachment_count} attached)")
        
        # Save as draft
        print("  💾 Saving draft...", end="", flush=True)
        mail.Save()
        print(" ✅")
        
        return mail, subject
        
    except Exception as e:
        print(f" ❌")
        raise Exception(f"Failed to create Outlook draft: {e}")


def sanitize_filename(value):
    return re.sub(r"[^A-Za-z0-9_.-]", "_", str(value))


def main(template_path='email_template.txt', excel_path='recipients.xlsx', attachments_dir='attachments', output_dir='generated_emails', conditionals_path='conditional_lines.json', use_outlook=True, create_eml_backup=False):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with open(template_path, 'r', encoding='utf-8') as f:
        template_text = f.read()

    # Load conditional lines from JSON
    conditional_lines = {}
    if os.path.exists(conditionals_path):
        with open(conditionals_path, 'r', encoding='utf-8') as f:
            conditional_lines = json.load(f)
        print(f"Loaded conditional lines: {list(conditional_lines.keys())}")

    variables = extract_variables(template_text)
    # Also consider variables from conditional lines
    for line in conditional_lines.values():
        variables.extend(extract_variables(line))
    variables = sorted(list(set(variables))) # Get unique variables
    print(f"Found variables: {variables}")

    df = pd.read_excel(excel_path)
    attachment_columns = [c for c in df.columns if c.lower().startswith('attachment')]

    total = len(df)
    generated_files = []
    outlook_drafts_created = 0
    
    # Check if we should use Outlook
    if use_outlook and not OUTLOOK_AVAILABLE:
        print("⚠️ Outlook not available, falling back to .eml files only")
        use_outlook = False
        create_eml_backup = True  # Force .eml creation if Outlook not available
    
    if use_outlook and create_eml_backup:
        print("🎯 Creating Outlook drafts + backup .eml files...")
    elif use_outlook:
        print("📧 Creating Outlook drafts only...")
    else:
        print("📁 Creating .eml files only...")

    for idx, row in df.iterrows():
        success = False
        subject = row.get('Subject') or f'Email_{idx+1}'
        first_name = row.get('FirstName', 'Unknown')
        
        # Get primary email for reference
        to_addresses = parse_email_addresses(row.get('Email') or row.get('email') or row.get('To'))
        primary_email = to_addresses[0] if to_addresses else 'no_email'
        
        # Create Outlook draft if available
        if use_outlook:
            try:
                mail, actual_subject = create_outlook_draft(row, template_text, variables, attachments_dir, attachment_columns, conditional_lines)
                outlook_drafts_created += 1
                success = True
                print(f"[{idx + 1}/{total}] ✅ Outlook Draft: '{actual_subject}' → {primary_email}")
            except Exception as e:
                print(f"[{idx + 1}/{total}] ❌ Outlook Draft Failed: {e}")
                print(f"    Falling back to .eml file...")
        
        # Create backup .eml file (conditional based on create_eml_backup parameter)
        if create_eml_backup or not success:  # Create .eml if requested or if Outlook failed
            try:
                # Create .eml message
                msg = create_email_message(row, template_text, variables, attachments_dir, attachment_columns, conditional_lines)
                
                # Create filename based on subject and recipient
                sanitized_subject = sanitize_filename(subject)[:50]  # Limit length
                sanitized_email = sanitize_filename(primary_email.replace('@', '_at_'))[:30]
                index_str = str(idx + 1).zfill(3)
                
                file_name = f"{index_str}_{sanitized_subject}_{sanitized_email}.eml"
                file_path = os.path.join(output_dir, file_name)
                
                with open(file_path, 'wb') as f:
                    f.write(msg.as_bytes())
                generated_files.append(file_name)
                
                if not success:  # Only show this if Outlook draft failed
                    print(f"[{idx + 1}/{total}] 📄 .eml Backup: {file_name}")
                
            except Exception as e:
                print(f"[{idx + 1}/{total}] ❌ .eml Creation Failed: {e}")

        # Print recipient info
        all_recipients = []
        if to_addresses:
            all_recipients.extend([f"To: {addr}" for addr in to_addresses])
        cc_addresses = parse_email_addresses(row.get('CC') or row.get('cc'))
        if cc_addresses:
            all_recipients.extend([f"CC: {addr}" for addr in cc_addresses])
        bcc_addresses = parse_email_addresses(row.get('BCC') or row.get('bcc'))
        if bcc_addresses:
            all_recipients.extend([f"BCC: {addr}" for addr in bcc_addresses])
        
        if all_recipients and not success:  # Only show detailed info if Outlook failed
            print(f"    Recipients: {', '.join(all_recipients)}")

    print("\n" + "="*60)
    print("✅ GENERATION COMPLETE!")
    print("="*60)
    
    if use_outlook:
        print(f"📧 Outlook Drafts Created: {outlook_drafts_created}")
        print(f"   ➤ Check your Outlook Drafts folder")
    
    print(f"📄 Backup .eml Files: {len(generated_files)}")
    print(f"   ➤ Saved to: {output_dir}")
    
    if generated_files:
        print(f"\n📁 Generated backup files:")
        for name in generated_files[:5]:  # Show first 5
            print(f"   • {name}")
        if len(generated_files) > 5:
            print(f"   ... and {len(generated_files) - 5} more")
    
    if use_outlook and outlook_drafts_created > 0:
        print(f"\n🎯 Next Steps:")
        print(f"   1. Open Outlook and check your Drafts folder")
        print(f"   2. Review each draft email")
        print(f"   3. Remove draft warning notices")
        print(f"   4. Send when ready")
    
    print(f"\n📋 Summary: {outlook_drafts_created if use_outlook else 0} Outlook drafts + {len(generated_files)} backup files")


if __name__ == '__main__':
    main()

