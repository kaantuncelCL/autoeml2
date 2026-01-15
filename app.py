import streamlit as st
import pandas as pd
import json
import os
import re
from pathlib import Path
import tempfile
import shutil
from typing import Dict, List, Optional, Any, Tuple
from streamlit_quill import st_quill
import html
import base64
import traceback

# Import the existing email generator functionality
import email_file_generator as efg
from template_manager import TemplateManager
from datetime import datetime, timedelta

# Import error handling utilities
from error_handler import (
    ErrorHandler, SafeOperation, error_handler,
    validate_email_address, validate_file_size, 
    validate_template_syntax, create_diagnostic_report
)

# Import recovery utilities
try:
    from recovery_utils import (
        session_recovery, display_error_dashboard, 
        display_diagnostic_panel, ApplicationDiagnostics
    )
    RECOVERY_AVAILABLE = True
except ImportError:
    RECOVERY_AVAILABLE = False
    print("Recovery utilities not available")

# Configure page
st.set_page_config(
    page_title="Email Generator Workflow",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state with error handling
try:
    if 'template_text' not in st.session_state:
        st.session_state.template_text = ""
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None
    if 'template_variables' not in st.session_state:
        st.session_state.template_variables = []
    if 'conditional_lines' not in st.session_state:
        st.session_state.conditional_lines = {}
    if 'attachments_dir' not in st.session_state:
        st.session_state.attachments_dir = None
    if 'output_dir' not in st.session_state:
        st.session_state.output_dir = Path.cwd() / "generated_emails"
    if 'current_step' not in st.session_state:
        st.session_state.current_step = 1
    if 'template_manager' not in st.session_state:
        st.session_state.template_manager = TemplateManager()
    if 'template_mode' not in st.session_state:
        st.session_state.template_mode = "plain"  # "plain" or "rich"
    if 'template_html' not in st.session_state:
        st.session_state.template_html = ""
    if 'rich_text_content' not in st.session_state:
        st.session_state.rich_text_content = ""
    if 'attachment_mode' not in st.session_state:
        st.session_state.attachment_mode = "global"  # "global" or "per_recipient"
    if 'attachment_identifier_column' not in st.session_state:
        st.session_state.attachment_identifier_column = None
    if 'per_recipient_attachments_base' not in st.session_state:
        st.session_state.per_recipient_attachments_base = None
    if 'generated_emails' not in st.session_state:
        st.session_state.generated_emails = []
    if 'error_handler' not in st.session_state:
        st.session_state.error_handler = error_handler
    if 'last_auto_save' not in st.session_state:
        st.session_state.last_auto_save = datetime.now()
    if 'safe_mode' not in st.session_state:
        st.session_state.safe_mode = False
except Exception as e:
    st.error(f"Failed to initialize session state: {str(e)}")
    error_handler.log_error(e, "Session Initialization", severity="CRITICAL")

@SafeOperation(error_handler, "Template Variable Extraction")
def extract_variables(template_text: str) -> List[str]:
    """Extract variables from template text with error handling"""
    try:
        if not template_text:
            return []
        
        # Validate template syntax first
        is_valid, errors = validate_template_syntax(template_text)
        if not is_valid:
            error_msg = "Template syntax errors: " + "; ".join(errors)
            raise ValueError(error_msg)
        
        # Extract variables
        all_vars = re.findall(r"\[([^\]]+)\]", template_text)
        # Filter out conditional placeholders
        regular_vars = [var for var in all_vars if not var.startswith("Conditional:")]
        return regular_vars
    except Exception as e:
        error_handler.log_error(e, "Extract Variables")
        return []

def convert_html_to_plain(html_content: str) -> str:
    """Convert HTML content to plain text while preserving placeholders."""
    if not html_content:
        return ""
    
    # Remove HTML tags but preserve content
    import re
    # First, replace <br> and </p> with newlines
    text = html_content.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
    text = text.replace('</p>', '\n').replace('</div>', '\n')
    
    # Remove all other HTML tags
    text = re.sub(r'<[^>]+>', '', text)
    
    # Decode HTML entities
    text = html.unescape(text)
    
    # Clean up excessive newlines
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    return text.strip()

def convert_plain_to_html(plain_text: str) -> str:
    """Convert plain text to HTML format."""
    if not plain_text:
        return ""
    
    # Escape HTML special characters but preserve placeholders
    lines = plain_text.split('\n')
    html_lines = []
    
    for line in lines:
        # Don't escape content within brackets (our placeholders)
        parts = re.split(r'(\[[^\]]+\])', line)
        escaped_parts = []
        for part in parts:
            if part.startswith('[') and part.endswith(']'):
                # This is a placeholder, keep it as-is
                escaped_parts.append(part)
            else:
                # Regular text, escape HTML
                escaped_parts.append(html.escape(part))
        html_lines.append(''.join(escaped_parts))
    
    # Join lines with <br> tags
    return '<br>'.join(html_lines)

def create_html_preview(html_content: str) -> str:
    """Create a preview of the HTML email with basic styling."""
    preview_html = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                font-size: 14px;
                line-height: 1.6;
                padding: 20px;
                background-color: #ffffff;
            }}
            a {{
                color: #0066cc;
                text-decoration: none;
            }}
            a:hover {{
                text-decoration: underline;
            }}
            ul, ol {{
                margin-left: 20px;
            }}
            .placeholder {{
                background-color: #fffacd;
                padding: 2px 4px;
                border-radius: 3px;
                font-family: monospace;
            }}
        </style>
    </head>
    <body>
        {html_content}
    </body>
    </html>
    """
    return preview_html

def validate_excel_columns(df: pd.DataFrame, required_vars: List[str]) -> Dict[str, Any]:
    """Validate Excel columns against template variables with comprehensive error checking"""
    try:
        if df is None or df.empty:
            return {
                'has_email': False,
                'missing_vars': required_vars,
                'available_cols': [],
                'is_valid': False,
                'error': 'DataFrame is empty or None'
            }
        
        missing_vars = []
        available_cols = df.columns.tolist()
        invalid_emails = []
        
        # Check for email column
        email_cols = ['Email', 'email', 'To', 'to']
        email_col = None
        for col in email_cols:
            if col in available_cols:
                email_col = col
                break
        
        has_email = email_col is not None
        
        # Validate email addresses if email column exists
        if has_email:
            for idx, row in df.iterrows():
                email = row.get(email_col, '')
                # Handle NaN and None values
                if pd.isna(email) or email == '':
                    continue  # Skip empty email rows
                email = str(email).strip()
                if not email:
                    continue  # Skip whitespace-only emails
                is_valid, error = validate_email_address(email)
                if not is_valid:
                    invalid_emails.append(f"Row {idx+1}: {error}")
        
        # Check for required template variables
        for var in required_vars:
            if var not in available_cols:
                missing_vars.append(var)
        
        # Check for Subject column
        has_subject = 'Subject' in available_cols or 'subject' in available_cols
        
        return {
            'has_email': has_email,
            'has_subject': has_subject,
            'missing_vars': missing_vars,
            'available_cols': available_cols,
            'invalid_emails': invalid_emails[:5],  # Show first 5 invalid emails
            'is_valid': has_email and len(missing_vars) == 0 and len(invalid_emails) == 0
        }
    except Exception as e:
        error_handler.log_error(e, "Excel Validation")
        return {
            'has_email': False,
            'missing_vars': required_vars,
            'available_cols': [],
            'is_valid': False,
            'error': str(e)
        }

def create_preview_email(row: pd.Series, template_text: str, variables: List[str], conditional_lines: Dict[str, str], is_html: bool = False) -> str:
    """Create a preview of what the email would look like for a given row"""
    # Use HTML template if in rich text mode
    if is_html and st.session_state.template_mode == "rich" and st.session_state.template_html:
        template_to_use = st.session_state.template_html
    else:
        template_to_use = template_text
    
    # Fill template with row data
    preview = efg.fill_template(template_to_use, row, variables)
    
    # Handle conditional placeholders
    conditional_placeholders = re.findall(r"\[Conditional:([^\]]+)\]", preview)
    for placeholder_key in conditional_placeholders:
        placeholder_tag = f"[Conditional:{placeholder_key}]"
        replacement_text = ""
        if placeholder_key in conditional_lines and placeholder_key in row and row[placeholder_key] == 1:
            raw_text = conditional_lines[placeholder_key]
            replacement_text = efg.fill_template(raw_text, row, variables)
        preview = preview.replace(placeholder_tag, replacement_text)
    
    return preview

def main():
    st.title("📧 Email Generator Workflow")
    st.markdown("Generate personalized emails from templates with Outlook integration")
    
    # Sidebar for navigation
    with st.sidebar:
        st.header("Workflow Steps")
        
        # Step indicators
        steps = [
            "1. Template Setup",
            "2. Excel Data Upload", 
            "3. Variable Mapping",
            "4. Attachments",
            "5. Conditional Content",
            "6. Preview & Generate"
        ]
        
        for i, step in enumerate(steps, 1):
            if i == st.session_state.current_step:
                st.markdown(f"**▶️ {step}**")
            elif i < st.session_state.current_step:
                st.markdown(f"✅ {step}")
            else:
                st.markdown(f"⏸️ {step}")
        
        st.divider()
        
        # Output directory setting
        st.subheader("Output Settings")
        output_path = st.text_input("Output Directory", value=str(st.session_state.output_dir))
        if output_path != str(st.session_state.output_dir):
            st.session_state.output_dir = Path(output_path)
        
        # Create output directory if it doesn't exist
        st.session_state.output_dir.mkdir(exist_ok=True)
        st.success(f"Output: {st.session_state.output_dir}")
        
        st.divider()
        
        # Template Library Section
        st.header("📚 Template Library")
        
        # List available templates
        templates = st.session_state.template_manager.list_templates()
        
        if templates:
            st.subheader("Available Templates")
            for template in templates:
                format_icon = "🎨" if template.get('format_type', 'plain') == 'rich' else "📄"
                with st.expander(f"{format_icon} {template['name']}" + (" 🔀" if template['has_conditionals'] else "")):
                    st.markdown(f"**Description:** {template['description']}")
                    st.markdown(f"**Created:** {template['created_date']}")
                    st.markdown(f"**Variables:** {template['variable_count']}")
                    st.markdown(f"**Format:** {template.get('format_type', 'plain').title()} Text")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(f"Load", key=f"load_{template['filename']}", use_container_width=True):
                            result = st.session_state.template_manager.load_template(template['filename'])
                            if result['success']:
                                data = result['data']
                                # Load template text
                                st.session_state.template_text = data['template_text']
                                # Load format type
                                st.session_state.template_mode = data.get('format_type', 'plain')
                                # Load HTML if available
                                if st.session_state.template_mode == 'rich' and 'template_html' in data:
                                    st.session_state.template_html = data['template_html']
                                    st.session_state.rich_text_content = data['template_html']
                                else:
                                    st.session_state.template_html = ""
                                    st.session_state.rich_text_content = ""
                                # Recompute variables from the loaded template text to ensure consistency
                                st.session_state.template_variables = extract_variables(data['template_text'])
                                # Load conditional lines if present
                                if 'conditional_keys' in data:
                                    # Prepare default conditional lines for the loaded template
                                    for key in data['conditional_keys']:
                                        if key not in st.session_state.conditional_lines:
                                            st.session_state.conditional_lines[key] = f"As a valued {key} member, you have access to exclusive benefits."
                                st.success(f"Loaded template: {template['name']}")
                                st.session_state.current_step = 1  # Go to template setup
                                st.rerun()
                            else:
                                st.error(result['message'])
                    
                    with col2:
                        # Use a popover for delete confirmation
                        with st.popover("Delete", use_container_width=True):
                            st.warning(f"Delete '{template['name']}'?")
                            if st.button("Confirm Delete", key=f"confirm_delete_{template['filename']}", type="primary"):
                                result = st.session_state.template_manager.delete_template(template['filename'])
                                if result['success']:
                                    st.success(result['message'])
                                    st.rerun()
                                else:
                                    st.error(result['message'])
        else:
            st.info("No saved templates yet. Create and save your first template!")
        
        # Save current template section
        st.divider()
        st.subheader("💾 Save Current Template")
        
        if st.session_state.template_text:
            template_name = st.text_input(
                "Template Name",
                placeholder="e.g., Sales Outreach",
                key="save_template_name"
            )
            
            template_description = st.text_area(
                "Description (optional)",
                placeholder="Brief description of this template",
                height=60,
                key="save_template_description"
            )
            
            if st.button("💾 Save to Library", use_container_width=True):
                if template_name:
                    result = st.session_state.template_manager.save_template(
                        name=template_name,
                        template_text=st.session_state.template_text,
                        description=template_description,
                        format_type=st.session_state.template_mode,
                        template_html=st.session_state.template_html if st.session_state.template_mode == "rich" else None
                    )
                    if result['success']:
                        st.success(result['message'])
                        st.balloons()
                        st.rerun()
                    else:
                        st.error(result['message'])
                else:
                    st.error("Please enter a template name")
        else:
            st.info("Create a template first to save it to the library")

    # Main content area
    if st.session_state.current_step == 1:
        step_1_template_setup()
    elif st.session_state.current_step == 2:
        step_2_excel_upload()
    elif st.session_state.current_step == 3:
        step_3_variable_mapping()
    elif st.session_state.current_step == 4:
        step_4_attachments()
    elif st.session_state.current_step == 5:
        step_5_conditional_content()
    elif st.session_state.current_step == 6:
        step_6_preview_generate()

def step_1_template_setup():
    st.header("1. Template Setup")
    st.markdown("Create or edit your email template with variable placeholders like `[FirstName]`")
    
    # Template mode toggle
    mode_col1, mode_col2, mode_col3 = st.columns([1, 2, 2])
    with mode_col1:
        template_mode = st.radio(
            "Template Mode",
            options=["plain", "rich"],
            format_func=lambda x: "Plain Text" if x == "plain" else "Rich Text",
            index=0 if st.session_state.template_mode == "plain" else 1,
            help="Choose between plain text or rich text editing with formatting options"
        )
        
        if template_mode != st.session_state.template_mode:
            st.session_state.template_mode = template_mode
            # Convert content when switching modes
            if template_mode == "rich" and st.session_state.template_text:
                # Convert plain text to HTML
                st.session_state.template_html = convert_plain_to_html(st.session_state.template_text)
                st.session_state.rich_text_content = st.session_state.template_html
            elif template_mode == "plain" and st.session_state.template_html:
                # Convert HTML to plain text
                st.session_state.template_text = convert_html_to_plain(st.session_state.template_html)
            st.rerun()
    
    with mode_col2:
        if st.session_state.template_mode == "rich":
            st.info("📝 Rich Text Mode: Use the toolbar to format your email. Variables like [FirstName] work in rich text too!")
        else:
            st.info("📄 Plain Text Mode: Simple text editing with variable placeholders.")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if st.session_state.template_mode == "plain":
            # Plain text editor
            template_text = st.text_area(
                "Email Template",
                value=st.session_state.template_text,
                height=400,
                help="Use [VariableName] for placeholders. Use [Conditional:KeyName] for conditional content."
            )
            
            if template_text != st.session_state.template_text:
                st.session_state.template_text = template_text
                st.session_state.template_variables = extract_variables(template_text) if template_text else []
                st.rerun()
        else:
            # Rich text editor
            st.markdown("### Email Template Editor")
            st.markdown("💡 **Tip:** Type `[FirstName]`, `[Company]` etc. for variable placeholders")
            
            # Quill editor configuration
            quill_toolbar = [
                ['bold', 'italic', 'underline', 'strike'],
                [{'size': ['small', False, 'large', 'huge']}],
                [{'color': []}, {'background': []}],
                [{'list': 'ordered'}, {'list': 'bullet'}],
                ['link', 'blockquote'],
                [{'align': []}],
                ['clean']
            ]
            
            # Rich text editor
            content = st_quill(
                value=st.session_state.rich_text_content,
                html=True,
                toolbar=quill_toolbar,
                key="rich_editor"
            )
            
            if content != st.session_state.rich_text_content:
                st.session_state.rich_text_content = content
                st.session_state.template_html = content
                # Extract plain text version for variable detection
                plain_version = convert_html_to_plain(content)
                st.session_state.template_text = plain_version
                st.session_state.template_variables = extract_variables(plain_version) if plain_version else []
            
            # HTML Preview
            if st.session_state.template_html:
                st.markdown("### 👁️ HTML Preview")
                # Create an iframe-like preview
                preview_html = create_html_preview(st.session_state.template_html)
                st.components.v1.html(preview_html, height=300, scrolling=True)
    
    with col2:
        st.subheader("Detected Variables")
        if st.session_state.template_variables:
            for var in st.session_state.template_variables:
                if var.startswith('Conditional:'):
                    st.markdown(f"🔀 `{var}` (conditional)")
                else:
                    st.markdown(f"📝 `{var}`")
        else:
            st.info("No variables detected")
        
        # Template actions
        st.subheader("Quick Actions")
        
        # Load from file
        file_types = ['txt', 'html'] if st.session_state.template_mode == "rich" else ['txt']
        uploaded_template = st.file_uploader(
            f"Upload Template (.{'/'.join(file_types)})", 
            type=file_types, 
            key="upload_template"
        )
        if uploaded_template:
            content = uploaded_template.read().decode('utf-8')
            if st.session_state.template_mode == "rich" and uploaded_template.name.endswith('.html'):
                st.session_state.template_html = content
                st.session_state.rich_text_content = content
                st.session_state.template_text = convert_html_to_plain(content)
            else:
                st.session_state.template_text = content
                if st.session_state.template_mode == "rich":
                    st.session_state.template_html = convert_plain_to_html(content)
                    st.session_state.rich_text_content = st.session_state.template_html
            st.session_state.template_variables = extract_variables(st.session_state.template_text)
            st.success("✅ Template loaded from file!")
            st.rerun()
        
        # Quick save to library
        if st.session_state.template_text:
            st.info("💡 Use the Template Library in the sidebar to save this template")
        
        # Export template
        export_format = st.selectbox(
            "Export Format",
            options=["txt", "html"] if st.session_state.template_mode == "rich" else ["txt"],
            key="export_format"
        )
        
        if st.button("📥 Export Template"):
            if st.session_state.template_text:
                timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
                if export_format == "html" and st.session_state.template_mode == "rich":
                    template_path = st.session_state.output_dir / f"template_{timestamp}.html"
                    with open(template_path, 'w', encoding='utf-8') as f:
                        f.write(st.session_state.template_html)
                else:
                    template_path = st.session_state.output_dir / f"template_{timestamp}.txt"
                    with open(template_path, 'w', encoding='utf-8') as f:
                        f.write(st.session_state.template_text)
                st.success(f"Exported to: {template_path}")
            else:
                st.warning("No template to export")
        
        # Formatting tips
        if st.session_state.template_mode == "rich":
            with st.expander("📖 Rich Text Tips"):
                st.markdown("""
                **Variable Placeholders:**
                - Type `[FirstName]`, `[Company]` etc. directly
                - They'll be highlighted in the preview
                
                **Formatting:**
                - Select text and use toolbar buttons
                - Bold: **Ctrl/Cmd + B**
                - Italic: *Ctrl/Cmd + I*
                - Underline: Ctrl/Cmd + U
                
                **Links:**
                - Select text and click the link button
                - Enter the URL in the popup
                
                **Lists:**
                - Use the list buttons in the toolbar
                - Press Enter for new items
                """)
    
    # Navigation
    st.divider()
    col1, col2, col3 = st.columns([1, 1, 1])
    with col3:
        # Check if we have content based on mode
        has_content = (st.session_state.template_text if st.session_state.template_mode == "plain" 
                      else st.session_state.template_html)
        if st.button("Next: Excel Upload ▶️", disabled=not has_content):
            st.session_state.current_step = 2
            st.rerun()

def step_2_excel_upload():
    st.header("2. Excel Data Upload")
    st.markdown("Upload your Excel file containing recipient data")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
    
    if uploaded_file:
        try:
            # Read Excel file
            df = pd.read_excel(uploaded_file)
            st.session_state.excel_data = df
            
            st.success(f"✅ Loaded {len(df)} rows from Excel file")
            
            # Display data preview
            st.subheader("Data Preview")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Show column info
            st.subheader("Column Information")
            col_info = []
            for col in df.columns:
                non_null_count = df[col].count()
                col_info.append({
                    'Column': col,
                    'Non-null Count': non_null_count,
                    'Data Type': str(df[col].dtype),
                    'Sample Values': ', '.join([str(x) for x in df[col].dropna().head(3).tolist()])
                })
            
            st.dataframe(pd.DataFrame(col_info), use_container_width=True)
            
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
            st.session_state.excel_data = None
    
    elif st.session_state.excel_data is not None:
        st.info("Excel data already loaded")
        st.dataframe(st.session_state.excel_data.head(5), use_container_width=True)
    
    # Navigation
    st.divider()
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        if st.button("◀️ Back: Template"):
            st.session_state.current_step = 1
            st.rerun()
    with col3:
        if st.button("Next: Variable Mapping ▶️", disabled=st.session_state.excel_data is None):
            st.session_state.current_step = 3
            st.rerun()

def step_3_variable_mapping():
    st.header("3. Variable Mapping Verification")
    st.markdown("Verify that your Excel columns match the template variables")
    
    if st.session_state.excel_data is None:
        st.error("Please upload Excel data first")
        return
    
    # Validate columns
    validation = validate_excel_columns(st.session_state.excel_data, st.session_state.template_variables)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Template Variables")
        if st.session_state.template_variables:
            for var in st.session_state.template_variables:
                if var in validation['available_cols']:
                    st.success(f"✅ `{var}` - Found in Excel")
                else:
                    st.error(f"❌ `{var}` - Missing from Excel")
        else:
            st.info("No template variables detected")
    
    with col2:
        st.subheader("Excel Columns")
        for col in validation['available_cols']:
            if col in st.session_state.template_variables:
                st.success(f"✅ `{col}` - Used in template")
            elif col.lower() in ['email', 'to', 'cc', 'bcc', 'from', 'subject']:
                st.info(f"📧 `{col}` - Email field")
            elif col in st.session_state.conditional_lines:
                st.info(f"🔀 `{col}` - Conditional flag")
            else:
                st.warning(f"⚠️ `{col}` - Not used")
    
    # Email validation
    st.subheader("Email Configuration Check")
    if validation['has_email']:
        st.success("✅ Email column found")
    else:
        st.error("❌ No email column found (need 'Email', 'email', 'To', or 'to')")
    
    # Show mapping issues
    if validation['missing_vars']:
        st.error("❌ Missing variables in Excel data:")
        for var in validation['missing_vars']:
            st.write(f"- `{var}`")
        st.markdown("**Fix:** Add these columns to your Excel file or remove the variables from your template")
    
    # Show invalid email addresses
    if validation.get('invalid_emails'):
        st.error("❌ Invalid email addresses found:")
        for email_error in validation['invalid_emails']:
            st.write(f"- {email_error}")
        st.markdown("**Fix:** Correct the email addresses in your Excel file")
    
    # Navigation
    st.divider()
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        if st.button("◀️ Back: Excel Upload"):
            st.session_state.current_step = 2
            st.rerun()
    with col3:
        if st.button("Next: Attachments ▶️", disabled=not validation['is_valid']):
            st.session_state.current_step = 4
            st.rerun()

def step_4_attachments():
    st.header("4. Attachments Configuration")
    st.markdown("Configure attachments for your emails")
    
    # Attachment mode selection
    st.subheader("Attachment Mode")
    attachment_mode = st.radio(
        "Select how to attach files:",
        options=["global", "per_recipient"],
        format_func=lambda x: "Global Attachments (same for all)" if x == "global" else "Per-Recipient Attachments (individual folders)",
        index=0 if st.session_state.attachment_mode == "global" else 1,
        help="Global: All recipients get the same attachments. Per-Recipient: Each recipient gets attachments from their own folder."
    )
    
    if attachment_mode != st.session_state.attachment_mode:
        st.session_state.attachment_mode = attachment_mode
        st.rerun()
    
    if st.session_state.attachment_mode == "global":
        # Global attachments mode
        st.subheader("Global Attachments")
        st.markdown("Select a folder containing files that will be attached to ALL emails")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            attachments_path = st.text_input(
                "Attachments Folder Path",
                value=str(st.session_state.attachments_dir) if st.session_state.attachments_dir else "",
                help="Enter the path to the folder containing your attachment files"
            )
            
            if attachments_path and Path(attachments_path).exists():
                st.session_state.attachments_dir = Path(attachments_path)
                
                # List files in directory
                files = list(st.session_state.attachments_dir.glob("*"))
                files = [f for f in files if f.is_file()]
                
                if files:
                    st.success(f"✅ Found {len(files)} files")
                    
                    # Show file list
                    file_info = []
                    for file_path in files:
                        size_mb = file_path.stat().st_size / (1024 * 1024)
                        file_info.append({
                            'Filename': file_path.name,
                            'Size (MB)': f"{size_mb:.2f}",
                            'Type': file_path.suffix
                        })
                    
                    st.dataframe(pd.DataFrame(file_info), use_container_width=True)
                else:
                    st.warning("⚠️ No files found in directory")
            elif attachments_path:
                st.error("❌ Directory not found")
        
        with col2:
            st.subheader("Quick Actions")
            
            # Create attachments directory
            if st.button("📁 Create Attachments Folder"):
                new_dir = st.session_state.output_dir / "attachments"
                new_dir.mkdir(exist_ok=True)
                st.session_state.attachments_dir = new_dir
                st.success(f"Created: {new_dir}")
                st.rerun()
            
            # Upload files
            uploaded_attachments = st.file_uploader(
                "Upload Attachment Files",
                accept_multiple_files=True,
                help="Upload files to add to the attachments folder"
            )
            
            if uploaded_attachments and st.session_state.attachments_dir:
                for uploaded_file in uploaded_attachments:
                    file_path = st.session_state.attachments_dir / uploaded_file.name
                    with open(file_path, 'wb') as f:
                        f.write(uploaded_file.read())
                    st.success(f"Saved: {uploaded_file.name}")
    
    else:
        # Per-recipient attachments mode
        st.subheader("Per-Recipient Attachments")
        st.markdown("Each recipient gets attachments from their individual folder")
        
        if st.session_state.excel_data is None:
            st.error("Please upload Excel data first to configure per-recipient attachments")
            col1, col2, col3 = st.columns([1, 1, 1])
            with col1:
                if st.button("◀️ Back: Variable Mapping"):
                    st.session_state.current_step = 3
                    st.rerun()
            return
        
        # Base folder selection
        col1, col2 = st.columns([2, 1])
        
        with col1:
            base_path = st.text_input(
                "Base Attachments Folder",
                value=str(st.session_state.per_recipient_attachments_base) if st.session_state.per_recipient_attachments_base else "",
                help="Enter the base folder path containing recipient subfolders (e.g., attachments/)"
            )
            
            if base_path:
                base_path_obj = Path(base_path)
                if base_path_obj.exists():
                    st.session_state.per_recipient_attachments_base = base_path_obj
                    st.success(f"✅ Base folder exists: {base_path}")
                else:
                    st.error(f"❌ Base folder not found: {base_path}")
        
        with col2:
            st.subheader("Quick Actions")
            if st.button("📁 Create Base Folder"):
                new_base = st.session_state.output_dir / "attachments"
                new_base.mkdir(exist_ok=True)
                st.session_state.per_recipient_attachments_base = new_base
                st.success(f"Created: {new_base}")
                st.rerun()
        
        # Column selection for folder identifier
        if st.session_state.per_recipient_attachments_base and st.session_state.per_recipient_attachments_base.exists():
            st.subheader("Folder Identifier Column")
            
            available_columns = st.session_state.excel_data.columns.tolist()
            
            # Suggest common identifier columns
            suggested_columns = ['FirstName', 'LastName', 'Email', 'Name', 'ID', 'CustomerID']
            suggested_available = [col for col in suggested_columns if col in available_columns]
            
            identifier_column = st.selectbox(
                "Select the Excel column to use as folder names:",
                options=available_columns,
                index=available_columns.index(st.session_state.attachment_identifier_column) if st.session_state.attachment_identifier_column in available_columns else (available_columns.index(suggested_available[0]) if suggested_available else 0),
                help="This column's values will be used as subfolder names for each recipient"
            )
            
            if identifier_column != st.session_state.attachment_identifier_column:
                st.session_state.attachment_identifier_column = identifier_column
                st.rerun()
            
            # Show folder mapping preview
            if st.session_state.attachment_identifier_column:
                st.subheader("Folder Mapping Preview")
                
                # Create folder mapping
                folder_mapping = []
                for idx, row in st.session_state.excel_data.iterrows():
                    identifier_value = str(row[st.session_state.attachment_identifier_column])
                    # Clean identifier for folder name (remove special characters)
                    clean_identifier = re.sub(r'[^\w\s-]', '', identifier_value).strip()
                    if clean_identifier:
                        folder_path = st.session_state.per_recipient_attachments_base / clean_identifier
                        
                        # Check folder existence and count files
                        exists = folder_path.exists()
                        file_count = 0
                        total_size_mb = 0
                        files_list = []
                        
                        if exists:
                            files = [f for f in folder_path.glob("*") if f.is_file()]
                            file_count = len(files)
                            total_size_mb = sum(f.stat().st_size for f in files) / (1024 * 1024)
                            files_list = [f.name for f in files[:5]]  # Show first 5 files
                        
                        folder_mapping.append({
                            'Recipient': row.get('Email', row.get('email', f'Row {idx+1}')),
                            'Identifier': identifier_value,
                            'Folder': clean_identifier,
                            'Path': str(folder_path),
                            'Exists': '✅' if exists else '❌',
                            'Files': file_count,
                            'Size (MB)': f"{total_size_mb:.2f}" if exists else "0",
                            'Sample Files': ', '.join(files_list[:3]) + ('...' if len(files_list) > 3 else '') if files_list else 'None'
                        })
                
                # Display mapping table
                if folder_mapping:
                    df_mapping = pd.DataFrame(folder_mapping)
                    
                    # Summary statistics
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        total_recipients = len(folder_mapping)
                        st.metric("Total Recipients", total_recipients)
                    with col2:
                        folders_exist = sum(1 for m in folder_mapping if m['Exists'] == '✅')
                        st.metric("Folders Found", f"{folders_exist}/{total_recipients}")
                    with col3:
                        total_files = sum(m['Files'] for m in folder_mapping)
                        st.metric("Total Files", total_files)
                    with col4:
                        total_size = sum(float(m['Size (MB)']) for m in folder_mapping)
                        st.metric("Total Size", f"{total_size:.2f} MB")
                    
                    # Show detailed mapping
                    st.dataframe(df_mapping, use_container_width=True, height=400)
                    
                    # Warnings for missing folders
                    missing_folders = [m for m in folder_mapping if m['Exists'] == '❌']
                    if missing_folders:
                        with st.expander(f"⚠️ Missing Folders ({len(missing_folders)})", expanded=True):
                            st.warning("The following recipients don't have attachment folders:")
                            for missing in missing_folders[:10]:
                                st.write(f"- **{missing['Recipient']}**: Expected folder `{missing['Folder']}`")
                            if len(missing_folders) > 10:
                                st.write(f"... and {len(missing_folders) - 10} more")
                            
                            if st.button("📁 Create All Missing Folders"):
                                created_count = 0
                                for missing in missing_folders:
                                    folder_path = Path(missing['Path'])
                                    folder_path.mkdir(parents=True, exist_ok=True)
                                    created_count += 1
                                st.success(f"Created {created_count} folders!")
                                st.rerun()
                    
                    # Large attachment warnings
                    large_attachments = [m for m in folder_mapping if m['Files'] > 0 and float(m['Size (MB)']) > 10]
                    if large_attachments:
                        with st.expander(f"⚠️ Large Attachments ({len(large_attachments)})"):
                            st.warning("The following recipients have attachments over 10 MB:")
                            for large in large_attachments[:5]:
                                st.write(f"- **{large['Recipient']}**: {large['Size (MB)']} MB ({large['Files']} files)")
                            if len(large_attachments) > 5:
                                st.write(f"... and {len(large_attachments) - 5} more")
    
    # Individual attachments info (legacy)
    st.divider()
    st.subheader("Excel Column Attachments (Optional)")
    st.markdown("You can also specify individual attachments in your Excel file using columns like `Attachment1`, `Attachment2`, etc.")
    
    if st.session_state.excel_data is not None:
        attachment_cols = [col for col in st.session_state.excel_data.columns if col.lower().startswith('attachment')]
        if attachment_cols:
            st.info(f"Found attachment columns in Excel: {', '.join(attachment_cols)}")
        else:
            st.info("No attachment columns found in Excel data")
    
    # Navigation
    st.divider()
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        if st.button("◀️ Back: Variable Mapping"):
            st.session_state.current_step = 3
            st.rerun()
    with col3:
        if st.button("Next: Conditional Content ▶️"):
            st.session_state.current_step = 5
            st.rerun()

def step_5_conditional_content():
    st.header("5. Conditional Content Configuration")
    st.markdown("Configure conditional content that appears based on flags in your Excel data")
    
    # JSON editor for conditional lines
    st.subheader("Conditional Content Rules")
    st.markdown("Define text snippets that will be included when the corresponding Excel column is set to `1`")
    
    # Load existing conditional lines
    try:
        with open('conditional_lines.json', 'r', encoding='utf-8') as f:
            default_conditional = json.load(f)
    except:
        default_conditional = st.session_state.conditional_lines
    
    # JSON input
    conditional_json = st.text_area(
        "Conditional Lines (JSON format)",
        value=json.dumps(default_conditional, indent=2),
        height=200,
        help="Define key-value pairs where the key matches an Excel column and the value is the text to include"
    )
    
    try:
        conditional_data = json.loads(conditional_json)
        st.session_state.conditional_lines = conditional_data
        
        # Show preview of conditional rules
        if conditional_data:
            st.success("✅ JSON is valid")
            
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Conditional Rules")
                for key, value in conditional_data.items():
                    st.markdown(f"**{key}:**")
                    st.markdown(f"> {value}")
                    st.divider()
            
            with col2:
                st.subheader("Excel Column Check")
                if st.session_state.excel_data is not None:
                    excel_cols = st.session_state.excel_data.columns.tolist()
                    for key in conditional_data.keys():
                        if key in excel_cols:
                            st.success(f"✅ `{key}` - Found in Excel")
                        else:
                            st.warning(f"⚠️ `{key}` - Not found in Excel")
                else:
                    st.info("Upload Excel data to verify columns")
        else:
            st.info("No conditional rules defined")
            
    except json.JSONDecodeError as e:
        st.error(f"❌ Invalid JSON: {str(e)}")
    
    # Save conditional lines
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("💾 Save Conditional Rules"):
            with open('conditional_lines.json', 'w', encoding='utf-8') as f:
                json.dump(st.session_state.conditional_lines, f, indent=2)
            st.success("Conditional rules saved!")
    
    # Navigation
    st.divider()
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        if st.button("◀️ Back: Attachments"):
            st.session_state.current_step = 4
            st.rerun()
    with col3:
        if st.button("Next: Preview & Generate ▶️"):
            st.session_state.current_step = 6
            st.rerun()

def step_6_preview_generate():
    st.header("6. Preview & Generate Emails")
    st.markdown("Preview your emails and generate Outlook drafts")
    
    if st.session_state.excel_data is None:
        st.error("Please upload Excel data first")
        return
    
    # Email preview
    st.subheader("Email Preview")
    
    # Select row for preview
    preview_row_idx = st.selectbox(
        "Select row to preview:",
        range(len(st.session_state.excel_data)),
        format_func=lambda x: f"Row {x+1}: {st.session_state.excel_data.iloc[x].get('FirstName', st.session_state.excel_data.iloc[x].get('Email', f'Row {x+1}'))}"
    )
    
    if preview_row_idx is not None:
        row = st.session_state.excel_data.iloc[preview_row_idx]
        
        # Show email preview
        col1, col2 = st.columns([2, 1])
        
        with col1:
            # Generate preview
            is_html_mode = st.session_state.template_mode == "rich"
            preview_text = create_preview_email(
                row, 
                st.session_state.template_text, 
                st.session_state.template_variables,
                st.session_state.conditional_lines,
                is_html=is_html_mode
            )
            
            st.markdown("**Email Preview:**")
            if is_html_mode:
                # Show HTML preview for rich text mode
                st.markdown("*Rich Text Preview:*")
                preview_html = create_html_preview(preview_text)
                st.components.v1.html(preview_html, height=300, scrolling=True)
                
                # Option to view raw HTML
                with st.expander("View Raw HTML"):
                    st.code(preview_text, language="html")
            else:
                # Show plain text preview
                st.text_area("", value=preview_text, height=300, disabled=True)
        
        with col2:
            st.markdown("**Email Details:**")
            
            # Email addresses
            email_field = row.get('Email') or row.get('email') or row.get('To') or 'No email'
            st.write(f"**To:** {email_field}")
            
            if 'CC' in row or 'cc' in row:
                cc_field = row.get('CC') or row.get('cc')
                st.write(f"**CC:** {cc_field}")
            
            if 'BCC' in row or 'bcc' in row:
                bcc_field = row.get('BCC') or row.get('bcc')
                st.write(f"**BCC:** {bcc_field}")
            
            subject_field = row.get('Subject') or 'No Subject'
            st.write(f"**Subject:** {subject_field}")
            
            # Show active conditional flags
            active_flags = []
            for key in st.session_state.conditional_lines.keys():
                if key in row and row[key] == 1:
                    active_flags.append(key)
            
            if active_flags:
                st.write(f"**Active Flags:** {', '.join(active_flags)}")
            
            # Attachments info
            st.write("**Attachments:**")
            attachments_to_add = []
            
            if st.session_state.attachment_mode == "global":
                # Global attachments mode
                if st.session_state.attachments_dir:
                    files = list(st.session_state.attachments_dir.glob("*"))
                    files = [f for f in files if f.is_file()]
                    attachments_to_add = files
                    st.write(f"  • {len(files)} global files")
                    if files:
                        with st.expander("View attachments"):
                            for f in files[:10]:
                                size_mb = f.stat().st_size / (1024 * 1024)
                                st.write(f"  - {f.name} ({size_mb:.2f} MB)")
                            if len(files) > 10:
                                st.write(f"  ... and {len(files) - 10} more")
                else:
                    st.write("  • No global attachments")
            else:
                # Per-recipient attachments mode
                if st.session_state.per_recipient_attachments_base and st.session_state.attachment_identifier_column:
                    identifier_value = str(row.get(st.session_state.attachment_identifier_column, "")).strip()
                    if identifier_value:
                        clean_identifier = re.sub(r'[^\w\s-]', '', identifier_value).strip()
                        if clean_identifier:
                            recipient_folder = st.session_state.per_recipient_attachments_base / clean_identifier
                            if recipient_folder.exists():
                                files = [f for f in recipient_folder.glob("*") if f.is_file()]
                                attachments_to_add = files
                                st.write(f"  • {len(files)} files from {clean_identifier}/")
                                if files:
                                    with st.expander("View attachments"):
                                        total_size_mb = 0
                                        for f in files[:10]:
                                            size_mb = f.stat().st_size / (1024 * 1024)
                                            total_size_mb += size_mb
                                            st.write(f"  - {f.name} ({size_mb:.2f} MB)")
                                        if len(files) > 10:
                                            st.write(f"  ... and {len(files) - 10} more")
                                        st.write(f"  **Total: {total_size_mb:.2f} MB**")
                                        if total_size_mb > 25:
                                            st.warning("⚠️ Large total size!")
                            else:
                                st.warning(f"  • Folder not found: {clean_identifier}/")
                                # Check if global fallback exists
                                if st.session_state.attachments_dir:
                                    files = list(st.session_state.attachments_dir.glob("*"))
                                    files = [f for f in files if f.is_file()]
                                    if files:
                                        st.info(f"  • Will use {len(files)} global files as fallback")
                        else:
                            st.warning("  • Invalid folder identifier")
                else:
                    st.warning("  • Per-recipient mode not configured")
    
    # Generation settings
    st.subheader("Generation Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        use_outlook = st.checkbox(
            "Create .msg Files (Editable in Outlook)",
            value=True,
            help="Create editable .msg files that can be opened and modified in Microsoft Outlook (requires pywin32)"
        )
        
        create_eml_backup = st.checkbox(
            "Create .eml Backup Files", 
            value=False,
            help="Create .eml files as backup (readable but not editable)"
        )
    
    with col2:
        # Check Outlook availability
        try:
            import win32com.client
            st.success("✅ Outlook integration available")
            outlook_available = True
        except ImportError:
            st.warning("⚠️ Outlook integration not available - will create .eml files instead")
            outlook_available = False
    
    # Generate emails
    st.divider()
    
    if st.button("🚀 Generate All Emails", type="primary"):
        if not use_outlook and not create_eml_backup:
            st.error("Please select at least one output option")
            return
        
        # Show progress
        progress_bar = st.progress(0)
        status_text = st.empty()
        results_container = st.container()
        
        try:
            # Prepare parameters
            if st.session_state.template_mode == "rich":
                # Save as HTML template for rich text mode
                template_path = st.session_state.output_dir / "temp_template.html"
                with open(template_path, 'w', encoding='utf-8') as f:
                    f.write(st.session_state.template_html)
            else:
                template_path = st.session_state.output_dir / "temp_template.txt"
                with open(template_path, 'w', encoding='utf-8') as f:
                    f.write(st.session_state.template_text)
            
            excel_path = st.session_state.output_dir / "temp_data.xlsx"
            st.session_state.excel_data.to_excel(excel_path, index=False)
            
            conditionals_path = "conditional_lines.json"
            with open(conditionals_path, 'w', encoding='utf-8') as f:
                json.dump(st.session_state.conditional_lines, f, indent=2)
            
            # Call the main function from email_file_generator
            status_text.text("Generating emails...")
            
            # Prepare attachment parameters based on mode
            if st.session_state.attachment_mode == "per_recipient":
                attachments_dir = str(st.session_state.attachments_dir) if st.session_state.attachments_dir else None
                per_recipient_base = str(st.session_state.per_recipient_attachments_base) if st.session_state.per_recipient_attachments_base else None
                identifier_column = st.session_state.attachment_identifier_column
            else:
                attachments_dir = str(st.session_state.attachments_dir) if st.session_state.attachments_dir else None
                per_recipient_base = None
                identifier_column = None
            
            result = efg.main(
                template_path=str(template_path),
                excel_path=str(excel_path),
                attachments_dir=attachments_dir,
                output_dir=str(st.session_state.output_dir),
                conditionals_path=conditionals_path,
                use_outlook=use_outlook and outlook_available,
                create_eml_backup=create_eml_backup,
                is_html_template=(st.session_state.template_mode == "rich"),
                attachment_mode=st.session_state.attachment_mode,
                per_recipient_base=per_recipient_base,
                identifier_column=identifier_column
            )
            
            progress_bar.progress(1.0)
            status_text.text("✅ Generation complete!")
            
            with results_container:
                # Check if result is a dictionary (new format) or boolean (old format)
                if isinstance(result, dict):
                    if result['success']:
                        st.success(f"Successfully generated emails for {result['success_count']} out of {result['total_count']} recipients")
                        
                        if result['error_count'] > 0:
                            st.warning(f"⚠️ {result['error_count']} emails failed to generate. Check console for details.")
                        
                        if result.get('outlook_drafts_created'):
                            st.info("📧 Outlook drafts have been created in your Drafts folder")
                        
                        if result.get('eml_files_created'):
                            st.info(f"📄 Backup .eml files saved to: {result.get('output_dir', st.session_state.output_dir)}")
                    else:
                        st.error(f"Failed to generate emails. {result.get('error_count', 0)} errors occurred.")
                else:
                    # Fallback for old boolean return format
                    if result:
                        st.success(f"Successfully generated emails for {len(st.session_state.excel_data)} recipients")
                    else:
                        st.error("Failed to generate emails")
                    
                    if use_outlook and outlook_available:
                        st.success(f"📧 Created {len(st.session_state.excel_data)} editable .msg files in: {st.session_state.output_dir}")
                    
                    if create_eml_backup:
                        st.info(f"📄 Backup .eml files saved to: {st.session_state.output_dir}")
            
            # Cleanup temp files
            if template_path.exists():
                template_path.unlink()
            if excel_path.exists():
                excel_path.unlink()
                
        except Exception as e:
            progress_bar.progress(0)
            status_text.text("❌ Generation failed")
            st.error(f"Error during generation: {str(e)}")
    
    # Navigation
    st.divider()
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        if st.button("◀️ Back: Conditional Content"):
            st.session_state.current_step = 5
            st.rerun()
    with col3:
        if st.button("🔄 Start Over"):
            # Reset session state
            for key in ['template_text', 'excel_data', 'template_variables', 'conditional_lines', 'attachments_dir']:
                if key in st.session_state:
                    del st.session_state[key]
            st.session_state.current_step = 1
            st.rerun()

if __name__ == "__main__":
    main()
