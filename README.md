# AutoEML2 - Email Draft Generator

A Streamlit-based tool for generating personalized, editable email drafts from templates and Excel data. Create bulk emails that open as editable drafts in Outlook, ready to review and send.

## Features

- **Template-based emails** with variable placeholders (`[FirstName]`, `[Company]`, etc.)
- **Excel integration** for recipient data management
- **Editable drafts** - Generated .eml files open in compose mode, not as received emails
- **ZIP download** - Bundle all generated emails for easy download
- **Rich text support** - Plain text or HTML email templates
- **Conditional content** - Include/exclude paragraphs based on Excel flags
- **Attachments** - Global attachments or per-recipient file folders
- **Cross-platform** - Works on Windows, Mac, and Linux

## Quick Start

### 1. Install Dependencies

```powershell
# Windows PowerShell
python -m venv venv
.\venv\Scripts\Activate
pip install streamlit pandas openpyxl streamlit-quill pywin32
```

```bash
# Mac/Linux
python3 -m venv venv
source venv/bin/activate
pip install streamlit pandas openpyxl streamlit-quill
```

### 2. Run the Application

```powershell
streamlit run app.py --server.address localhost --server.port 8501
```

Open http://localhost:8501 in your browser.

## How It Works

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│  Email Template │  +  │   Excel Data    │  =  │  Draft Emails   │
│  with [Variables]│     │  (Recipients)   │     │  (.eml files)   │
└─────────────────┘     └─────────────────┘     └─────────────────┘
```

1. Create a template with placeholders like `[FirstName]`
2. Upload Excel with columns matching your placeholders
3. Generate emails - each row becomes a personalized draft
4. Download ZIP and open .eml files in Outlook to review/send

---

## Detailed User Guide

### Step 1: Gather Your Information

Before using the tool, collect the following:

#### Required Information

| Item | Description | Example |
|------|-------------|---------|
| **Recipient emails** | Email addresses to send to | john@company.com |
| **Email subject** | Subject line (can include variables) | Meeting Request - [Company] |
| **Email body** | Your message content | See template section below |

#### Optional Information

| Item | Description | Example |
|------|-------------|---------|
| **Personalization fields** | Data that varies per recipient | FirstName, Company, Title |
| **CC/BCC addresses** | Additional recipients | manager@company.com |
| **Attachments** | Files to include | proposal.pdf, brochure.pdf |
| **Conditional content** | Paragraphs shown only for some recipients | Special offer text |

---

### Step 2: Prepare Your Excel File

Create an Excel file (.xlsx) with your recipient data.

#### Required Columns

| Column Name | Description |
|-------------|-------------|
| `Email` | Recipient's email address (required) |
| `Subject` | Email subject line |

#### Common Optional Columns

| Column Name | Description |
|-------------|-------------|
| `FirstName` | Recipient's first name |
| `LastName` | Recipient's last name |
| `Company` | Company name |
| `Title` | Job title |
| `CC` | CC email addresses (comma-separated) |
| `BCC` | BCC email addresses |

#### Example Excel Layout

| Email | Subject | FirstName | Company | Title |
|-------|---------|-----------|---------|-------|
| john@acme.com | Partnership Opportunity | John | Acme Corp | CEO |
| jane@globex.com | Partnership Opportunity | Jane | Globex Inc | Director |
| bob@initech.com | Partnership Opportunity | Bob | Initech | Manager |

---

### Step 3: Create Your Email Template

Write your email using **placeholders** in square brackets. Each placeholder must match a column name in your Excel file.

#### Template Syntax

```
[ColumnName]  →  Replaced with value from that Excel column
```

#### Example Template

```
Dear [FirstName],

I hope this message finds you well. I'm reaching out regarding a potential
partnership opportunity between our companies.

Given [Company]'s leadership in the industry, I believe there could be
significant synergies in working together.

As [Title] at [Company], you would be the ideal person to discuss this with.

Would you be available for a brief call next week?

Best regards,
Your Name
```

#### What Gets Generated

For the first row (John at Acme Corp), the email becomes:

```
Dear John,

I hope this message finds you well. I'm reaching out regarding a potential
partnership opportunity between our companies.

Given Acme Corp's leadership in the industry, I believe there could be
significant synergies in working together.

As CEO at Acme Corp, you would be the ideal person to discuss this with.

Would you be available for a brief call next week?

Best regards,
Your Name
```

---

### Step 4: Using Conditional Content (Optional)

Conditional content lets you include paragraphs only for certain recipients.

#### How It Works

1. Add a column in Excel with values `1` (include) or `0`/blank (exclude)
2. Add the conditional placeholder in your template: `[Conditional:ColumnName]`
3. Define the conditional text in the app

#### Example

**Excel:**

| Email | FirstName | SpecialOffer |
|-------|-----------|--------------|
| john@acme.com | John | 1 |
| jane@globex.com | Jane | 0 |

**Template:**
```
Dear [FirstName],

Thank you for your interest in our services.

[Conditional:SpecialOffer]

Please let me know if you have any questions.
```

**Conditional Text for "SpecialOffer":**
```
As a valued partner, we're pleased to offer you an exclusive 20% discount
on your first order. Use code PARTNER20 at checkout.
```

**Result:** John's email includes the discount paragraph; Jane's does not.

---

### Step 5: Adding Attachments (Optional)

#### Global Attachments (Same files for everyone)

1. Create a folder with your attachment files
2. In Step 4 of the app, select "Global Attachments"
3. Choose your attachments folder
4. All recipients receive the same files

#### Per-Recipient Attachments (Different files per person)

1. Create a base folder (e.g., `attachments/`)
2. Inside, create subfolders named after an identifier (e.g., recipient name):
   ```
   attachments/
   ├── John/
   │   ├── proposal_john.pdf
   │   └── quote_john.xlsx
   ├── Jane/
   │   ├── proposal_jane.pdf
   │   └── quote_jane.xlsx
   └── Bob/
       └── proposal_bob.pdf
   ```
3. In the app, select "Per-Recipient Attachments"
4. Choose the base folder and identifier column (e.g., `FirstName`)

---

### Step 6: Generate and Download

1. **Preview** - Review a sample email to verify placeholders are replaced correctly
2. **Generate** - Click "Generate All Emails"
3. **Download ZIP** - Click the download button to get all .eml files

---

## Using Generated Emails

### Opening in Outlook (Windows)

1. Extract the ZIP file
2. Double-click any `.eml` file
3. It opens as an **editable draft** (compose window)
4. Review, edit if needed, then click **Send**

### Compatibility Notes

| Email Client | Support Level |
|--------------|---------------|
| **Classic Outlook** (Windows) | Full support - opens as editable draft |
| **Thunderbird** | Full support - opens in compose mode |
| **New Outlook** (Windows 11) | Limited - may need to drag to Drafts folder first |
| **Outlook Web** | Not supported for direct editing |
| **Apple Mail** | Opens as draft, editing may be limited |

### If Emails Don't Open as Drafts

The `.eml` files include the `X-Unsent: 1` header which signals "draft mode." If your email client ignores this:

1. **Workaround for New Outlook:** Drag the .eml file into your Drafts folder, then open it from there
2. **Alternative:** Use Classic Outlook or Thunderbird

---

## Troubleshooting

### "Variable not found" warnings

- Ensure your Excel column names **exactly match** the placeholders (case-sensitive)
- `[FirstName]` requires a column named `FirstName`, not `firstname` or `First Name`

### Emails open as received mail, not drafts

- This is a limitation of "New Outlook" on Windows 11
- Use Classic Outlook or drag files to Drafts folder first

### Attachments not appearing

- Verify the attachment folder path is correct
- Check file permissions
- For per-recipient mode, ensure folder names match the identifier column values

### Special characters in names

- Avoid special characters in folder names for per-recipient attachments
- Characters like `< > : " / \ | ? *` are automatically replaced with underscores

---

## File Structure

```
autoeml2/
├── app.py                      # Main Streamlit application
├── email_file_generator.py     # Core email generation engine
├── template_manager.py         # Template save/load functionality
├── error_handler.py            # Error handling utilities
├── recovery_utils.py           # Session recovery
├── conditional_lines.json      # Conditional content definitions
├── templates/                  # Saved email templates
├── generated_emails/           # Output directory for .eml files
└── README.md                   # This file
```

---

## License

MIT License - See LICENSE file for details.
