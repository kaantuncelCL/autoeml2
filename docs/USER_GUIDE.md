# AutoEML2 User Guide

A step-by-step guide to creating personalized email drafts.

---

## Table of Contents

1. [Pre-Flight Checklist](#pre-flight-checklist)
2. [Preparing Your Data](#preparing-your-data)
3. [Writing Your Template](#writing-your-template)
4. [Using the Application](#using-the-application)
5. [Advanced Features](#advanced-features)
6. [Examples](#examples)

---

## Pre-Flight Checklist

Before you begin, gather the following information:

### Required Items

- [ ] **Recipient list** - Who are you emailing?
- [ ] **Email addresses** - Collected in a spreadsheet
- [ ] **Subject line** - What's the email about?
- [ ] **Email body** - What do you want to say?

### Optional Items

- [ ] **Personalization data** - Names, companies, titles, etc.
- [ ] **CC/BCC recipients** - Anyone to copy?
- [ ] **Attachments** - Files to include?
- [ ] **Conditional content** - Different text for different recipients?

### Technical Requirements

- [ ] Python 3.8+ installed
- [ ] Excel file (.xlsx) with your data
- [ ] Outlook, Thunderbird, or compatible email client

---

## Preparing Your Data

### Excel File Structure

Your Excel file is the data source. Each **row** = one email. Each **column** = one piece of data.

#### Minimum Required Columns

```
┌─────────────────────┬──────────────────────────┐
│ Column Name         │ Purpose                  │
├─────────────────────┼──────────────────────────┤
│ Email               │ Recipient email address  │
│ Subject             │ Email subject line       │
└─────────────────────┴──────────────────────────┘
```

#### Recommended Columns

```
┌─────────────────────┬──────────────────────────┬─────────────────────┐
│ Column Name         │ Purpose                  │ Example             │
├─────────────────────┼──────────────────────────┼─────────────────────┤
│ Email               │ Primary recipient        │ john@company.com    │
│ Subject             │ Subject line             │ Meeting Request     │
│ FirstName           │ Recipient's first name   │ John                │
│ LastName            │ Recipient's last name    │ Smith               │
│ Company             │ Company name             │ Acme Corporation    │
│ Title               │ Job title                │ Vice President      │
│ CC                  │ CC recipients            │ assistant@co.com    │
│ BCC                 │ BCC recipients           │ records@myco.com    │
│ From                │ Sender address           │ me@mycompany.com    │
└─────────────────────┴──────────────────────────┴─────────────────────┘
```

### Column Naming Rules

1. **No spaces** - Use `FirstName` not `First Name`
2. **Case-sensitive** - `FirstName` ≠ `firstname` ≠ `FIRSTNAME`
3. **Letters and numbers only** - Avoid special characters
4. **Match your template** - Column names must exactly match placeholders

### Example Excel File

**File: recipients.xlsx**

| Email | Subject | FirstName | LastName | Company | Title | CC |
|-------|---------|-----------|----------|---------|-------|-----|
| john.smith@acme.com | Q1 Review Meeting | John | Smith | Acme Corp | CEO | assistant@acme.com |
| jane.doe@globex.com | Q1 Review Meeting | Jane | Doe | Globex Inc | CFO | |
| bob.jones@initech.com | Q1 Review Meeting | Bob | Jones | Initech | COO | admin@initech.com |

---

## Writing Your Template

### Basic Syntax

Use square brackets to mark placeholders:

```
[ColumnName]
```

The tool replaces each placeholder with the corresponding value from your Excel file.

### Template Structure

```
Dear [FirstName],

[Your email body here with [Company] and other [Placeholder] values]

Best regards,
[Your name]
```

### Example Templates

#### Simple Meeting Request

```
Dear [FirstName],

I hope this email finds you well.

I would like to schedule a meeting to discuss [Company]'s upcoming projects.
As [Title], your insights would be invaluable.

Would you have 30 minutes available next week?

Best regards,
Sarah Johnson
Director of Partnerships
```

#### Sales Follow-Up

```
Hi [FirstName],

Thank you for taking the time to speak with me last week about [Company]'s
needs in the enterprise software space.

Based on our conversation, I've put together a customized proposal that
addresses the specific challenges you mentioned.

Key highlights:
- 40% reduction in processing time
- Seamless integration with your existing systems
- Dedicated support team for [Company]

I'd love to walk you through the details. Are you available for a follow-up
call this Thursday or Friday?

Best,
Mike Chen
Account Executive
```

#### Event Invitation

```
Dear [FirstName],

You're invited!

[Company] is cordially invited to our Annual Industry Summit on March 15th.

As [Title] at [Company], we believe you would find our keynote session on
"Future of Technology" particularly relevant.

Event Details:
- Date: March 15, 2025
- Time: 9:00 AM - 4:00 PM
- Location: Grand Hotel Conference Center

Please RSVP by March 1st.

We hope to see you there!

Best regards,
Events Team
```

---

## Using the Application

### Step-by-Step Walkthrough

#### Step 1: Template Setup

1. Open the application (http://localhost:8501)
2. Choose template mode:
   - **Plain Text** - Simple text with basic formatting
   - **Rich Text** - HTML editor with bold, italic, lists, etc.
3. Enter or paste your template
4. Variables are automatically detected and highlighted

#### Step 2: Upload Excel Data

1. Click "Upload Excel File"
2. Select your .xlsx file
3. Preview the data to verify it loaded correctly
4. Check that column names are recognized

#### Step 3: Variable Mapping

1. Review detected variables from your template
2. Verify each variable matches an Excel column
3. Green checkmarks = matched, Red X = not found
4. Fix any mismatches by editing template or Excel

#### Step 4: Attachments (Optional)

**Global Attachments:**
- All recipients receive the same files
- Select a folder containing your attachments

**Per-Recipient Attachments:**
- Different files for each recipient
- Create folders named after an identifier (e.g., FirstName)
- Select the base folder and identifier column

#### Step 5: Conditional Content (Optional)

1. Add conditional placeholders: `[Conditional:FlagName]`
2. Define the text for each condition
3. In Excel, use `1` to include, `0` or blank to exclude

#### Step 6: Preview & Generate

1. Preview emails for different recipients
2. Verify personalization looks correct
3. Check "Create .eml Draft Files"
4. Check "Create ZIP bundle for download"
5. Click "Generate All Emails"
6. Download the ZIP file

---

## Advanced Features

### Conditional Content

Include different content for different recipients based on flags.

#### Setup

1. **Excel:** Add a column with `1` or `0` values

   | Email | FirstName | VIPCustomer |
   |-------|-----------|-------------|
   | john@co.com | John | 1 |
   | jane@co.com | Jane | 0 |

2. **Template:** Add conditional placeholder

   ```
   Dear [FirstName],

   Thank you for being a customer.

   [Conditional:VIPCustomer]

   Best regards
   ```

3. **App:** Define conditional text

   ```
   As a VIP customer, you qualify for our exclusive 25% discount.
   Use code VIP25 on your next order.
   ```

4. **Result:**
   - John's email includes the VIP paragraph
   - Jane's email does not

### Multiple Conditionals

You can use multiple conditional blocks:

```
Dear [FirstName],

[Conditional:NewCustomer]

[Conditional:RenewalDue]

[Conditional:SpecialOffer]

Best regards
```

### Per-Recipient Attachments

Send different files to each person.

#### Folder Structure

```
attachments/
├── John/
│   ├── proposal_john.pdf
│   └── pricing_john.xlsx
├── Jane/
│   ├── proposal_jane.pdf
│   └── contract_jane.pdf
└── Bob/
    └── proposal_bob.pdf
```

#### Excel Setup

| Email | FirstName | ... |
|-------|-----------|-----|
| john@co.com | John | ... |
| jane@co.com | Jane | ... |
| bob@co.com | Bob | ... |

#### App Configuration

- Attachment Mode: Per-Recipient
- Base Folder: `attachments/`
- Identifier Column: `FirstName`

The tool matches the `FirstName` value to folder names.

---

## Examples

### Example 1: Simple Newsletter

**Excel (newsletter_list.xlsx):**

| Email | FirstName |
|-------|-----------|
| subscriber1@email.com | Alex |
| subscriber2@email.com | Jordan |
| subscriber3@email.com | Taylor |

**Template:**

```
Hi [FirstName],

Here's your weekly newsletter!

This week's highlights:
- New product launch announcement
- Industry insights and trends
- Upcoming webinar schedule

Read more on our website.

Cheers,
The Newsletter Team
```

---

### Example 2: Invoice Reminders with Conditionals

**Excel (invoices.xlsx):**

| Email | FirstName | Company | InvoiceNumber | Amount | Overdue |
|-------|-----------|---------|---------------|--------|---------|
| ap@acme.com | John | Acme Corp | INV-001 | $5,000 | 0 |
| billing@globex.com | Jane | Globex | INV-002 | $3,500 | 1 |

**Template:**

```
Dear [FirstName],

This is a reminder regarding Invoice #[InvoiceNumber] for [Company].

Amount Due: [Amount]

[Conditional:Overdue]

Please process payment at your earliest convenience.

Best regards,
Accounts Receivable
```

**Conditional Text for "Overdue":**

```
IMPORTANT: This invoice is now past due. Please prioritize payment to
avoid any service interruptions. If you have already sent payment,
please disregard this notice.
```

---

### Example 3: Job Application Follow-ups with Attachments

**Excel (applications.xlsx):**

| Email | FirstName | Position | Company |
|-------|-----------|----------|---------|
| hr@techcorp.com | Sarah | Software Engineer | TechCorp |
| recruiting@startup.io | Mike | Product Manager | StartupIO |

**Attachment Folders:**

```
applications/
├── Sarah/
│   ├── resume_techcorp.pdf
│   └── portfolio.pdf
└── Mike/
    ├── resume_startup.pdf
    └── references.pdf
```

**Template:**

```
Dear [FirstName],

I am writing to follow up on my application for the [Position] role at [Company].

I remain very enthusiastic about the opportunity and believe my experience
aligns well with what you're looking for.

I have attached my updated resume and supporting documents for your review.

Thank you for your consideration. I look forward to hearing from you.

Best regards,
[Your Name]
```

---

## Tips & Best Practices

### Do's

- **Test with a small batch first** - Generate 2-3 emails to verify everything looks right
- **Preview multiple rows** - Check different recipients to catch edge cases
- **Use clear column names** - `FirstName` is better than `FN` or `col1`
- **Keep backups** - Save your Excel file and template before generating

### Don'ts

- **Don't use spaces in column names** - Use `FirstName` not `First Name`
- **Don't forget the Subject column** - Emails need subject lines
- **Don't skip the preview step** - Always verify before bulk generation
- **Don't ignore warnings** - Yellow warnings usually indicate missing data

### Common Mistakes

| Mistake | Solution |
|---------|----------|
| `[First Name]` not replaced | Remove space: `[FirstName]` |
| Some emails blank | Check for empty cells in Excel |
| Wrong attachment | Verify folder names match identifier values |
| Email opens as received | Use Classic Outlook or drag to Drafts |

---

## Getting Help

If you encounter issues:

1. Check the [Troubleshooting](#troubleshooting) section in the README
2. Review error messages in the application
3. Verify Excel column names match template placeholders exactly
4. Report issues at: https://github.com/kaantuncelCL/autoeml2/issues
