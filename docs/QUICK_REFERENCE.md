# AutoEML2 Quick Reference

Print this page for a handy reference while preparing your email campaign.

---

## Data Gathering Checklist

### Before You Start

Copy this checklist and fill it out:

```
Campaign Name: _______________________
Date: _______________________
Number of Recipients: _______________________

REQUIRED DATA
─────────────────────────────────────────────
[ ] Email addresses for all recipients
[ ] Subject line text
[ ] Email body content written

PERSONALIZATION FIELDS (check all that apply)
─────────────────────────────────────────────
[ ] First Name
[ ] Last Name
[ ] Full Name
[ ] Company Name
[ ] Job Title
[ ] Department
[ ] Custom Field 1: _____________
[ ] Custom Field 2: _____________
[ ] Custom Field 3: _____________

ADDITIONAL RECIPIENTS
─────────────────────────────────────────────
[ ] CC addresses needed?  Y / N
[ ] BCC addresses needed? Y / N

ATTACHMENTS
─────────────────────────────────────────────
[ ] No attachments needed
[ ] Same attachments for everyone (Global)
[ ] Different attachments per person (Per-Recipient)

Attachment files:
  1. _______________________
  2. _______________________
  3. _______________________

CONDITIONAL CONTENT
─────────────────────────────────────────────
[ ] No conditional content needed
[ ] Conditional content needed:

  Flag 1 Name: _____________
  When to include: _____________

  Flag 2 Name: _____________
  When to include: _____________
```

---

## Excel Column Reference

### Standard Column Names

| Column | Required? | Description |
|--------|-----------|-------------|
| `Email` | **Yes** | Recipient email |
| `Subject` | **Yes** | Email subject |
| `FirstName` | No | First name |
| `LastName` | No | Last name |
| `Company` | No | Company name |
| `Title` | No | Job title |
| `CC` | No | CC addresses |
| `BCC` | No | BCC addresses |
| `From` | No | Sender address |

### Custom Columns

Add any column you want! Just match the name in your template.

```
Excel Column:    MyCustomField
Template:        [MyCustomField]
```

---

## Template Syntax

### Variables

```
[ColumnName]  →  Replaced with Excel value
```

### Conditionals

```
[Conditional:FlagName]  →  Replaced with text if flag = 1
```

### Examples

```
Dear [FirstName],                    ← Simple variable
Meeting at [Company] headquarters    ← Variable in sentence
[Conditional:VIPCustomer]            ← Conditional block
```

---

## Common Patterns

### Professional Email

```
Dear [FirstName],

[Opening paragraph about purpose]

[Body content with [Company] and [Title] references]

[Conditional:SpecialOffer]

[Closing and call to action]

Best regards,
[Your Name]
[Your Title]
```

### Casual Email

```
Hi [FirstName],

[Friendly opening]

[Main content]

[Conditional:PersonalNote]

Cheers,
[Your Name]
```

---

## File Naming for Per-Recipient Attachments

### Folder Structure

```
base_folder/
├── {Identifier1}/
│   └── files...
├── {Identifier2}/
│   └── files...
└── {Identifier3}/
    └── files...
```

### Example with FirstName

```
attachments/
├── John/
│   └── proposal.pdf
├── Jane/
│   └── proposal.pdf
└── Bob/
    └── proposal.pdf
```

Identifier column in Excel: `FirstName`

---

## Troubleshooting Quick Fixes

| Problem | Quick Fix |
|---------|-----------|
| Variable not replaced | Check spelling & case match exactly |
| Email opens as received mail | Use Classic Outlook or drag to Drafts |
| Attachment not found | Verify folder name matches identifier |
| Empty email body | Check template isn't blank |
| Wrong recipient count | Check for empty rows in Excel |

---

## Keyboard Shortcuts (in App)

| Action | Keys |
|--------|------|
| Navigate steps | Click step numbers in sidebar |
| Preview next row | Change dropdown selection |
| Generate emails | Click primary button |

---

## Output Files

| File Type | Extension | Opens As |
|-----------|-----------|----------|
| Draft email | `.eml` | Editable draft in Outlook |
| All emails | `.zip` | ZIP containing all .eml files |

---

## Support

- Documentation: See README.md and USER_GUIDE.md
- Issues: https://github.com/kaantuncelCL/autoeml2/issues
