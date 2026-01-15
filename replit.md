# Overview

An advanced email automation tool that generates personalized email files from templates and recipient data. The system uses Excel spreadsheets to store recipient information and supports both plain text and rich HTML templates with placeholder variables to create customized emails. It features a rich text editor with formatting options, conditional content insertion based on flags in the data, attachment handling, and outputs editable .msg files that can be opened and modified in Microsoft Outlook (with optional .eml backup files).

# User Preferences

Preferred communication style: Simple, everyday language.

# System Architecture

## Frontend Architecture
- **Streamlit Web Interface**: Multi-step workflow application with session state management
- **Rich Text Editor**: Integrated Quill editor for HTML email creation with formatting toolbar
- **Template Mode Toggle**: Switch between Plain Text and Rich Text editing modes
- **Progressive Workflow**: Step-by-step process for template creation, data upload, variable mapping, and email generation
- **Real-time Validation**: Immediate feedback on template variables and Excel column matching
- **Live HTML Preview**: Real-time preview of formatted emails in rich text mode
- **File Upload Handling**: Support for Excel files, text/HTML templates, and attachment directories

## Backend Architecture
- **Template Processing Engine**: Regex-based variable extraction and substitution system
- **Conditional Content System**: JSON-driven conditional line insertion based on Excel flags
- **Email Generation Module**: Creates .eml files with proper MIME formatting, HTML conversion, and attachment handling
- **Data Validation Layer**: Ensures required columns exist and validates email addresses
- **File Management**: Handles temporary files, output directories, and attachment processing

## Data Processing
- **Excel Data Handler**: Pandas-based spreadsheet processing with column validation
- **Template Variable System**: Bracket-notation placeholders `[VariableName]` for dynamic content insertion
- **Dual Template Format**: Support for both plain text and rich HTML templates
- **Rich Text Editor**: Quill-based WYSIWYG editor with formatting toolbar (bold, italic, underline, font sizes, colors, lists, links)
- **Format Preservation**: Maintains HTML formatting through template save/load cycle
- **Conditional Logic**: Flag-based content inclusion using JSON configuration files
- **Text Formatting**: Markdown-style formatting conversion to HTML (bold, italic, underline) for plain text mode

## Email Generation
- **MSG File Creation**: Creates editable .msg files using Outlook COM automation (Windows only)
- **MIME Message Creation**: Proper email structure with headers, body, and attachments for .eml backup files
- **Dual Format Support**: Handles both plain text templates with auto-HTML conversion and pre-formatted HTML templates
- **HTML Email Support**: Direct HTML template usage in rich text mode, automatic conversion in plain text mode
- **Format Detection**: Automatically determines template format and processes accordingly
- **Multi-recipient Support**: Handles To, CC, and BCC fields with comma-separated addresses
- **Attachment Processing**: Automatic file attachment based on Excel column references
- **Outlook .msg Files**: Primary output format - editable Outlook messages saved to disk (requires Windows with Outlook)
- **EML Backup Files**: Optional backup format - standard email files readable by most email clients

# External Dependencies

## Core Libraries
- **Streamlit**: Web application framework for the user interface
- **Streamlit-Quill**: Rich text editor component for HTML email creation
- **Pandas**: Excel file processing and data manipulation
- **Python Standard Library**: Email message creation (email.message), file operations (pathlib, os), regex processing, HTML handling

## Optional Dependencies
- **pywin32**: Windows COM integration for Outlook automation (optional, with graceful fallback)

## File System Requirements
- **Input Files**: Excel spreadsheets (.xlsx), text templates (.txt), conditional rules (JSON)
- **Output Directory**: Generated .msg files (and optional .eml backups) stored in configurable output folder
- **Attachments Directory**: Optional folder for files to be attached to emails

## Data Format Dependencies
- **Excel Structure**: Requires specific column names (Email/To, optional CC/BCC, From, Subject)
- **Template Format**: Bracket notation for variables `[VariableName]` works in both plain and HTML templates
- **HTML Templates**: Full HTML support with inline styles, links, lists, and formatting
- **Template Storage**: JSON format stores template text, HTML version, format type, and metadata
- **JSON Configuration**: Conditional content rules stored in JSON format