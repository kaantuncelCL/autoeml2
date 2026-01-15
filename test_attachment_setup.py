#!/usr/bin/env python3
"""
Test setup for the per-recipient attachment functionality.
Creates sample data and folder structure for testing.
"""

import os
import pandas as pd
from pathlib import Path

def create_test_setup():
    """Create test data and folder structure for testing per-recipient attachments."""
    
    # Create base directories
    base_dir = Path("test_attachments")
    base_dir.mkdir(exist_ok=True)
    
    # Create sample Excel data
    data = {
        'FirstName': ['Alice', 'Bob', 'Charlie', 'Diana'],
        'LastName': ['Johnson', 'Smith', 'Brown', 'Wilson'],
        'Email': ['alice@example.com', 'bob@example.com', 'charlie@example.com', 'diana@example.com'],
        'Company': ['TechCorp', 'DataInc', 'WebSolutions', 'CloudServices'],
        'Subject': ['Q1 Report', 'Product Update', 'Partnership Proposal', 'Service Agreement'],
        'PremiumMember': [1, 0, 1, 0],  # Conditional flag
    }
    
    df = pd.DataFrame(data)
    df.to_excel('test_recipients.xlsx', index=False)
    print("✅ Created test_recipients.xlsx")
    
    # Create per-recipient folders and sample files
    for firstname in data['FirstName']:
        recipient_dir = base_dir / firstname
        recipient_dir.mkdir(exist_ok=True)
        
        # Create sample files for each recipient
        # Alice gets contracts
        if firstname == 'Alice':
            with open(recipient_dir / 'Q1_Report_2025.pdf', 'w') as f:
                f.write(f"Sample Q1 Report for {firstname}")
            with open(recipient_dir / 'Financial_Summary.xlsx', 'w') as f:
                f.write(f"Financial data for {firstname}")
            print(f"  📁 Created folder: {recipient_dir} with 2 files")
        
        # Bob gets product specs
        elif firstname == 'Bob':
            with open(recipient_dir / 'Product_Specs_v2.pdf', 'w') as f:
                f.write(f"Product specifications for {firstname}")
            with open(recipient_dir / 'Release_Notes.txt', 'w') as f:
                f.write(f"Release notes for {firstname}")
            with open(recipient_dir / 'Demo_Video_Link.txt', 'w') as f:
                f.write(f"Demo video link for {firstname}")
            print(f"  📁 Created folder: {recipient_dir} with 3 files")
        
        # Charlie gets partnership docs
        elif firstname == 'Charlie':
            with open(recipient_dir / 'Partnership_Agreement.docx', 'w') as f:
                f.write(f"Partnership agreement for {firstname}")
            with open(recipient_dir / 'Terms_and_Conditions.pdf', 'w') as f:
                f.write(f"Terms for {firstname}")
            print(f"  📁 Created folder: {recipient_dir} with 2 files")
        
        # Diana's folder is intentionally left empty to test warning
        else:
            print(f"  📁 Created folder: {recipient_dir} (empty - for testing)")
    
    # Create global attachments folder
    global_dir = Path("global_attachments")
    global_dir.mkdir(exist_ok=True)
    
    # Add some global files
    with open(global_dir / 'Company_Brochure.pdf', 'w') as f:
        f.write("Company brochure - sent to all recipients")
    with open(global_dir / 'Newsletter_Jan_2025.pdf', 'w') as f:
        f.write("Monthly newsletter - sent to all recipients")
    print(f"✅ Created global_attachments with 2 files")
    
    # Create sample email template
    template = """Dear [FirstName] [LastName],

I hope this email finds you well. I'm reaching out from [Company] regarding: [Subject].

We have prepared some important documents for your review. Please find them attached to this email.

[Conditional:PremiumMember]

Best regards,
Your Sales Team
"""
    
    with open('test_template.txt', 'w') as f:
        f.write(template)
    print("✅ Created test_template.txt")
    
    # Create conditional lines
    conditional_lines = {
        "PremiumMember": "As a valued Premium Member, you have access to exclusive benefits including priority support and advanced features."
    }
    
    import json
    with open('test_conditional_lines.json', 'w') as f:
        json.dump(conditional_lines, f, indent=2)
    print("✅ Created test_conditional_lines.json")
    
    print("\n" + "="*50)
    print("TEST SETUP COMPLETE!")
    print("="*50)
    print("\nTest the implementation with:")
    print("1. Template: test_template.txt")
    print("2. Excel data: test_recipients.xlsx")
    print("3. Per-recipient attachments: test_attachments/")
    print("4. Global attachments: global_attachments/")
    print("5. Conditional lines: test_conditional_lines.json")
    print("\nFolder structure created:")
    print("  test_attachments/Alice/ (2 files)")
    print("  test_attachments/Bob/ (3 files)")
    print("  test_attachments/Charlie/ (2 files)")
    print("  test_attachments/Diana/ (empty - for testing warnings)")
    print("  global_attachments/ (2 files)")
    print("\nYou can now:")
    print("1. Run the Streamlit app")
    print("2. Load the test template")
    print("3. Upload test_recipients.xlsx")
    print("4. Switch between Global and Per-Recipient attachment modes")
    print("5. Select 'FirstName' as the identifier column in Per-Recipient mode")
    print("6. Preview and generate emails to test the functionality")

if __name__ == "__main__":
    create_test_setup()