#!/usr/bin/env python3
"""
Test script to verify rich text editor features in the Email Generator
Tests HTML template handling, variable replacement, and formatting preservation
"""

import sys
import os
import json
import pandas as pd
from pathlib import Path
import tempfile

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import email_file_generator as efg
from template_manager import TemplateManager

def test_html_template_creation():
    """Test creating and processing HTML templates with rich formatting"""
    print("\n=== Testing HTML Template Creation ===")
    
    # Create a rich HTML template with formatting
    html_template = """
    <h2>Welcome [FirstName]!</h2>
    <p>We're excited to have you join <b>[Company]</b> as our new <i>[Position]</i>.</p>
    
    <p style="color: blue;">Here are your <u>onboarding details</u>:</p>
    <ul>
        <li>Start Date: <strong>[StartDate]</strong></li>
        <li>Department: <span style="color: green;">[Department]</span></li>
        <li>Manager: [Manager]</li>
    </ul>
    
    <p>Please visit our <a href="https://example.com">company portal</a> to complete your paperwork.</p>
    
    [Conditional:Premium]
    
    <p style="font-size: 18px;">We look forward to working with you!</p>
    <br>
    <p><em>Best regards,<br>
    The [Company] Team</em></p>
    """
    
    # Create test data
    test_data = pd.DataFrame({
        'FirstName': ['Alice', 'Bob', 'Charlie'],
        'Company': ['TechCorp', 'DataInc', 'WebSoft'],
        'Position': ['Developer', 'Analyst', 'Designer'],
        'StartDate': ['Jan 15', 'Jan 20', 'Feb 1'],
        'Department': ['Engineering', 'Analytics', 'Design'],
        'Manager': ['John Smith', 'Jane Doe', 'Mike Wilson'],
        'Email': ['alice@test.com', 'bob@test.com', 'charlie@test.com'],
        'Subject': ['Welcome to TechCorp', 'Welcome to DataInc', 'Welcome to WebSoft'],
        'Premium': [1, 0, 1]
    })
    
    # Save template and data to temp files
    with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False) as f:
        f.write(html_template)
        template_path = f.name
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        test_data.to_excel(f.name, index=False)
        excel_path = f.name
    
    # Create conditional lines
    conditional_lines = {
        'Premium': '<p style="background-color: gold; padding: 10px;"><b>Premium Package:</b> As a premium member, you get access to exclusive benefits including gym membership and enhanced healthcare!</p>'
    }
    
    conditionals_path = 'test_conditionals.json'
    with open(conditionals_path, 'w') as f:
        json.dump(conditional_lines, f)
    
    # Test with output directory
    output_dir = Path('test_rich_text_output')
    output_dir.mkdir(exist_ok=True)
    
    try:
        # Extract variables
        variables = efg.extract_variables(html_template)
        print(f"✅ Extracted variables: {variables}")
        
        # Process first row for testing
        row = test_data.iloc[0]
        
        # Test HTML template processing
        msg = efg.create_email_message(
            row=row,
            template_text=html_template,
            variables=variables,
            attachments_dir=None,
            attachment_columns=[],
            conditional_lines=conditional_lines,
            is_html_template=True
        )
        
        # Save the email
        email_file = output_dir / f"test_email_{row['FirstName']}.eml"
        with open(email_file, 'w', encoding='utf-8') as f:
            f.write(str(msg))
        
        print(f"✅ Created HTML email: {email_file}")
        
        # Verify the email contains HTML content
        email_content = str(msg)
        assert '<h2>Welcome Alice!</h2>' in email_content
        assert '<b>TechCorp</b>' in email_content
        assert 'Premium Package' in email_content  # Conditional content
        print("✅ HTML formatting preserved in email")
        
        # Test plain text template for comparison
        plain_template = """
Welcome [FirstName]!

We're excited to have you join [Company] as our new [Position].

Here are your onboarding details:
- Start Date: [StartDate]
- Department: [Department]
- Manager: [Manager]

[Conditional:Premium]

We look forward to working with you!

Best regards,
The [Company] Team
        """
        
        msg_plain = efg.create_email_message(
            row=row,
            template_text=plain_template,
            variables=variables,
            attachments_dir=None,
            attachment_columns=[],
            conditional_lines={'Premium': 'Premium Package: You get exclusive benefits!'},
            is_html_template=False
        )
        
        email_file_plain = output_dir / f"test_email_{row['FirstName']}_plain.eml"
        with open(email_file_plain, 'w', encoding='utf-8') as f:
            f.write(str(msg_plain))
        
        print(f"✅ Created plain text email: {email_file_plain}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error in HTML template test: {e}")
        return False
    finally:
        # Cleanup temp files
        if os.path.exists(template_path):
            os.remove(template_path)
        if os.path.exists(excel_path):
            os.remove(excel_path)
        if os.path.exists(conditionals_path):
            os.remove(conditionals_path)

def test_template_manager_with_html():
    """Test saving and loading HTML templates with the TemplateManager"""
    print("\n=== Testing Template Manager with HTML ===")
    
    tm = TemplateManager("test_templates")
    
    # Create test HTML template
    html_content = """
    <div style="font-family: Arial, sans-serif;">
        <h1 style="color: navy;">Hello [Name]!</h1>
        <p>This is a <b>rich text</b> template with <i>formatting</i>.</p>
        <ul>
            <li>Item 1: [Item1]</li>
            <li>Item 2: [Item2]</li>
        </ul>
    </div>
    """
    
    plain_content = "Hello [Name]!\nThis is a rich text template with formatting.\n- Item 1: [Item1]\n- Item 2: [Item2]"
    
    try:
        # Save HTML template
        result = tm.save_template(
            name="Test Rich Template",
            template_text=plain_content,
            description="A test template with HTML formatting",
            format_type="rich",
            template_html=html_content
        )
        
        assert result['success'], f"Failed to save template: {result.get('message')}"
        print("✅ Saved HTML template successfully")
        
        # List templates
        templates = tm.list_templates()
        rich_templates = [t for t in templates if t.get('format_type') == 'rich']
        assert len(rich_templates) > 0, "No rich text templates found"
        print(f"✅ Found {len(rich_templates)} rich text template(s)")
        
        # Load the template back
        test_template = rich_templates[0]
        result = tm.load_template(test_template['filename'])
        assert result['success'], f"Failed to load template: {result.get('message')}"
        
        data = result['data']
        assert data['format_type'] == 'rich', "Format type not preserved"
        assert 'template_html' in data, "HTML content not saved"
        assert '<h1' in data['template_html'], "HTML tags not preserved"
        print("✅ Loaded HTML template with formatting preserved")
        
        # Clean up
        tm.delete_template(test_template['filename'])
        print("✅ Cleaned up test template")
        
        return True
        
    except Exception as e:
        print(f"❌ Error in template manager test: {e}")
        return False

def test_variable_preservation_in_html():
    """Test that variable placeholders work correctly in HTML content"""
    print("\n=== Testing Variable Preservation in HTML ===")
    
    html_with_vars = """
    <p>Dear <strong>[CustomerName]</strong>,</p>
    <p>Your order #<span style="color: red;">[OrderNumber]</span> has been shipped!</p>
    <p>Tracking: <a href="[TrackingURL]">[TrackingNumber]</a></p>
    """
    
    test_row = pd.Series({
        'CustomerName': 'John Doe',
        'OrderNumber': '12345',
        'TrackingURL': 'https://track.example.com/12345',
        'TrackingNumber': 'TRK-12345',
        'Email': 'john@example.com',
        'Subject': 'Order Shipped'
    })
    
    variables = efg.extract_variables(html_with_vars)
    print(f"Variables found: {variables}")
    
    # Fill template
    filled = efg.fill_template(html_with_vars, test_row, variables)
    
    # Verify replacements
    assert 'John Doe' in filled
    assert '12345' in filled
    assert 'TRK-12345' in filled
    assert '[CustomerName]' not in filled
    
    print("✅ All variables correctly replaced in HTML template")
    return True

def main():
    """Run all tests"""
    print("\n" + "="*50)
    print("RICH TEXT EDITOR FEATURE TESTS")
    print("="*50)
    
    tests_passed = 0
    tests_total = 3
    
    # Run tests
    if test_html_template_creation():
        tests_passed += 1
    
    if test_template_manager_with_html():
        tests_passed += 1
    
    if test_variable_preservation_in_html():
        tests_passed += 1
    
    # Summary
    print("\n" + "="*50)
    print(f"TEST SUMMARY: {tests_passed}/{tests_total} tests passed")
    
    if tests_passed == tests_total:
        print("✅ ALL TESTS PASSED - Rich text features working correctly!")
    else:
        print(f"⚠️ {tests_total - tests_passed} test(s) failed")
    
    print("="*50)
    
    return tests_passed == tests_total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)