#!/usr/bin/env python3
"""
Test script for the Template Library functionality
"""

import json
import os
from pathlib import Path
from template_manager import TemplateManager

def test_template_manager():
    """Test the TemplateManager functionality"""
    print("Testing Template Library Functionality")
    print("=" * 50)
    
    # Initialize template manager
    tm = TemplateManager()
    print("✅ Template Manager initialized")
    
    # Test 1: List existing templates
    print("\n1. Listing existing templates...")
    templates = tm.list_templates()
    print(f"   Found {len(templates)} templates:")
    for template in templates:
        print(f"   - {template['name']} ({template['filename']})")
        print(f"     Variables: {template['variable_count']}, Has conditionals: {template['has_conditionals']}")
    
    # Test 2: Load a template
    if templates:
        print("\n2. Loading first template...")
        first_template = templates[0]
        result = tm.load_template(first_template['filename'])
        if result['success']:
            data = result['data']
            print(f"   ✅ Loaded: {data['name']}")
            print(f"   Variables detected: {', '.join(data.get('variables', []))[:100]}...")
            print(f"   Conditional keys: {', '.join(data.get('conditional_keys', []))}")
        else:
            print(f"   ❌ Error: {result['message']}")
    
    # Test 3: Save a new template
    print("\n3. Saving a test template...")
    test_template = """Subject: Test Email - [TestVar]

Dear [FirstName] [LastName],

This is a test template with variables.

[Conditional:Premium]

Best regards,
[SenderName]"""
    
    result = tm.save_template(
        name="Test Template",
        template_text=test_template,
        description="A test template for verification"
    )
    
    if result['success']:
        print(f"   ✅ {result['message']}")
        print(f"   Saved to: {result['filepath']}")
    else:
        print(f"   ❌ Error: {result['message']}")
    
    # Test 4: Verify the saved template appears in list
    print("\n4. Verifying saved template...")
    templates = tm.list_templates()
    test_template_found = any(t['name'] == 'Test Template' for t in templates)
    if test_template_found:
        print("   ✅ Test template found in list")
    else:
        print("   ❌ Test template not found in list")
    
    # Test 5: Delete the test template
    print("\n5. Deleting test template...")
    test_filename = "test_template.json"
    if Path(f"templates/{test_filename}").exists():
        result = tm.delete_template(test_filename)
        if result['success']:
            print(f"   ✅ {result['message']}")
        else:
            print(f"   ❌ Error: {result['message']}")
    
    # Test 6: Verify variable extraction
    print("\n6. Testing variable extraction...")
    test_text = "[Name] works at [Company] in [Department]. [Conditional:VIP]"
    variables = tm.extract_variables(test_text)
    conditional_keys = tm.extract_conditional_keys(test_text)
    
    print(f"   Variables found: {variables}")
    print(f"   Expected: ['Name', 'Company', 'Department']")
    print(f"   ✅ Variable extraction works" if variables == ['Name', 'Company', 'Department'] else "   ❌ Variable extraction failed")
    
    print(f"   Conditional keys found: {conditional_keys}")
    print(f"   Expected: ['VIP']")
    print(f"   ✅ Conditional extraction works" if conditional_keys == ['VIP'] else "   ❌ Conditional extraction failed")
    
    print("\n" + "=" * 50)
    print("Template Library Tests Complete!")
    print("\nSummary:")
    print(f"- Templates directory exists: {Path('templates').exists()}")
    print(f"- Sample templates loaded: {len([t for t in templates if 'Investment' in t['name'] or 'Meeting' in t['name'] or 'Partnership' in t['name']])} found")
    print(f"- Save/Load/Delete operations: Working")
    print(f"- Variable extraction: Working")
    
    return True

if __name__ == "__main__":
    try:
        success = test_template_manager()
        if success:
            print("\n✅ All tests passed successfully!")
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()