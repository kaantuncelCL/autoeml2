#!/usr/bin/env python3
"""Test the template library fixes"""

import json
import os
from pathlib import Path
from template_manager import TemplateManager
import email_file_generator as efg

def test_template_manager():
    """Test the template manager functionality"""
    print("Testing Template Manager Fixes...")
    
    # Initialize template manager
    tm = TemplateManager("test_templates")
    
    # Test 1: Save a new template
    print("\n1. Testing save_template (new template)...")
    template_text = """Dear [FirstName] [LastName],

Thank you for your interest in our [Product] product.

[Conditional:Premium]

Best regards,
[SenderName]"""
    
    result = tm.save_template(
        name="Test Template",
        template_text=template_text,
        description="A test template for verification"
    )
    print(f"   Result: {result['message']}")
    print(f"   Overwrite: {result.get('overwrite', False)}")
    assert result['success'], "Failed to save new template"
    assert not result.get('overwrite', False), "Should not be an overwrite for new template"
    
    # Test 2: Save same template again (overwrite check)
    print("\n2. Testing save_template (overwrite existing)...")
    result = tm.save_template(
        name="Test Template",
        template_text=template_text + "\nUpdated!",
        description="Updated test template"
    )
    print(f"   Result: {result['message']}")
    print(f"   Overwrite: {result.get('overwrite', False)}")
    assert result['success'], "Failed to overwrite template"
    assert result.get('overwrite', False), "Should be an overwrite for existing template"
    
    # Test 3: Load template and verify variables are extracted correctly
    print("\n3. Testing load_template and variable extraction...")
    result = tm.load_template("test_template.json")
    assert result['success'], "Failed to load template"
    
    loaded_data = result['data']
    loaded_text = loaded_data['template_text']
    
    # Recompute variables from loaded text (as done in app.py)
    recomputed_vars = efg.extract_variables(loaded_text)
    stored_vars = loaded_data.get('variables', [])
    
    print(f"   Stored variables: {stored_vars}")
    print(f"   Recomputed variables: {recomputed_vars}")
    
    # Variables should be properly extracted (without Conditional: prefix)
    expected_vars = ['FirstName', 'LastName', 'Product', 'SenderName']
    for var in expected_vars:
        assert var in recomputed_vars, f"Variable {var} not found in recomputed variables"
    
    # Verify conditional keys are extracted
    conditional_keys = loaded_data.get('conditional_keys', [])
    print(f"   Conditional keys: {conditional_keys}")
    assert 'Premium' in conditional_keys, "Conditional key 'Premium' not found"
    
    # Test 4: Export template
    print("\n4. Testing export_template...")
    export_path = Path("test_templates") / "exported_template.json"
    result = tm.export_template("test_template.json", str(export_path))
    print(f"   Result: {result['message']}")
    assert result['success'], "Failed to export template"
    assert export_path.exists(), "Exported file does not exist"
    
    # Verify exported content
    with open(export_path, 'r') as f:
        exported_data = json.load(f)
    assert exported_data['name'] == "Test Template", "Exported template name mismatch"
    
    # Test 5: List templates
    print("\n5. Testing list_templates...")
    templates = tm.list_templates()
    print(f"   Found {len(templates)} template(s)")
    assert len(templates) > 0, "No templates found"
    
    for template in templates:
        print(f"   - {template['name']}: {template['variable_count']} vars, conditionals: {template['has_conditionals']}")
    
    # Test 6: Delete template
    print("\n6. Testing delete_template...")
    result = tm.delete_template("test_template.json")
    print(f"   Result: {result['message']}")
    assert result['success'], "Failed to delete template"
    
    # Cleanup
    print("\n7. Cleaning up test files...")
    if export_path.exists():
        export_path.unlink()
    
    # Remove test directory if empty
    test_dir = Path("test_templates")
    if test_dir.exists() and not any(test_dir.iterdir()):
        test_dir.rmdir()
    
    print("\n✅ All tests passed successfully!")
    return True

if __name__ == "__main__":
    try:
        test_template_manager()
    except AssertionError as e:
        print(f"\n❌ Test failed: {e}")
        exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        exit(1)