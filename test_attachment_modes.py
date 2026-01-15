#!/usr/bin/env python3
"""Test script to verify both Global and Per-Recipient attachment modes work correctly."""

import os
import pandas as pd
import json
from pathlib import Path
import shutil
import email_file_generator as efg

def setup_test_data():
    """Setup test data for testing attachment modes."""
    
    # Create test directories
    test_dir = Path("test_attachment_run")
    test_dir.mkdir(exist_ok=True)
    
    # Create global attachments directory
    global_attachments = test_dir / "global_attachments"
    global_attachments.mkdir(exist_ok=True)
    
    # Create sample global attachment files
    (global_attachments / "global_doc.txt").write_text("This is a global document")
    (global_attachments / "global_report.pdf").write_bytes(b"PDF content placeholder")
    
    # Create per-recipient attachments directories
    per_recipient_base = test_dir / "per_recipient_attachments"
    per_recipient_base.mkdir(exist_ok=True)
    
    # Create folders for specific recipients
    alice_folder = per_recipient_base / "Alice"
    alice_folder.mkdir(exist_ok=True)
    (alice_folder / "alice_invoice.txt").write_text("Invoice for Alice")
    (alice_folder / "alice_contract.pdf").write_bytes(b"Alice contract PDF")
    
    bob_folder = per_recipient_base / "Bob"
    bob_folder.mkdir(exist_ok=True)
    (bob_folder / "bob_proposal.txt").write_text("Proposal for Bob")
    
    # Create test Excel data
    test_data = pd.DataFrame({
        'FirstName': ['Alice', 'Bob', 'Charlie'],
        'LastName': ['Smith', 'Jones', 'Brown'],
        'Email': ['alice@example.com', 'bob@example.com', 'charlie@example.com'],
        'Company': ['Tech Corp', 'Sales Inc', 'Marketing Ltd'],
        'Subject': ['Meeting Request', 'Project Update', 'Partnership Proposal']
    })
    
    excel_path = test_dir / "test_recipients.xlsx"
    test_data.to_excel(excel_path, index=False)
    
    # Create test template
    template_text = """Dear [FirstName] [LastName],

I hope this email finds you well.

I'm reaching out to you from [Company] regarding the subject: [Subject].

This is a test email to verify attachment functionality.

Best regards,
Test Sender"""
    
    template_path = test_dir / "test_template.txt"
    template_path.write_text(template_text)
    
    # Create empty conditional lines file
    conditionals_path = test_dir / "conditional_lines.json"
    with open(conditionals_path, 'w') as f:
        json.dump({}, f)
    
    return {
        'test_dir': test_dir,
        'template_path': template_path,
        'excel_path': excel_path,
        'global_attachments': global_attachments,
        'per_recipient_base': per_recipient_base,
        'conditionals_path': conditionals_path
    }

def test_global_attachments(paths):
    """Test global attachment mode."""
    print("\n=== Testing Global Attachment Mode ===")
    
    output_dir = paths['test_dir'] / "output_global"
    output_dir.mkdir(exist_ok=True)
    
    result = efg.main(
        template_path=str(paths['template_path']),
        excel_path=str(paths['excel_path']),
        attachments_dir=str(paths['global_attachments']),
        output_dir=str(output_dir),
        conditionals_path=str(paths['conditionals_path']),
        use_outlook=False,  # Don't create Outlook drafts for testing
        create_eml_backup=True,
        is_html_template=False,
        attachment_mode="global",
        per_recipient_base=None,
        identifier_column=None
    )
    
    print("\nGlobal Mode Test Results:")
    if isinstance(result, dict):
        print(f"  Success: {result['success']}")
        print(f"  Successful emails: {result['success_count']}")
        print(f"  Failed emails: {result['error_count']}")
        print(f"  Total processed: {result['total_count']}")
        
        # Check that .eml files were created
        eml_files = list(output_dir.glob("*.eml"))
        print(f"  .eml files created: {len(eml_files)}")
        
        # Verify each .eml file has attachments
        for eml_file in eml_files[:1]:  # Check first file
            print(f"\n  Checking {eml_file.name}:")
            with open(eml_file, 'r', encoding='utf-8') as f:
                content = f.read()
                if "global_doc.txt" in content:
                    print("    ✓ Found global_doc.txt attachment")
                if "global_report.pdf" in content:
                    print("    ✓ Found global_report.pdf attachment")
    else:
        print(f"  Legacy format result: {result}")
    
    return result

def test_per_recipient_attachments(paths):
    """Test per-recipient attachment mode."""
    print("\n=== Testing Per-Recipient Attachment Mode ===")
    
    output_dir = paths['test_dir'] / "output_per_recipient"
    output_dir.mkdir(exist_ok=True)
    
    result = efg.main(
        template_path=str(paths['template_path']),
        excel_path=str(paths['excel_path']),
        attachments_dir=str(paths['global_attachments']),  # Fallback for recipients without folders
        output_dir=str(output_dir),
        conditionals_path=str(paths['conditionals_path']),
        use_outlook=False,  # Don't create Outlook drafts for testing
        create_eml_backup=True,
        is_html_template=False,
        attachment_mode="per_recipient",
        per_recipient_base=str(paths['per_recipient_base']),
        identifier_column="FirstName"
    )
    
    print("\nPer-Recipient Mode Test Results:")
    if isinstance(result, dict):
        print(f"  Success: {result['success']}")
        print(f"  Successful emails: {result['success_count']}")
        print(f"  Failed emails: {result['error_count']}")
        print(f"  Total processed: {result['total_count']}")
        
        # Check that .eml files were created
        eml_files = list(output_dir.glob("*.eml"))
        print(f"  .eml files created: {len(eml_files)}")
        
        # Verify Alice's email has her specific attachments
        alice_file = None
        for eml_file in eml_files:
            if "Alice" in eml_file.name:
                alice_file = eml_file
                break
        
        if alice_file:
            print(f"\n  Checking Alice's email ({alice_file.name}):")
            with open(alice_file, 'r', encoding='utf-8') as f:
                content = f.read()
                if "alice_invoice.txt" in content:
                    print("    ✓ Found alice_invoice.txt attachment")
                if "alice_contract.pdf" in content:
                    print("    ✓ Found alice_contract.pdf attachment")
        
        # Verify Bob's email has his specific attachments
        bob_file = None
        for eml_file in eml_files:
            if "Bob" in eml_file.name:
                bob_file = eml_file
                break
        
        if bob_file:
            print(f"\n  Checking Bob's email ({bob_file.name}):")
            with open(bob_file, 'r', encoding='utf-8') as f:
                content = f.read()
                if "bob_proposal.txt" in content:
                    print("    ✓ Found bob_proposal.txt attachment")
        
        # Verify Charlie's email (should use global attachments as fallback)
        charlie_file = None
        for eml_file in eml_files:
            if "Charlie" in eml_file.name:
                charlie_file = eml_file
                break
        
        if charlie_file:
            print(f"\n  Checking Charlie's email ({charlie_file.name}) - should have global attachments as fallback:")
            with open(charlie_file, 'r', encoding='utf-8') as f:
                content = f.read()
                if "global_doc.txt" in content or "global_report.pdf" in content:
                    print("    ✓ Found global attachments (fallback worked)")
                else:
                    print("    ⚠ No attachments found (Charlie has no per-recipient folder)")
    else:
        print(f"  Legacy format result: {result}")
    
    return result

def cleanup_test_data(test_dir):
    """Clean up test data after testing."""
    if test_dir.exists():
        shutil.rmtree(test_dir)
        print(f"\n✓ Cleaned up test directory: {test_dir}")

def main():
    """Run all attachment mode tests."""
    print("=" * 60)
    print("Testing Email Generator Attachment Modes")
    print("=" * 60)
    
    try:
        # Setup test data
        paths = setup_test_data()
        print(f"\n✓ Test data created in: {paths['test_dir']}")
        
        # Test global attachment mode
        global_result = test_global_attachments(paths)
        
        # Test per-recipient attachment mode
        per_recipient_result = test_per_recipient_attachments(paths)
        
        # Summary
        print("\n" + "=" * 60)
        print("Test Summary")
        print("=" * 60)
        
        global_success = isinstance(global_result, dict) and global_result['success'] or global_result == True
        per_recipient_success = isinstance(per_recipient_result, dict) and per_recipient_result['success'] or per_recipient_result == True
        
        if global_success:
            print("✅ Global attachment mode: PASSED")
        else:
            print("❌ Global attachment mode: FAILED")
        
        if per_recipient_success:
            print("✅ Per-recipient attachment mode: PASSED")
        else:
            print("❌ Per-recipient attachment mode: FAILED")
        
        # Clean up
        cleanup_test_data(paths['test_dir'])
        
        if global_success and per_recipient_success:
            print("\n🎉 All tests passed successfully!")
            return 0
        else:
            print("\n⚠️ Some tests failed. Please review the output above.")
            return 1
            
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    exit(main())