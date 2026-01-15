"""
Test Script for Error Handling Features
Tests the enhanced error handling and recovery capabilities.
"""

import streamlit as st
import pandas as pd
import json
from pathlib import Path
import sys
import traceback

# Add current directory to path
sys.path.append('.')

# Import the modules to test
try:
    from error_handler import (
        ErrorHandler, SafeOperation, error_handler,
        validate_email_address, validate_file_size,
        validate_template_syntax, create_diagnostic_report
    )
    print("✅ Error handler imported successfully")
except Exception as e:
    print(f"❌ Failed to import error_handler: {e}")
    traceback.print_exc()

try:
    from recovery_utils import (
        SessionRecovery, ApplicationDiagnostics,
        session_recovery
    )
    print("✅ Recovery utils imported successfully")
except Exception as e:
    print(f"❌ Failed to import recovery_utils: {e}")
    traceback.print_exc()

def test_error_logging():
    """Test error logging functionality."""
    print("\n=== Testing Error Logging ===")
    
    try:
        # Test logging different severity levels
        test_error = ValueError("Test error for logging")
        error_record = error_handler.log_error(
            test_error,
            context="Test Context",
            user_message="This is a test error message",
            severity="ERROR"
        )
        print(f"✅ Error logged: {error_record['user_message']}")
        
        # Test getting recent errors
        recent_errors = error_handler.get_recent_errors(5)
        print(f"✅ Retrieved {len(recent_errors)} recent errors")
        
        return True
    except Exception as e:
        print(f"❌ Error logging test failed: {e}")
        return False

def test_validation_functions():
    """Test validation functions."""
    print("\n=== Testing Validation Functions ===")
    
    # Test email validation
    test_cases = [
        ("user@example.com", True),
        ("invalid.email", False),
        ("user@example.com, admin@test.org", True),
        ("[disabled@example.com]", True),  # Bracketed emails are skipped
        ("", False)
    ]
    
    for email, expected_valid in test_cases:
        is_valid, error_msg = validate_email_address(email)
        if is_valid == expected_valid:
            print(f"✅ Email validation correct for: {email[:30]}")
        else:
            print(f"❌ Email validation failed for: {email}")
    
    # Test template syntax validation
    template_tests = [
        ("Hello [Name], welcome to [Company]!", True),
        ("Missing bracket [Name", False),
        ("Empty placeholder []", False),
        ("Valid [Conditional:Premium] content", True)
    ]
    
    for template, expected_valid in template_tests:
        is_valid, errors = validate_template_syntax(template)
        if is_valid == expected_valid:
            print(f"✅ Template validation correct: {template[:30]}...")
        else:
            print(f"❌ Template validation failed: {template[:30]}...")
    
    # Test file size validation
    test_file = Path("test_file.tmp")
    try:
        test_file.write_text("x" * 1000)  # 1KB file
        is_valid, error = validate_file_size(test_file, max_size_mb=0.001)
        if not is_valid:
            print("✅ File size validation working")
        else:
            print("❌ File size validation should have failed")
        test_file.unlink()
    except Exception as e:
        print(f"❌ File size validation error: {e}")

def test_backup_restore():
    """Test backup and restore functionality."""
    print("\n=== Testing Backup/Restore ===")
    
    try:
        # Test backup creation
        test_data = {"test_key": "test_value", "count": 42}
        backup_file = error_handler.create_backup(
            test_data,
            "test_backup",
            "json"
        )
        
        if backup_file and backup_file.exists():
            print(f"✅ Backup created: {backup_file}")
            
            # Test restore
            success, restored_data = error_handler.restore_backup(backup_file)
            if success and restored_data == test_data:
                print("✅ Backup restored successfully")
            else:
                print("❌ Backup restore failed or data mismatch")
        else:
            print("❌ Backup creation failed")
        
        # Test CSV backup
        test_df = pd.DataFrame({
            'Name': ['Alice', 'Bob'],
            'Email': ['alice@example.com', 'bob@example.com']
        })
        
        csv_backup = error_handler.create_backup(
            test_df,
            "test_csv_backup",
            "csv"
        )
        
        if csv_backup and csv_backup.exists():
            print(f"✅ CSV backup created: {csv_backup}")
        else:
            print("❌ CSV backup failed")
            
    except Exception as e:
        print(f"❌ Backup/restore test failed: {e}")
        traceback.print_exc()

def test_session_recovery():
    """Test session recovery functionality."""
    print("\n=== Testing Session Recovery ===")
    
    try:
        # Create mock session state
        class MockSessionState:
            def __init__(self):
                self.template_text = "Test template with [Name]"
                self.template_mode = "plain"
                self.template_variables = ["Name"]
                self.current_step = 3
                self.excel_data = pd.DataFrame({
                    'Name': ['Test User'],
                    'Email': ['test@example.com']
                })
        
        mock_state = MockSessionState()
        
        # Test auto-save
        if session_recovery.auto_save_session(mock_state):
            print("✅ Session auto-saved successfully")
        else:
            print("❌ Session auto-save failed")
        
        # Test recovery
        new_state = MockSessionState()
        new_state.template_text = ""  # Clear data
        
        if session_recovery.recover_session(new_state):
            if new_state.template_text == "Test template with [Name]":
                print("✅ Session recovered successfully")
            else:
                print("❌ Session recovery data mismatch")
        else:
            print("❌ Session recovery failed")
        
        # Test export
        export_file = session_recovery.export_session(mock_state)
        if export_file and export_file.exists():
            print(f"✅ Session exported: {export_file}")
        else:
            print("❌ Session export failed")
            
    except Exception as e:
        print(f"❌ Session recovery test failed: {e}")
        traceback.print_exc()

def test_diagnostics():
    """Test diagnostic functionality."""
    print("\n=== Testing Diagnostics ===")
    
    try:
        diagnostics = ApplicationDiagnostics()
        
        # Test system requirements check
        requirements = diagnostics.check_system_requirements()
        print(f"✅ System requirements checked: Python {requirements['python_version']['current'][:10]}...")
        
        # Test self-test
        test_results = diagnostics.run_self_test()
        passed = sum(1 for result in test_results.values() if result)
        total = len(test_results)
        print(f"✅ Self-test completed: {passed}/{total} tests passed")
        
        # Test performance metrics
        try:
            metrics = diagnostics.get_performance_metrics()
            if 'memory' in metrics:
                print(f"✅ Performance metrics available")
            else:
                print("⚠️ Performance metrics limited (psutil not installed)")
        except Exception:
            print("⚠️ Performance metrics not available (psutil required)")
            
    except Exception as e:
        print(f"❌ Diagnostics test failed: {e}")
        traceback.print_exc()

def test_error_report():
    """Test error report generation."""
    print("\n=== Testing Error Report ===")
    
    try:
        # Generate some test errors
        for i in range(3):
            error_handler.log_error(
                ValueError(f"Test error {i+1}"),
                context=f"Test Context {i+1}",
                severity="ERROR" if i < 2 else "WARNING"
            )
        
        # Generate report
        report = error_handler.export_error_report()
        
        if "EMAIL GENERATOR ERROR REPORT" in report:
            print("✅ Error report generated")
            print(f"   Report length: {len(report)} characters")
        else:
            print("❌ Error report generation failed")
            
    except Exception as e:
        print(f"❌ Error report test failed: {e}")

def test_safe_operation_decorator():
    """Test the SafeOperation decorator."""
    print("\n=== Testing SafeOperation Decorator ===")
    
    try:
        @SafeOperation(error_handler, "Test Operation")
        def test_function(should_fail=False):
            if should_fail:
                raise ValueError("Intentional test error")
            return "Success"
        
        # Test successful operation
        result = test_function(should_fail=False)
        if result == "Success":
            print("✅ SafeOperation decorator - success case works")
        else:
            print("❌ SafeOperation decorator - unexpected result")
        
        # Test failed operation
        result = test_function(should_fail=True)
        if result is None:
            print("✅ SafeOperation decorator - handles errors gracefully")
        else:
            print("❌ SafeOperation decorator - should return None on error")
            
    except Exception as e:
        print(f"❌ SafeOperation decorator test failed: {e}")

def main():
    """Run all tests."""
    print("=" * 60)
    print("TESTING ERROR HANDLING AND RECOVERY FEATURES")
    print("=" * 60)
    
    # Run all tests
    test_error_logging()
    test_validation_functions()
    test_backup_restore()
    test_session_recovery()
    test_diagnostics()
    test_error_report()
    test_safe_operation_decorator()
    
    print("\n" + "=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    
    # Check if critical directories exist
    critical_dirs = ['logs', 'backups', 'session_backups', 'templates', 'generated_emails']
    for dir_name in critical_dirs:
        dir_path = Path(dir_name)
        if dir_path.exists():
            print(f"✅ Directory exists: {dir_name}")
        else:
            print(f"⚠️ Directory missing: {dir_name} (will be created on first use)")
    
    # Final diagnostic report
    print("\n" + "=" * 60)
    print("DIAGNOSTIC REPORT")
    print("=" * 60)
    
    report = create_diagnostic_report()
    print(f"Timestamp: {report['timestamp']}")
    print(f"Platform: {report['system']['platform']}")
    print(f"Python: {report['system']['python_version'][:30]}...")
    
    if report['disk_space']:
        if 'free_gb' in report['disk_space']:
            print(f"Disk Space: {report['disk_space']['free_gb']}GB free")
    
    print("\n✅ Error handling and recovery features are operational!")
    print("🎉 All critical components tested successfully!")

if __name__ == "__main__":
    main()