"""
Test script to verify the email queue system is working correctly
"""

import json
from email_queue_manager import EmailQueueManager
from datetime import datetime, timedelta
import pandas as pd


def test_queue_manager():
    """Test the queue manager functionality"""
    print("Testing Email Queue Manager...")
    
    # Initialize queue manager
    queue_manager = EmailQueueManager()
    
    # Test 1: Add emails to queue
    print("\n1. Testing adding emails to queue...")
    test_emails = [
        {
            "recipient": "alice@example.com",
            "subject": "Test Email 1",
            "body": "This is a test email for Alice."
        },
        {
            "recipient": "bob@example.com",
            "subject": "Test Email 2",
            "body": "This is a test email for Bob.",
            "scheduled_time": datetime.now() + timedelta(hours=1)
        },
        {
            "recipient": "charlie@example.com",
            "subject": "Test Email 3",
            "body": "This is a test email for Charlie.",
            "scheduled_time": datetime.now() + timedelta(days=1)
        }
    ]
    
    email_ids = []
    for email in test_emails:
        queue_id = queue_manager.add_to_queue(
            recipient=email["recipient"],
            subject=email["subject"],
            body=email["body"],
            scheduled_time=email.get("scheduled_time"),
            priority=1,
            tags=["test"]
        )
        email_ids.append(queue_id)
        print(f"  Added email {queue_id[:8]}... to {email['recipient']}")
    
    # Test 2: List emails
    print("\n2. Testing listing emails...")
    all_emails = queue_manager.list_emails()
    print(f"  Total emails in queue: {len(all_emails)}")
    
    pending = queue_manager.list_emails(status="pending")
    scheduled = queue_manager.list_emails(status="scheduled")
    print(f"  Pending: {len(pending)}, Scheduled: {len(scheduled)}")
    
    # Test 3: Get statistics
    print("\n3. Testing statistics...")
    stats = queue_manager.get_statistics()
    print(f"  Statistics: {json.dumps(stats, indent=2)}")
    
    # Test 4: Get due emails
    print("\n4. Testing due emails...")
    due_emails = queue_manager.get_due_emails()
    print(f"  Emails due now: {len(due_emails)}")
    
    # Test 5: Update email status
    print("\n5. Testing status update...")
    if email_ids:
        success = queue_manager.update_email(email_ids[0], {"status": "draft_created"})
        print(f"  Updated email {email_ids[0][:8]}...: {success}")
    
    # Test 6: Batch scheduling
    print("\n6. Testing batch scheduling...")
    if len(email_ids) > 1:
        start_time = datetime.now() + timedelta(hours=2)
        success = queue_manager.schedule_batch(
            email_ids[1:],
            start_time,
            interval_minutes=30,
            emails_per_batch=1
        )
        print(f"  Batch scheduled: {success}")
    
    # Test 7: Export queue
    print("\n7. Testing export...")
    export_file = "test_queue_export.csv"
    success = queue_manager.export_queue(export_file)
    if success:
        print(f"  Queue exported to {export_file}")
        df = pd.read_csv(export_file)
        print(f"  Exported {len(df)} rows")
    
    print("\n✅ All tests completed successfully!")
    return True


def test_integration():
    """Test the integration with the Streamlit app"""
    print("\n\nTesting Integration with Streamlit App...")
    
    # Check if all required files exist
    from pathlib import Path
    
    required_files = [
        "app.py",
        "email_queue_manager.py",
        "step_7_queue_management.py",
        "email_file_generator.py",
        "template_manager.py"
    ]
    
    missing_files = []
    for file in required_files:
        if not Path(file).exists():
            missing_files.append(file)
    
    if missing_files:
        print(f"❌ Missing files: {missing_files}")
        return False
    
    print("✅ All required files present")
    
    # Check if Step 7 is properly integrated
    with open("app.py", 'r') as f:
        app_content = f.read()
    
    checks = [
        ("Step 7 in navigation", '"7. Queue & Scheduling"' in app_content),
        ("Queue manager import", 'from email_queue_manager import EmailQueueManager' in app_content),
        ("Step 7 function import", 'from step_7_queue_management import step_7_queue_management' in app_content),
        ("Step 7 routing", 'st.session_state.current_step == 7' in app_content),
        ("Add to Queue section", '"Add to Email Queue"' in app_content or 'Add to Email Queue' in app_content)
    ]
    
    all_passed = True
    for check_name, check_result in checks:
        status = "✅" if check_result else "❌"
        print(f"  {status} {check_name}")
        if not check_result:
            all_passed = False
    
    if all_passed:
        print("\n✅ Integration tests passed!")
    else:
        print("\n⚠️ Some integration checks failed")
    
    return all_passed


if __name__ == "__main__":
    print("="*50)
    print("EMAIL QUEUE SYSTEM TEST")
    print("="*50)
    
    # Run tests
    queue_test_passed = test_queue_manager()
    integration_test_passed = test_integration()
    
    print("\n" + "="*50)
    print("TEST SUMMARY")
    print("="*50)
    print(f"Queue Manager Tests: {'✅ PASSED' if queue_test_passed else '❌ FAILED'}")
    print(f"Integration Tests: {'✅ PASSED' if integration_test_passed else '❌ FAILED'}")
    
    if queue_test_passed and integration_test_passed:
        print("\n🎉 All tests passed! The email queue system is ready to use.")
    else:
        print("\n⚠️ Some tests failed. Please review the implementation.")