import pandas as pd
from pathlib import Path

# Create sample data
data = {
    'Email': [
        'john.smith@example.com',
        'sarah.johnson@company.com',
        'michael.chen@business.org',
        'emma.wilson@enterprise.net',
        'david.brown@corporation.com'
    ],
    'FirstName': ['John', 'Sarah', 'Michael', 'Emma', 'David'],
    'Company': ['TechCorp', 'DataSystems', 'CloudWorks', 'InnovateLab', 'DigitalHub'],
    'Subject': [
        'Investment Opportunity - TechCorp',
        'Partnership Proposal - DataSystems',
        'Strategic Alliance - CloudWorks',
        'Business Opportunity - InnovateLab',
        'Collaboration Request - DigitalHub'
    ],
    'CC': ['manager@example.com', '', 'team@business.org', '', 'supervisor@corporation.com'],
    'SenderName': ['Brian', 'Brian', 'Brian', 'Brian', 'Brian'],
    'MeetingDate': ['2025-01-15', '2025-01-17', '2025-01-20', '2025-01-22', '2025-01-25'],
    'Discount': [10, 15, 20, 25, 30],
    'SpecialOffer': [1, 0, 1, 0, 1],
    'MeetingReminder': [0, 1, 1, 0, 1],
    'BrochureInfo': [1, 1, 0, 1, 0]
}

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
output_path = Path('sample_recipients.xlsx')
df.to_excel(output_path, index=False)

print(f"✅ Sample Excel file created: {output_path}")
print(f"📊 Contains {len(df)} sample recipients")
print("\nColumns:")
for col in df.columns:
    print(f"  - {col}")

# Also create a sample email template if needed
template_path = Path('sample_template.txt')
template_content = """Dear [FirstName],

I hope this message finds you well. I'm reaching out from our team regarding an exciting opportunity with [Company].

[Conditional:SpecialOffer]

[Conditional:MeetingReminder]

[Conditional:BrochureInfo]

We believe this could be a valuable partnership for both our organizations.

Please let me know if you'd like to schedule a call to discuss this further.

Best regards,
[SenderName]"""

with open(template_path, 'w') as f:
    f.write(template_content)

print(f"\n✅ Sample template created: {template_path}")