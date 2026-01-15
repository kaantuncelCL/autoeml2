#!/usr/bin/env python3
"""
Create a zip file of the email generator project for local testing.
Excludes runtime-generated directories and temporary files.
"""

import zipfile
import os
from pathlib import Path

def should_exclude(file_path):
    """Check if a file/directory should be excluded from the zip."""
    exclude_dirs = {
        'logs', 'backups', 'session_backups', 'generated_emails', 
        'email_queue', 'test_rich_text_output', '__pycache__', '.git'
    }
    exclude_files = {
        'uv.lock', '.gitignore', 'create_project_zip.py'
    }
    exclude_extensions = {'.pyc', '.pyo', '.log'}
    
    path = Path(file_path)
    
    # Check if any parent directory is in exclude_dirs
    for part in path.parts:
        if part in exclude_dirs:
            return True
    
    # Check if filename is in exclude_files
    if path.name in exclude_files:
        return True
    
    # Check if file extension should be excluded
    if path.suffix in exclude_extensions:
        return True
    
    return False

def create_project_zip():
    """Create a zip file of the project."""
    zip_filename = 'email_generator_project.zip'
    
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Walk through all files and directories
        for root, dirs, files in os.walk('.'):
            # Remove excluded directories from dirs list to prevent walking into them
            dirs[:] = [d for d in dirs if not should_exclude(os.path.join(root, d))]
            
            for file in files:
                file_path = os.path.join(root, file)
                
                # Skip excluded files
                if should_exclude(file_path):
                    continue
                
                # Add file to zip
                # Use relative path and remove leading './'
                arc_name = file_path[2:] if file_path.startswith('./') else file_path
                zipf.write(file_path, arc_name)
                print(f"Added: {arc_name}")
    
    # Get zip file size
    zip_size = os.path.getsize(zip_filename) / (1024 * 1024)  # MB
    print(f"\n✅ Created {zip_filename} ({zip_size:.1f} MB)")
    print(f"📦 Contains project files for local testing")
    print(f"\nTo use locally:")
    print(f"1. Extract the zip file")
    print(f"2. Install dependencies: pip install -r requirements.txt (or use pyproject.toml)")
    print(f"3. Run the app: streamlit run app.py --server.port 5000")

if __name__ == "__main__":
    create_project_zip()