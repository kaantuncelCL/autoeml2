"""
Centralized Error Handling and Logging System
Provides robust error handling, logging, and recovery utilities for the Email Generator app.
"""

import logging
import traceback
import json
import os
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any, Tuple
import streamlit as st
from functools import wraps
import shutil
import pandas as pd


class ErrorHandler:
    """Centralized error handling and logging system."""
    
    def __init__(self, log_dir: str = "logs", max_log_size: int = 10 * 1024 * 1024):
        """
        Initialize the error handler.
        
        Args:
            log_dir: Directory for log files
            max_log_size: Maximum log file size before rotation (default 10MB)
        """
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        self.log_file = self.log_dir / "app.log"
        self.error_log_file = self.log_dir / "errors.json"
        self.max_log_size = max_log_size
        
        # Initialize logging
        self._setup_logging()
        
        # Initialize error history
        self.error_history = self._load_error_history()
        
        # Initialize backup directory
        self.backup_dir = Path("backups")
        self.backup_dir.mkdir(exist_ok=True)
    
    def _setup_logging(self):
        """Setup logging configuration."""
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Setup file handler with rotation
        if self.log_file.exists() and self.log_file.stat().st_size > self.max_log_size:
            # Rotate log file
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            rotated_file = self.log_dir / f"app_{timestamp}.log"
            self.log_file.rename(rotated_file)
            
            # Keep only last 5 rotated logs
            self._cleanup_old_logs()
        
        # Create file handler
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        
        # Setup logger
        self.logger = logging.getLogger('EmailGenerator')
        self.logger.setLevel(logging.DEBUG)
        self.logger.addHandler(file_handler)
        
        # Also log to console in debug mode
        if os.environ.get('DEBUG'):
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)
    
    def _cleanup_old_logs(self):
        """Keep only the 5 most recent rotated log files."""
        log_files = sorted(
            [f for f in self.log_dir.glob("app_*.log")],
            key=lambda x: x.stat().st_mtime,
            reverse=True
        )
        
        # Delete old log files
        for old_log in log_files[5:]:
            try:
                old_log.unlink()
            except Exception:
                pass
    
    def _load_error_history(self) -> List[Dict]:
        """Load error history from JSON file."""
        if self.error_log_file.exists():
            try:
                with open(self.error_log_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return []
        return []
    
    def _save_error_history(self):
        """Save error history to JSON file."""
        try:
            # Keep only last 100 errors
            self.error_history = self.error_history[-100:]
            
            with open(self.error_log_file, 'w', encoding='utf-8') as f:
                json.dump(self.error_history, f, indent=2)
        except Exception as e:
            self.logger.error(f"Failed to save error history: {str(e)}")
    
    def log_error(self, error: Exception, context: str = "", user_message: str = "", 
                  severity: str = "ERROR") -> Dict:
        """
        Log an error with context and user-friendly message.
        
        Args:
            error: The exception that occurred
            context: Context where the error occurred
            user_message: User-friendly error message
            severity: ERROR, WARNING, or CRITICAL
        
        Returns:
            Error record dictionary
        """
        error_record = {
            "timestamp": datetime.now().isoformat(),
            "severity": severity,
            "context": context,
            "error_type": type(error).__name__,
            "error_message": str(error),
            "user_message": user_message or self._get_user_friendly_message(error, context),
            "traceback": traceback.format_exc() if severity == "CRITICAL" else "",
            "recovery_suggestions": self._get_recovery_suggestions(error, context)
        }
        
        # Log to file
        self.logger.log(
            getattr(logging, severity),
            f"{context}: {error_record['error_message']}"
        )
        
        # Add to error history
        self.error_history.append(error_record)
        self._save_error_history()
        
        return error_record
    
    def _get_user_friendly_message(self, error: Exception, context: str) -> str:
        """Generate user-friendly error message based on error type and context."""
        error_messages = {
            "FileNotFoundError": "The specified file could not be found. Please check the file path.",
            "PermissionError": "Permission denied. Please check file permissions or try running as administrator.",
            "pd.errors.EmptyDataError": "The Excel file appears to be empty. Please check the file contents.",
            "KeyError": "Required data column is missing. Please check your Excel file structure.",
            "ValueError": "Invalid data format detected. Please check your input data.",
            "json.JSONDecodeError": "Invalid JSON format. Please check the configuration file.",
            "ConnectionError": "Connection failed. Please check your network connection.",
            "TimeoutError": "Operation timed out. Please try again.",
            "MemoryError": "Not enough memory to complete the operation. Try processing fewer items.",
            "OSError": "System error occurred. Please check disk space and file permissions."
        }
        
        # Context-specific messages
        if "Excel" in context or "excel" in context:
            if isinstance(error, (pd.errors.EmptyDataError, KeyError)):
                return "Excel file validation failed. Please ensure the file contains the required columns (Email, Subject) and is not empty."
            elif isinstance(error, FileNotFoundError):
                return "Excel file not found. Please upload a valid Excel file."
        
        if "Template" in context or "template" in context:
            if isinstance(error, ValueError):
                return "Template parsing error. Please check for matching brackets [] in your template."
            elif isinstance(error, FileNotFoundError):
                return "Template file not found. Please create or upload a template."
        
        if "Attachment" in context or "attachment" in context:
            if isinstance(error, FileNotFoundError):
                return "One or more attachment files could not be found. Please check the attachment directory."
            elif isinstance(error, PermissionError):
                return "Cannot access attachment files. Please check file permissions."
        
        if "Outlook" in context:
            return "Outlook integration error. Please ensure Outlook is installed and running."
        
        # Default messages by error type
        error_type = type(error).__name__
        return error_messages.get(error_type, f"An unexpected error occurred: {str(error)}")
    
    def _get_recovery_suggestions(self, error: Exception, context: str) -> List[str]:
        """Get recovery suggestions based on error type and context."""
        suggestions = []
        
        if isinstance(error, FileNotFoundError):
            suggestions.extend([
                "Verify the file path is correct",
                "Check if the file was moved or deleted",
                "Try uploading the file again"
            ])
        
        elif isinstance(error, PermissionError):
            suggestions.extend([
                "Check file permissions",
                "Close the file if it's open in another application",
                "Try saving to a different location"
            ])
        
        elif isinstance(error, (pd.errors.EmptyDataError, KeyError)):
            suggestions.extend([
                "Ensure the Excel file has the required columns",
                "Check that the file is not empty",
                "Verify column names match the template variables"
            ])
        
        elif isinstance(error, ValueError):
            if "template" in context.lower():
                suggestions.extend([
                    "Check for matching brackets [] in the template",
                    "Ensure all placeholders are properly closed",
                    "Try using the template validator"
                ])
            else:
                suggestions.extend([
                    "Check the data format",
                    "Ensure numeric fields contain valid numbers",
                    "Verify date formats are correct"
                ])
        
        elif isinstance(error, MemoryError):
            suggestions.extend([
                "Try processing fewer emails at once",
                "Close other applications to free up memory",
                "Restart the application"
            ])
        
        # Add general suggestions
        suggestions.extend([
            "Check the error log for more details",
            "Try restoring from a recent backup",
            "Reset the application if the issue persists"
        ])
        
        return suggestions[:5]  # Return top 5 suggestions
    
    def create_backup(self, data: Any, backup_name: str, backup_type: str = "json") -> Optional[Path]:
        """
        Create a backup of data.
        
        Args:
            data: Data to backup
            backup_name: Name for the backup
            backup_type: Type of backup (json, csv, text)
        
        Returns:
            Path to backup file or None if failed
        """
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_file = self.backup_dir / f"{backup_name}_{timestamp}.{backup_type}"
            
            if backup_type == "json":
                with open(backup_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2)
            elif backup_type == "csv" and isinstance(data, pd.DataFrame):
                data.to_csv(backup_file, index=False)
            elif backup_type == "text":
                with open(backup_file, 'w', encoding='utf-8') as f:
                    f.write(str(data))
            else:
                # Copy file
                if isinstance(data, (str, Path)):
                    shutil.copy2(data, backup_file)
            
            self.logger.info(f"Created backup: {backup_file}")
            
            # Clean up old backups (keep last 10)
            self._cleanup_old_backups(backup_name)
            
            return backup_file
            
        except Exception as e:
            self.logger.error(f"Failed to create backup: {str(e)}")
            return None
    
    def _cleanup_old_backups(self, backup_name: str):
        """Keep only the 10 most recent backups for a given name."""
        backup_files = sorted(
            [f for f in self.backup_dir.glob(f"{backup_name}_*")],
            key=lambda x: x.stat().st_mtime,
            reverse=True
        )
        
        for old_backup in backup_files[10:]:
            try:
                old_backup.unlink()
            except Exception:
                pass
    
    def restore_backup(self, backup_file: Path) -> Tuple[bool, Any]:
        """
        Restore data from a backup file.
        
        Args:
            backup_file: Path to backup file
        
        Returns:
            Tuple of (success, data or error message)
        """
        try:
            if not backup_file.exists():
                return False, "Backup file not found"
            
            ext = backup_file.suffix.lower()
            
            if ext == ".json":
                with open(backup_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            elif ext == ".csv":
                data = pd.read_csv(backup_file)
            elif ext in [".txt", ".html"]:
                with open(backup_file, 'r', encoding='utf-8') as f:
                    data = f.read()
            else:
                return False, f"Unsupported backup format: {ext}"
            
            self.logger.info(f"Restored backup from: {backup_file}")
            return True, data
            
        except Exception as e:
            error_msg = f"Failed to restore backup: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg
    
    def get_recent_errors(self, limit: int = 10) -> List[Dict]:
        """Get recent errors from history."""
        return self.error_history[-limit:]
    
    def clear_error_history(self):
        """Clear error history."""
        self.error_history = []
        self._save_error_history()
        self.logger.info("Error history cleared")
    
    def export_error_report(self) -> str:
        """Export error report as formatted text."""
        report = ["=" * 80]
        report.append("EMAIL GENERATOR ERROR REPORT")
        report.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append("=" * 80)
        report.append("")
        
        if not self.error_history:
            report.append("No errors recorded.")
        else:
            for i, error in enumerate(self.error_history[-20:], 1):
                report.append(f"Error #{i}")
                report.append("-" * 40)
                report.append(f"Time: {error['timestamp']}")
                report.append(f"Severity: {error['severity']}")
                report.append(f"Context: {error['context']}")
                report.append(f"Error: {error['error_message']}")
                report.append(f"User Message: {error['user_message']}")
                
                if error.get('recovery_suggestions'):
                    report.append("Recovery Suggestions:")
                    for suggestion in error['recovery_suggestions']:
                        report.append(f"  - {suggestion}")
                
                if error.get('traceback'):
                    report.append("Traceback:")
                    report.append(error['traceback'])
                
                report.append("")
        
        return "\n".join(report)


class SafeOperation:
    """Decorator for safe operation execution with error handling."""
    
    def __init__(self, error_handler: ErrorHandler, context: str, 
                 show_user_message: bool = True, create_backup: bool = False):
        """
        Initialize safe operation decorator.
        
        Args:
            error_handler: ErrorHandler instance
            context: Context for error logging
            show_user_message: Whether to show error message to user
            create_backup: Whether to create backup before operation
        """
        self.error_handler = error_handler
        self.context = context
        self.show_user_message = show_user_message
        self.create_backup = create_backup
    
    def __call__(self, func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                # Create backup if requested
                if self.create_backup and 'backup_data' in kwargs:
                    backup_data = kwargs.pop('backup_data')
                    backup_name = kwargs.pop('backup_name', self.context)
                    self.error_handler.create_backup(backup_data, backup_name)
                
                # Execute function
                result = func(*args, **kwargs)
                return result
                
            except Exception as e:
                # Log error
                error_record = self.error_handler.log_error(
                    e, 
                    context=self.context,
                    severity="ERROR" if not isinstance(e, (MemoryError, OSError)) else "CRITICAL"
                )
                
                # Show user message if in Streamlit context
                if self.show_user_message and 'streamlit' in str(type(st)):
                    st.error(f"❌ {error_record['user_message']}")
                    
                    # Show recovery suggestions
                    if error_record.get('recovery_suggestions'):
                        with st.expander("💡 Recovery Suggestions"):
                            for suggestion in error_record['recovery_suggestions']:
                                st.write(f"• {suggestion}")
                
                # Re-raise or return None based on severity
                if error_record['severity'] == "CRITICAL":
                    raise
                return None
        
        return wrapper


# Global error handler instance
error_handler = ErrorHandler()


def validate_email_address(email: str) -> Tuple[bool, str]:
    """
    Validate email address format.
    
    Args:
        email: Email address to validate
    
    Returns:
        Tuple of (is_valid, error_message)
    """
    import re
    
    if not email or not email.strip():
        return False, "Email address is empty"
    
    # Basic email regex pattern
    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    
    # Handle multiple emails (comma-separated)
    emails = [e.strip() for e in email.split(',')]
    invalid_emails = []
    
    for single_email in emails:
        # Skip bracketed emails (manual circuit breaker)
        if single_email.startswith('[') and single_email.endswith(']'):
            continue
        
        if not re.match(email_pattern, single_email):
            invalid_emails.append(single_email)
    
    if invalid_emails:
        return False, f"Invalid email format: {', '.join(invalid_emails)}"
    
    return True, ""


def validate_file_size(file_path: Path, max_size_mb: float = 25) -> Tuple[bool, str]:
    """
    Validate file size.
    
    Args:
        file_path: Path to file
        max_size_mb: Maximum allowed size in MB
    
    Returns:
        Tuple of (is_valid, error_message)
    """
    try:
        if not file_path.exists():
            return False, f"File not found: {file_path}"
        
        file_size_mb = file_path.stat().st_size / (1024 * 1024)
        
        if file_size_mb > max_size_mb:
            return False, f"File too large: {file_size_mb:.2f}MB (max: {max_size_mb}MB)"
        
        return True, ""
        
    except Exception as e:
        return False, f"Error checking file size: {str(e)}"


def validate_template_syntax(template_text: str) -> Tuple[bool, List[str]]:
    """
    Validate template syntax for proper bracket matching.
    
    Args:
        template_text: Template text to validate
    
    Returns:
        Tuple of (is_valid, list of errors)
    """
    errors = []
    
    # Check for unclosed brackets
    open_brackets = template_text.count('[')
    close_brackets = template_text.count(']')
    
    if open_brackets != close_brackets:
        errors.append(f"Mismatched brackets: {open_brackets} opening, {close_brackets} closing")
    
    # Check for empty placeholders
    if '[]' in template_text:
        errors.append("Empty placeholder [] found")
    
    # Check for nested placeholders
    import re
    placeholders = re.findall(r'\[([^\]]*\[.*?\].*?)\]', template_text)
    if placeholders:
        errors.append(f"Nested placeholders found: {placeholders}")
    
    # Check for special characters in placeholders
    all_placeholders = re.findall(r'\[([^\]]+)\]', template_text)
    for placeholder in all_placeholders:
        if not placeholder.replace(':', '').replace('_', '').replace(' ', '').isalnum():
            if not placeholder.startswith('Conditional:'):
                errors.append(f"Invalid characters in placeholder: [{placeholder}]")
    
    return len(errors) == 0, errors


def create_diagnostic_report() -> Dict[str, Any]:
    """Create a diagnostic report of the system state."""
    import platform
    import sys
    
    report = {
        "timestamp": datetime.now().isoformat(),
        "system": {
            "platform": platform.platform(),
            "python_version": sys.version,
            "streamlit_version": st.__version__ if 'st' in globals() else "Unknown"
        },
        "directories": {},
        "disk_space": {},
        "recent_errors": []
    }
    
    # Check directories
    for dir_name in ["logs", "backups", "templates", "generated_emails", "email_queue"]:
        dir_path = Path(dir_name)
        report["directories"][dir_name] = {
            "exists": dir_path.exists(),
            "writable": os.access(dir_path, os.W_OK) if dir_path.exists() else False
        }
    
    # Check disk space
    try:
        import shutil
        total, used, free = shutil.disk_usage("/")
        report["disk_space"] = {
            "total_gb": total // (2**30),
            "used_gb": used // (2**30),
            "free_gb": free // (2**30),
            "percentage_used": (used / total) * 100
        }
    except Exception:
        report["disk_space"] = {"error": "Could not determine disk space"}
    
    # Get recent errors
    if 'error_handler' in globals():
        report["recent_errors"] = error_handler.get_recent_errors(5)
    
    return report