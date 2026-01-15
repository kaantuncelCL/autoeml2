"""
Recovery and Diagnostic Utilities for Email Generator
Provides auto-save, recovery, and diagnostic features.
"""

import streamlit as st
import json
import pickle
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Any, Optional
import pandas as pd
import shutil
from error_handler import error_handler


class SessionRecovery:
    """Handles session auto-save and recovery."""
    
    def __init__(self, save_dir: str = "session_backups"):
        """Initialize session recovery system."""
        self.save_dir = Path(save_dir)
        self.save_dir.mkdir(exist_ok=True)
        self.auto_save_file = self.save_dir / "auto_save.json"
        self.last_save_time = None
    
    def should_auto_save(self, interval_minutes: int = 5) -> bool:
        """Check if it's time to auto-save."""
        if self.last_save_time is None:
            return True
        
        elapsed = datetime.now() - self.last_save_time
        return elapsed > timedelta(minutes=interval_minutes)
    
    def auto_save_session(self, session_state: Any) -> bool:
        """Auto-save critical session data."""
        try:
            save_data = {
                'timestamp': datetime.now().isoformat(),
                'template_text': getattr(session_state, 'template_text', ''),
                'template_mode': getattr(session_state, 'template_mode', 'plain'),
                'template_html': getattr(session_state, 'template_html', ''),
                'template_variables': getattr(session_state, 'template_variables', []),
                'conditional_lines': getattr(session_state, 'conditional_lines', {}),
                'attachment_mode': getattr(session_state, 'attachment_mode', 'global'),
                'current_step': getattr(session_state, 'current_step', 1)
            }
            
            # Save Excel data separately if exists
            if hasattr(session_state, 'excel_data') and session_state.excel_data is not None:
                excel_file = self.save_dir / "auto_save_excel.csv"
                session_state.excel_data.to_csv(excel_file, index=False)
                save_data['has_excel_data'] = True
            else:
                save_data['has_excel_data'] = False
            
            # Write JSON data
            with open(self.auto_save_file, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, indent=2)
            
            self.last_save_time = datetime.now()
            return True
            
        except Exception as e:
            error_handler.log_error(e, "Auto-save Session")
            return False
    
    def recover_session(self, session_state: Any) -> bool:
        """Recover session from auto-save."""
        try:
            if not self.auto_save_file.exists():
                return False
            
            # Check if save is recent (within 24 hours)
            if self.auto_save_file.stat().st_mtime < (datetime.now().timestamp() - 86400):
                return False
            
            with open(self.auto_save_file, 'r', encoding='utf-8') as f:
                save_data = json.load(f)
            
            # Restore session data
            session_state.template_text = save_data.get('template_text', '')
            session_state.template_mode = save_data.get('template_mode', 'plain')
            session_state.template_html = save_data.get('template_html', '')
            session_state.template_variables = save_data.get('template_variables', [])
            session_state.conditional_lines = save_data.get('conditional_lines', {})
            session_state.attachment_mode = save_data.get('attachment_mode', 'global')
            session_state.current_step = save_data.get('current_step', 1)
            
            # Restore Excel data if exists
            if save_data.get('has_excel_data'):
                excel_file = self.save_dir / "auto_save_excel.csv"
                if excel_file.exists():
                    session_state.excel_data = pd.read_csv(excel_file)
            
            return True
            
        except Exception as e:
            error_handler.log_error(e, "Session Recovery")
            return False
    
    def clear_auto_save(self):
        """Clear auto-save files."""
        try:
            if self.auto_save_file.exists():
                self.auto_save_file.unlink()
            
            excel_file = self.save_dir / "auto_save_excel.csv"
            if excel_file.exists():
                excel_file.unlink()
            
            return True
        except Exception as e:
            error_handler.log_error(e, "Clear Auto-save")
            return False
    
    def export_session(self, session_state: Any) -> Optional[Path]:
        """Export current session to a file."""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            export_file = self.save_dir / f"session_export_{timestamp}.json"
            
            export_data = {
                'timestamp': datetime.now().isoformat(),
                'template_text': getattr(session_state, 'template_text', ''),
                'template_mode': getattr(session_state, 'template_mode', 'plain'),
                'template_html': getattr(session_state, 'template_html', ''),
                'template_variables': getattr(session_state, 'template_variables', []),
                'conditional_lines': getattr(session_state, 'conditional_lines', {}),
                'attachment_mode': getattr(session_state, 'attachment_mode', 'global'),
                'current_step': getattr(session_state, 'current_step', 1)
            }
            
            with open(export_file, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2)
            
            return export_file
            
        except Exception as e:
            error_handler.log_error(e, "Export Session")
            return None
    
    def import_session(self, file_path: Path, session_state: Any) -> bool:
        """Import session from a file."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                import_data = json.load(f)
            
            # Restore session data
            session_state.template_text = import_data.get('template_text', '')
            session_state.template_mode = import_data.get('template_mode', 'plain')
            session_state.template_html = import_data.get('template_html', '')
            session_state.template_variables = import_data.get('template_variables', [])
            session_state.conditional_lines = import_data.get('conditional_lines', {})
            session_state.attachment_mode = import_data.get('attachment_mode', 'global')
            session_state.current_step = import_data.get('current_step', 1)
            
            return True
            
        except Exception as e:
            error_handler.log_error(e, "Import Session")
            return False


class ApplicationDiagnostics:
    """Provides diagnostic tools for the application."""
    
    @staticmethod
    def check_system_requirements() -> Dict[str, Any]:
        """Check if system meets all requirements."""
        import platform
        import sys
        
        requirements = {
            'python_version': {
                'required': '3.7+',
                'current': sys.version,
                'met': sys.version_info >= (3, 7)
            },
            'platform': {
                'current': platform.platform(),
                'is_windows': platform.system() == 'Windows'
            },
            'modules': {}
        }
        
        # Check required modules
        required_modules = [
            'streamlit',
            'pandas',
            'openpyxl',
            'streamlit_quill'
        ]
        
        for module_name in required_modules:
            try:
                __import__(module_name)
                requirements['modules'][module_name] = {'installed': True}
            except ImportError:
                requirements['modules'][module_name] = {'installed': False}
        
        # Check optional modules
        optional_modules = ['win32com.client']
        for module_name in optional_modules:
            try:
                __import__(module_name)
                requirements['modules'][module_name] = {
                    'installed': True,
                    'optional': True
                }
            except ImportError:
                requirements['modules'][module_name] = {
                    'installed': False,
                    'optional': True
                }
        
        # Check disk space
        try:
            total, used, free = shutil.disk_usage("/")
            requirements['disk_space'] = {
                'free_gb': free // (2**30),
                'sufficient': free > (100 * 1024 * 1024)  # 100MB minimum
            }
        except Exception:
            requirements['disk_space'] = {'error': 'Could not check disk space'}
        
        # Check directory permissions
        directories = ['logs', 'backups', 'templates', 'generated_emails', 'email_queue']
        requirements['directories'] = {}
        
        for dir_name in directories:
            dir_path = Path(dir_name)
            try:
                dir_path.mkdir(exist_ok=True)
                test_file = dir_path / '.test'
                test_file.touch()
                test_file.unlink()
                requirements['directories'][dir_name] = {'writable': True}
            except Exception:
                requirements['directories'][dir_name] = {'writable': False}
        
        return requirements
    
    @staticmethod
    def run_self_test() -> Dict[str, bool]:
        """Run self-test on all major components."""
        tests = {}
        
        # Test template parsing
        try:
            from email_file_generator import extract_variables
            test_template = "Hello [Name], welcome to [Company]!"
            vars = extract_variables(test_template)
            tests['template_parsing'] = vars == ['Name', 'Company']
        except Exception:
            tests['template_parsing'] = False
        
        # Test Excel handling
        try:
            test_df = pd.DataFrame({
                'Email': ['test@example.com'],
                'Name': ['Test User']
            })
            tests['excel_handling'] = True
        except Exception:
            tests['excel_handling'] = False
        
        # Test file operations
        try:
            test_file = Path('test_file.tmp')
            test_file.write_text('test')
            content = test_file.read_text()
            test_file.unlink()
            tests['file_operations'] = content == 'test'
        except Exception:
            tests['file_operations'] = False
        
        # Test JSON operations
        try:
            test_data = {'test': 'data'}
            json_str = json.dumps(test_data)
            loaded = json.loads(json_str)
            tests['json_operations'] = loaded == test_data
        except Exception:
            tests['json_operations'] = False
        
        # Test error logging
        try:
            error_handler.logger.info("Test log entry")
            tests['error_logging'] = True
        except Exception:
            tests['error_logging'] = False
        
        return tests
    
    @staticmethod
    def get_performance_metrics() -> Dict[str, Any]:
        """Get current performance metrics."""
        import psutil
        import os
        
        metrics = {}
        
        # Memory usage
        try:
            process = psutil.Process(os.getpid())
            memory_info = process.memory_info()
            metrics['memory'] = {
                'rss_mb': memory_info.rss / (1024 * 1024),
                'percent': process.memory_percent()
            }
        except Exception:
            metrics['memory'] = {'error': 'Could not get memory info'}
        
        # CPU usage
        try:
            metrics['cpu'] = {
                'percent': psutil.cpu_percent(interval=1)
            }
        except Exception:
            metrics['cpu'] = {'error': 'Could not get CPU info'}
        
        # Session size estimate
        try:
            session_size = 0
            if hasattr(st.session_state, 'excel_data') and st.session_state.excel_data is not None:
                session_size += st.session_state.excel_data.memory_usage(deep=True).sum()
            metrics['session_size_mb'] = session_size / (1024 * 1024)
        except Exception:
            metrics['session_size_mb'] = 0
        
        return metrics


def display_error_dashboard():
    """Display error dashboard in Streamlit."""
    st.subheader("🔍 Error Dashboard")
    
    # Recent errors
    recent_errors = error_handler.get_recent_errors(5)
    
    if recent_errors:
        st.warning(f"Found {len(recent_errors)} recent errors")
        
        for i, error in enumerate(recent_errors, 1):
            with st.expander(f"Error {i}: {error['context']} - {error['timestamp'][:19]}"):
                st.write(f"**Severity:** {error['severity']}")
                st.write(f"**Type:** {error['error_type']}")
                st.write(f"**Message:** {error['user_message']}")
                
                if error.get('recovery_suggestions'):
                    st.write("**Recovery Suggestions:**")
                    for suggestion in error['recovery_suggestions']:
                        st.write(f"• {suggestion}")
                
                if st.button(f"Copy Error Details", key=f"copy_error_{i}"):
                    error_text = json.dumps(error, indent=2)
                    st.code(error_text, language='json')
    else:
        st.success("✅ No recent errors")
    
    # Actions
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Clear Error History", use_container_width=True):
            error_handler.clear_error_history()
            st.success("Error history cleared")
            st.rerun()
    
    with col2:
        if st.button("Export Error Report", use_container_width=True):
            report = error_handler.export_error_report()
            st.download_button(
                "Download Report",
                report,
                "error_report.txt",
                "text/plain",
                use_container_width=True
            )


def display_diagnostic_panel():
    """Display diagnostic panel in Streamlit."""
    st.subheader("🛠️ System Diagnostics")
    
    diagnostics = ApplicationDiagnostics()
    
    # System requirements check
    with st.expander("System Requirements"):
        requirements = diagnostics.check_system_requirements()
        
        # Python version
        python_ok = requirements['python_version']['met']
        if python_ok:
            st.success(f"✅ Python {requirements['python_version']['required']} or higher")
        else:
            st.error(f"❌ Python {requirements['python_version']['required']} required")
        
        # Modules
        st.write("**Required Modules:**")
        for module, info in requirements['modules'].items():
            if not info.get('optional', False):
                if info['installed']:
                    st.write(f"✅ {module}")
                else:
                    st.write(f"❌ {module} (not installed)")
        
        # Disk space
        if 'free_gb' in requirements.get('disk_space', {}):
            free_gb = requirements['disk_space']['free_gb']
            if requirements['disk_space']['sufficient']:
                st.success(f"✅ Disk space: {free_gb}GB free")
            else:
                st.error(f"❌ Low disk space: {free_gb}GB free")
    
    # Self-test
    with st.expander("Component Self-Test"):
        if st.button("Run Self-Test", use_container_width=True):
            with st.spinner("Running tests..."):
                test_results = diagnostics.run_self_test()
            
            for test_name, passed in test_results.items():
                if passed:
                    st.success(f"✅ {test_name.replace('_', ' ').title()}")
                else:
                    st.error(f"❌ {test_name.replace('_', ' ').title()}")
    
    # Performance metrics
    with st.expander("Performance Metrics"):
        try:
            metrics = diagnostics.get_performance_metrics()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                if 'memory' in metrics and 'rss_mb' in metrics['memory']:
                    st.metric("Memory Usage", f"{metrics['memory']['rss_mb']:.1f} MB")
            with col2:
                if 'cpu' in metrics and 'percent' in metrics['cpu']:
                    st.metric("CPU Usage", f"{metrics['cpu']['percent']}%")
            with col3:
                if 'session_size_mb' in metrics:
                    st.metric("Session Size", f"{metrics['session_size_mb']:.1f} MB")
        except ImportError:
            st.info("Install psutil for performance metrics: pip install psutil")


# Global session recovery instance
session_recovery = SessionRecovery()