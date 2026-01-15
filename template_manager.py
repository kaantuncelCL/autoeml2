"""
Template Manager Module
Handles saving, loading, deleting, and listing email templates
"""

import json
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Any
import re

class TemplateManager:
    def __init__(self, templates_dir: str = "templates"):
        """Initialize the template manager with a templates directory."""
        self.templates_dir = Path(templates_dir)
        self.templates_dir.mkdir(exist_ok=True)
    
    def extract_variables(self, template_text: str) -> List[str]:
        """Extract all variable placeholders from template text."""
        # Extract regular variables [VarName]
        all_vars = re.findall(r"\[([^\]]+)\]", template_text)
        # Filter out conditional placeholders
        regular_vars = [var for var in all_vars if not var.startswith("Conditional:")]
        return regular_vars
    
    def extract_conditional_keys(self, template_text: str) -> List[str]:
        """Extract conditional placeholder keys from template text."""
        conditionals = re.findall(r"\[Conditional:([^\]]+)\]", template_text)
        return conditionals
    
    def save_template(self, name: str, template_text: str, description: str = "", format_type: str = "plain", template_html: Optional[str] = None) -> Dict[str, Any]:
        """
        Save a template to a JSON file.
        
        Args:
            name: Template name
            template_text: The template content
            description: Optional description of the template
            
        Returns:
            Dict with success status and message
        """
        try:
            # Create filename from name (sanitize for filesystem)
            filename = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in name)
            filename = filename.lower().replace(' ', '_') + '.json'
            filepath = self.templates_dir / filename
            
            # Check if template already exists
            is_overwrite = filepath.exists()
            
            # Extract variables and conditional keys
            variables = self.extract_variables(template_text)
            conditional_keys = self.extract_conditional_keys(template_text)
            
            # Create template data
            template_data = {
                "name": name,
                "description": description,
                "created_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "template_text": template_text,
                "variables": variables,
                "conditional_keys": conditional_keys,
                "format_type": format_type,
                "template_html": template_html  # Store HTML version if available
            }
            
            # If updating existing template, preserve original creation date
            if is_overwrite:
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        existing_data = json.load(f)
                        template_data["created_date"] = existing_data.get("created_date", template_data["created_date"])
                        template_data["modified_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                except:
                    pass  # If we can't read the existing file, just use new creation date
            
            # Save to file
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, indent=2, ensure_ascii=False)
            
            # Return appropriate message
            if is_overwrite:
                return {
                    "success": True,
                    "message": f"Template '{name}' updated successfully (overwrote existing template)",
                    "filepath": str(filepath),
                    "overwrite": True
                }
            else:
                return {
                    "success": True,
                    "message": f"Template '{name}' saved successfully",
                    "filepath": str(filepath),
                    "overwrite": False
                }
            
        except Exception as e:
            return {
                "success": False,
                "message": f"Error saving template: {str(e)}"
            }
    
    def load_template(self, filename: str) -> Dict[str, Any]:
        """
        Load a template from a JSON file.
        
        Args:
            filename: The template filename
            
        Returns:
            Dict with template data or error info
        """
        try:
            filepath = self.templates_dir / filename
            
            if not filepath.exists():
                return {
                    "success": False,
                    "message": f"Template file not found: {filename}"
                }
            
            with open(filepath, 'r', encoding='utf-8') as f:
                template_data = json.load(f)
            
            return {
                "success": True,
                "data": template_data
            }
            
        except json.JSONDecodeError:
            return {
                "success": False,
                "message": f"Error reading template: Invalid JSON format"
            }
        except Exception as e:
            return {
                "success": False,
                "message": f"Error loading template: {str(e)}"
            }
    
    def delete_template(self, filename: str) -> Dict[str, Any]:
        """
        Delete a template file.
        
        Args:
            filename: The template filename to delete
            
        Returns:
            Dict with success status and message
        """
        try:
            filepath = self.templates_dir / filename
            
            if not filepath.exists():
                return {
                    "success": False,
                    "message": f"Template file not found: {filename}"
                }
            
            # Load template to get its name for the message
            with open(filepath, 'r', encoding='utf-8') as f:
                template_data = json.load(f)
                template_name = template_data.get('name', filename)
            
            # Delete the file
            filepath.unlink()
            
            return {
                "success": True,
                "message": f"Template '{template_name}' deleted successfully"
            }
            
        except Exception as e:
            return {
                "success": False,
                "message": f"Error deleting template: {str(e)}"
            }
    
    def list_templates(self) -> List[Dict[str, Any]]:
        """
        List all available templates.
        
        Returns:
            List of template metadata dictionaries
        """
        templates = []
        
        try:
            # Get all JSON files in templates directory
            for filepath in sorted(self.templates_dir.glob("*.json")):
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        template_data = json.load(f)
                    
                    # Add filename to the data
                    template_data['filename'] = filepath.name
                    
                    # Create summary info
                    template_info = {
                        'filename': filepath.name,
                        'name': template_data.get('name', filepath.stem),
                        'description': template_data.get('description', ''),
                        'created_date': template_data.get('created_date', ''),
                        'variable_count': len(template_data.get('variables', [])),
                        'has_conditionals': len(template_data.get('conditional_keys', [])) > 0,
                        'format_type': template_data.get('format_type', 'plain')  # Default to plain for older templates
                    }
                    
                    templates.append(template_info)
                    
                except (json.JSONDecodeError, KeyError):
                    # Skip invalid template files
                    continue
                    
        except Exception as e:
            print(f"Error listing templates: {str(e)}")
        
        return templates
    
    def get_template_details(self, filename: str) -> Optional[Dict[str, Any]]:
        """
        Get detailed information about a specific template.
        
        Args:
            filename: The template filename
            
        Returns:
            Template data dictionary or None if not found
        """
        result = self.load_template(filename)
        if result['success']:
            return result['data']
        return None
    
    def export_template(self, filename: str, export_path: str) -> Dict[str, Any]:
        """
        Export a template to a different location.
        
        Args:
            filename: The template filename
            export_path: Path where to export the template
            
        Returns:
            Dict with success status and message
        """
        try:
            source = self.templates_dir / filename
            destination = Path(export_path)
            
            if not source.exists():
                return {
                    "success": False,
                    "message": f"Template file not found: {filename}"
                }
            
            # Copy the file
            with open(source, 'r', encoding='utf-8') as f:
                content = f.read()
            
            with open(destination, 'w', encoding='utf-8') as f:
                f.write(content)
            
            return {
                "success": True,
                "message": f"Template exported to {destination}"
            }
            
        except Exception as e:
            return {
                "success": False,
                "message": f"Error exporting template: {str(e)}"
            }
    
    def update_template(self, filename: str, template_text: str, description: Optional[str] = None, format_type: Optional[str] = None, template_html: Optional[str] = None) -> Dict[str, Any]:
        """
        Update an existing template.
        
        Args:
            filename: The template filename to update
            template_text: New template content
            description: Optional new description
            
        Returns:
            Dict with success status and message
        """
        try:
            filepath = self.templates_dir / filename
            
            if not filepath.exists():
                return {
                    "success": False,
                    "message": f"Template file not found: {filename}"
                }
            
            # Load existing template
            with open(filepath, 'r', encoding='utf-8') as f:
                template_data = json.load(f)
            
            # Update fields
            template_data['template_text'] = template_text
            template_data['variables'] = self.extract_variables(template_text)
            template_data['conditional_keys'] = self.extract_conditional_keys(template_text)
            template_data['modified_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            if description is not None:
                template_data['description'] = description
            
            if format_type is not None:
                template_data['format_type'] = format_type
            
            if template_html is not None:
                template_data['template_html'] = template_html
            
            # Save updated template
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, indent=2, ensure_ascii=False)
            
            return {
                "success": True,
                "message": f"Template '{template_data['name']}' updated successfully"
            }
            
        except Exception as e:
            return {
                "success": False,
                "message": f"Error updating template: {str(e)}"
            }