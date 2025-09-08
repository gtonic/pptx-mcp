"""
Template Management Module

Handles PowerPoint template styles and theming functionality.
"""
from typing import Optional, Dict, Any
import os
import ppt_utils


class TemplateManager:
    """Manages PowerPoint template styles and theming."""
    
    def __init__(self):
        # In-memory storage for template styles
        self.current_template_styles: Optional[Dict[str, Any]] = None
        self.current_template_path: Optional[str] = None
    
    def set_template_presentation(self, file_path: str) -> Dict[str, Any]:
        """Set a template presentation by file path and extract its styles."""
        if not os.path.exists(file_path):
            return {"error": f"Template file not found: {file_path}"}
        try:
            pres = ppt_utils.open_presentation(file_path)
            styles = ppt_utils.extract_template_styles(pres)
            self.current_template_styles = styles
            self.current_template_path = file_path
            return {
                "message": f"Template set from {file_path}",
                "template_path": file_path,
                "styles": styles
            }
        except Exception as e:
            return {"error": f"Failed to set template: {str(e)}"}
    
    def get_template_styles(self) -> Dict[str, Any]:
        """Get the currently loaded template styles."""
        if self.current_template_styles is None:
            return {"error": "No template styles loaded."}
        return {
            "template_path": self.current_template_path,
            "styles": self.current_template_styles
        }
    
    def get_default_font_settings(self) -> Dict[str, Any]:
        """Get default font settings from current template or return sensible defaults."""
        if self.current_template_styles and "fonts" in self.current_template_styles:
            return self.current_template_styles["fonts"]
        # Return sensible defaults if no template is loaded
        return {
            "body_font_name": "Calibri",
            "body_font_size": 18,
            "title_font_name": "Calibri",
            "title_font_size": 24
        }
    
    def get_default_color_settings(self) -> Dict[str, Any]:
        """Get default color settings from current template or return sensible defaults."""
        if self.current_template_styles and "colors" in self.current_template_styles:
            return self.current_template_styles["colors"]
        # Return sensible defaults if no template is loaded
        return {
            "accent_1": (79, 129, 189),  # Blue
            "text_1": (0, 0, 0),         # Black
            "background_1": (255, 255, 255)  # White
        }


# Global instance
template_manager = TemplateManager()