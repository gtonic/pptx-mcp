"""
Template Management Module

Handles PowerPoint template styles and theming functionality.
Now integrated with semantic styling for AI-friendly color and font selection.
"""
from typing import Optional, Dict, Any, Union, List
import os
import ppt_utils
from semantic_styles import style_resolver, SemanticStyleResolver


class TemplateManager:
    """Manages PowerPoint template styles and theming."""
    
    def __init__(self):
        # In-memory storage for template styles
        self.current_template_styles: Optional[Dict[str, Any]] = None
        self.current_template_path: Optional[str] = None
        self._style_resolver: SemanticStyleResolver = style_resolver
    
    def set_template_presentation(self, file_path: str) -> Dict[str, Any]:
        """Set a template presentation by file path and extract its styles."""
        if not os.path.exists(file_path):
            return {"error": f"Template file not found: {file_path}"}
        try:
            pres = ppt_utils.open_presentation(file_path)
            styles = ppt_utils.extract_template_styles(pres)
            self.current_template_styles = styles
            self.current_template_path = file_path
            
            # Update semantic style resolver with template colors and fonts
            self._style_resolver.update_from_template(
                styles.get("colors", {}),
                styles.get("fonts", {})
            )
            
            return {
                "message": f"Template set from {file_path}",
                "template_path": file_path,
                "styles": styles,
                "semantic_colors": self._style_resolver.get_color_palette(),
                "semantic_fonts": self._style_resolver.get_font_styles()
            }
        except Exception as e:
            return {"error": f"Failed to set template: {str(e)}"}
    
    def get_template_styles(self) -> Dict[str, Any]:
        """Get the currently loaded template styles."""
        if self.current_template_styles is None:
            return {
                "error": "No template styles loaded.",
                "semantic_colors": self._style_resolver.get_color_palette(),
                "semantic_fonts": self._style_resolver.get_font_styles()
            }
        return {
            "template_path": self.current_template_path,
            "styles": self.current_template_styles,
            "semantic_colors": self._style_resolver.get_color_palette(),
            "semantic_fonts": self._style_resolver.get_font_styles()
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
    
    def resolve_color(self, color_input: Union[str, List[int], None]) -> Optional[List[int]]:
        """
        Resolve a color input to RGB values.
        
        Accepts either:
        - A semantic tag string (e.g., "accent", "critical", "success")
        - An RGB list (e.g., [255, 0, 0])
        - None (returns None)
        
        Args:
            color_input: Semantic tag, RGB list, or None
            
        Returns:
            RGB color as list [r, g, b] or None
        """
        return self._style_resolver.resolve_color_input(color_input)
    
    def resolve_font(
        self,
        font_tag: Optional[str] = None,
        font_name: Optional[str] = None,
        font_size: Optional[int] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None
    ) -> Dict[str, Any]:
        """
        Resolve font settings from semantic tag and/or explicit values.
        
        Args:
            font_tag: Optional semantic font tag (e.g., "heading", "body", "title")
            font_name: Optional explicit font name (overrides tag)
            font_size: Optional explicit font size (overrides tag)
            bold: Optional explicit bold setting (overrides tag)
            italic: Optional explicit italic setting (overrides tag)
            
        Returns:
            Font properties dict with resolved values
        """
        return self._style_resolver.resolve_font_input(
            font_tag=font_tag,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic
        )
    
    def get_semantic_color_tags(self) -> List[str]:
        """Get list of all available semantic color tags."""
        return self._style_resolver.get_available_color_tags()
    
    def get_semantic_font_tags(self) -> List[str]:
        """Get list of all available semantic font tags."""
        return self._style_resolver.get_available_font_tags()
    
    def get_color_palette(self) -> Dict[str, List[int]]:
        """
        Get the complete color palette for the current theme.
        
        Returns:
            Dict mapping semantic tags to RGB colors
        """
        return self._style_resolver.get_color_palette()
    
    def get_font_styles(self) -> Dict[str, Dict[str, Any]]:
        """
        Get all font styles for the current theme.
        
        Returns:
            Dict mapping semantic font tags to font properties
        """
        return self._style_resolver.get_font_styles()


# Global instance
template_manager = TemplateManager()