"""
Semantic Styling Module

Provides semantic styling tags for AI-generated PowerPoint content.
Instead of requiring hard-coded color values or font names, this module
allows AI models to use semantic tags like 'accent', 'critical', 'success', etc.
that are automatically mapped to the current template's theme colors.

Goal: Consistent, professional styling without AI models needing to know
specific color codes or font names.
"""
from typing import Optional, Dict, Any, List, Tuple, Union
from dataclasses import dataclass, field
from enum import Enum


class SemanticColorTag(str, Enum):
    """Semantic color tags for styling elements."""
    # Primary colors
    PRIMARY = "primary"
    SECONDARY = "secondary"
    ACCENT = "accent"
    
    # Status/feedback colors
    SUCCESS = "success"
    WARNING = "warning"
    CRITICAL = "critical"
    INFO = "info"
    
    # Neutral colors
    NEUTRAL = "neutral"
    NEUTRAL_LIGHT = "neutral_light"
    NEUTRAL_DARK = "neutral_dark"
    
    # Text colors
    TEXT = "text"
    TEXT_LIGHT = "text_light"
    TEXT_MUTED = "text_muted"
    TEXT_INVERTED = "text_inverted"
    
    # Background colors
    BACKGROUND = "background"
    BACKGROUND_ALT = "background_alt"
    
    # Emphasis colors
    HIGHLIGHT = "highlight"
    EMPHASIS = "emphasis"


class SemanticFontTag(str, Enum):
    """Semantic font tags for text styling."""
    TITLE = "title"
    HEADING = "heading"
    BODY = "body"
    CAPTION = "caption"
    CODE = "code"


@dataclass
class SemanticStyle:
    """Complete semantic style definition."""
    color: Optional[List[int]] = None
    font_name: Optional[str] = None
    font_size: Optional[int] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None


@dataclass
class SemanticTheme:
    """
    A complete semantic theme mapping.
    
    This defines how semantic tags map to actual values.
    """
    name: str = "default"
    
    # Color mappings (semantic tag -> RGB tuple)
    colors: Dict[str, Tuple[int, int, int]] = field(default_factory=dict)
    
    # Font mappings (semantic tag -> font properties)
    fonts: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    
    def get_color(self, tag: Union[str, SemanticColorTag]) -> Optional[Tuple[int, int, int]]:
        """Get RGB color for a semantic tag."""
        tag_str = tag.value if isinstance(tag, SemanticColorTag) else str(tag).lower()
        return self.colors.get(tag_str)
    
    def get_font(self, tag: Union[str, SemanticFontTag]) -> Optional[Dict[str, Any]]:
        """Get font properties for a semantic tag."""
        tag_str = tag.value if isinstance(tag, SemanticFontTag) else str(tag).lower()
        return self.fonts.get(tag_str)


def create_default_theme() -> SemanticTheme:
    """
    Create the default semantic theme.
    
    This theme provides sensible defaults that work well across
    different templates and produces professional-looking output.
    """
    return SemanticTheme(
        name="default",
        colors={
            # Primary colors
            SemanticColorTag.PRIMARY.value: (79, 129, 189),      # Blue
            SemanticColorTag.SECONDARY.value: (119, 147, 60),    # Green
            SemanticColorTag.ACCENT.value: (128, 100, 162),      # Purple
            
            # Status/feedback colors
            SemanticColorTag.SUCCESS.value: (0, 176, 80),        # Green
            SemanticColorTag.WARNING.value: (255, 192, 0),       # Yellow/Orange
            SemanticColorTag.CRITICAL.value: (192, 0, 0),        # Red
            SemanticColorTag.INFO.value: (0, 112, 192),          # Blue
            
            # Neutral colors
            SemanticColorTag.NEUTRAL.value: (127, 127, 127),     # Gray
            SemanticColorTag.NEUTRAL_LIGHT.value: (217, 217, 217),  # Light gray
            SemanticColorTag.NEUTRAL_DARK.value: (64, 64, 64),   # Dark gray
            
            # Text colors
            SemanticColorTag.TEXT.value: (0, 0, 0),              # Black
            SemanticColorTag.TEXT_LIGHT.value: (64, 64, 64),     # Dark gray
            SemanticColorTag.TEXT_MUTED.value: (127, 127, 127),  # Gray
            SemanticColorTag.TEXT_INVERTED.value: (255, 255, 255),  # White
            
            # Background colors
            SemanticColorTag.BACKGROUND.value: (255, 255, 255),  # White
            SemanticColorTag.BACKGROUND_ALT.value: (242, 242, 242),  # Light gray
            
            # Emphasis colors
            SemanticColorTag.HIGHLIGHT.value: (255, 255, 0),     # Yellow
            SemanticColorTag.EMPHASIS.value: (79, 129, 189),     # Blue (same as primary)
        },
        fonts={
            SemanticFontTag.TITLE.value: {
                "font_name": "Calibri Light",
                "font_size": 44,
                "bold": False,
            },
            SemanticFontTag.HEADING.value: {
                "font_name": "Calibri",
                "font_size": 28,
                "bold": True,
            },
            SemanticFontTag.BODY.value: {
                "font_name": "Calibri",
                "font_size": 18,
                "bold": False,
            },
            SemanticFontTag.CAPTION.value: {
                "font_name": "Calibri",
                "font_size": 12,
                "bold": False,
                "italic": True,
            },
            SemanticFontTag.CODE.value: {
                "font_name": "Consolas",
                "font_size": 14,
                "bold": False,
            },
        }
    )


class SemanticStyleResolver:
    """
    Resolves semantic styling tags to actual values.
    
    This class manages the mapping between semantic tags and actual
    styling values, taking into account the current template theme.
    
    Usage:
        resolver = SemanticStyleResolver()
        
        # Get color from semantic tag
        rgb = resolver.resolve_color("accent")  # Returns [128, 100, 162]
        
        # Get font settings from semantic tag
        font = resolver.resolve_font("heading")  # Returns font dict
        
        # Resolve any color input (semantic tag OR RGB)
        rgb = resolver.resolve_color_input("critical")  # Returns [192, 0, 0]
        rgb = resolver.resolve_color_input([255, 0, 0])  # Returns [255, 0, 0]
    """
    
    def __init__(self):
        self._default_theme = create_default_theme()
        self._template_theme: Optional[SemanticTheme] = None
    
    def update_from_template(self, template_colors: Dict[str, Any], template_fonts: Dict[str, Any]) -> None:
        """
        Update the semantic theme based on template colors and fonts.
        
        This maps the template's theme colors to semantic tags, allowing
        automatic adaptation to different templates.
        
        Args:
            template_colors: Color dict from template_manager.get_default_color_settings()
            template_fonts: Font dict from template_manager.get_default_font_settings()
        """
        theme_colors = dict(self._default_theme.colors)  # Start with defaults
        theme_fonts = dict(self._default_theme.fonts)    # Start with defaults
        
        # Map template colors to semantic colors
        # PowerPoint themes typically have accent_1 through accent_6
        if "accent_1" in template_colors:
            theme_colors[SemanticColorTag.PRIMARY.value] = template_colors["accent_1"]
            theme_colors[SemanticColorTag.ACCENT.value] = template_colors["accent_1"]
            theme_colors[SemanticColorTag.EMPHASIS.value] = template_colors["accent_1"]
        
        if "accent_2" in template_colors:
            theme_colors[SemanticColorTag.SECONDARY.value] = template_colors["accent_2"]
        
        if "accent_3" in template_colors:
            # Use accent_3 for success (often green in themes)
            theme_colors[SemanticColorTag.SUCCESS.value] = template_colors["accent_3"]
        
        if "accent_6" in template_colors:
            # Use accent_6 for critical/warning (often warm colors)
            theme_colors[SemanticColorTag.WARNING.value] = template_colors["accent_6"]
        
        # Map text colors
        if "text_1" in template_colors:
            theme_colors[SemanticColorTag.TEXT.value] = template_colors["text_1"]
            theme_colors[SemanticColorTag.TEXT_LIGHT.value] = template_colors["text_1"]
        
        if "text_2" in template_colors:
            theme_colors[SemanticColorTag.TEXT_MUTED.value] = template_colors["text_2"]
        
        # Map background colors
        if "background_1" in template_colors:
            theme_colors[SemanticColorTag.BACKGROUND.value] = template_colors["background_1"]
            theme_colors[SemanticColorTag.TEXT_INVERTED.value] = template_colors["background_1"]
        
        if "background_2" in template_colors:
            theme_colors[SemanticColorTag.BACKGROUND_ALT.value] = template_colors["background_2"]
        
        # Map template fonts to semantic fonts
        if "title_font_name" in template_fonts:
            theme_fonts[SemanticFontTag.TITLE.value]["font_name"] = template_fonts["title_font_name"]
        if "title_font_size" in template_fonts and template_fonts["title_font_size"]:
            theme_fonts[SemanticFontTag.TITLE.value]["font_size"] = int(template_fonts["title_font_size"])
        
        if "body_font_name" in template_fonts:
            theme_fonts[SemanticFontTag.BODY.value]["font_name"] = template_fonts["body_font_name"]
            theme_fonts[SemanticFontTag.HEADING.value]["font_name"] = template_fonts["body_font_name"]
            theme_fonts[SemanticFontTag.CAPTION.value]["font_name"] = template_fonts["body_font_name"]
        if "body_font_size" in template_fonts and template_fonts["body_font_size"]:
            theme_fonts[SemanticFontTag.BODY.value]["font_size"] = int(template_fonts["body_font_size"])
        
        self._template_theme = SemanticTheme(
            name="template",
            colors=theme_colors,
            fonts=theme_fonts
        )
    
    def clear_template(self) -> None:
        """Clear the template theme, reverting to defaults."""
        self._template_theme = None
    
    @property
    def current_theme(self) -> SemanticTheme:
        """Get the current active theme."""
        return self._template_theme or self._default_theme
    
    def resolve_color(self, tag: Union[str, SemanticColorTag]) -> Optional[List[int]]:
        """
        Resolve a semantic color tag to RGB values.
        
        Args:
            tag: Semantic color tag (e.g., "accent", "critical", SemanticColorTag.SUCCESS)
            
        Returns:
            RGB color as list [r, g, b] or None if tag is unknown
        """
        color = self.current_theme.get_color(tag)
        if color:
            return list(color)
        return None
    
    def resolve_font(self, tag: Union[str, SemanticFontTag]) -> Optional[Dict[str, Any]]:
        """
        Resolve a semantic font tag to font properties.
        
        Args:
            tag: Semantic font tag (e.g., "title", "body", SemanticFontTag.HEADING)
            
        Returns:
            Font properties dict or None if tag is unknown
        """
        return self.current_theme.get_font(tag)
    
    def resolve_color_input(self, color_input: Union[str, List[int], None]) -> Optional[List[int]]:
        """
        Resolve any color input to RGB values.
        
        This method accepts either:
        - A semantic tag string (e.g., "accent", "critical")
        - An RGB list (e.g., [255, 0, 0])
        - None (returns None)
        
        Args:
            color_input: Semantic tag, RGB list, or None
            
        Returns:
            RGB color as list [r, g, b] or None
        """
        if color_input is None:
            return None
        
        if isinstance(color_input, str):
            # Try to resolve as semantic tag
            resolved = self.resolve_color(color_input)
            if resolved:
                return resolved
            # If not a known tag, return None (invalid tag)
            return None
        
        if isinstance(color_input, (list, tuple)) and len(color_input) == 3:
            # Direct RGB values - validate and return
            try:
                return [int(c) for c in color_input]
            except (ValueError, TypeError):
                return None
        
        return None
    
    def resolve_font_input(
        self, 
        font_tag: Optional[str] = None,
        font_name: Optional[str] = None,
        font_size: Optional[int] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None
    ) -> Dict[str, Any]:
        """
        Resolve font input, combining semantic tag with overrides.
        
        Args:
            font_tag: Optional semantic font tag (e.g., "heading", "body")
            font_name: Optional explicit font name (overrides tag)
            font_size: Optional explicit font size (overrides tag)
            bold: Optional explicit bold setting (overrides tag)
            italic: Optional explicit italic setting (overrides tag)
            
        Returns:
            Font properties dict with resolved values
        """
        result: Dict[str, Any] = {}
        
        # Start with semantic tag defaults if provided
        if font_tag:
            tag_props = self.resolve_font(font_tag)
            if tag_props:
                result.update(tag_props)
        
        # Override with explicit values
        if font_name is not None:
            result["font_name"] = font_name
        if font_size is not None:
            result["font_size"] = font_size
        if bold is not None:
            result["bold"] = bold
        if italic is not None:
            result["italic"] = italic
        
        return result
    
    def get_available_color_tags(self) -> List[str]:
        """Get list of all available semantic color tags."""
        return [tag.value for tag in SemanticColorTag]
    
    def get_available_font_tags(self) -> List[str]:
        """Get list of all available semantic font tags."""
        return [tag.value for tag in SemanticFontTag]
    
    def get_color_palette(self) -> Dict[str, List[int]]:
        """
        Get the complete color palette for the current theme.
        
        Returns:
            Dict mapping semantic tags to RGB colors
        """
        return {
            tag: list(color) 
            for tag, color in self.current_theme.colors.items()
        }
    
    def get_font_styles(self) -> Dict[str, Dict[str, Any]]:
        """
        Get all font styles for the current theme.
        
        Returns:
            Dict mapping semantic font tags to font properties
        """
        return dict(self.current_theme.fonts)


# Global instance
style_resolver = SemanticStyleResolver()
