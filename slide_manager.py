"""
Slide Management Module

Handles slide creation, manipulation, and content management.
"""
from typing import Optional, Dict, Any, List
import os
import ppt_utils
from presentation_manager import presentation_manager
from template_manager import template_manager


class SlideManager:
    """Manages slide operations within presentations."""
    
    def add_slide(self, layout_index: int = 1, title: Optional[str] = None, presentation_id: Optional[str] = None) -> Dict[str, Any]:
        """Add a new slide to the presentation."""
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            if not (0 <= layout_index < len(pres.slide_layouts)):
                return {
                    "error": f"Invalid layout index: {layout_index}. Available: 0-{len(pres.slide_layouts) - 1}",
                    "available_layouts": {i: l.name for i, l in enumerate(pres.slide_layouts)}
                }
            slide, info = ppt_utils.add_slide(pres, layout_index, title)
            return {
                "message": f"Added slide with layout '{info['layout_name']}'",
                "slide_index": len(pres.slides) - 1,
                **info
            }
        except (ValueError, KeyError) as e:
            return {"error": str(e)}
    
    def add_textbox(
        self,
        slide_index: int,
        left: float,
        top: float,
        width: float,
        height: float,
        text: str,
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[List[int]] = None,
        alignment: Optional[str] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add a textbox to a slide, using template styles as defaults if available."""
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Invalid slide index: {slide_index}"}
            slide = pres.slides[slide_index]
            
            # Use template styles as defaults if not provided
            font_defaults = template_manager.get_default_font_settings()
            color_defaults = template_manager.get_default_color_settings()
            
            if font_name is None:
                font_name = font_defaults.get("body_font_name")
            if font_size is None:
                font_size = font_defaults.get("body_font_size")
            if color is None:
                accent = color_defaults.get("accent_1")
                if accent:
                    color = list(accent)
            
            ppt_utils.add_textbox(
                slide, left, top, width, height, text,
                font_size=font_size, font_name=font_name, bold=bold,
                italic=italic, color=color, alignment=alignment
            )
            return {"message": f"Added textbox to slide {slide_index}", "shape_index": len(slide.shapes) - 1}
        except (ValueError, KeyError) as e:
            return {"error": str(e)}
    
    def add_shape(
        self,
        slide_index: int,
        shape_type: str,
        left: float,
        top: float,
        width: float,
        height: float,
        fill_color: Optional[List[int]] = None,
        line_color: Optional[List[int]] = None,
        line_width: Optional[float] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add an auto shape to a slide, using template styles as defaults if available."""
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Invalid slide index: {slide_index}"}
            slide = pres.slides[slide_index]
            
            # Use template styles as defaults if not provided
            color_defaults = template_manager.get_default_color_settings()
            
            if fill_color is None:
                accent = color_defaults.get("accent_1")
                if accent:
                    fill_color = list(accent)
            if line_color is None:
                text1 = color_defaults.get("text_1")
                if text1:
                    line_color = list(text1)
            
            ppt_utils.add_shape(
                slide, shape_type, left, top, width, height,
                fill_color=fill_color, line_color=line_color, line_width=line_width
            )
            return {"message": f"Added {shape_type} shape to slide {slide_index}", "shape_index": len(slide.shapes) - 1}
        except (ValueError, KeyError) as e:
            return {"error": str(e)}
    
    def add_line(
        self,
        slide_index: int,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        line_color: Optional[List[int]] = None,
        line_width: Optional[float] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add a straight line to a slide."""
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Invalid slide index: {slide_index}"}
            slide = pres.slides[slide_index]
            
            # Use template default colors if not provided
            if line_color is None:
                color_defaults = template_manager.get_default_color_settings()
                text1 = color_defaults.get("text_1")
                if text1:
                    line_color = list(text1)
            
            ppt_utils.add_line(
                slide, x1, y1, x2, y2,
                line_color=line_color, line_width=line_width
            )
            return {
                "message": f"Added line to slide {slide_index}",
                "shape_index": len(slide.shapes) - 1
            }
        except (ValueError, KeyError) as e:
            return {"error": str(e)}


    def add_chart(
        self,
        slide_index: int,
        chart_type: str,
        left: float,
        top: float,
        width: float,
        height: float,
        data: Dict[str, Any],
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add a chart to a slide."""
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Invalid slide index: {slide_index}"}
            slide = pres.slides[slide_index]
            
            chart = ppt_utils.add_chart(slide, chart_type, left, top, width, height, data)
            return {
                "message": f"Added {chart_type} chart to slide {slide_index}",
                "shape_index": len(slide.shapes) - 1
            }
        except (ValueError, KeyError) as e:
            return {"error": str(e)}
    
    def add_table(
        self,
        slide_index: int,
        left: float,
        top: float,
        rows: int,
        cols: int,
        data: Optional[List[List[str]]] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add a table to a slide."""
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Invalid slide index: {slide_index}"}
            slide = pres.slides[slide_index]
            
            table = ppt_utils.add_table(slide, left, top, rows, cols, data)
            return {
                "message": f"Added {rows}x{cols} table to slide {slide_index}",
                "shape_index": len(slide.shapes) - 1
            }
        except (ValueError, KeyError) as e:
            return {"error": str(e)}
    
    def add_image(
        self,
        slide_index: int,
        image_path: str,
        left: float,
        top: float,
        width: Optional[float] = None,
        height: Optional[float] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add an image to a slide."""
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Invalid slide index: {slide_index}"}
            slide = pres.slides[slide_index]
            
            # Ensure image path is in /data
            if not image_path.startswith("/data/"):
                image_path = os.path.join("/data", os.path.basename(image_path))
            
            picture = ppt_utils.add_image_from_path(slide, image_path, left, top, width, height)
            return {
                "message": f"Added image to slide {slide_index}",
                "shape_index": len(slide.shapes) - 1
            }
        except (ValueError, KeyError, FileNotFoundError) as e:
            return {"error": str(e)}
    
    def add_bullet_points(
        self,
        slide_index: int,
        placeholder_idx: int,
        bullet_points: List[str],
        font_size: Optional[int] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add bullet points to a placeholder on a slide."""
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Invalid slide index: {slide_index}"}
            slide = pres.slides[slide_index]
            
            ppt_utils.create_bullet_points(slide, placeholder_idx, bullet_points, font_size)
            return {
                "message": f"Added {len(bullet_points)} bullet points to slide {slide_index}",
                "placeholder_index": placeholder_idx
            }
        except (ValueError, KeyError) as e:
            return {"error": str(e)}


# Global instance
slide_manager = SlideManager()