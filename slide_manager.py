"""
Slide Management Module

Handles slide creation, manipulation, and content management with validation and performance monitoring.
"""
from typing import Optional, Dict, Any, List
import os
import ppt_utils
from presentation_manager import presentation_manager
from template_manager import template_manager
from input_validator import validator, ValidationError
from performance_optimizer import performance_monitor
from text_autofit import (
    text_autofit_engine, 
    AutoFitStrategy, 
    ContainerDimensions,
    AutoFitResult,
    AutoFitConfig
)


# Default auto-fit configuration for slide manager
DEFAULT_AUTOFIT_CONFIG = AutoFitConfig()
class SlideManager:
    """Manages slide operations within presentations."""
    
    @performance_monitor.track_operation("add_slide")
    def add_slide(self, layout_index: int = 1, title: Optional[str] = None, presentation_id: Optional[str] = None) -> Dict[str, Any]:
        """Add a new slide to the presentation."""
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            if not (0 <= layout_index < len(pres.slide_layouts)):
                return {
                    "error": f"Invalid layout index: {layout_index}. Available: 0-{len(pres.slide_layouts) - 1}",
                    "available_layouts": {i: l.name for i, l in enumerate(pres.slide_layouts)}
                }
            
            # Validate title if provided
            if title:
                title = validator.validate_text(title, max_length=validator.MAX_TITLE_LENGTH)
            
            slide, info = ppt_utils.add_slide(pres, layout_index, title)
            return {
                "message": f"Added slide with layout '{info['layout_name']}'",
                "slide_index": len(pres.slides) - 1,
                **info
            }
        except (ValueError, KeyError, ValidationError) as e:
            return {"error": str(e)}
    
    @performance_monitor.track_operation("add_textbox")
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
            
            # Validate inputs
            slide_index = validator.validate_slide_index(slide_index, len(pres.slides))
            left, top, width, height = validator.validate_dimensions(left, top, width, height)
            text = validator.validate_text(text)
            
            if color:
                color = validator.validate_color(color)
            
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
        except (ValueError, KeyError, ValidationError) as e:
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


    @performance_monitor.track_operation("add_chart")
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
            
            # Validate inputs
            slide_index = validator.validate_slide_index(slide_index, len(pres.slides))
            left, top, width, height = validator.validate_dimensions(left, top, width, height)
            data = validator.validate_chart_data(data)
            
            slide = pres.slides[slide_index]
            
            chart = ppt_utils.add_chart(slide, chart_type, left, top, width, height, data)
            return {
                "message": f"Added {chart_type} chart to slide {slide_index}",
                "shape_index": len(slide.shapes) - 1
            }
        except (ValueError, KeyError, ValidationError) as e:
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
    
    @performance_monitor.track_operation("add_auto_fit_text")
    def add_auto_fit_text(
        self,
        slide_index: int,
        left: float,
        top: float,
        width: float,
        height: float,
        text: str,
        strategy: str = "smart",
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[List[int]] = None,
        alignment: Optional[str] = None,
        create_new_slides: bool = True,
        slide_title_template: Optional[str] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Add text with intelligent auto-fit to a slide.
        
        When AI-generated content is extensive, this method automatically adjusts
        font size, uses multi-column layout, or splits content across slides.
        
        Args:
            slide_index: Index of the target slide
            left: Left position in inches
            top: Top position in inches
            width: Width in inches
            height: Height in inches
            text: Text content to add (can be very long)
            strategy: Auto-fit strategy ('smart', 'shrink_font', 'multi_column', 'split_slides')
            font_size: Optional preferred font size in points
            font_name: Optional font name
            bold: Optional bold formatting
            italic: Optional italic formatting
            color: Optional text color as RGB list [r, g, b]
            alignment: Optional text alignment (left, center, right, justify)
            create_new_slides: Whether to create new slides if content is split
            slide_title_template: Title template for new slides (use {page} for page number)
            presentation_id: Optional ID of the target presentation
            
        Returns:
            Dictionary with auto-fit results including strategy used and any new slides created
        """
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            
            # Validate inputs
            slide_index = validator.validate_slide_index(slide_index, len(pres.slides))
            left, top, width, height = validator.validate_dimensions(left, top, width, height)
            text = validator.validate_text(text)
            
            if color:
                color = validator.validate_color(color)
            
            # Use template styles as defaults if not provided
            font_defaults = template_manager.get_default_font_settings()
            color_defaults = template_manager.get_default_color_settings()
            
            if font_name is None:
                font_name = font_defaults.get("body_font_name")
            if color is None:
                accent = color_defaults.get("accent_1")
                if accent:
                    color = list(accent)
            
            # Parse strategy
            strategy_map = {
                "smart": AutoFitStrategy.SMART,
                "shrink_font": AutoFitStrategy.SHRINK_FONT,
                "multi_column": AutoFitStrategy.MULTI_COLUMN,
                "split_slides": AutoFitStrategy.SPLIT_SLIDES
            }
            fit_strategy = strategy_map.get(strategy.lower(), AutoFitStrategy.SMART)
            
            # Create container dimensions
            container = ContainerDimensions(
                width=width,
                height=height,
                slide_width=pres.slide_width.inches,
                slide_height=pres.slide_height.inches
            )
            
            # Calculate auto-fit
            result = text_autofit_engine.auto_fit(
                text=text,
                container=container,
                strategy=fit_strategy,
                preferred_font_size=font_size
            )
            
            created_shapes = []
            created_slides = []
            
            if result.strategy == AutoFitStrategy.MULTI_COLUMN:
                # Create multi-column layout on the same slide
                slide = pres.slides[slide_index]
                column_gap = 0.3  # Gap between columns in inches
                
                for col_idx, col_text in enumerate(result.text_segments):
                    col_left = left + col_idx * (result.column_width + column_gap)
                    
                    ppt_utils.add_textbox(
                        slide, col_left, top, result.column_width, height, col_text,
                        font_size=result.font_size, font_name=font_name, bold=bold,
                        italic=italic, color=color, alignment=alignment
                    )
                    created_shapes.append({
                        "slide_index": slide_index,
                        "shape_index": len(slide.shapes) - 1,
                        "column": col_idx
                    })
                
            elif result.strategy == AutoFitStrategy.SPLIT_SLIDES and len(result.text_segments) > 1:
                # Split content across multiple slides
                for seg_idx, segment_text in enumerate(result.text_segments):
                    if seg_idx == 0:
                        # Use the specified slide for first segment
                        current_slide = pres.slides[slide_index]
                        current_slide_index = slide_index
                    else:
                        if create_new_slides:
                            # Create a new slide for subsequent segments
                            # Find the layout index for the original slide's layout
                            layout_idx = 1  # Default blank layout
                            original_layout = pres.slides[slide_index].slide_layout
                            for i, layout in enumerate(pres.slide_layouts):
                                if layout == original_layout:
                                    layout_idx = i
                                    break
                            
                            # Add the new slide
                            new_slide, _ = ppt_utils.add_slide(pres, layout_idx)
                            current_slide = new_slide
                            current_slide_index = len(pres.slides) - 1
                            
                            # Set title if template provided
                            if slide_title_template:
                                title_text = slide_title_template.replace("{page}", str(seg_idx + 1))
                                if current_slide.shapes.title:
                                    current_slide.shapes.title.text = title_text
                            
                            created_slides.append(current_slide_index)
                        else:
                            # Don't create new slides, put all on original
                            current_slide = pres.slides[slide_index]
                            current_slide_index = slide_index
                            # Adjust top position for stacking using configurable gap
                            top = top + height + DEFAULT_AUTOFIT_CONFIG.stacking_gap
                    
                    ppt_utils.add_textbox(
                        current_slide, left, top, width, height, segment_text,
                        font_size=result.font_size, font_name=font_name, bold=bold,
                        italic=italic, color=color, alignment=alignment
                    )
                    created_shapes.append({
                        "slide_index": current_slide_index,
                        "shape_index": len(current_slide.shapes) - 1,
                        "segment": seg_idx
                    })
            else:
                # Single textbox with adjusted font size
                slide = pres.slides[slide_index]
                ppt_utils.add_textbox(
                    slide, left, top, width, height, result.text_segments[0],
                    font_size=result.font_size, font_name=font_name, bold=bold,
                    italic=italic, color=color, alignment=alignment
                )
                created_shapes.append({
                    "slide_index": slide_index,
                    "shape_index": len(slide.shapes) - 1
                })
            
            return {
                "message": f"Added auto-fit text using '{result.strategy.value}' strategy",
                "strategy_used": result.strategy.value,
                "font_size": result.font_size,
                "columns": result.columns,
                "slides_used": result.slides_needed,
                "recommendation": result.recommendation,
                "shapes_created": created_shapes,
                "new_slides_created": created_slides
            }
            
        except (ValueError, KeyError, ValidationError) as e:
            return {"error": str(e)}


# Global instance
slide_manager = SlideManager()