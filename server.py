# server.py
from mcp.server.fastmcp import FastMCP
from typing import Optional, Dict, Any, List
import os
import ppt_utils

mcp = FastMCP("pptx")
app = mcp.sse_app()

# In-memory storage for presentations
presentations = {}
current_presentation_id = None

def get_current_presentation():
    """Get the current presentation object or raise an error if none is loaded."""
    if current_presentation_id is None or current_presentation_id not in presentations:
        raise ValueError("No presentation is currently loaded. Please create or open a presentation first.")
    return presentations[current_presentation_id]

# ---- Presentation Tools ----

@mcp.tool()
def create_presentation(id: Optional[str] = None) -> Dict:
    """Create a new PowerPoint presentation."""
    global current_presentation_id
    pres = ppt_utils.create_presentation()
    if id is None:
        id = f"presentation_{len(presentations) + 1}"
    presentations[id] = pres
    current_presentation_id = id
    return {
        "presentation_id": id,
        "message": f"Created new presentation with ID: {id}",
        "slide_count": len(pres.slides)
    }

@mcp.tool()
def open_presentation(file_path: str, id: Optional[str] = None) -> Dict:
    """Open an existing PowerPoint presentation from a file."""
    global current_presentation_id
    if not os.path.exists(file_path):
        return {"error": f"File not found: {file_path}"}
    try:
        pres = ppt_utils.open_presentation(file_path)
    except Exception as e:
        return {"error": f"Failed to open presentation: {str(e)}"}
    if id is None:
        id = f"presentation_{len(presentations) + 1}"
    presentations[id] = pres
    current_presentation_id = id
    return {
        "presentation_id": id,
        "message": f"Opened presentation from {file_path} with ID: {id}",
        "slide_count": len(pres.slides)
    }

@mcp.tool()
def save_presentation(file_path: str, presentation_id: Optional[str] = None) -> Dict:
    """Save a presentation to a file."""
    try:
        pres = get_current_presentation() if presentation_id is None else presentations[presentation_id]
        saved_path = ppt_utils.save_presentation(pres, file_path)
        return {"message": f"Presentation saved to {saved_path}", "file_path": saved_path}
    except (ValueError, KeyError) as e:
        return {"error": str(e)}

@mcp.tool()
def get_presentation_info(presentation_id: Optional[str] = None) -> Dict:
    """Get information about a presentation."""
    try:
        pres = get_current_presentation() if presentation_id is None else presentations[presentation_id]
        info = ppt_utils.get_presentation_info(pres)
        info["presentation_id"] = current_presentation_id if presentation_id is None else presentation_id
        return info
    except (ValueError, KeyError) as e:
        return {"error": str(e)}

@mcp.tool()
def set_core_properties(
    title: Optional[str] = None, subject: Optional[str] = None, author: Optional[str] = None,
    keywords: Optional[str] = None, comments: Optional[str] = None, presentation_id: Optional[str] = None
) -> Dict:
    """Set core document properties."""
    try:
        pres = get_current_presentation() if presentation_id is None else presentations[presentation_id]
        updated_props = ppt_utils.set_core_properties(
            pres, title=title, subject=subject, author=author, keywords=keywords, comments=comments
        )
        return {"message": "Core properties updated successfully", "core_properties": updated_props}
    except (ValueError, KeyError) as e:
        return {"error": str(e)}

# ---- Slide Tools ----

@mcp.tool()
def add_slide(layout_index: int = 1, title: Optional[str] = None, presentation_id: Optional[str] = None) -> Dict:
    """Add a new slide to the presentation."""
    try:
        pres = get_current_presentation() if presentation_id is None else presentations[presentation_id]
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

# ---- Shape and Text Tools ----

@mcp.tool()
def add_textbox(
    slide_index: int, left: float, top: float, width: float, height: float, text: str,
    font_size: Optional[int] = None, font_name: Optional[str] = None, bold: Optional[bool] = None,
    italic: Optional[bool] = None, color: Optional[List[int]] = None, alignment: Optional[str] = None,
    presentation_id: Optional[str] = None
) -> Dict:
    """Add a textbox to a slide."""
    try:
        pres = get_current_presentation() if presentation_id is None else presentations[presentation_id]
        if not (0 <= slide_index < len(pres.slides)):
            return {"error": f"Invalid slide index: {slide_index}"}
        slide = pres.slides[slide_index]
        ppt_utils.add_textbox(
            slide, left, top, width, height, text,
            font_size=font_size, font_name=font_name, bold=bold,
            italic=italic, color=color, alignment=alignment
        )
        return {"message": f"Added textbox to slide {slide_index}", "shape_index": len(slide.shapes) - 1}
    except (ValueError, KeyError) as e:
        return {"error": str(e)}

@mcp.tool()
def add_shape(
    slide_index: int, shape_type: str, left: float, top: float, width: float, height: float,
    fill_color: Optional[List[int]] = None, line_color: Optional[List[int]] = None,
    line_width: Optional[float] = None, presentation_id: Optional[str] = None
) -> Dict:
    """Add an auto shape to a slide."""
    try:
        pres = get_current_presentation() if presentation_id is None else presentations[presentation_id]
        if not (0 <= slide_index < len(pres.slides)):
            return {"error": f"Invalid slide index: {slide_index}"}
        slide = pres.slides[slide_index]
        ppt_utils.add_shape(
            slide, shape_type, left, top, width, height,
            fill_color=fill_color, line_color=line_color, line_width=line_width
        )
        return {"message": f"Added {shape_type} shape to slide {slide_index}", "shape_index": len(slide.shapes) - 1}
    except (ValueError, KeyError) as e:
        return {"error": str(e)}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
