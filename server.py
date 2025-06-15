from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from starlette.responses import JSONResponse
from mcp.server.fastmcp import FastMCP
from typing import Optional, Dict, Any, List
import os
import ppt_utils

# Create the MCP app as before
mcp = FastMCP("pptx")
mcp_app = mcp.sse_app()

# Create the FastAPI app
app = FastAPI()

# Enable CORS for all origins (for development; restrict in production as needed)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Add the /.well-known/wingman endpoint to the FastAPI app
@app.get("/.well-known/wingman")
async def wingman_well_known():
    return {}

# Add all MCP routes to the FastAPI app at root
app.router.routes.extend(mcp_app.routes)

import shutil
from fastapi import UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse

# In-memory storage for presentations
presentations = {}
current_presentation_id = None

# In-memory storage for template styles
current_template_styles = None
current_template_path = None

def get_current_presentation():
    """Get the current presentation object or raise an error if none is loaded."""
    if current_presentation_id is None or current_presentation_id not in presentations:
        raise ValueError("No presentation is currently loaded. Please create or open a presentation first.")
    return presentations[current_presentation_id]
# ---- Template Tools ----

@mcp.tool()
def set_template_presentation(file_path: str) -> Dict:
    """
    Set a template presentation by file path and extract its styles.
    """
    global current_template_styles, current_template_path
    if not os.path.exists(file_path):
        return {"error": f"Template file not found: {file_path}"}
    try:
        pres = ppt_utils.open_presentation(file_path)
        styles = ppt_utils.extract_template_styles(pres)
        current_template_styles = styles
        current_template_path = file_path
        return {
            "message": f"Template set from {file_path}",
            "template_path": file_path,
            "styles": styles
        }
    except Exception as e:
        return {"error": f"Failed to set template: {str(e)}"}

@mcp.tool()
def get_template_styles() -> Dict:
    """
    Get the currently loaded template styles.
    """
    if current_template_styles is None:
        return {"error": "No template styles loaded."}
    return {
        "template_path": current_template_path,
        "styles": current_template_styles
    }

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
    # Ensure file_path is in /data
    if not file_path.startswith("/data/"):
        file_path = os.path.join("/data", os.path.basename(file_path))
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
    # Ensure file_path is in /data
    if not file_path.startswith("/data/"):
        file_path = os.path.join("/data", os.path.basename(file_path))
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
    """Add a textbox to a slide, using template styles as defaults if available."""
    try:
        pres = get_current_presentation() if presentation_id is None else presentations[presentation_id]
        if not (0 <= slide_index < len(pres.slides)):
            return {"error": f"Invalid slide index: {slide_index}"}
        slide = pres.slides[slide_index]

        # Use template styles as defaults if not provided
        defaults = current_template_styles["fonts"] if current_template_styles and "fonts" in current_template_styles else {}
        if font_name is None:
            font_name = defaults.get("body_font_name")
        if font_size is None:
            font_size = defaults.get("body_font_size")
        # Optionally, set a default color from template theme colors (e.g., "accent1")
        if color is None and current_template_styles and "colors" in current_template_styles:
            accent = current_template_styles["colors"].get("accent_1")
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

@mcp.tool()
def add_shape(
    slide_index: int, shape_type: str, left: float, top: float, width: float, height: float,
    fill_color: Optional[List[int]] = None, line_color: Optional[List[int]] = None,
    line_width: Optional[float] = None, presentation_id: Optional[str] = None
) -> Dict:
    """Add an auto shape to a slide, using template styles as defaults if available."""
    try:
        pres = get_current_presentation() if presentation_id is None else presentations[presentation_id]
        if not (0 <= slide_index < len(pres.slides)):
            return {"error": f"Invalid slide index: {slide_index}"}
        slide = pres.slides[slide_index]

        # Use template styles as defaults if not provided
        if current_template_styles and "colors" in current_template_styles:
            colors = current_template_styles["colors"]
            if fill_color is None:
                accent = colors.get("accent_1")
                if accent:
                    fill_color = list(accent)
            if line_color is None:
                text1 = colors.get("text_1")
                if text1:
                    line_color = list(text1)

        ppt_utils.add_shape(
            slide, shape_type, left, top, width, height,
            fill_color=fill_color, line_color=line_color, line_width=line_width
        )
        return {"message": f"Added {shape_type} shape to slide {slide_index}", "shape_index": len(slide.shapes) - 1}
    except (ValueError, KeyError) as e:
        return {"error": str(e)}

@mcp.tool()
def add_line(
    slide_index: int,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    line_color: Optional[List[int]] = None,
    line_width: Optional[float] = None,
    presentation_id: Optional[str] = None
) -> Dict:
    """Add a straight line to a slide."""
    try:
        pres = get_current_presentation() if presentation_id is None else presentations[presentation_id]
        if not (0 <= slide_index < len(pres.slides)):
            return {"error": f"Invalid slide index: {slide_index}"}
        slide = pres.slides[slide_index]
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


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
