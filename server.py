"""
PowerPoint MCP Server

A Model Context Protocol server for creating and manipulating Microsoft PowerPoint presentations.
This server provides professional PowerPoint generation capabilities using the FastMCP framework.
"""
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from mcp.server.fastmcp import FastMCP
from typing import Optional, Dict, Any, List
import logging

# Import our modular components
from presentation_manager import presentation_manager
from template_manager import template_manager
from slide_manager import slide_manager

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create the MCP app with better naming and description
mcp = FastMCP(
    name="PowerPoint MCP Server",
    instructions="A professional PowerPoint presentation generation server supporting templates, styling, and advanced content creation."
)
mcp_app = mcp.sse_app()

# Create the FastAPI app with metadata
app = FastAPI(
    title="PowerPoint MCP Server",
    description="Professional PowerPoint presentation generation via Model Context Protocol",
    version="2.0.0"
)

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
    """Wingman discovery endpoint."""
    return {}

# Add all MCP routes to the FastAPI app at root
app.router.routes.extend(mcp_app.routes)
# ============================================================================
# TEMPLATE MANAGEMENT TOOLS
# ============================================================================

@mcp.tool()
def set_template_presentation(file_path: str) -> Dict[str, Any]:
    """
    Set a template presentation by file path and extract its styles.
    
    Args:
        file_path: Path to the template presentation file
        
    Returns:
        Dictionary with template information and extracted styles
    """
    logger.info(f"Setting template from: {file_path}")
    return template_manager.set_template_presentation(file_path)

@mcp.tool()
def get_template_styles() -> Dict[str, Any]:
    """
    Get the currently loaded template styles.
    
    Returns:
        Dictionary containing current template path and extracted styles
    """
    return template_manager.get_template_styles()


# ============================================================================
# PRESENTATION MANAGEMENT TOOLS  
# ============================================================================

@mcp.tool()
def create_presentation(id: Optional[str] = None) -> Dict[str, Any]:
    """
    Create a new PowerPoint presentation.
    
    Args:
        id: Optional unique identifier for the presentation
        
    Returns:
        Dictionary with presentation ID, confirmation message, and slide count
    """
    logger.info(f"Creating new presentation with ID: {id}")
    return presentation_manager.create_presentation(id)

@mcp.tool()
def open_presentation(file_path: str, id: Optional[str] = None) -> Dict[str, Any]:
    """
    Open an existing PowerPoint presentation from a file.
    
    Args:
        file_path: Path to the presentation file to open
        id: Optional unique identifier to assign to the opened presentation
        
    Returns:
        Dictionary with presentation ID, confirmation message, and slide count
    """
    logger.info(f"Opening presentation from: {file_path}")
    return presentation_manager.open_presentation(file_path, id)

@mcp.tool()
def save_presentation(file_path: str, presentation_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Save a presentation to a file.
    
    Args:
        file_path: Path where the presentation will be saved
        presentation_id: Optional ID of the presentation to save
        
    Returns:
        Dictionary with confirmation message and file path
    """
    logger.info(f"Saving presentation to: {file_path}")
    return presentation_manager.save_presentation(file_path, presentation_id)

@mcp.tool()
def get_presentation_info(presentation_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Get information about a presentation.
    
    Args:
        presentation_id: Optional ID of the presentation to inspect
        
    Returns:
        Dictionary containing presentation metadata, slide count, layouts, and properties
    """
    return presentation_manager.get_presentation_info(presentation_id)

@mcp.tool()
def set_core_properties(
    title: Optional[str] = None,
    subject: Optional[str] = None,
    author: Optional[str] = None,
    keywords: Optional[str] = None,
    comments: Optional[str] = None,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Set core document properties.
    
    Args:
        title: Presentation title
        subject: Presentation subject
        author: Presentation author
        keywords: Keywords associated with the presentation
        comments: Comments about the presentation
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and updated properties
    """
    logger.info(f"Setting core properties for presentation: {presentation_id}")
    return presentation_manager.set_core_properties(
        title=title, subject=subject, author=author, 
        keywords=keywords, comments=comments, presentation_id=presentation_id
    )

# ============================================================================
# SLIDE MANAGEMENT TOOLS
# ============================================================================

@mcp.tool()
def add_slide(layout_index: int = 1, title: Optional[str] = None, presentation_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Add a new slide to the presentation.
    
    Args:
        layout_index: Index of the slide layout to use (0-based)
        title: Optional title to set for the new slide
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message, slide index, and layout information
    """
    logger.info(f"Adding slide with layout index: {layout_index}")
    return slide_manager.add_slide(layout_index, title, presentation_id)


# ============================================================================
# SHAPE AND TEXT TOOLS
# ============================================================================

@mcp.tool()
def add_textbox(
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
    """
    Add a textbox to a slide with professional formatting options.
    
    Args:
        slide_index: Index of the target slide
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        text: Text content for the textbox
        font_size: Optional font size in points
        font_name: Optional font name (defaults to template or Calibri)
        bold: Optional bold formatting
        italic: Optional italic formatting
        color: Optional text color as RGB list [r, g, b]
        alignment: Optional text alignment (left, center, right, justify)
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and shape index
    """
    logger.info(f"Adding textbox to slide {slide_index}")
    return slide_manager.add_textbox(
        slide_index, left, top, width, height, text,
        font_size=font_size, font_name=font_name, bold=bold,
        italic=italic, color=color, alignment=alignment,
        presentation_id=presentation_id
    )

@mcp.tool()
def add_shape(
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
    """
    Add an auto shape to a slide with professional styling.
    
    Args:
        slide_index: Index of the target slide
        shape_type: Type of shape (rectangle, oval, triangle, arrow, etc.)
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        fill_color: Optional fill color as RGB list [r, g, b]
        line_color: Optional line color as RGB list [r, g, b]
        line_width: Optional line width in points
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and shape index
    """
    logger.info(f"Adding {shape_type} shape to slide {slide_index}")
    return slide_manager.add_shape(
        slide_index, shape_type, left, top, width, height,
        fill_color=fill_color, line_color=line_color,
        line_width=line_width, presentation_id=presentation_id
    )

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
) -> Dict[str, Any]:
    """
    Add a straight line to a slide.
    
    Args:
        slide_index: Index of the target slide
        x1: Starting X coordinate in inches
        y1: Starting Y coordinate in inches
        x2: Ending X coordinate in inches
        y2: Ending Y coordinate in inches
        line_color: Optional line color as RGB list [r, g, b]
        line_width: Optional line width in points
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and shape index
    """
    logger.info(f"Adding line to slide {slide_index}")
    return slide_manager.add_line(
        slide_index, x1, y1, x2, y2,
        line_color=line_color, line_width=line_width,
        presentation_id=presentation_id
    )


# ============================================================================
# ADVANCED CONTENT TOOLS
# ============================================================================

@mcp.tool()
def add_chart(
    slide_index: int,
    chart_type: str,
    left: float,
    top: float,
    width: float,
    height: float,
    data: Dict[str, Any],
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add a chart to a slide for professional data visualization.
    
    Args:
        slide_index: Index of the target slide
        chart_type: Type of chart ('column', 'line', 'pie', 'bar', 'area')
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        data: Chart data in format {'categories': [...], 'series': [{'name': '...', 'values': [...]}]}
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and shape index
    """
    logger.info(f"Adding {chart_type} chart to slide {slide_index}")
    return slide_manager.add_chart(
        slide_index, chart_type, left, top, width, height, data, presentation_id
    )

@mcp.tool()
def add_table(
    slide_index: int,
    left: float,
    top: float,
    rows: int,
    cols: int,
    data: Optional[List[List[str]]] = None,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add a professional table to a slide.
    
    Args:
        slide_index: Index of the target slide
        left: Left position in inches
        top: Top position in inches
        rows: Number of rows
        cols: Number of columns
        data: Optional table data as list of lists (first row will be styled as header)
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and shape index
    """
    logger.info(f"Adding {rows}x{cols} table to slide {slide_index}")
    return slide_manager.add_table(
        slide_index, left, top, rows, cols, data, presentation_id
    )

@mcp.tool()
def add_image(
    slide_index: int,
    image_path: str,
    left: float,
    top: float,
    width: Optional[float] = None,
    height: Optional[float] = None,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add an image to a slide with automatic sizing and positioning.
    
    Args:
        slide_index: Index of the target slide
        image_path: Path to the image file (will be resolved to /data/ directory)
        left: Left position in inches
        top: Top position in inches
        width: Optional width in inches (maintains aspect ratio if height not specified)
        height: Optional height in inches (maintains aspect ratio if width not specified)
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and shape index
    """
    logger.info(f"Adding image to slide {slide_index}")
    return slide_manager.add_image(
        slide_index, image_path, left, top, width, height, presentation_id
    )

@mcp.tool()
def add_bullet_points(
    slide_index: int,
    placeholder_idx: int,
    bullet_points: List[str],
    font_size: Optional[int] = None,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add professional bullet points to a placeholder on a slide.
    
    Args:
        slide_index: Index of the target slide
        placeholder_idx: Index of the placeholder to use for bullet points
        bullet_points: List of bullet point text
        font_size: Optional font size in points
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and placeholder index
    """
    logger.info(f"Adding {len(bullet_points)} bullet points to slide {slide_index}")
    return slide_manager.add_bullet_points(
        slide_index, placeholder_idx, bullet_points, font_size, presentation_id
    )


# ============================================================================
# SERVER STARTUP
# ============================================================================

if __name__ == "__main__":
    import uvicorn
    logger.info("Starting PowerPoint MCP Server...")
    uvicorn.run(app, host="0.0.0.0", port=8000)
