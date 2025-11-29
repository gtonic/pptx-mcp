"""
PowerPoint MCP Server

A Model Context Protocol server for creating and manipulating Microsoft PowerPoint presentations.
This server provides professional PowerPoint generation capabilities using the FastMCP framework.
"""
from mcp.server.fastmcp import FastMCP
from typing import Optional, Dict, Any, List, Union
import logging

# Import our modular components
from presentation_manager import presentation_manager
from template_manager import template_manager
from slide_manager import slide_manager
from layout_manager import layout_manager
from input_validator import validator, ValidationError
from performance_optimizer import performance_monitor
from diagram_renderer import diagram_renderer
from business_diagrams import business_diagrams

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create the MCP app with better naming and description
# streamable_http_path="/mcp" is the default endpoint for Streamable HTTP
mcp = FastMCP(
    name="PowerPoint MCP Server",
    instructions="A professional PowerPoint presentation generation server supporting templates, styling, and advanced content creation.",
    host="0.0.0.0",
    port=8081,
    streamable_http_path="/mcp",
)
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
        Dictionary containing current template path, extracted styles,
        and available semantic color/font tags with their current values
    """
    return template_manager.get_template_styles()


@mcp.tool()
def get_semantic_tags() -> Dict[str, Any]:
    """
    Get all available semantic styling tags and their current values.
    
    This tool returns the semantic color and font tags that can be used
    instead of hard-coded values. The values are mapped to the current
    template/theme or sensible defaults.
    
    Returns:
        Dictionary with:
            - color_tags: List of available semantic color tags
            - font_tags: List of available semantic font tags
            - color_palette: Current color values for each semantic tag
            - font_styles: Current font settings for each semantic tag
            
    Example color tags:
        - 'primary': Main brand/theme color
        - 'secondary': Secondary accent color
        - 'accent': Highlight/emphasis color
        - 'success': Positive/success state (green)
        - 'warning': Warning/caution state (yellow/orange)
        - 'critical': Error/danger state (red)
        - 'info': Informational content (blue)
        - 'neutral': Neutral/gray elements
        - 'text': Default text color
        - 'background': Background color
        
    Example font tags:
        - 'title': Large title text style
        - 'heading': Section heading style
        - 'body': Main body text style
        - 'caption': Small caption/note style
        - 'code': Monospace code style
    """
    return {
        "color_tags": template_manager.get_semantic_color_tags(),
        "font_tags": template_manager.get_semantic_font_tags(),
        "color_palette": template_manager.get_color_palette(),
        "font_styles": template_manager.get_font_styles()
    }


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
    font_style: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[Union[str, List[int]]] = None,
    alignment: Optional[str] = None,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add a textbox to a slide with professional formatting options.
    
    Supports semantic styling tags for colors and fonts, allowing AI models to
    use meaningful names instead of hard-coded values.
    
    Args:
        slide_index: Index of the target slide
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        text: Text content for the textbox
        font_size: Optional font size in points
        font_name: Optional font name (defaults to template or Calibri)
        font_style: Optional semantic font style tag ('title', 'heading', 'body', 'caption', 'code')
        bold: Optional bold formatting
        italic: Optional italic formatting
        color: Optional text color - can be a semantic tag ('primary', 'accent', 'critical', 
               'success', 'warning', 'text', 'neutral') or RGB list [r, g, b]
        alignment: Optional text alignment (left, center, right, justify)
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and shape index
        
    Example with semantic tags:
        add_textbox(0, 1, 1, 4, 1, "Important!", color="critical", font_style="heading")
    """
    logger.info(f"Adding textbox to slide {slide_index}")
    return slide_manager.add_textbox(
        slide_index, left, top, width, height, text,
        font_size=font_size, font_name=font_name, font_style=font_style, bold=bold,
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
    fill_color: Optional[Union[str, List[int]]] = None,
    line_color: Optional[Union[str, List[int]]] = None,
    line_width: Optional[float] = None,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add an auto shape to a slide with professional styling.
    
    Supports semantic color tags for fill and line colors.
    
    Args:
        slide_index: Index of the target slide
        shape_type: Type of shape (rectangle, oval, triangle, arrow, etc.)
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        fill_color: Optional fill color - can be a semantic tag ('primary', 'accent', 
                   'success', 'warning', 'critical', 'neutral') or RGB list [r, g, b]
        line_color: Optional line color - same options as fill_color
        line_width: Optional line width in points
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and shape index
        
    Example with semantic tags:
        add_shape(0, "rectangle", 1, 1, 2, 2, fill_color="accent", line_color="neutral")
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
    line_color: Optional[Union[str, List[int]]] = None,
    line_width: Optional[float] = None,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add a straight line to a slide.
    
    Supports semantic color tags for line color.
    
    Args:
        slide_index: Index of the target slide
        x1: Starting X coordinate in inches
        y1: Starting Y coordinate in inches
        x2: Ending X coordinate in inches
        y2: Ending Y coordinate in inches
        line_color: Optional line color - can be a semantic tag ('neutral', 'accent', 
                   'text') or RGB list [r, g, b]
        line_width: Optional line width in points
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation message and shape index
        
    Example with semantic tags:
        add_line(0, 1, 1, 5, 1, line_color="neutral")
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
# INTELLIGENT TEXT AUTO-FIT TOOLS
# ============================================================================

@mcp.tool()
def add_auto_fit_text(
    slide_index: int,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    strategy: str = "smart",
    font_size: Optional[int] = None,
    font_name: Optional[str] = None,
    font_style: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[Union[str, List[int]]] = None,
    alignment: Optional[str] = None,
    create_new_slides: bool = True,
    slide_title_template: Optional[str] = None,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add text with intelligent auto-fit to a slide.
    
    When AI-generated content is extensive, this tool automatically adjusts:
    - Font size based on text length and container dimensions
    - Multi-column layout for better readability
    - Slide splitting for very long content
    
    Supports semantic styling tags for colors and fonts.
    Goal: Maximum readability and sensible slide division for large data sets.
    
    Args:
        slide_index: Index of the target slide
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        text: Text content to add (can be very long AI-generated content)
        strategy: Auto-fit strategy:
            - 'smart': Automatically choose best strategy (default)
            - 'shrink_font': Only adjust font size
            - 'multi_column': Split into multiple columns
            - 'split_slides': Split across multiple slides
        font_size: Optional preferred font size in points (will be adjusted if needed)
        font_name: Optional font name (defaults to template or Calibri)
        font_style: Optional semantic font style tag ('title', 'heading', 'body', 'caption', 'code')
        bold: Optional bold formatting
        italic: Optional italic formatting
        color: Optional text color - can be a semantic tag ('primary', 'accent', 'critical',
               'success', 'warning', 'text', 'neutral') or RGB list [r, g, b]
        alignment: Optional text alignment (left, center, right, justify)
        create_new_slides: Whether to create new slides if content is split (default: True)
        slide_title_template: Title template for new slides (use {page} for page number)
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with:
            - strategy_used: The auto-fit strategy that was applied
            - font_size: The final font size used
            - columns: Number of columns (if multi-column)
            - slides_used: Number of slides used
            - recommendation: Explanation of what was done
            - shapes_created: List of created shapes
            - new_slides_created: List of any new slide indices
            
    Example with semantic tags:
        add_auto_fit_text(
            slide_index=0,
            left=0.5, top=1.5, width=9.0, height=5.0,
            text="Very long AI-generated content...",
            strategy="smart",
            color="text",
            font_style="body",
            create_new_slides=True,
            slide_title_template="Content (Page {page})"
        )
    """
    logger.info(f"Adding auto-fit text to slide {slide_index} with strategy: {strategy}")
    return slide_manager.add_auto_fit_text(
        slide_index, left, top, width, height, text,
        strategy=strategy, font_size=font_size, font_name=font_name, font_style=font_style,
        bold=bold, italic=italic, color=color, alignment=alignment,
        create_new_slides=create_new_slides, slide_title_template=slide_title_template,
        presentation_id=presentation_id
    )


# ============================================================================
# HIGH-LEVEL LAYOUT ENGINE TOOLS
# ============================================================================

@mcp.tool()
def add_grid_layout(
    slide_index: int,
    elements: List[Dict[str, Any]],
    rows: int = 2,
    cols: int = 2,
    gap: float = 0.2,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Create elements arranged in a grid layout without specifying coordinates.
    
    The layout engine automatically calculates positions based on slide dimensions.
    This is ideal for AI/LLM-generated content that works with structural descriptions.
    
    Supports semantic styling tags for colors in addition to RGB values.
    
    Args:
        slide_index: Index of the target slide
        elements: List of element dictionaries. Each element can have:
            - content (required): Text content for the element
            - element_type: 'textbox' or 'shape' (default: 'textbox')
            - shape_type: For shapes: 'rectangle', 'rounded_rectangle', 'oval', etc.
            - fill_color: Semantic tag ('primary', 'accent', 'success', etc.) or RGB list [r, g, b]
            - text_color: Semantic tag ('text', 'text_inverted', etc.) or RGB list [r, g, b]
            - line_color: Semantic tag or RGB list [r, g, b]
            - font_size: Font size in points
            - bold: Boolean for bold text
        rows: Number of rows in the grid (default: 2)
        cols: Number of columns in the grid (default: 2)
        gap: Gap between cells in inches (default: 0.2)
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation, layout info, and created shape indices
        
    Example with semantic tags:
        add_grid_layout(0, [
            {"content": "Q1", "fill_color": "primary"},
            {"content": "Q2", "fill_color": "secondary"},
            {"content": "Q3", "fill_color": "success"},
            {"content": "Q4", "fill_color": "accent"}
        ], rows=2, cols=2)
    """
    logger.info(f"Creating grid layout ({rows}x{cols}) with {len(elements)} elements on slide {slide_index}")
    return layout_manager.create_grid_layout(
        slide_index, elements, rows, cols, gap, 
        presentation_id=presentation_id
    )


@mcp.tool()
def add_list_layout(
    slide_index: int,
    elements: List[Dict[str, Any]],
    direction: str = "vertical",
    gap: float = 0.2,
    alignment: str = "left",
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Create elements arranged in a list layout without specifying coordinates.
    
    The layout engine automatically calculates positions based on slide dimensions.
    Perfect for bullet-point style content or horizontal feature comparisons.
    
    Supports semantic styling tags for colors in addition to RGB values.
    
    Args:
        slide_index: Index of the target slide
        elements: List of element dictionaries. Each element can have:
            - content (required): Text content for the element
            - element_type: 'textbox' or 'shape' (default: 'textbox')
            - shape_type: For shapes: 'rectangle', 'rounded_rectangle', 'oval', etc.
            - fill_color: Semantic tag ('primary', 'accent', etc.) or RGB list [r, g, b]
            - text_color: Semantic tag ('text', 'text_inverted', etc.) or RGB list [r, g, b]
            - font_size: Font size in points
            - bold: Boolean for bold text
        direction: 'vertical' or 'horizontal' (default: 'vertical')
        gap: Gap between items in inches (default: 0.2)
        alignment: For vertical: 'left', 'center', 'right'
                   For horizontal: 'top', 'middle', 'bottom' (default: 'left')
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation, layout info, and created shape indices
        
    Example with semantic tags:
        add_list_layout(0, [
            {"content": "Performance improvements", "fill_color": "success"},
            {"content": "New user interface", "fill_color": "info"},
            {"content": "Enhanced security", "fill_color": "primary"}
        ], direction="vertical", alignment="left")
    """
    logger.info(f"Creating {direction} list layout with {len(elements)} elements on slide {slide_index}")
    return layout_manager.create_list_layout(
        slide_index, elements, direction, gap, alignment,
        presentation_id=presentation_id
    )


@mcp.tool()
def add_hierarchy_layout(
    slide_index: int,
    root: Dict[str, Any],
    level_gap: float = 0.8,
    sibling_gap: float = 0.3,
    show_connectors: bool = True,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Create elements arranged in a hierarchical/tree structure without specifying coordinates.
    
    The layout engine automatically calculates positions for org charts, taxonomies,
    and other tree-like structures. Connectors are drawn automatically.
    
    Args:
        slide_index: Index of the target slide
        root: Root element dictionary with:
            - content (required): Text content for the node
            - children: Optional list of child nodes (same structure)
            - element_type: 'textbox' or 'shape' (default: 'textbox')
            - shape_type: For shapes: 'rectangle', 'rounded_rectangle', etc.
            - fill_color: RGB list [r, g, b]
        level_gap: Vertical gap between levels in inches (default: 0.8)
        sibling_gap: Horizontal gap between siblings in inches (default: 0.3)
        show_connectors: Whether to draw connecting lines (default: True)
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation, layout info, shapes, and connectors
        
    Example:
        add_hierarchy_layout(0, {
            "content": "CEO",
            "children": [
                {"content": "VP Sales", "children": [
                    {"content": "Sales Team A"},
                    {"content": "Sales Team B"}
                ]},
                {"content": "VP Engineering", "children": [
                    {"content": "Frontend"},
                    {"content": "Backend"}
                ]}
            ]
        })
    """
    logger.info(f"Creating hierarchy layout on slide {slide_index}")
    return layout_manager.create_hierarchy_layout(
        slide_index, root, level_gap, sibling_gap, show_connectors,
        presentation_id=presentation_id
    )


@mcp.tool()
def add_flow_layout(
    slide_index: int,
    steps: List[Dict[str, Any]],
    direction: str = "horizontal",
    gap: float = 0.4,
    show_connectors: bool = True,
    connector_style: str = "arrow",
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Create elements arranged as a flow/process diagram without specifying coordinates.
    
    The layout engine automatically calculates positions for process flows,
    workflows, and step-by-step diagrams with connecting arrows.
    
    Args:
        slide_index: Index of the target slide
        steps: List of step dictionaries. Each step can have:
            - content (required): Text content for the step
            - shape_type: Shape type (default: 'rounded_rectangle')
            - fill_color: RGB list [r, g, b]
            - text_color: RGB list [r, g, b]
            - font_size: Font size in points
        direction: 'horizontal' (left to right) or 'vertical' (top to bottom)
        gap: Gap between steps in inches, includes connector space (default: 0.4)
        show_connectors: Whether to draw connecting arrows (default: True)
        connector_style: 'arrow', 'line', or 'none' (default: 'arrow')
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with confirmation, layout info, shapes, and connectors
        
    Example:
        add_flow_layout(0, [
            {"content": "Start"},
            {"content": "Process Data"},
            {"content": "Analyze Results"},
            {"content": "Generate Report"},
            {"content": "End"}
        ], direction="horizontal", connector_style="arrow")
    """
    logger.info(f"Creating {direction} flow layout with {len(steps)} steps on slide {slide_index}")
    return layout_manager.create_flow_layout(
        slide_index, steps, direction, gap, show_connectors, connector_style,
        presentation_id=presentation_id
    )

@mcp.tool()
def get_performance_report() -> Dict[str, Any]:
    """
    Get comprehensive performance report and recommendations.
    
    Returns:
        Dictionary with performance metrics, memory usage, and optimization recommendations
    """
    logger.info("Generating performance report")
    return performance_monitor.get_performance_report()

@mcp.tool()
def optimize_for_large_presentation(slide_count: int) -> Dict[str, Any]:
    """
    Get optimization recommendations for large presentations.
    
    Args:
        slide_count: Number of slides in the presentation
        
    Returns:
        Dictionary with optimization suggestions and batch processing recommendations
    """
    logger.info(f"Getting optimization recommendations for {slide_count} slides")
    return performance_monitor.optimize_large_presentation(slide_count)

@mcp.tool()
def cleanup_memory() -> Dict[str, Any]:
    """
    Force memory cleanup to free resources.
    
    Returns:
        Dictionary with cleanup confirmation
    """
    import time
    logger.info("Performing memory cleanup")
    performance_monitor.cleanup_memory()
    return {"message": "Memory cleanup completed", "timestamp": time.time()}


# ============================================================================
# DIAGRAM TOOLS (Mermaid/PlantUML)
# ============================================================================

@mcp.tool()
def add_mermaid_diagram(
    slide_index: int,
    mermaid_code: str,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Render a Mermaid diagram as editable PowerPoint shapes.
    
    This tool parses Mermaid flowchart syntax and converts it to native
    PowerPoint vector shapes (not images!), allowing further editing.
    
    Supported Mermaid syntax:
    - Flowcharts: graph TD/LR/BT/RL
    - Node shapes: [rect], (rounded), {diamond}, ((circle)), [[database]]
    - Edges: -->, ---, -.->
    - Edge labels: -->|label|
    
    Args:
        slide_index: Index of the target slide
        mermaid_code: Mermaid diagram code (e.g., "graph TD\\nA[Start] --> B[End]")
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with rendering results including node count and edge count
        
    Example:
        add_mermaid_diagram(0, '''
            graph TD
            A[Start] --> B{Decision}
            B -->|Yes| C[Process]
            B -->|No| D[End]
        ''')
    """
    logger.info(f"Rendering Mermaid diagram on slide {slide_index}")
    return diagram_renderer.render_mermaid(
        slide_index=slide_index,
        mermaid_code=mermaid_code,
        presentation_id=presentation_id
    )


@mcp.tool()
def add_plantuml_diagram(
    slide_index: int,
    plantuml_code: str,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Render a PlantUML diagram as editable PowerPoint shapes.
    
    This tool parses PlantUML activity diagram syntax and converts it to native
    PowerPoint vector shapes (not images!), allowing further editing.
    
    Supported PlantUML syntax:
    - Activity diagrams: start, stop, :action;
    - Conditionals: if/then/else/endif
    - Arrows between nodes: A --> B
    
    Args:
        slide_index: Index of the target slide
        plantuml_code: PlantUML diagram code
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with rendering results including node count and edge count
        
    Example:
        add_plantuml_diagram(0, '''
            @startuml
            start
            :Initialize;
            if (Valid?) then (yes)
                :Process;
            else (no)
                :Error;
            endif
            stop
            @enduml
        ''')
    """
    logger.info(f"Rendering PlantUML diagram on slide {slide_index}")
    return diagram_renderer.render_plantuml(
        slide_index=slide_index,
        plantuml_code=plantuml_code,
        presentation_id=presentation_id
    )


@mcp.tool()
def add_diagram(
    slide_index: int,
    diagram_code: str,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Automatically detect and render a diagram from Mermaid or PlantUML code.
    
    This tool automatically detects whether the input is Mermaid or PlantUML
    syntax and renders it as editable PowerPoint vector shapes.
    
    Perfect for AI/LLM workflows that generate diagrams in text-based DSLs.
    The diagrams are converted to native PowerPoint shapes (not images!).
    
    Args:
        slide_index: Index of the target slide
        diagram_code: Diagram code in Mermaid or PlantUML syntax
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with rendering results including:
        - detected_type: 'mermaid' or 'plantuml'
        - diagram_type: Type of diagram (flow, hierarchy)
        - node_count: Number of nodes rendered
        - edge_count: Number of connections rendered
        - shapes: List of created shape indices
        
    Example (Mermaid):
        add_diagram(0, '''
            graph LR
            A[Input] --> B[Process] --> C[Output]
        ''')
        
    Example (PlantUML):
        add_diagram(0, '''
            @startuml
            start
            :Step 1;
            :Step 2;
            stop
            @enduml
        ''')
    """
    logger.info(f"Rendering auto-detected diagram on slide {slide_index}")
    return diagram_renderer.render_auto(
        slide_index=slide_index,
        diagram_code=diagram_code,
        presentation_id=presentation_id
    )


# ============================================================================
# SPECIALIZED BUSINESS DIAGRAM TOOLS
# ============================================================================

@mcp.tool()
def create_swot_analysis(
    slide_index: int,
    strengths: List[str],
    weaknesses: List[str],
    opportunities: List[str],
    threats: List[str],
    title: Optional[str] = None,
    show_labels: bool = True,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Create a professional SWOT analysis diagram from structured data.
    
    Automatically generates a 2x2 grid layout with professional styling:
    - Top-left: Strengths (green)
    - Top-right: Weaknesses (red)
    - Bottom-left: Opportunities (blue)
    - Bottom-right: Threats (yellow/orange)
    
    The AI only needs to provide the raw data lists - the visual arrangement,
    colors, and positioning are handled automatically.
    
    Args:
        slide_index: Index of the target slide
        strengths: List of strength items (e.g., ["Strong brand", "Loyal customers"])
        weaknesses: List of weakness items (e.g., ["High costs", "Limited reach"])
        opportunities: List of opportunity items (e.g., ["New markets", "Digital growth"])
        threats: List of threat items (e.g., ["Competition", "Regulations"])
        title: Optional title for the SWOT diagram
        show_labels: Whether to show category labels (Strengths, Weaknesses, etc.)
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with creation results including shape indices and item counts
        
    Example:
        create_swot_analysis(
            slide_index=0,
            strengths=["Strong brand recognition", "Skilled workforce", "Patent portfolio"],
            weaknesses=["High production costs", "Limited global presence"],
            opportunities=["Emerging markets in Asia", "E-commerce expansion"],
            threats=["Aggressive competitors", "Supply chain disruptions"],
            title="Company SWOT Analysis"
        )
    """
    logger.info(f"Creating SWOT analysis on slide {slide_index}")
    return business_diagrams.create_swot_analysis(
        slide_index=slide_index,
        strengths=strengths,
        weaknesses=weaknesses,
        opportunities=opportunities,
        threats=threats,
        title=title,
        show_labels=show_labels,
        presentation_id=presentation_id
    )


@mcp.tool()
def create_timeline(
    slide_index: int,
    events: List[Dict[str, Any]],
    direction: str = "horizontal",
    title: Optional[str] = None,
    show_connector: bool = True,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Create a professional timeline diagram from a list of events.
    
    Automatically generates a timeline with:
    - Central connector line
    - Event markers (circles) at each point
    - Labels with dates and descriptions
    - Alternating positions for readability (horizontal mode)
    
    The AI only needs to provide the event data - layout and styling are automatic.
    
    Args:
        slide_index: Index of the target slide
        events: List of event dictionaries, each containing:
            - label (required): Event name/title
            - date (optional): Date or time period string
            - description (optional): Additional details
            - color (optional): Semantic tag ('success', 'warning') or RGB list [r, g, b]
        direction: 'horizontal' (left to right) or 'vertical' (top to bottom)
        title: Optional title for the timeline
        show_connector: Whether to show the connecting line between events
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with creation results including shape indices and event count
        
    Example:
        create_timeline(
            slide_index=0,
            events=[
                {"label": "Project Kickoff", "date": "Jan 2024", "color": "success"},
                {"label": "Phase 1 Complete", "date": "Mar 2024"},
                {"label": "Beta Launch", "date": "May 2024"},
                {"label": "GA Release", "date": "Jul 2024", "color": "accent"}
            ],
            direction="horizontal",
            title="Product Roadmap 2024"
        )
    """
    logger.info(f"Creating {direction} timeline on slide {slide_index}")
    return business_diagrams.create_timeline(
        slide_index=slide_index,
        events=events,
        direction=direction,
        title=title,
        show_connector=show_connector,
        presentation_id=presentation_id
    )


@mcp.tool()
def create_org_chart(
    slide_index: int,
    root: Dict[str, Any],
    title: Optional[str] = None,
    show_connectors: bool = True,
    compact: bool = False,
    presentation_id: Optional[str] = None
) -> Dict[str, Any]:
    """
    Create a professional organization chart from hierarchical data.
    
    Automatically generates an org chart with:
    - Rectangular boxes for each person/role
    - Connecting lines between hierarchical levels
    - Automatic positioning based on tree structure
    - Name and title displayed in each box
    
    The AI only needs to provide the organizational hierarchy - all positioning
    and styling is handled automatically.
    
    Args:
        slide_index: Index of the target slide
        root: Root node of the organization hierarchy:
            {
                "name": "Person Name",
                "title": "Job Title (optional)",
                "color": semantic tag or RGB list (optional),
                "children": [
                    {
                        "name": "Child Name",
                        "title": "Child Title",
                        "children": [...] (nested children)
                    }
                ]
            }
        title: Optional title for the org chart
        show_connectors: Whether to show connecting lines between levels
        compact: Whether to use compact layout (smaller boxes, tighter spacing)
        presentation_id: Optional ID of the target presentation
        
    Returns:
        Dictionary with creation results including shape indices and level count
        
    Example:
        create_org_chart(
            slide_index=0,
            root={
                "name": "Sarah Johnson",
                "title": "CEO",
                "children": [
                    {
                        "name": "Mike Chen",
                        "title": "VP Engineering",
                        "children": [
                            {"name": "Alice Wong", "title": "Tech Lead"},
                            {"name": "Bob Smith", "title": "Senior Dev"}
                        ]
                    },
                    {
                        "name": "Emily Brown",
                        "title": "VP Marketing",
                        "children": [
                            {"name": "Carol Davis", "title": "Marketing Manager"}
                        ]
                    }
                ]
            },
            title="Company Organization"
        )
    """
    logger.info(f"Creating org chart on slide {slide_index}")
    return business_diagrams.create_org_chart(
        slide_index=slide_index,
        root=root,
        title=title,
        show_connectors=show_connectors,
        compact=compact,
        presentation_id=presentation_id
    )


# ============================================================================
# SERVER STARTUP
# ============================================================================

if __name__ == "__main__":
    logger.info("Starting PowerPoint MCP Server...")
    # Use the built-in run method with streamable-http transport
    mcp.run(transport="streamable-http")
