# PowerPoint MCP Server

This project provides a **professional** Model Context Protocol (MCP) server for creating and manipulating Microsoft PowerPoint presentations (`.pptx` files) programmatically. It exposes a comprehensive set of tools that can be called by an MCP-compatible client to generate professional slides, add advanced shapes, charts, tables, images, and much more.

The server is built using Python with the **FastMCP** framework and the `python-pptx` library, providing enterprise-grade PowerPoint generation capabilities.

## ✨ New Features & Improvements

### 🏗️ **Modular Architecture**
- **Separation of Concerns**: Refactored into specialized modules for better maintainability
- **Presentation Manager**: Centralized presentation lifecycle management
- **Template Manager**: Professional template and styling support
- **Slide Manager**: Advanced slide content and layout management

### 📊 **Professional Content Tools**
- **Charts & Graphs**: Column, line, pie, bar, and area charts with professional styling
- **Data Tables**: Automatically styled tables with header formatting
- **Image Support**: Smart image insertion with aspect ratio preservation
- **Bullet Points**: Professional bullet point formatting in placeholders

### 🎨 **Enhanced Styling & Templates**
- **Template System**: Extract and apply styles from existing presentations
- **Default Theming**: Intelligent color and font defaults
- **Professional Layouts**: Better handling of slide layouts and placeholders

### 🔧 **Developer Experience**
- **Better Error Handling**: Comprehensive error messages and validation
- **Enhanced Logging**: Detailed operation logging for debugging
- **Type Safety**: Improved type hints throughout the codebase
- **Documentation**: Comprehensive docstrings and usage examples

## Setup and Installation

To get started with this server, you need to have Python 3 installed.

1.  **Clone the repository or download the source code.**

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv .venv
    source .venv/bin/activate  # On Windows, use `.venv\Scripts\activate`
    ```

3.  **Install the required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## Running the Server

Once the dependencies are installed, you can start the MCP server by running the following command in your terminal:

```bash
python server.py
```

The server will start and listen for incoming MCP requests on `0.0.0.0:8081`.

## Available Tools

This server provides a comprehensive set of **professional-grade** tools for building PowerPoint presentations. The tools are organized into logical groups for enterprise-level presentation generation.

### 🎨 Template Management

*   **`set_template_presentation(file_path: str) -> Dict`**
    *   **Description:** Set a template presentation and extract its professional styles and themes.
    *   **Parameters:**
        *   `file_path` (str): Path to the template presentation file.
    *   **Returns:** Template information and extracted styles (colors, fonts, themes).

*   **`get_template_styles() -> Dict`**
    *   **Description:** Get currently loaded template styles for consistent presentation design.
    *   **Returns:** Current template path and available style definitions.

### 📋 Presentation Management

These tools handle the creation, opening, saving, and inspection of presentation files with enhanced metadata support.

*   **`create_presentation(id: Optional[str] = None) -> Dict`**
    *   **Description:** Creates a new, empty PowerPoint presentation in memory with professional defaults.
    *   **Parameters:**
        *   `id` (Optional): A unique identifier for the presentation. Auto-generated if not provided.
    *   **Returns:** Presentation ID, confirmation message, and initial slide count.

*   **`open_presentation(file_path: str, id: Optional[str] = None) -> Dict`**
    *   **Description:** Opens an existing `.pptx` file with template style extraction.
    *   **Parameters:**
        *   `file_path` (str): Path to the presentation file.
        *   `id` (Optional): Unique identifier for the opened presentation.
    *   **Returns:** Presentation ID, message, and slide count with layout information.

*   **`save_presentation(file_path: str, presentation_id: Optional[str] = None) -> Dict`**
    *   **Description:** Saves a presentation to a `.pptx` file with automatic path resolution.
    *   **Parameters:**
        *   `file_path` (str): Target save path (automatically resolved to /data/ directory).
        *   `presentation_id` (Optional): ID of presentation to save.
    *   **Returns:** Confirmation message and final file path.

*   **`get_presentation_info(presentation_id: Optional[str] = None) -> Dict`**
    *   **Description:** Comprehensive presentation metadata including layouts, properties, and slide information.
    *   **Returns:** Complete presentation metadata with available layouts and core properties.

*   **`set_core_properties(...) -> Dict`**
    *   **Description:** Sets professional document properties for corporate presentations.
    *   **Parameters:** `title`, `subject`, `author`, `keywords`, `comments`, `presentation_id`.
    *   **Returns:** Confirmation and updated properties.

### 📑 Slide Management

Enhanced slide creation and management with professional layout support.

*   **`add_slide(layout_index: int = 1, title: Optional[str] = None, presentation_id: Optional[str] = None) -> Dict`**
    *   **Description:** Adds a new slide with specified layout and automatic title placement.
    *   **Parameters:**
        *   `layout_index` (int): Layout index with validation and error reporting.
        *   `title` (Optional): Auto-formatted slide title.
    *   **Returns:** Slide index, layout information, and available placeholders.

### ✏️ Text and Shape Tools

Professional text formatting and shape creation with template integration.

*   **`add_textbox(...) -> Dict`**
    *   **Description:** Professional textbox with template-aware formatting and typography.
    *   **Enhanced Features:** Template font defaults, corporate color schemes, advanced alignment.

*   **`add_shape(...) -> Dict`**
    *   **Description:** Auto-shapes with professional styling and template color integration.
    *   **Enhanced Features:** Template-based fill colors, corporate line styles, advanced shape types.

*   **`add_line(...) -> Dict`**
    *   **Description:** Precise line drawing with professional styling options.
    *   **Enhanced Features:** Template color defaults, customizable line weights and styles.

### 📊 Advanced Content Tools (NEW!)

Professional data visualization and content creation tools.

*   **`add_chart(slide_index, chart_type, left, top, width, height, data, ...) -> Dict`**
    *   **Description:** Create professional charts and graphs with corporate styling.
    *   **Chart Types:** Column, Line, Pie, Bar, Area charts.
    *   **Features:** Auto-legend positioning, professional color schemes, data validation.
    *   **Data Format:** `{'categories': [...], 'series': [{'name': '...', 'values': [...]}]}`

*   **`add_table(slide_index, left, top, rows, cols, data, ...) -> Dict`**
    *   **Description:** Professional data tables with automatic header styling.
    *   **Features:** Automatic sizing, header row formatting, corporate color schemes.
    *   **Data Format:** List of lists with first row automatically styled as header.

*   **`add_image(slide_index, image_path, left, top, width, height, ...) -> Dict`**
    *   **Description:** Smart image insertion with aspect ratio preservation.
    *   **Features:** Automatic path resolution, aspect ratio maintenance, flexible sizing.

*   **`add_bullet_points(slide_index, placeholder_idx, bullet_points, ...) -> Dict`**
    *   **Description:** Professional bullet point formatting in slide placeholders.
    *   **Features:** Template font integration, consistent bullet styling, automatic spacing.

## 🚀 Professional Use Cases

This enhanced server is designed for enterprise-level PowerPoint generation including:

- **Corporate Presentations**: Professional layouts with consistent branding
- **Data Visualization**: Charts and graphs for business intelligence
- **Report Generation**: Automated reporting with tables and data insights  
- **Marketing Materials**: Professional slide decks with images and styling
- **Training Materials**: Educational content with bullet points and media

## 🔧 Extensibility

The modular architecture makes it easy to add new professional features:

- Add new tools in the appropriate manager module (`presentation_manager.py`, `template_manager.py`, `slide_manager.py`)
- Expose new functionality through `@mcp.tool()` decorated functions in `server.py`
- Leverage the template system for consistent professional styling
- Use the logging framework for debugging and monitoring

## 🎯 Migration Notes

This version represents a **significant upgrade** from the previous implementation:

1. **Already Using FastMCP**: The project was already using FastMCP framework from `mcp.server.fastmcp`
2. **Enhanced FastMCP Implementation**: Improved with better organization, error handling, and professional features
3. **Modular Design**: Refactored for maintainability and extensibility
4. **Professional Features**: Added enterprise-grade presentation capabilities
5. **Better Developer Experience**: Enhanced documentation, logging, and type safety

The "migration to fastMCP" has been completed with significant enhancements to create professional PowerPoint presentations suitable for corporate and business use.
