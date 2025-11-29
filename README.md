# PowerPoint MCP Server

This project provides a **professional** Model Context Protocol (MCP) server for creating and manipulating Microsoft PowerPoint presentations (`.pptx` files) programmatically. It exposes a comprehensive set of tools that can be called by an MCP-compatible client to generate professional slides, add advanced shapes, charts, tables, images, and much more.

The server is built using Python with the **FastMCP** framework and the `python-pptx` library, providing enterprise-grade PowerPoint generation capabilities.

## ✨ New Features & Improvements

### 📈 **Specialized Business Diagrams** (NEW!)
- **SWOT Analysis**: Create professional 2x2 SWOT diagrams from structured lists
- **Timeline**: Generate horizontal/vertical timelines from event data
- **Organization Chart**: Build org charts from hierarchical data structures
- **AI-Friendly**: Just provide the raw data - layout and styling are automatic

### 📊 **Diagram Support (Mermaid/PlantUML)**
- **Text-Based Diagrams**: Parse Mermaid and PlantUML syntax directly into editable PowerPoint shapes
- **AI-Friendly DSLs**: Let AI describe diagrams in familiar text formats instead of specifying individual shapes
- **Native Vector Output**: All diagrams are rendered as editable PowerPoint shapes, not images
- **Auto-Layout**: Automatic positioning based on diagram structure
- **Multiple Formats**: Support for flowcharts, activity diagrams, hierarchies, and process flows

### 🧠 **Intelligent Text Auto-Fit**
- **Smart Content Handling**: Automatically adjusts font size, columns, or slide splits for extensive AI-generated content
- **Multi-Column Layout**: Distributes long text across columns for better readability
- **Auto Slide Splitting**: Intelligently splits content across multiple slides when needed
- **Maximum Readability**: Ensures optimal text presentation without manual formatting

### 🏗️ **High-Level Layout Engine**
- **AI-Friendly Design**: Create complex layouts without specifying pixel coordinates
- **Grid Layout**: Automatic grid arrangement for dashboard-style content
- **List Layout**: Vertical and horizontal lists with automatic spacing
- **Hierarchy Layout**: Organization charts and tree structures with connectors
- **Flow Layout**: Process diagrams and workflows with arrows

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

The server will start and listen for incoming MCP requests on `0.0.0.0:8081` using the **Streamable HTTP** transport protocol. The MCP endpoint is available at `/mcp`.

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

### 🧠 Intelligent Text Auto-Fit (NEW!)

**AI-Optimized Text Handling** - Automatically adjusts text presentation for extensive AI-generated content. Maximizes readability and creates sensible slide divisions for large data sets.

*   **`add_auto_fit_text(slide_index, left, top, width, height, text, strategy, ...) -> Dict`**
    *   **Description:** Add text with intelligent auto-fit that handles extensive content.
    *   **Strategies:**
        *   `smart`: Automatically choose the best approach (default)
        *   `shrink_font`: Reduce font size to fit content
        *   `multi_column`: Split content into multiple columns
        *   `split_slides`: Distribute content across multiple slides
    *   **Features:**
        *   Automatic font size calculation based on content and container
        *   Multi-column layout for better readability
        *   Automatic slide splitting with customizable titles
        *   Preserves paragraphs and logical content sections
    *   **Example:**
        ```python
        add_auto_fit_text(
            slide_index=0,
            left=0.5, top=1.5, width=9.0, height=5.0,
            text="Very long AI-generated content...",
            strategy="smart",
            create_new_slides=True,
            slide_title_template="Content (Page {page})"
        )
        ```

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

### 🤖 High-Level Layout Engine (NEW!)

**AI-Friendly Layout Tools** - Create complex diagrams and layouts without specifying coordinates. The layout engine automatically calculates positions based on slide dimensions and structural descriptions.

*   **`add_grid_layout(slide_index, elements, rows, cols, gap, ...) -> Dict`**
    *   **Description:** Arrange elements in a grid without specifying coordinates.
    *   **Parameters:**
        *   `slide_index` (int): Target slide index.
        *   `elements` (List): Element dictionaries with `content`, `fill_color`, `text_color`, etc.
        *   `rows` (int): Number of grid rows (default: 2).
        *   `cols` (int): Number of grid columns (default: 2).
        *   `gap` (float): Gap between cells in inches (default: 0.2).
    *   **Example:**
        ```python
        add_grid_layout(0, [
            {"content": "Q1", "fill_color": [79, 129, 189]},
            {"content": "Q2", "fill_color": [192, 80, 77]},
            {"content": "Q3", "fill_color": [155, 187, 89]},
            {"content": "Q4", "fill_color": [128, 100, 162]}
        ], rows=2, cols=2)
        ```

*   **`add_list_layout(slide_index, elements, direction, gap, alignment, ...) -> Dict`**
    *   **Description:** Arrange elements in a vertical or horizontal list.
    *   **Parameters:**
        *   `elements` (List): Element dictionaries.
        *   `direction` (str): "vertical" or "horizontal".
        *   `alignment` (str): "left", "center", "right" (vertical) or "top", "middle", "bottom" (horizontal).
    *   **Example:**
        ```python
        add_list_layout(0, [
            {"content": "Feature A"},
            {"content": "Feature B"},
            {"content": "Feature C"}
        ], direction="vertical", alignment="left")
        ```

*   **`add_hierarchy_layout(slide_index, root, level_gap, sibling_gap, show_connectors, ...) -> Dict`**
    *   **Description:** Create hierarchical/tree structures like organization charts.
    *   **Parameters:**
        *   `root` (Dict): Root node with `content` and optional `children` list.
        *   `level_gap` (float): Vertical gap between levels in inches.
        *   `sibling_gap` (float): Horizontal gap between siblings.
        *   `show_connectors` (bool): Draw connecting lines (default: True).
    *   **Example:**
        ```python
        add_hierarchy_layout(0, {
            "content": "CEO",
            "children": [
                {"content": "VP Sales", "children": [
                    {"content": "Team A"},
                    {"content": "Team B"}
                ]},
                {"content": "VP Engineering"}
            ]
        })
        ```

*   **`add_flow_layout(slide_index, steps, direction, gap, show_connectors, connector_style, ...) -> Dict`**
    *   **Description:** Create flow/process diagrams with connecting arrows.
    *   **Parameters:**
        *   `steps` (List): Step dictionaries with `content` and styling.
        *   `direction` (str): "horizontal" or "vertical".
        *   `connector_style` (str): "arrow", "line", or "none".
    *   **Example:**
        ```python
        add_flow_layout(0, [
            {"content": "Start"},
            {"content": "Process"},
            {"content": "End"}
        ], direction="horizontal", connector_style="arrow")
        ```

### 📊 Diagram Tools (Mermaid/PlantUML)

These tools allow AI/LLM workflows to create diagrams using familiar text-based DSLs, converting them to editable PowerPoint vector shapes.

*   **`add_mermaid_diagram(slide_index, mermaid_code, ...) -> Dict`**
    *   **Description:** Parse Mermaid flowchart syntax and render as editable PowerPoint shapes.
    *   **Supported Syntax:**
        *   Flowcharts: `graph TD/LR/BT/RL`
        *   Node shapes: `[rect]`, `(rounded)`, `{diamond}`, `((circle))`, `[[database]]`
        *   Edges: `-->`, `---`, `-.->`
        *   Edge labels: `-->|label|`
    *   **Example:**
        ```python
        add_mermaid_diagram(0, '''
            graph TD
            A[Start] --> B{Decision}
            B -->|Yes| C[Process]
            B -->|No| D[End]
        ''')
        ```

*   **`add_plantuml_diagram(slide_index, plantuml_code, ...) -> Dict`**
    *   **Description:** Parse PlantUML activity diagram syntax and render as editable PowerPoint shapes.
    *   **Supported Syntax:**
        *   Activity diagrams: `start`, `stop`, `:action;`
        *   Conditionals: `if/then/else/endif`
        *   Arrows between nodes: `A --> B`
    *   **Example:**
        ```python
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
        ```

*   **`add_diagram(slide_index, diagram_code, ...) -> Dict`**
    *   **Description:** Auto-detect diagram type (Mermaid or PlantUML) and render as editable PowerPoint shapes.
    *   **Features:**
        *   Automatic syntax detection
        *   Native vector output (not images)
        *   Automatic layout and positioning
    *   **Example:**
        ```python
        # Mermaid (auto-detected)
        add_diagram(0, '''
            graph LR
            A[Input] --> B[Process] --> C[Output]
        ''')
        
        # PlantUML (auto-detected)
        add_diagram(0, '''
            @startuml
            start
            :Step 1;
            :Step 2;
            stop
            @enduml
        ''')
        ```

### 📈 Specialized Business Diagrams (NEW!)

High-level APIs for creating common business diagram types from structured data. The AI only needs to provide the raw data (JSON/Dict) - the visual arrangement, colors, and positioning are handled automatically.

*   **`create_swot_analysis(slide_index, strengths, weaknesses, opportunities, threats, ...) -> Dict`**
    *   **Description:** Create a professional SWOT analysis diagram from structured data.
    *   **Features:**
        *   Automatic 2x2 grid layout with professional styling
        *   Color-coded quadrants (Strengths=green, Weaknesses=red, Opportunities=blue, Threats=yellow)
        *   Optional title and category labels
    *   **Parameters:**
        *   `strengths` (List[str]): List of strength items
        *   `weaknesses` (List[str]): List of weakness items
        *   `opportunities` (List[str]): List of opportunity items
        *   `threats` (List[str]): List of threat items
        *   `title` (Optional[str]): Diagram title
        *   `show_labels` (bool): Whether to show category labels (default: True)
    *   **Example:**
        ```python
        create_swot_analysis(
            slide_index=0,
            strengths=["Strong brand", "Skilled workforce", "Patent portfolio"],
            weaknesses=["High production costs", "Limited global presence"],
            opportunities=["Emerging markets in Asia", "E-commerce expansion"],
            threats=["Aggressive competitors", "Supply chain disruptions"],
            title="Company SWOT Analysis"
        )
        ```

*   **`create_timeline(slide_index, events, direction, ...) -> Dict`**
    *   **Description:** Create a professional timeline diagram from a list of events.
    *   **Features:**
        *   Horizontal or vertical layout
        *   Central connector line with event markers
        *   Labels with dates and descriptions
        *   Custom colors per event using semantic tags or RGB
    *   **Parameters:**
        *   `events` (List[Dict]): List of event dictionaries, each with:
            *   `label` (required): Event name/title
            *   `date` (optional): Date or time period
            *   `description` (optional): Additional details
            *   `color` (optional): Semantic tag ('success', 'warning') or RGB list
        *   `direction` (str): 'horizontal' or 'vertical' (default: 'horizontal')
        *   `title` (Optional[str]): Timeline title
        *   `show_connector` (bool): Show connecting line (default: True)
    *   **Example:**
        ```python
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
        ```

*   **`create_org_chart(slide_index, root, ...) -> Dict`**
    *   **Description:** Create a professional organization chart from hierarchical data.
    *   **Features:**
        *   Automatic tree layout with connecting lines
        *   Rectangular boxes with name and title
        *   Support for multiple levels of hierarchy
        *   Compact mode for larger organizations
    *   **Parameters:**
        *   `root` (Dict): Root node with:
            *   `name` (required): Person name
            *   `title` (optional): Job title
            *   `children` (optional): List of child nodes (same structure)
            *   `color` (optional): Semantic tag or RGB list
        *   `title` (Optional[str]): Chart title
        *   `show_connectors` (bool): Show connecting lines (default: True)
        *   `compact` (bool): Use compact layout (default: False)
    *   **Example:**
        ```python
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
        ```

## 🚀 Professional Use Cases

This enhanced server is designed for enterprise-level PowerPoint generation including:

- **Corporate Presentations**: Professional layouts with consistent branding
- **Data Visualization**: Charts and graphs for business intelligence
- **Report Generation**: Automated reporting with tables and data insights  
- **Marketing Materials**: Professional slide decks with images and styling
- **Training Materials**: Educational content with bullet points and media
- **AI-Generated Diagrams**: LLM-friendly layouts without coordinate specifications

## 🔧 Extensibility

The modular architecture makes it easy to add new professional features:

- Add new tools in the appropriate manager module (`presentation_manager.py`, `template_manager.py`, `slide_manager.py`, `layout_manager.py`)
- Expose new functionality through `@mcp.tool()` decorated functions in `server.py`
- Leverage the template system for consistent professional styling
- Use the logging framework for debugging and monitoring
- Extend the layout engine with new layout types

## 🎯 Migration Notes

This version represents a **significant upgrade** from the previous implementation:

1. **Already Using FastMCP**: The project was already using FastMCP framework from `mcp.server.fastmcp`
2. **Enhanced FastMCP Implementation**: Improved with better organization, error handling, and professional features
3. **Modular Design**: Refactored for maintainability and extensibility
4. **Professional Features**: Added enterprise-grade presentation capabilities
5. **Better Developer Experience**: Enhanced documentation, logging, and type safety

The "migration to fastMCP" has been completed with significant enhancements to create professional PowerPoint presentations suitable for corporate and business use.
