# PowerPoint MCP Server

This project provides a Model Context Protocol (MCP) server for creating and manipulating Microsoft PowerPoint presentations (`.pptx` files) programmatically. It exposes a set of tools that can be called by an MCP-compatible client to generate slides, add shapes, text, and more.

The server is built using Python with the `FastMCP` framework and the `python-pptx` library.

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

The server will start and listen for incoming MCP requests on `0.0.0.0:8000`.

## Available Tools

This server provides a rich set of tools for building PowerPoint presentations. The tools are organized into logical groups.

### Presentation Management

These tools handle the creation, opening, saving, and inspection of presentation files.

*   **`create_presentation(id: Optional[str] = None) -> Dict`**
    *   **Description:** Creates a new, empty PowerPoint presentation in memory.
    *   **Parameters:**
        *   `id` (Optional): A unique identifier for the presentation. If not provided, one will be generated automatically.
    *   **Returns:** A dictionary containing the `presentation_id`, a confirmation message, and the initial slide count.

*   **`open_presentation(file_path: str, id: Optional[str] = None) -> Dict`**
    *   **Description:** Opens an existing `.pptx` file from the local filesystem.
    *   **Parameters:**
        *   `file_path` (str): The path to the presentation file to open.
        *   `id` (Optional): A unique identifier to assign to the opened presentation.
    *   **Returns:** A dictionary with the `presentation_id`, a message, and the slide count.

*   **`save_presentation(file_path: str, presentation_id: Optional[str] = None) -> Dict`**
    *   **Description:** Saves a presentation from memory to a `.pptx` file.
    *   **Parameters:**
        *   `file_path` (str): The path where the presentation will be saved.
        *   `presentation_id` (Optional): The ID of the presentation to save. If not provided, the currently active presentation will be used.
    *   **Returns:** A confirmation message and the final file path.

*   **`get_presentation_info(presentation_id: Optional[str] = None) -> Dict`**
    *   **Description:** Retrieves metadata about a presentation, including slide count, available slide layouts, and core properties.
    *   **Parameters:**
        *   `presentation_id` (Optional): The ID of the presentation to inspect.
    *   **Returns:** A dictionary containing the presentation's metadata.

*   **`set_core_properties(...) -> Dict`**
    *   **Description:** Sets the core document properties of a presentation (e.g., title, author, subject).
    *   **Parameters:** `title`, `subject`, `author`, `keywords`, `comments`, `presentation_id`.
    *   **Returns:** A confirmation message and the updated properties.

### Slide Management

These tools are used to add and manage slides within a presentation.

*   **`add_slide(layout_index: int = 1, title: Optional[str] = None, presentation_id: Optional[str] = None) -> Dict`**
    *   **Description:** Adds a new slide to the presentation using a specified layout.
    *   **Parameters:**
        *   `layout_index` (int): The index of the slide layout to use (0-based).
        *   `title` (Optional): The title to set for the new slide.
        *   `presentation_id` (Optional): The ID of the target presentation.
    *   **Returns:** A dictionary with a confirmation message, the new slide's index, and layout information.

### Shape and Text Tools

These tools allow you to draw shapes and add text to slides.

*   **`add_shape(...) -> Dict`**
    *   **Description:** Adds an auto-shape (e.g., rectangle, oval, arrow) to a specified slide.
    *   **Parameters:** `slide_index`, `shape_type`, `left`, `top`, `width`, `height`, `fill_color`, `line_color`, `line_width`, `presentation_id`.
    *   **Returns:** A confirmation message and the index of the newly created shape.

*   **`add_textbox(...) -> Dict`**
    *   **Description:** Adds a textbox to a slide and allows for text formatting.
    *   **Parameters:** `slide_index`, `left`, `top`, `width`, `height`, `text`, `font_size`, `font_name`, `bold`, `italic`, `color`, `alignment`, `presentation_id`.
    *   **Returns:** A confirmation message and the index of the new textbox shape.

This project is designed to be extensible. New tools can be easily added by defining new functions in `ppt_utils.py` and exposing them through new `@mcp.tool()` decorated functions in `server.py`.
