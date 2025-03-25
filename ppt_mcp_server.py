#!/usr/bin/env python
"""
MCP Server for PowerPoint manipulation using python-pptx.
"""
import os
import json
import tempfile
from typing import Dict, List, Optional, Any, Union
from mcp.server.fastmcp import FastMCP

import ppt_utils

# Initialize the FastMCP server
app = FastMCP(
    name="ppt-mcp-server",
    description="MCP Server for PowerPoint manipulation using python-pptx",
    version="1.0.0"
)

# Global state to store presentations in memory
presentations = {}
current_presentation_id = None

# ---- Helper Functions ----

def get_current_presentation():
    """Get the current presentation object or raise an error if none is loaded."""
    if current_presentation_id is None or current_presentation_id not in presentations:
        raise ValueError("No presentation is currently loaded. Please create or open a presentation first.")
    return presentations[current_presentation_id]

# ---- Presentation Tools ----

@app.tool()
def create_presentation(id: Optional[str] = None) -> Dict:
    """Create a new PowerPoint presentation."""
    global current_presentation_id
    
    # Create a new presentation
    pres = ppt_utils.create_presentation()
    
    # Generate an ID if not provided
    if id is None:
        id = f"presentation_{len(presentations) + 1}"
    
    # Store the presentation
    presentations[id] = pres
    current_presentation_id = id
    
    return {
        "presentation_id": id,
        "message": f"Created new presentation with ID: {id}",
        "slide_count": len(pres.slides)
    }

@app.tool()
def open_presentation(file_path: str, id: Optional[str] = None) -> Dict:
    """Open an existing PowerPoint presentation from a file."""
    global current_presentation_id
    
    # Check if file exists
    if not os.path.exists(file_path):
        return {
            "error": f"File not found: {file_path}"
        }
    
    # Open the presentation
    try:
        pres = ppt_utils.open_presentation(file_path)
    except Exception as e:
        return {
            "error": f"Failed to open presentation: {str(e)}"
        }
    
    # Generate an ID if not provided
    if id is None:
        id = f"presentation_{len(presentations) + 1}"
    
    # Store the presentation
    presentations[id] = pres
    current_presentation_id = id
    
    return {
        "presentation_id": id,
        "message": f"Opened presentation from {file_path} with ID: {id}",
        "slide_count": len(pres.slides)
    }

@app.tool()
def save_presentation(file_path: str, presentation_id: Optional[str] = None) -> Dict:
    """Save a presentation to a file."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    # Save the presentation
    try:
        saved_path = ppt_utils.save_presentation(presentations[pres_id], file_path)
        return {
            "message": f"Presentation saved to {saved_path}",
            "file_path": saved_path
        }
    except Exception as e:
        return {
            "error": f"Failed to save presentation: {str(e)}"
        }

@app.tool()
def get_presentation_info(presentation_id: Optional[str] = None) -> Dict:
    """Get information about a presentation."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Get slide layouts
    layouts = ppt_utils.get_slide_layouts(pres)
    
    # Get core properties
    core_props = ppt_utils.get_core_properties(pres)
    
    return {
        "presentation_id": pres_id,
        "slide_count": len(pres.slides),
        "slide_layouts": layouts,
        "core_properties": core_props
    }

@app.tool()
def set_core_properties(
    title: Optional[str] = None,
    subject: Optional[str] = None,
    author: Optional[str] = None,
    keywords: Optional[str] = None,
    comments: Optional[str] = None,
    presentation_id: Optional[str] = None
) -> Dict:
    """Set core document properties."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Set core properties
    try:
        ppt_utils.set_core_properties(
            pres, title=title, subject=subject, author=author, 
            keywords=keywords, comments=comments
        )
        
        # Get updated properties
        updated_props = ppt_utils.get_core_properties(pres)
        
        return {
            "message": "Core properties updated successfully",
            "core_properties": updated_props
        }
    except Exception as e:
        return {
            "error": f"Failed to set core properties: {str(e)}"
        }

# ---- Slide Tools ----

@app.tool()
def add_slide(
    layout_index: int = 1,
    title: Optional[str] = None,
    presentation_id: Optional[str] = None
) -> Dict:
    """Add a new slide to the presentation."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    try:
        # Check if layout index is valid
        if layout_index < 0 or layout_index >= len(pres.slide_layouts):
            return {
                "error": f"Invalid layout index: {layout_index}. Available layouts: 0-{len(pres.slide_layouts) - 1}"
            }
        
        # Add the slide
        slide, layout = ppt_utils.add_slide(pres, layout_index)
        
        # Set the title if provided
        if title and slide.shapes.title:
            ppt_utils.set_title(slide, title)
        
        # Get placeholders
        placeholders = ppt_utils.get_placeholders(slide)
        
        return {
            "message": f"Added slide with layout '{layout.name}'",
            "slide_index": len(pres.slides) - 1,
            "layout_name": layout.name,
            "placeholders": placeholders
        }
    except Exception as e:
        return {
            "error": f"Failed to add slide: {str(e)}"
        }

@app.tool()
def get_slide_info(slide_index: int, presentation_id: Optional[str] = None) -> Dict:
    """Get information about a specific slide."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Check if slide index is valid
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
        }
    
    slide = pres.slides[slide_index]
    
    # Get placeholders
    placeholders = ppt_utils.get_placeholders(slide)
    
    # Get shapes information
    shapes_info = []
    for i, shape in enumerate(slide.shapes):
        shape_info = {
            "index": i,
            "name": shape.name,
            "shape_type": str(shape.shape_type),
            "width": shape.width.inches,
            "height": shape.height.inches,
            "left": shape.left.inches,
            "top": shape.top.inches
        }
        shapes_info.append(shape_info)
    
    return {
        "slide_index": slide_index,
        "placeholders": placeholders,
        "shapes": shapes_info
    }

@app.tool()
def populate_placeholder(
    slide_index: int,
    placeholder_idx: int,
    text: str,
    presentation_id: Optional[str] = None
) -> Dict:
    """Populate a placeholder with text."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Check if slide index is valid
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
        }
    
    slide = pres.slides[slide_index]
    
    try:
        # Check if placeholder exists
        if placeholder_idx not in [p.placeholder_format.idx for p in slide.placeholders]:
            return {
                "error": f"Placeholder with index {placeholder_idx} not found in slide {slide_index}"
            }
        
        # Populate the placeholder
        ppt_utils.populate_placeholder(slide, placeholder_idx, text)
        
        return {
            "message": f"Populated placeholder {placeholder_idx} in slide {slide_index}"
        }
    except Exception as e:
        return {
            "error": f"Failed to populate placeholder: {str(e)}"
        }

@app.tool()
def add_bullet_points(
    slide_index: int,
    placeholder_idx: int,
    bullet_points: List[str],
    presentation_id: Optional[str] = None
) -> Dict:
    """Add bullet points to a placeholder."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Check if slide index is valid
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
        }
    
    slide = pres.slides[slide_index]
    
    try:
        # Check if placeholder exists
        if placeholder_idx not in [p.placeholder_format.idx for p in slide.placeholders]:
            return {
                "error": f"Placeholder with index {placeholder_idx} not found in slide {slide_index}"
            }
        
        # Get the placeholder
        placeholder = slide.placeholders[placeholder_idx]
        
        # Add bullet points
        ppt_utils.add_bullet_points(placeholder, bullet_points)
        
        return {
            "message": f"Added {len(bullet_points)} bullet points to placeholder {placeholder_idx} in slide {slide_index}"
        }
    except Exception as e:
        return {
            "error": f"Failed to add bullet points: {str(e)}"
        }

# ---- Text Tools ----

@app.tool()
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
) -> Dict:
    """Add a textbox to a slide."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Check if slide index is valid
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
        }
    
    slide = pres.slides[slide_index]
    
    try:
        # Add the textbox
        textbox = ppt_utils.add_textbox(slide, left, top, width, height, text)
        
        # Format the text if formatting options are provided
        if any([font_size, font_name, bold, italic, color, alignment]):
            ppt_utils.format_text(
                textbox.text_frame,
                font_size=font_size,
                font_name=font_name,
                bold=bold,
                italic=italic,
                color=tuple(color) if color else None,
                alignment=alignment
            )
        
        return {
            "message": f"Added textbox to slide {slide_index}",
            "shape_index": len(slide.shapes) - 1
        }
    except Exception as e:
        return {
            "error": f"Failed to add textbox: {str(e)}"
        }

# ---- Image Tools ----

@app.tool()
def add_image(
    slide_index: int,
    image_path: str,
    left: float,
    top: float,
    width: Optional[float] = None,
    height: Optional[float] = None,
    presentation_id: Optional[str] = None
) -> Dict:
    """Add an image to a slide."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Check if slide index is valid
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
        }
    
    slide = pres.slides[slide_index]
    
    # Check if image file exists
    if not os.path.exists(image_path):
        return {
            "error": f"Image file not found: {image_path}"
        }
    
    try:
        # Add the image
        picture = ppt_utils.add_image(slide, image_path, left, top, width, height)
        
        return {
            "message": f"Added image to slide {slide_index}",
            "shape_index": len(slide.shapes) - 1,
            "width": picture.width.inches,
            "height": picture.height.inches
        }
    except Exception as e:
        return {
            "error": f"Failed to add image: {str(e)}"
        }

@app.tool()
def add_image_from_base64(
    slide_index: int,
    base64_string: str,
    left: float,
    top: float,
    width: Optional[float] = None,
    height: Optional[float] = None,
    presentation_id: Optional[str] = None
) -> Dict:
    """Add an image from a base64 encoded string to a slide."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Check if slide index is valid
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
        }
    
    slide = pres.slides[slide_index]
    
    try:
        # Add the image
        picture = ppt_utils.add_image_from_base64(slide, base64_string, left, top, width, height)
        
        return {
            "message": f"Added image to slide {slide_index}",
            "shape_index": len(slide.shapes) - 1,
            "width": picture.width.inches,
            "height": picture.height.inches
        }
    except Exception as e:
        return {
            "error": f"Failed to add image: {str(e)}"
        }

# ---- Table Tools ----

@app.tool()
def add_table(
    slide_index: int,
    rows: int,
    cols: int,
    left: float,
    top: float,
    width: float,
    height: float,
    data: Optional[List[List[str]]] = None,
    presentation_id: Optional[str] = None
) -> Dict:
    """Add a table to a slide."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Check if slide index is valid
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
        }
    
    slide = pres.slides[slide_index]
    
    try:
        # Add the table
        table = ppt_utils.add_table(slide, rows, cols, left, top, width, height)
        
        # Populate the table if data is provided
        if data:
            for row_idx, row_data in enumerate(data):
                if row_idx >= rows:
                    break
                    
                for col_idx, cell_text in enumerate(row_data):
                    if col_idx >= cols:
                        break
                        
                    ppt_utils.set_cell_text(table, row_idx, col_idx, cell_text)
        
        return {
            "message": f"Added {rows}x{cols} table to slide {slide_index}",
            "shape_index": len(slide.shapes) - 1
        }
    except Exception as e:
        return {
            "error": f"Failed to add table: {str(e)}"
        }

@app.tool()
def format_table_cell(
    slide_index: int,
    shape_index: int,
    row: int,
    col: int,
    font_size: Optional[int] = None,
    font_name: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[List[int]] = None,
    bg_color: Optional[List[int]] = None,
    alignment: Optional[str] = None,
    vertical_alignment: Optional[str] = None,
    presentation_id: Optional[str] = None
) -> Dict:
    """Format a table cell."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Check if slide index is valid
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
        }
    
    slide = pres.slides[slide_index]
    
    # Check if shape index is valid
    if shape_index < 0 or shape_index >= len(slide.shapes):
        return {
            "error": f"Invalid shape index: {shape_index}. Available shapes: 0-{len(slide.shapes) - 1}"
        }
    
    shape = slide.shapes[shape_index]
    
    try:
        # Check if shape is a table
        if not hasattr(shape, 'table'):
            return {
                "error": f"Shape at index {shape_index} is not a table"
            }
        
        table = shape.table
        
        # Check if row and column indices are valid
        if row < 0 or row >= len(table.rows):
            return {
                "error": f"Invalid row index: {row}. Available rows: 0-{len(table.rows) - 1}"
            }
            
        if col < 0 or col >= len(table.columns):
            return {
                "error": f"Invalid column index: {col}. Available columns: 0-{len(table.columns) - 1}"
            }
        
        # Get the cell
        cell = table.cell(row, col)
        
        # Format the cell
        ppt_utils.format_table_cell(
            cell,
            font_size=font_size,
            font_name=font_name,
            bold=bold,
            italic=italic,
            color=tuple(color) if color else None,
            bg_color=tuple(bg_color) if bg_color else None,
            alignment=alignment,
            vertical_alignment=vertical_alignment
        )
        
        return {
            "message": f"Formatted cell at row {row}, column {col} in table at shape index {shape_index} on slide {slide_index}"
        }
    except Exception as e:
        return {
            "error": f"Failed to format table cell: {str(e)}"
        }

# ---- Shape Tools ----

@app.tool()
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
) -> Dict:
    """Add an auto shape to a slide."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Check if slide index is valid
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
        }
    
    slide = pres.slides[slide_index]
    
    try:
        # Add the shape
        shape = ppt_utils.add_shape(slide, shape_type, left, top, width, height)
        
        # Format the shape if formatting options are provided
        if any([fill_color, line_color, line_width]):
            ppt_utils.format_shape(
                shape,
                fill_color=tuple(fill_color) if fill_color else None,
                line_color=tuple(line_color) if line_color else None,
                line_width=line_width
            )
        
        return {
            "message": f"Added {shape_type} shape to slide {slide_index}",
            "shape_index": len(slide.shapes) - 1
        }
    except Exception as e:
        return {
            "error": f"Failed to add shape: {str(e)}"
        }

# ---- Chart Tools ----

@app.tool()
def add_chart(
    slide_index: int,
    chart_type: str,
    left: float,
    top: float,
    width: float,
    height: float,
    categories: List[str],
    series_names: List[str],
    series_values: List[List[float]],
    has_legend: bool = True,
    legend_position: str = "right",
    has_data_labels: bool = False,
    title: Optional[str] = None,
    presentation_id: Optional[str] = None
) -> Dict:
    """Add a chart to a slide."""
    # Use the specified presentation or the current one
    pres_id = presentation_id if presentation_id is not None else current_presentation_id
    
    if pres_id is None or pres_id not in presentations:
        return {
            "error": "No presentation is currently loaded or the specified ID is invalid"
        }
    
    pres = presentations[pres_id]
    
    # Check if slide index is valid
    if slide_index < 0 or slide_index >= len(pres.slides):
        return {
            "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
        }
    
    slide = pres.slides[slide_index]
    
    # Validate series data
    if len(series_names) != len(series_values):
        return {
            "error": f"Number of series names ({len(series_names)}) must match number of series values ({len(series_values)})"
        }
    
    try:
        # Add the chart
        chart = ppt_utils.add_chart(
            slide, chart_type, left, top, width, height,
            categories, series_names, series_values
        )
        
        # Format the chart
        ppt_utils.format_chart(
            chart,
            has_legend=has_legend,
            legend_position=legend_position,
            has_data_labels=has_data_labels,
            title=title
        )
        
        return {
            "message": f"Added {chart_type} chart to slide {slide_index}",
            "shape_index": len(slide.shapes) - 1
        }
    except Exception as e:
        return {
            "error": f"Failed to add chart: {str(e)}"
        }

# ---- Main Execution ----

if __name__ == "__main__":
    # Run the FastMCP server
    app.run(transport='stdio')
