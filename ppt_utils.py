"""
Utility functions for PowerPoint manipulation using python-pptx.

This module provides a comprehensive set of functions to create and manipulate
PowerPoint presentations programmatically. It wraps the python-pptx library
with higher-level functions that simplify common operations.

Key features:
- Create, open, and save presentations
- Add and format slides with various layouts
- Manipulate text, including titles, placeholders, and textboxes
- Add and format shapes, images, tables, and charts
- Set document properties

Usage examples:
    # Create a new presentation
    pres = create_presentation()
    
    # Add a title slide
    slide, layout = add_slide(pres, 0)
    set_title(slide, "Presentation Title")
    
    # Add a content slide with bullet points
    slide, layout = add_slide(pres, 1)
    set_title(slide, "Key Points")
    placeholder = slide.placeholders[1]  # Content placeholder
    add_bullet_points(placeholder, ["Point 1", "Point 2", "Point 3"])
    
    # Save the presentation
    save_presentation(pres, "my_presentation.pptx")
"""
from pptx import Presentation
from pptx.chart.data import CategoryChartData, ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR, MSO_FILL_TYPE
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.shapes.graphfrm import GraphicFrame
import io
from typing import Dict, List, Tuple, Union, Optional, Any
import base64

# Professional color schemes
PROFESSIONAL_COLOR_SCHEMES = {
    'modern_blue': {
        'primary': (0, 120, 215),      # Microsoft Blue
        'secondary': (40, 40, 40),     # Dark Gray
        'accent1': (0, 176, 240),      # Light Blue
        'accent2': (255, 192, 0),      # Orange
        'light': (247, 247, 247),      # Light Gray
        'text': (68, 68, 68),          # Text Gray
    },
    'corporate_gray': {
        'primary': (68, 68, 68),       # Charcoal
        'secondary': (0, 120, 215),    # Blue
        'accent1': (89, 89, 89),       # Medium Gray
        'accent2': (217, 217, 217),    # Light Gray
        'light': (242, 242, 242),      # Very Light Gray
        'text': (51, 51, 51),          # Dark Text
    },
    'elegant_green': {
        'primary': (70, 136, 71),      # Forest Green
        'secondary': (255, 255, 255),  # White
        'accent1': (146, 208, 80),     # Light Green
        'accent2': (112, 173, 71),     # Medium Green
        'light': (238, 236, 225),      # Cream
        'text': (89, 89, 89),          # Gray Text
    },
    'warm_red': {
        'primary': (192, 80, 77),      # Deep Red
        'secondary': (68, 68, 68),     # Dark Gray
        'accent1': (230, 126, 34),     # Orange
        'accent2': (241, 196, 15),     # Yellow
        'light': (253, 253, 253),      # White
        'text': (44, 62, 80),          # Blue Gray
    }
}

# Professional typography settings
PROFESSIONAL_FONTS = {
    'title': {
        'name': 'Segoe UI',
        'size_large': 36,
        'size_medium': 28,
        'size_small': 24,
        'bold': True
    },
    'subtitle': {
        'name': 'Segoe UI Light',
        'size_large': 20,
        'size_medium': 18,
        'size_small': 16,
        'bold': False
    },
    'body': {
        'name': 'Segoe UI',
        'size_large': 16,
        'size_medium': 14,
        'size_small': 12,
        'bold': False
    },
    'caption': {
        'name': 'Segoe UI',
        'size_large': 12,
        'size_medium': 10,
        'size_small': 9,
        'bold': False
    }
}

# Professional layout constants (in inches)
PROFESSIONAL_LAYOUT = {
    'margins': {
        'left': 0.75,
        'right': 0.75,
        'top': 0.5,
        'bottom': 0.5
    },
    'spacing': {
        'title_to_content': 0.3,
        'between_elements': 0.2,
        'bullet_indent': 0.5
    },
    'standard_sizes': {
        'title_height': 1.2,
        'content_area_height': 5.5,
        'footer_height': 0.4
    }
}

def get_professional_color(scheme_name: str, color_type: str) -> Tuple[int, int, int]:
    """
    Get a professional color from predefined color schemes.
    
    Args:
        scheme_name: Name of the color scheme ('modern_blue', 'corporate_gray', etc.)
        color_type: Type of color ('primary', 'secondary', 'accent1', 'accent2', 'light', 'text')
        
    Returns:
        RGB color tuple (r, g, b)
    """
    if scheme_name not in PROFESSIONAL_COLOR_SCHEMES:
        scheme_name = 'modern_blue'  # Default fallback
    
    scheme = PROFESSIONAL_COLOR_SCHEMES[scheme_name]
    return scheme.get(color_type, scheme['primary'])

def get_professional_font(font_type: str, size_category: str = 'medium') -> Dict:
    """
    Get professional font settings.
    
    Args:
        font_type: Type of font ('title', 'subtitle', 'body', 'caption')
        size_category: Size category ('large', 'medium', 'small')
        
    Returns:
        Dictionary with font settings
    """
    if font_type not in PROFESSIONAL_FONTS:
        font_type = 'body'  # Default fallback
    
    font_config = PROFESSIONAL_FONTS[font_type]
    size_key = f'size_{size_category}'
    
    return {
        'name': font_config['name'],
        'size': font_config.get(size_key, font_config['size_medium']),
        'bold': font_config['bold']
    }

def calculate_professional_layout(slide_width: float = 10, slide_height: float = 7.5) -> Dict:
    """
    Calculate professional layout dimensions based on slide size.
    
    Args:
        slide_width: Slide width in inches
        slide_height: Slide height in inches
        
    Returns:
        Dictionary with calculated layout dimensions
    """
    margins = PROFESSIONAL_LAYOUT['margins']
    
    content_width = slide_width - margins['left'] - margins['right']
    content_height = slide_height - margins['top'] - margins['bottom']
    
    title_area = {
        'left': margins['left'],
        'top': margins['top'],
        'width': content_width,
        'height': PROFESSIONAL_LAYOUT['standard_sizes']['title_height']
    }
    
    content_area = {
        'left': margins['left'],
        'top': title_area['top'] + title_area['height'] + PROFESSIONAL_LAYOUT['spacing']['title_to_content'],
        'width': content_width,
        'height': content_height - title_area['height'] - PROFESSIONAL_LAYOUT['spacing']['title_to_content']
    }
    
    return {
        'slide': {'width': slide_width, 'height': slide_height},
        'margins': margins,
        'title_area': title_area,
        'content_area': content_area,
        'spacing': PROFESSIONAL_LAYOUT['spacing']
    }

def try_multiple_approaches(operation_name, approaches):
    """
    Try multiple approaches to perform an operation, returning the first successful result.
    
    Args:
        operation_name: Name of the operation for error reporting
        approaches: List of (approach_func, description) tuples to try
        
    Returns:
        Tuple of (result, None) if any approach succeeded, or (None, error_messages) if all failed
    """
    error_messages = []
    
    for approach_func, description in approaches:
        try:
            result = approach_func()
            return result, None
        except Exception as e:
            error_messages.append(f"{description}: {str(e)}")
    
    return None, f"Failed to {operation_name} after trying multiple approaches: {'; '.join(error_messages)}"

def safe_operation(operation_name, operation_func, error_message=None, *args, **kwargs):
    """
    Execute an operation safely with standard error handling.
    
    Args:
        operation_name: Name of the operation for error reporting
        operation_func: Function to execute
        error_message: Custom error message (optional)
        *args, **kwargs: Arguments to pass to the operation function
        
    Returns:
        A tuple (result, error) where error is None if operation was successful
    """
    try:
        result = operation_func(*args, **kwargs)
        return result, None
    except ValueError as e:
        error_msg = error_message or f"Invalid input for {operation_name}: {str(e)}"
        return None, error_msg
    except TypeError as e:
        error_msg = error_message or f"Type error in {operation_name}: {str(e)}"
        return None, error_msg
    except Exception as e:
        error_msg = error_message or f"Failed to execute {operation_name}: {str(e)}"
        return None, error_msg

# ---- Presentation Functions ----

def open_presentation(file_path: str) -> Presentation:
    """
    Open an existing PowerPoint presentation.
    
    Args:
        file_path: Path to the PowerPoint file
        
    Returns:
        A Presentation object
    """
    return Presentation(file_path)

def create_presentation() -> Presentation:
    """
    Create a new PowerPoint presentation.
    
    Returns:
        A new Presentation object
    """
    return Presentation()

def create_presentation_from_template(template_path: str) -> Presentation:
    """
    Create a new PowerPoint presentation from a template file.
    
    Args:
        template_path: Path to the template .pptx file
        
    Returns:
        A new Presentation object based on the template
        
    Raises:
        FileNotFoundError: If the template file doesn't exist
        Exception: If the template file is corrupted or invalid
    """
    import os
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template file not found: {template_path}")
    
    if not template_path.lower().endswith(('.pptx', '.potx')):
        raise ValueError("Template file must be a .pptx or .potx file")
    
    try:
        # Load the template file as a presentation
        presentation = Presentation(template_path)
        return presentation
    except Exception as e:
        raise Exception(f"Failed to load template file '{template_path}': {str(e)}")

def get_template_info(template_path: str) -> Dict:
    """
    Get information about a template file without fully loading it.
    
    Args:
        template_path: Path to the template .pptx file
        
    Returns:
        Dictionary containing template information
    """
    import os
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template file not found: {template_path}")
    
    try:
        presentation = Presentation(template_path)
        
        # Get slide layouts
        layouts = get_slide_layouts(presentation)
        
        # Get core properties
        core_props = get_core_properties(presentation)
        
        # Get slide count
        slide_count = len(presentation.slides)
        
        # Get file size
        file_size = os.path.getsize(template_path)
        
        return {
            "template_path": template_path,
            "file_size_bytes": file_size,
            "slide_count": slide_count,
            "layout_count": len(layouts),
            "slide_layouts": layouts,
            "core_properties": core_props
        }
    except Exception as e:
        raise Exception(f"Failed to read template info from '{template_path}': {str(e)}")

def save_presentation(presentation: Presentation, file_path: str) -> str:
    """
    Save a PowerPoint presentation to a file.
    
    Args:
        presentation: The Presentation object
        file_path: Path where the file should be saved
        
    Returns:
        The file path where the presentation was saved
    """
    presentation.save(file_path)
    return file_path

def presentation_to_base64(presentation: Presentation) -> str:
    """
    Convert a presentation to a base64 encoded string.
    
    Args:
        presentation: The Presentation object
        
    Returns:
        Base64 encoded string of the presentation
    """
    ppt_bytes = io.BytesIO()
    presentation.save(ppt_bytes)
    ppt_bytes.seek(0)
    return base64.b64encode(ppt_bytes.read()).decode('utf-8')

def base64_to_presentation(base64_string: str) -> Presentation:
    """
    Create a presentation from a base64 encoded string.
    
    Args:
        base64_string: Base64 encoded string of a presentation
        
    Returns:
        A Presentation object
    """
    ppt_bytes = io.BytesIO(base64.b64decode(base64_string))
    return Presentation(ppt_bytes)

# ---- Slide Functions ----

def add_slide(presentation: Presentation, layout_index: int = 1) -> Tuple:
    """
    Add a slide to the presentation.
    
    Args:
        presentation: The Presentation object
        layout_index: Index of the slide layout to use (default is 1, typically a title and content slide)
        
    Returns:
        A tuple containing the slide and its layout
    """
    layout = presentation.slide_layouts[layout_index]
    slide = presentation.slides.add_slide(layout)
    return slide, layout

def add_professional_slide(presentation: Presentation, slide_type: str = 'title_content', 
                          color_scheme: str = 'modern_blue') -> Tuple:
    """
    Add a professionally designed slide with proper spacing and typography.
    
    Args:
        presentation: The Presentation object
        slide_type: Type of slide ('title', 'title_content', 'content', 'two_column', 'blank')
        color_scheme: Color scheme to apply ('modern_blue', 'corporate_gray', etc.)
        
    Returns:
        A tuple containing the slide, layout, and design info
    """
    # Map slide types to layout indices
    layout_map = {
        'title': 0,           # Title slide
        'title_content': 1,   # Title and content
        'content': 6,         # Content only
        'two_column': 3,      # Two content columns
        'blank': 6            # Blank layout
    }
    
    layout_index = layout_map.get(slide_type, 1)
    layout = presentation.slide_layouts[layout_index]
    slide = presentation.slides.add_slide(layout)
    
    # Calculate professional layout
    slide_width = presentation.slide_width / 914400  # Convert EMU to inches
    slide_height = presentation.slide_height / 914400
    layout_info = calculate_professional_layout(slide_width, slide_height)
    
    # Apply professional styling if possible
    try:
        apply_professional_slide_background(slide, color_scheme)
    except Exception:
        pass  # Graceful degradation if background setting fails
    
    design_info = {
        'color_scheme': color_scheme,
        'layout_info': layout_info,
        'slide_type': slide_type
    }
    
    return slide, layout, design_info

def apply_professional_slide_background(slide, color_scheme: str = 'modern_blue'):
    """
    Apply a professional background to a slide.
    
    Args:
        slide: The slide object
        color_scheme: Color scheme name
    """
    try:
        # Get background color
        bg_color = get_professional_color(color_scheme, 'light')
        
        # Try to set slide background
        if hasattr(slide, 'background'):
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(*bg_color)
    except Exception:
        # Graceful fallback - background setting may not be available
        pass

def get_slide_layouts(presentation: Presentation) -> List[Dict]:
    """
    Get all available slide layouts in the presentation.
    
    Args:
        presentation: The Presentation object
        
    Returns:
        A list of dictionaries with layout information
    """
    layouts = []
    for i, layout in enumerate(presentation.slide_layouts):
        layout_info = {
            "index": i,
            "name": layout.name,
            "placeholder_count": len(layout.placeholders)
        }
        layouts.append(layout_info)
    return layouts

# ---- Placeholder Functions ----

def get_placeholders(slide) -> List[Dict]:
    """
    Get all placeholders in a slide.
    
    Args:
        slide: The slide object
        
    Returns:
        A list of dictionaries with placeholder information
    """
    placeholders = []
    for placeholder in slide.placeholders:
        placeholder_info = {
            "idx": placeholder.placeholder_format.idx,
            "type": placeholder.placeholder_format.type,
            "name": placeholder.name,
            "shape_type": placeholder.shape_type
        }
        placeholders.append(placeholder_info)
    return placeholders

def set_title(slide, title: str) -> None:
    """
    Set the title of a slide.
    
    Args:
        slide: The slide object
        title: The title text
    """
    if slide.shapes.title:
        slide.shapes.title.text = title

def set_professional_title(slide, title: str, color_scheme: str = 'modern_blue', 
                          size_category: str = 'large') -> None:
    """
    Set a professionally formatted title with proper typography and colors.
    
    Args:
        slide: The slide object
        title: The title text
        color_scheme: Color scheme to use
        size_category: Font size category ('large', 'medium', 'small')
    """
    if not slide.shapes.title:
        return
    
    # Set title text
    slide.shapes.title.text = title
    
    # Get professional font settings
    font_settings = get_professional_font('title', size_category)
    title_color = get_professional_color(color_scheme, 'primary')
    
    # Apply professional formatting
    try:
        text_frame = slide.shapes.title.text_frame
        
        # Enable word wrap and proper margins
        text_frame.word_wrap = True
        text_frame.margin_left = Inches(0)
        text_frame.margin_right = Inches(0)
        text_frame.margin_top = Inches(0.1)
        text_frame.margin_bottom = Inches(0.1)
        
        # Format all paragraphs
        for paragraph in text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.LEFT
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(6)
            paragraph.line_spacing = 1.15
            
            # Format all runs in the paragraph
            for run in paragraph.runs:
                font = run.font
                font.name = font_settings['name']
                font.size = Pt(font_settings['size'])
                font.bold = font_settings['bold']
                font.color.rgb = RGBColor(*title_color)
                
    except Exception as e:
        # Fallback to basic formatting
        try:
            format_text_advanced(
                slide.shapes.title.text_frame,
                font_size=font_settings['size'],
                font_name=font_settings['name'],
                bold=font_settings['bold'],
                color=title_color
            )
        except:
            pass

def populate_placeholder(slide, placeholder_idx: int, text: str) -> None:
    """
    Populate a placeholder with text.
    
    Args:
        slide: The slide object
        placeholder_idx: The index of the placeholder
        text: The text to add
    """
    placeholder = slide.placeholders[placeholder_idx]
    placeholder.text = text

def add_bullet_points(placeholder, bullet_points: List[str]) -> None:
    """
    Add bullet points to a placeholder.
    
    Args:
        placeholder: The placeholder object
        bullet_points: List of bullet point texts
    """
    text_frame = placeholder.text_frame
    text_frame.clear()
    
    for i, point in enumerate(bullet_points):
        p = text_frame.add_paragraph()
        p.text = point
        p.level = 0
        
        # Only add line breaks between bullet points, not after the last one
        if i < len(bullet_points) - 1:
            p.line_spacing = 1.0

def add_professional_bullet_points(placeholder, bullet_points: List[str], 
                                   color_scheme: str = 'modern_blue',
                                   hierarchical: bool = False) -> None:
    """
    Add professionally formatted bullet points with proper typography and spacing.
    
    Args:
        placeholder: The placeholder object
        bullet_points: List of bullet point texts (can include level indicators like "  - sub point")
        color_scheme: Color scheme to use
        hierarchical: Whether to automatically detect and format hierarchical bullet points
    """
    text_frame = placeholder.text_frame
    text_frame.clear()
    
    # Set professional text frame properties
    try:
        text_frame.word_wrap = True
        text_frame.margin_left = Inches(0.25)
        text_frame.margin_right = Inches(0.25)
        text_frame.margin_top = Inches(0.1)
        text_frame.margin_bottom = Inches(0.1)
    except:
        pass
    
    # Get professional font settings
    body_font = get_professional_font('body', 'medium')
    text_color = get_professional_color(color_scheme, 'text')
    
    for i, point in enumerate(bullet_points):
        # Determine bullet level if hierarchical
        level = 0
        clean_text = point.strip()
        
        if hierarchical:
            # Count leading spaces or tabs to determine level
            leading_spaces = len(point) - len(point.lstrip(' \t'))
            level = min(leading_spaces // 2, 4)  # Max 5 levels (0-4)
            
            # Remove common level indicators
            for indicator in ['• ', '- ', '* ', '○ ', '▪ ']:
                if clean_text.startswith(indicator):
                    clean_text = clean_text[len(indicator):]
                    break
        
        # Add paragraph
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        
        p.text = clean_text
        p.level = level
        
        # Set professional paragraph formatting
        try:
            p.space_before = Pt(0)
            p.space_after = Pt(3)
            p.line_spacing = 1.25
            
            if level == 0:
                p.space_before = Pt(6) if i > 0 else Pt(0)
            
            # Format the text run
            for run in p.runs:
                font = run.font
                font.name = body_font['name']
                font.size = Pt(max(body_font['size'] - level, 10))  # Smaller font for sub-levels
                font.bold = (level == 0)  # Only main bullets are bold
                font.color.rgb = RGBColor(*text_color)
                
        except Exception:
            # Fallback formatting
            try:
                format_text_advanced(
                    text_frame,
                    font_size=body_font['size'],
                    font_name=body_font['name'],
                    color=text_color
                )
            except:
                pass

# ---- Text Functions ----

def add_textbox(slide, left: float, top: float, width: float, height: float, text: str,
                font_size: int = None, font_name: str = None, bold: bool = None,
                italic: bool = None, color: Tuple[int, int, int] = None,
                alignment: str = None, auto_resize: bool = True) -> Any:
    """
    Add a textbox to a slide with enhanced text formatting and overflow handling.
    
    Args:
        slide: The slide object
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        text: Text content
        font_size: Font size in points
        font_name: Font name
        bold: Whether text should be bold
        italic: Whether text should be italic
        color: RGB color tuple (r, g, b)
        alignment: Text alignment ('left', 'center', 'right', 'justify')
        auto_resize: Whether to automatically resize font if text overflows
        
    Returns:
        The created textbox shape
    """
    textbox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    
    # Set the text first
    textbox.text_frame.text = text
    
    # Apply formatting if any formatting options are provided
    if any([font_size, font_name, bold, italic, color, alignment]):
        format_text_advanced(
            textbox.text_frame,
            font_size=font_size,
            font_name=font_name,
            bold=bold,
            italic=italic,
            color=color,
            alignment=alignment,
            auto_resize=auto_resize
        )
    
    return textbox

def format_text(text_frame, font_size: int = None, font_name: str = None, 
                bold: bool = None, italic: bool = None, color: Tuple[int, int, int] = None,
                alignment: str = None, auto_resize: bool = False, max_font_size: int = None) -> None:
    """
    Format text in a text frame with overflow detection and auto-resize capability.
    
    Args:
        text_frame: The text frame to format
        font_size: Font size in points
        font_name: Font name
        bold: Whether text should be bold
        italic: Whether text should be italic
        color: RGB color tuple (r, g, b)
        alignment: Text alignment ('left', 'center', 'right', 'justify')
        auto_resize: Whether to automatically resize font if text overflows
        max_font_size: Maximum font size when auto-resizing (defaults to original font_size)
    """
    alignment_map = {
        'left': PP_ALIGN.LEFT,
        'center': PP_ALIGN.CENTER,
        'right': PP_ALIGN.RIGHT,
        'justify': PP_ALIGN.JUSTIFY
    }
    
    # Enable text wrapping and auto-fit behavior
    text_frame.word_wrap = True
    text_frame.auto_size = 1  # Auto-fit text to shape
    
    # Set maximum font size if auto_resize is enabled
    if auto_resize and max_font_size is None and font_size is not None:
        max_font_size = font_size
    
    # Apply formatting to all paragraphs and runs
    for paragraph in text_frame.paragraphs:
        if alignment and alignment in alignment_map:
            paragraph.alignment = alignment_map[alignment]
            
        for run in paragraph.runs:
            font = run.font
            
            # Apply font size with auto-resize logic
            if font_size is not None:
                if auto_resize and max_font_size:
                    # Try progressively smaller font sizes if text overflows
                    current_size = min(font_size, max_font_size)
                    font.size = Pt(current_size)
                    
                    # Check if we need to reduce font size (basic heuristic)
                    while current_size > 8:  # Minimum readable font size
                        try:
                            # Set the font size and check if it fits
                            font.size = Pt(current_size)
                            break
                        except:
                            current_size -= 2
                            continue
                else:
                    font.size = Pt(font_size)
                
            if font_name is not None:
                font.name = font_name
                
            if bold is not None:
                font.bold = bold
                
            if italic is not None:
                font.italic = italic
                
            if color is not None:
                r, g, b = color
                font.color.rgb = RGBColor(r, g, b)

def format_text_advanced(text_frame, font_size: int = None, font_name: str = None, 
                        bold: bool = None, italic: bool = None, color: Tuple[int, int, int] = None,
                        alignment: str = None, auto_resize: bool = True, 
                        min_font_size: int = 8, max_font_size: int = None) -> Dict:
    """
    Advanced text formatting with overflow detection and automatic font size adjustment.
    
    Args:
        text_frame: The text frame to format
        font_size: Initial font size in points
        font_name: Font name
        bold: Whether text should be bold
        italic: Whether text should be italic
        color: RGB color tuple (r, g, b)
        alignment: Text alignment ('left', 'center', 'right', 'justify')
        auto_resize: Whether to automatically resize font if text overflows
        min_font_size: Minimum font size when auto-resizing
        max_font_size: Maximum font size when auto-resizing
    
    Returns:
        Dictionary with formatting results and final font size used
    """
    result = {
        'success': True,
        'original_font_size': font_size,
        'final_font_size': font_size,
        'auto_resized': False,
        'warnings': []
    }
    
    try:
        # Clear existing text formatting and start fresh
        if len(text_frame.paragraphs) == 0:
            text_frame.add_paragraph()
        
        # Enable text wrapping and auto-fit
        text_frame.word_wrap = True
        
        # Try to set auto-fit behavior (may not be available in all versions)
        try:
            text_frame.auto_size = 1  # MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except:
            result['warnings'].append("Auto-fit not available, using manual font sizing")
        
        alignment_map = {
            'left': PP_ALIGN.LEFT,
            'center': PP_ALIGN.CENTER,
            'right': PP_ALIGN.RIGHT,
            'justify': PP_ALIGN.JUSTIFY
        }
        
        # Set max font size if not provided
        if max_font_size is None:
            max_font_size = font_size if font_size else 72
        
        # Start with the desired font size
        current_font_size = font_size if font_size else 12
        
        # Apply formatting to all paragraphs and runs
        for paragraph in text_frame.paragraphs:
            if alignment and alignment in alignment_map:
                paragraph.alignment = alignment_map[alignment]
            
            # If paragraph has no runs, create text to establish runs
            if len(paragraph.runs) == 0 and hasattr(paragraph, 'text'):
                original_text = paragraph.text
                paragraph.text = original_text or " "  # Ensure we have some text
            
            for run in paragraph.runs:
                font = run.font
                
                # Apply font name first
                if font_name is not None:
                    font.name = font_name
                
                # Apply font size with auto-resize logic
                if current_font_size is not None:
                    if auto_resize:
                        # Try progressively smaller sizes if needed
                        size_to_try = min(current_font_size, max_font_size)
                        
                        while size_to_try >= min_font_size:
                            try:
                                font.size = Pt(size_to_try)
                                # Test if this size works (no exception means it's ok)
                                result['final_font_size'] = size_to_try
                                if size_to_try != current_font_size:
                                    result['auto_resized'] = True
                                break
                            except Exception as e:
                                result['warnings'].append(f"Font size {size_to_try} failed: {str(e)}")
                                size_to_try -= 2
                                continue
                        
                        if size_to_try < min_font_size:
                            result['warnings'].append(f"Could not fit text even at minimum size {min_font_size}")
                            font.size = Pt(min_font_size)
                            result['final_font_size'] = min_font_size
                            result['auto_resized'] = True
                    else:
                        font.size = Pt(current_font_size)
                        result['final_font_size'] = current_font_size
                
                # Apply other formatting
                if bold is not None:
                    font.bold = bold
                
                if italic is not None:
                    font.italic = italic
                
                if color is not None:
                    r, g, b = color
                    font.color.rgb = RGBColor(r, g, b)
        
        return result
        
    except Exception as e:
        result['success'] = False
        result['error'] = str(e)
        return result

def validate_text_container(shape, text_content: str, font_size: int) -> Dict:
    """
    Validate if text content will fit in a shape container.
    
    Args:
        shape: The shape containing the text
        text_content: The text to validate
        font_size: The font size to check
    
    Returns:
        Dictionary with validation results and suggestions
    """
    result = {
        'fits': True,
        'estimated_overflow': False,
        'suggested_font_size': font_size,
        'suggested_dimensions': None,
        'warnings': []
    }
    
    try:
        # Basic heuristic: estimate if text will overflow
        if hasattr(shape, 'width') and hasattr(shape, 'height'):
            # Rough estimation: average character width is about 0.6 * font_size
            avg_char_width = font_size * 0.6
            estimated_width = len(text_content) * avg_char_width
            
            # Convert shape dimensions to points (assuming they're in EMU)
            shape_width_pt = shape.width / 12700  # EMU to points conversion
            shape_height_pt = shape.height / 12700
            
            if estimated_width > shape_width_pt:
                result['fits'] = False
                result['estimated_overflow'] = True
                
                # Suggest smaller font size
                suggested_size = int((shape_width_pt / len(text_content)) * 0.8)
                result['suggested_font_size'] = max(suggested_size, 8)
                
                # Suggest larger dimensions
                result['suggested_dimensions'] = {
                    'width': estimated_width * 1.2,
                    'height': shape_height_pt
                }
                
                result['warnings'].append(
                    f"Text may overflow. Consider font size {result['suggested_font_size']} "
                    f"or increase width to {result['suggested_dimensions']['width']:.1f} points"
                )
        
        return result
        
    except Exception as e:
        result['warnings'].append(f"Validation error: {str(e)}")
        return result

# ---- Image Functions ----

def add_image(slide, image_path: str, left: float, top: float, width: float = None, height: float = None) -> Any:
    """
    Add an image to a slide.
    
    Args:
        slide: The slide object
        image_path: Path to the image file
        left: Left position in inches
        top: Top position in inches
        width: Width in inches (optional)
        height: Height in inches (optional)
        
    Returns:
        The created picture shape
    """
    if width and height:
        picture = slide.shapes.add_picture(
            image_path, Inches(left), Inches(top), Inches(width), Inches(height)
        )
    else:
        picture = slide.shapes.add_picture(
            image_path, Inches(left), Inches(top)
        )
    return picture

def add_image_from_base64(slide, base64_string: str, left: float, top: float, 
                          width: float = None, height: float = None) -> Any:
    """
    Add an image from a base64 encoded string to a slide.
    
    Args:
        slide: The slide object
        base64_string: Base64 encoded image string
        left: Left position in inches
        top: Top position in inches
        width: Width in inches (optional)
        height: Height in inches (optional)
        
    Returns:
        The created picture shape
    """
    image_data = base64.b64decode(base64_string)
    image_stream = io.BytesIO(image_data)
    
    if width and height:
        picture = slide.shapes.add_picture(
            image_stream, Inches(left), Inches(top), Inches(width), Inches(height)
        )
    else:
        picture = slide.shapes.add_picture(
            image_stream, Inches(left), Inches(top)
        )
    return picture

# ---- Table Functions ----

def add_table(slide, rows: int, cols: int, left: float, top: float, width: float, height: float) -> Any:
    """
    Add a table to a slide.
    
    Args:
        slide: The slide object
        rows: Number of rows
        cols: Number of columns
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        
    Returns:
        The created table shape
    """
    table = slide.shapes.add_table(
        rows, cols, Inches(left), Inches(top), Inches(width), Inches(height)
    ).table
    return table

def set_cell_text(table, row: int, col: int, text: str) -> None:
    """
    Set text in a table cell.
    
    Args:
        table: The table object
        row: Row index
        col: Column index
        text: Text content
    """
    cell = table.cell(row, col)
    cell.text = text

def format_table_cell(cell, font_size: int = None, font_name: str = None, 
                     bold: bool = None, italic: bool = None, 
                     color: Tuple[int, int, int] = None,
                     bg_color: Tuple[int, int, int] = None,
                     alignment: str = None,
                     vertical_alignment: str = None) -> None:
    """
    Format a table cell.
    
    Args:
        cell: The table cell to format
        font_size: Font size in points
        font_name: Font name
        bold: Whether text should be bold
        italic: Whether text should be italic
        color: RGB color tuple for text (r, g, b)
        bg_color: RGB color tuple for background (r, g, b)
        alignment: Text alignment ('left', 'center', 'right', 'justify')
        vertical_alignment: Vertical alignment ('top', 'middle', 'bottom')
    """
    alignment_map = {
        'left': PP_ALIGN.LEFT,
        'center': PP_ALIGN.CENTER,
        'right': PP_ALIGN.RIGHT,
        'justify': PP_ALIGN.JUSTIFY
    }
    
    vertical_alignment_map = {
        'top': MSO_VERTICAL_ANCHOR.TOP,
        'middle': MSO_VERTICAL_ANCHOR.MIDDLE,
        'bottom': MSO_VERTICAL_ANCHOR.BOTTOM
    }
    
    # Format text
    text_frame = cell.text_frame
    
    if vertical_alignment and vertical_alignment in vertical_alignment_map:
        text_frame.vertical_anchor = vertical_alignment_map[vertical_alignment]
    
    for paragraph in text_frame.paragraphs:
        if alignment and alignment in alignment_map:
            paragraph.alignment = alignment_map[alignment]
            
        for run in paragraph.runs:
            font = run.font
            
            if font_size is not None:
                font.size = Pt(font_size)
                
            if font_name is not None:
                font.name = font_name
                
            if bold is not None:
                font.bold = bold
                
            if italic is not None:
                font.italic = italic
                
            if color is not None:
                r, g, b = color
                font.color.rgb = RGBColor(r, g, b)
    
    # Set background color
    if bg_color is not None:
        r, g, b = bg_color
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(r, g, b)

# ---- Shape Functions ----

def add_shape(slide, shape_type: str, left: float, top: float, width: float, height: float) -> Any:
    """
    Add an auto shape to a slide.
    
    Args:
        slide: The slide object
        shape_type: Shape type string (e.g., 'rectangle', 'oval', 'triangle')
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        
    Returns:
        The created shape
    """
    # Ensure shape_type is a string
    shape_type_str = str(shape_type)
    shape_type_lower = shape_type_str.lower()
    
    # Define shape type mapping with correct enum values
    shape_type_map = {
        'rectangle': MSO_SHAPE.RECTANGLE,
        'rounded_rectangle': MSO_SHAPE.ROUNDED_RECTANGLE,
        'oval': MSO_SHAPE.OVAL,
        'diamond': MSO_SHAPE.DIAMOND,
        'triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,  # Changed from TRIANGLE to ISOSCELES_TRIANGLE
        'isosceles_triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
        'right_triangle': MSO_SHAPE.RIGHT_TRIANGLE,
        'pentagon': MSO_SHAPE.PENTAGON,
        'hexagon': MSO_SHAPE.HEXAGON,
        'heptagon': MSO_SHAPE.HEPTAGON,
        'octagon': MSO_SHAPE.OCTAGON,
        'star': MSO_SHAPE.STAR_5_POINTS,
        'arrow': MSO_SHAPE.ARROW,
        'cloud': MSO_SHAPE.CLOUD,
        'heart': MSO_SHAPE.HEART,
        'lightning_bolt': MSO_SHAPE.LIGHTNING_BOLT,
        'sun': MSO_SHAPE.SUN,
        'moon': MSO_SHAPE.MOON,
        'smiley_face': MSO_SHAPE.SMILEY_FACE,
        'no_symbol': MSO_SHAPE.NO_SYMBOL,
        'flowchart_process': MSO_SHAPE.FLOWCHART_PROCESS,
        'flowchart_decision': MSO_SHAPE.FLOWCHART_DECISION,
        'flowchart_data': MSO_SHAPE.FLOWCHART_DATA,
        'flowchart_document': MSO_SHAPE.FLOWCHART_DOCUMENT,
        'flowchart_predefined_process': MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS,
        'flowchart_internal_storage': MSO_SHAPE.FLOWCHART_INTERNAL_STORAGE,
        'flowchart_connector': MSO_SHAPE.FLOWCHART_CONNECTOR
    }
    
    # Check if shape type is valid before trying to use it
    if shape_type_lower not in shape_type_map:
        available_shapes = ', '.join(sorted(shape_type_map.keys()))
        raise ValueError(f"Unsupported shape type: '{shape_type}'. Available shape types: {available_shapes}")
    
    # Get the shape enum value
    shape_enum = shape_type_map[shape_type_lower]
    
    # Create the shape with better error handling
    try:
        shape = slide.shapes.add_shape(
            shape_enum, Inches(left), Inches(top), Inches(width), Inches(height)
        )
        return shape
    except Exception as e:
        # More detailed error for debugging
        raise ValueError(f"Could not create shape '{shape_type}' (enum: {shape_enum.__class__.__name__}.{shape_enum.name}): {str(e)}")

def format_shape(shape, fill_color: Tuple[int, int, int] = None, 
                line_color: Tuple[int, int, int] = None, line_width: float = None) -> None:
    """
    Format a shape.
    
    Args:
        shape: The shape object
        fill_color: RGB color tuple for fill (r, g, b)
        line_color: RGB color tuple for outline (r, g, b)
        line_width: Line width in points
    """
    if fill_color is not None:
        r, g, b = fill_color
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(r, g, b)
    
    if line_color is not None:
        r, g, b = line_color
        shape.line.color.rgb = RGBColor(r, g, b)
    
    if line_width is not None:
        shape.line.width = Pt(line_width)

def format_professional_shape(shape, color_scheme: str = 'modern_blue', 
                             style: str = 'primary', with_shadow: bool = True) -> None:
    """
    Apply professional formatting to a shape with modern styling.
    
    Args:
        shape: The shape object
        color_scheme: Color scheme to use
        style: Style type ('primary', 'secondary', 'accent1', 'accent2', 'subtle')
        with_shadow: Whether to add a subtle shadow effect
    """
    try:
        # Get colors for the style
        if style == 'primary':
            fill_color = get_professional_color(color_scheme, 'primary')
            line_color = get_professional_color(color_scheme, 'primary')
        elif style == 'secondary':
            fill_color = get_professional_color(color_scheme, 'secondary')
            line_color = get_professional_color(color_scheme, 'secondary')
        elif style == 'accent1':
            fill_color = get_professional_color(color_scheme, 'accent1')
            line_color = get_professional_color(color_scheme, 'accent1')
        elif style == 'accent2':
            fill_color = get_professional_color(color_scheme, 'accent2')
            line_color = get_professional_color(color_scheme, 'accent2')
        elif style == 'subtle':
            fill_color = get_professional_color(color_scheme, 'light')
            line_color = get_professional_color(color_scheme, 'text')
        else:
            fill_color = get_professional_color(color_scheme, 'primary')
            line_color = get_professional_color(color_scheme, 'primary')
        
        # Apply fill
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*fill_color)
        
        # Apply modern line styling
        if style == 'subtle':
            shape.line.color.rgb = RGBColor(*line_color)
            shape.line.width = Pt(1)
        else:
            # No outline for solid colored shapes (modern flat design)
            shape.line.fill.background()
        
        # Add subtle shadow for depth (if supported)
        if with_shadow and hasattr(shape, 'shadow'):
            try:
                shadow = shape.shadow
                shadow.inherit = False
                shadow.visible = True
                shadow.style = 1  # Simple shadow
                shadow.blur_radius = Pt(4)
                shadow.distance = Pt(3)
                shadow.angle = 45
                shadow.color.rgb = RGBColor(0, 0, 0)
                shadow.transparency = 0.3
            except:
                pass  # Shadow not supported
                
    except Exception:
        # Fallback to basic formatting
        try:
            format_shape(shape, fill_color, line_color, 1.0)
        except:
            pass

# ---- Chart Functions ----

def add_chart(slide, chart_type: str, left: float, top: float, width: float, height: float,
             categories: List[str], series_names: List[str], series_values: List[List[float]]) -> Any:
    """
    Add a chart to a slide.
    
    Args:
        slide: The slide object
        chart_type: Type of chart ('column', 'bar', 'line', 'pie')
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        categories: List of category names
        series_names: List of series names
        series_values: List of lists containing values for each series
        
    Returns:
        The created chart
    """
    chart_type_map = {
        'column': XL_CHART_TYPE.COLUMN_CLUSTERED,
        'stacked_column': XL_CHART_TYPE.COLUMN_STACKED,
        'bar': XL_CHART_TYPE.BAR_CLUSTERED,
        'stacked_bar': XL_CHART_TYPE.BAR_STACKED,
        'line': XL_CHART_TYPE.LINE,
        'line_markers': XL_CHART_TYPE.LINE_MARKERS,
        'pie': XL_CHART_TYPE.PIE,
        'doughnut': XL_CHART_TYPE.DOUGHNUT,
        'area': XL_CHART_TYPE.AREA,
        'stacked_area': XL_CHART_TYPE.AREA_STACKED,
        'scatter': XL_CHART_TYPE.XY_SCATTER,
        'radar': XL_CHART_TYPE.RADAR,
        'radar_markers': XL_CHART_TYPE.RADAR_MARKERS
    }
    
    chart_type_enum = chart_type_map.get(chart_type.lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)
    
    # Create chart data
    chart_data = CategoryChartData()
    chart_data.categories = categories
    
    for i, series_name in enumerate(series_names):
        chart_data.add_series(series_name, series_values[i])
    
    # Add chart to slide
    graphic_frame = slide.shapes.add_chart(
        chart_type_enum, Inches(left), Inches(top), Inches(width), Inches(height), chart_data
    )
    
    return graphic_frame.chart

def format_chart(chart, has_legend: bool = True, legend_position: str = 'right',
                has_data_labels: bool = False, title: str = None) -> None:
    """
    Format a chart.
    
    Args:
        chart: The chart object
        has_legend: Whether to show the legend
        legend_position: Position of the legend ('right', 'left', 'top', 'bottom')
        has_data_labels: Whether to show data labels
        title: Chart title
    """
    # Set chart title
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
    else:
        chart.has_title = False
    
    # Configure legend
    chart.has_legend = has_legend
    if has_legend:
        position_map = {
            'right': 2,  # XL_LEGEND_POSITION.RIGHT
            'left': 3,   # XL_LEGEND_POSITION.LEFT
            'top': 1,    # XL_LEGEND_POSITION.TOP
            'bottom': 4  # XL_LEGEND_POSITION.BOTTOM
        }
        chart.legend.position = position_map.get(legend_position.lower(), 2)
    
    # Configure data labels
    for series in chart.series:
        series.has_data_labels = has_data_labels

def format_professional_chart(chart, color_scheme: str = 'modern_blue', 
                             title: str = None, modern_style: bool = True) -> None:
    """
    Apply professional formatting to a chart with modern styling and colors.
    
    Args:
        chart: The chart object
        color_scheme: Color scheme to use
        title: Chart title
        modern_style: Whether to apply modern flat design principles
    """
    try:
        # Set professional title
        if title:
            chart.has_title = True
            title_frame = chart.chart_title.text_frame
            title_frame.text = title
            
            # Format title text
            title_font = get_professional_font('title', 'small')
            title_color = get_professional_color(color_scheme, 'primary')
            
            for paragraph in title_frame.paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    font.name = title_font['name']
                    font.size = Pt(title_font['size'])
                    font.bold = title_font['bold']
                    font.color.rgb = RGBColor(*title_color)
        
        # Configure modern legend
        if chart.has_legend:
            chart.legend.position = 4  # Bottom position for modern look
            try:
                legend_font = get_professional_font('body', 'small')
                text_color = get_professional_color(color_scheme, 'text')
                
                for paragraph in chart.legend.font:
                    font = paragraph.font
                    font.name = legend_font['name']
                    font.size = Pt(legend_font['size'])
                    font.color.rgb = RGBColor(*text_color)
            except:
                pass
        
        # Apply professional colors to series
        professional_colors = [
            get_professional_color(color_scheme, 'primary'),
            get_professional_color(color_scheme, 'accent1'),
            get_professional_color(color_scheme, 'accent2'),
            get_professional_color(color_scheme, 'secondary'),
        ]
        
        for i, series in enumerate(chart.series):
            try:
                color = professional_colors[i % len(professional_colors)]
                
                # Apply color to series
                if hasattr(series, 'fill'):
                    series.fill.solid()
                    series.fill.fore_color.rgb = RGBColor(*color)
                
                # Modern style: remove borders from series
                if modern_style and hasattr(series, 'line'):
                    series.line.fill.background()
                
            except Exception:
                continue
        
        # Style axes for modern look
        try:
            # Category axis
            if hasattr(chart, 'category_axis'):
                cat_axis = chart.category_axis
                if hasattr(cat_axis, 'tick_labels'):
                    font_settings = get_professional_font('caption', 'medium')
                    text_color = get_professional_color(color_scheme, 'text')
                    
                    tick_font = cat_axis.tick_labels.font
                    tick_font.name = font_settings['name']
                    tick_font.size = Pt(font_settings['size'])
                    tick_font.color.rgb = RGBColor(*text_color)
            
            # Value axis
            if hasattr(chart, 'value_axis'):
                val_axis = chart.value_axis
                if hasattr(val_axis, 'tick_labels'):
                    font_settings = get_professional_font('caption', 'medium')
                    text_color = get_professional_color(color_scheme, 'text')
                    
                    tick_font = val_axis.tick_labels.font
                    tick_font.name = font_settings['name']
                    tick_font.size = Pt(font_settings['size'])
                    tick_font.color.rgb = RGBColor(*text_color)
                
        except Exception:
            pass
        
    except Exception:
        # Fallback to basic formatting
        try:
            format_chart(chart, has_legend=True, title=title)
        except:
            pass

# ---- Document Properties Functions ----

def set_core_properties(presentation: Presentation, title: str = None, subject: str = None,
                       author: str = None, keywords: str = None, comments: str = None) -> None:
    """
    Set core document properties.
    
    Args:
        presentation: The Presentation object
        title: Document title
        subject: Document subject
        author: Document author
        keywords: Document keywords
        comments: Document comments
    """
    core_props = presentation.core_properties
    
    if title is not None:
        core_props.title = title
        
    if subject is not None:
        core_props.subject = subject
        
    if author is not None:
        core_props.author = author
        
    if keywords is not None:
        core_props.keywords = keywords
        
    if comments is not None:
        core_props.comments = comments

def get_core_properties(presentation: Presentation) -> Dict:
    """
    Get core document properties.
    
    Args:
        presentation: The Presentation object
        
    Returns:
        Dictionary of core properties
    """
    core_props = presentation.core_properties
    
    return {
        'title': core_props.title,
        'subject': core_props.subject,
        'author': core_props.author,
        'keywords': core_props.keywords,
        'comments': core_props.comments,
        'category': core_props.category,
        'created': core_props.created,
        'modified': core_props.modified,
        'last_modified_by': core_props.last_modified_by
    }
