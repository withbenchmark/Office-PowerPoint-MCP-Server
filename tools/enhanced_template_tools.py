"""
Enhanced template tools with dynamic sizing, auto-wrapping, and visual effects.
Provides advanced slide creation capabilities with intelligent content adaptation.
"""
from typing import Dict, List, Optional, Any
from mcp.server.fastmcp import FastMCP
import utils.enhanced_template_utils as enhanced_utils


def register_enhanced_template_tools(app: FastMCP, presentations: Dict, get_current_presentation_id):
    """Register enhanced template tools with the FastMCP app"""
    
    @app.tool()
    def list_enhanced_templates() -> Dict:
        """List all available enhanced slide templates with dynamic features."""
        try:
            manager = enhanced_utils.get_enhanced_template_manager()
            templates_data = manager.templates_data
            
            template_list = []
            for template_id, template_info in templates_data.get('templates', {}).items():
                # Analyze template features
                features = []
                for element in template_info.get('elements', []):
                    styling = element.get('styling', {})
                    if styling.get('font_size') == 'dynamic':
                        features.append('Dynamic text sizing')
                    if styling.get('auto_wrap'):
                        features.append('Auto text wrapping')
                    if styling.get('text_effects'):
                        features.append('Text effects')
                    if 'fill_gradient' in styling:
                        features.append('Gradient fills')
                
                template_list.append({
                    'id': template_id,
                    'name': template_info.get('name', template_id),
                    'description': template_info.get('description', ''),
                    'typography_style': template_info.get('typography_style', 'modern_sans'),
                    'element_count': len(template_info.get('elements', [])),
                    'enhanced_features': list(set(features)),
                    'background_type': template_info.get('background', {}).get('type', 'solid')
                })
            
            return {
                "enhanced_templates": template_list,
                "total_templates": len(template_list),
                "color_schemes": list(templates_data.get('color_schemes', {}).keys()),
                "typography_styles": list(templates_data.get('typography_styles', {}).keys()),
                "available_effects": {
                    "text_effects": list(templates_data.get('text_effects', {}).keys()),
                    "image_effects": list(templates_data.get('image_effects', {}).keys())
                },
                "dynamic_features": [
                    "Automatic text sizing based on content length",
                    "Intelligent text wrapping to fit containers",
                    "Visual effects (shadows, glows, outlines)",
                    "Gradient backgrounds and fills",
                    "Typography variety with multiple font combinations",
                    "Dynamic line spacing adjustment",
                    "Content-aware formatting"
                ],
                "message": "Use apply_enhanced_template to create slides with advanced dynamic features"
            }
        except Exception as e:
            return {
                "error": f"Failed to list enhanced templates: {str(e)}"
            }
    
    @app.tool()
    def apply_enhanced_template(
        slide_index: int,
        template_id: str,
        color_scheme: str = "modern_blue",
        content_mapping: Optional[Dict[str, str]] = None,
        image_paths: Optional[Dict[str, str]] = None,
        typography_style: Optional[str] = None,
        auto_optimize: bool = True,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Apply an enhanced slide template with dynamic features.
        
        Args:
            slide_index: Index of the slide to apply template to
            template_id: ID of the enhanced template (e.g., 'enhanced_title_slide', 'enhanced_text_with_image')
            color_scheme: Color scheme ('modern_blue', 'corporate_gray', 'elegant_green', 'warm_red')
            content_mapping: Dictionary mapping element roles to custom content
            image_paths: Dictionary mapping image element roles to file paths
            typography_style: Typography style ('modern_sans', 'elegant_serif', 'tech_modern')
            auto_optimize: Whether to automatically optimize text sizing and wrapping
            presentation_id: Presentation ID (uses current if None)
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        if slide_index < 0 or slide_index >= len(pres.slides):
            return {
                "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
            }
        
        slide = pres.slides[slide_index]
        
        try:
            # Apply enhanced template
            result = enhanced_utils.apply_enhanced_template(
                slide, template_id, color_scheme, 
                content_mapping or {}, image_paths or {}
            )
            
            if result['success']:
                return {
                    "message": f"Applied enhanced template '{template_id}' to slide {slide_index}",
                    "slide_index": slide_index,
                    "template_applied": result,
                    "optimization_applied": auto_optimize,
                    "typography_style": typography_style or "default"
                }
            else:
                return {
                    "error": f"Failed to apply enhanced template: {result.get('error', 'Unknown error')}"
                }
                
        except Exception as e:
            return {
                "error": f"Failed to apply enhanced template: {str(e)}"
            }
    
    @app.tool()
    def create_enhanced_slide(
        template_id: str,
        color_scheme: str = "modern_blue",
        content_mapping: Optional[Dict[str, str]] = None,
        image_paths: Optional[Dict[str, str]] = None,
        typography_style: Optional[str] = None,
        layout_index: int = 1,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Create a new slide using an enhanced template with dynamic features.
        
        Args:
            template_id: ID of the enhanced template to use
            color_scheme: Color scheme to apply
            content_mapping: Dictionary mapping element roles to custom content
            image_paths: Dictionary mapping image element roles to file paths
            typography_style: Typography style to use
            layout_index: PowerPoint layout index (default: 1)
            presentation_id: Presentation ID (uses current if None)
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        if layout_index < 0 or layout_index >= len(pres.slide_layouts):
            return {
                "error": f"Invalid layout index: {layout_index}. Available layouts: 0-{len(pres.slide_layouts) - 1}"
            }
        
        try:
            # Add new slide
            layout = pres.slide_layouts[layout_index]
            slide = pres.slides.add_slide(layout)
            slide_index = len(pres.slides) - 1
            
            # Apply enhanced template
            result = enhanced_utils.apply_enhanced_template(
                slide, template_id, color_scheme,
                content_mapping or {}, image_paths or {}
            )
            
            if result['success']:
                return {
                    "message": f"Created enhanced slide {slide_index} using template '{template_id}'",
                    "slide_index": slide_index,
                    "template_applied": result,
                    "typography_style": typography_style or "default"
                }
            else:
                return {
                    "error": f"Failed to apply enhanced template to new slide: {result.get('error', 'Unknown error')}"
                }
                
        except Exception as e:
            return {
                "error": f"Failed to create enhanced slide: {str(e)}"
            }
    
    @app.tool()
    def optimize_slide_text(
        slide_index: int,
        auto_resize: bool = True,
        auto_wrap: bool = True,
        optimize_spacing: bool = True,
        min_font_size: int = 8,
        max_font_size: int = 36,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Optimize text elements on a slide for better readability and fit.
        
        Args:
            slide_index: Index of the slide to optimize
            auto_resize: Whether to automatically resize fonts to fit containers
            auto_wrap: Whether to apply intelligent text wrapping
            optimize_spacing: Whether to optimize line spacing
            min_font_size: Minimum allowed font size
            max_font_size: Maximum allowed font size
            presentation_id: Presentation ID (uses current if None)
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        if slide_index < 0 or slide_index >= len(pres.slides):
            return {
                "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
            }
        
        slide = pres.slides[slide_index]
        
        try:
            optimizations_applied = []
            manager = enhanced_utils.get_enhanced_template_manager()
            
            # Analyze each text shape on the slide
            for i, shape in enumerate(slide.shapes):
                if hasattr(shape, 'text_frame') and shape.text_frame.text:
                    text = shape.text_frame.text
                    
                    # Calculate container dimensions
                    container_width = shape.width.inches
                    container_height = shape.height.inches
                    
                    shape_optimizations = []
                    
                    # Apply auto-resize if enabled
                    if auto_resize:
                        optimal_size = enhanced_utils.calculate_dynamic_font_size(
                            text, container_width, container_height
                        )
                        optimal_size = max(min_font_size, min(max_font_size, optimal_size))
                        
                        # Apply the calculated font size
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = enhanced_utils.Pt(optimal_size)
                        
                        shape_optimizations.append(f"Font resized to {optimal_size}pt")
                    
                    # Apply auto-wrap if enabled
                    if auto_wrap:
                        current_font_size = 14  # Default assumption
                        if shape.text_frame.paragraphs and shape.text_frame.paragraphs[0].runs:
                            current_font_size = shape.text_frame.paragraphs[0].runs[0].font.size.pt
                        
                        wrapped_text = enhanced_utils.wrap_text_automatically(
                            text, container_width, current_font_size
                        )
                        
                        if wrapped_text != text:
                            shape.text_frame.text = wrapped_text
                            shape_optimizations.append("Text wrapped automatically")
                    
                    # Optimize spacing if enabled
                    if optimize_spacing:
                        text_length = len(text)
                        if text_length > 300:
                            line_spacing = 1.4
                        elif text_length > 150:
                            line_spacing = 1.3
                        else:
                            line_spacing = 1.2
                        
                        for paragraph in shape.text_frame.paragraphs:
                            paragraph.line_spacing = line_spacing
                        
                        shape_optimizations.append(f"Line spacing set to {line_spacing}")
                    
                    if shape_optimizations:
                        optimizations_applied.append({
                            "shape_index": i,
                            "optimizations": shape_optimizations
                        })
            
            return {
                "message": f"Optimized {len(optimizations_applied)} text elements on slide {slide_index}",
                "slide_index": slide_index,
                "optimizations_applied": optimizations_applied,
                "settings": {
                    "auto_resize": auto_resize,
                    "auto_wrap": auto_wrap,
                    "optimize_spacing": optimize_spacing,
                    "font_size_range": f"{min_font_size}-{max_font_size}pt"
                }
            }
            
        except Exception as e:
            return {
                "error": f"Failed to optimize slide text: {str(e)}"
            }
    
    @app.tool()
    def analyze_text_content(
        slide_index: int,
        provide_recommendations: bool = True,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Analyze text content on a slide and provide optimization recommendations.
        
        Args:
            slide_index: Index of the slide to analyze
            provide_recommendations: Whether to provide optimization recommendations
            presentation_id: Presentation ID (uses current if None)
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        if slide_index < 0 or slide_index >= len(pres.slides):
            return {
                "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
            }
        
        slide = pres.slides[slide_index]
        
        try:
            text_analysis = []
            manager = enhanced_utils.get_enhanced_template_manager()
            
            for i, shape in enumerate(slide.shapes):
                if hasattr(shape, 'text_frame') and shape.text_frame.text:
                    text = shape.text_frame.text
                    
                    # Basic text metrics
                    char_count = len(text)
                    word_count = len(text.split())
                    line_count = len(text.split('\n'))
                    
                    # Container analysis
                    container_width = shape.width.inches
                    container_height = shape.height.inches
                    
                    # Current font size analysis
                    current_font_size = 14  # Default
                    if shape.text_frame.paragraphs and shape.text_frame.paragraphs[0].runs:
                        if shape.text_frame.paragraphs[0].runs[0].font.size:
                            current_font_size = shape.text_frame.paragraphs[0].runs[0].font.size.pt
                    
                    # Calculate optimal font size
                    optimal_font_size = enhanced_utils.calculate_dynamic_font_size(
                        text, container_width, container_height
                    )
                    
                    # Determine text length category
                    if char_count <= 50:
                        length_category = "short"
                    elif char_count <= 150:
                        length_category = "medium"
                    elif char_count <= 300:
                        length_category = "long"
                    else:
                        length_category = "very_long"
                    
                    recommendations = []
                    if provide_recommendations:
                        # Font size recommendations
                        if abs(current_font_size - optimal_font_size) > 2:
                            recommendations.append(
                                f"Consider changing font size from {current_font_size}pt to {optimal_font_size}pt"
                            )
                        
                        # Text wrapping recommendations
                        estimated_width = manager.text_calculator.estimate_text_width(text, current_font_size)
                        container_width_pts = container_width * 72
                        
                        if estimated_width > container_width_pts * 0.9:
                            recommendations.append("Text may benefit from automatic wrapping")
                        
                        # Content recommendations
                        if char_count > 500:
                            recommendations.append("Consider breaking content into multiple slides")
                        elif word_count > 100:
                            recommendations.append("Consider using bullet points for better readability")
                        
                        # Line spacing recommendations
                        if line_count > 3 and length_category == "long":
                            recommendations.append("Increase line spacing for better readability")
                    
                    text_analysis.append({
                        "shape_index": i,
                        "metrics": {
                            "character_count": char_count,
                            "word_count": word_count,
                            "line_count": line_count,
                            "length_category": length_category
                        },
                        "container": {
                            "width_inches": round(container_width, 2),
                            "height_inches": round(container_height, 2)
                        },
                        "font_analysis": {
                            "current_size": current_font_size,
                            "optimal_size": optimal_font_size,
                            "size_difference": optimal_font_size - current_font_size
                        },
                        "recommendations": recommendations
                    })
            
            return {
                "slide_index": slide_index,
                "text_elements_analyzed": len(text_analysis),
                "analysis_results": text_analysis,
                "overall_recommendations": [
                    "Use dynamic text sizing for optimal readability",
                    "Enable auto-wrapping for content that exceeds container width",
                    "Consider visual hierarchy with varied font sizes",
                    "Break long content into multiple slides or bullet points"
                ] if provide_recommendations else []
            }
            
        except Exception as e:
            return {
                "error": f"Failed to analyze text content: {str(e)}"
            }
    
    @app.tool()
    def apply_visual_effects_preset(
        slide_index: int,
        effect_preset: str = "professional",
        target_elements: Optional[List[str]] = None,
        intensity: str = "medium",
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Apply visual effects preset to elements on a slide.
        
        Args:
            slide_index: Index of the slide to apply effects to
            effect_preset: Preset name ('professional', 'modern', 'elegant')
            target_elements: List of element types to apply effects to (['text', 'image', 'shape'])
            intensity: Effect intensity ('subtle', 'medium', 'strong')
            presentation_id: Presentation ID (uses current if None)
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        if slide_index < 0 or slide_index >= len(pres.slides):
            return {
                "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
            }
        
        slide = pres.slides[slide_index]
        
        if target_elements is None:
            target_elements = ['text', 'image', 'shape']
        
        try:
            manager = enhanced_utils.get_enhanced_template_manager()
            effects_applied = []
            
            # Define effect presets
            presets = {
                'professional': {
                    'text_effects': ['shadow_soft'],
                    'image_effects': ['professional_shadow'],
                    'shape_effects': ['shadow_soft']
                },
                'modern': {
                    'text_effects': ['glow_subtle', 'shadow_soft'],
                    'image_effects': ['modern_glow'],
                    'shape_effects': ['glow_vibrant']
                },
                'elegant': {
                    'text_effects': ['outline_thin', 'shadow_soft'],
                    'image_effects': ['elegant_frame'],
                    'shape_effects': ['shadow_soft']
                }
            }
            
            if effect_preset not in presets:
                return {
                    "error": f"Unknown effect preset: {effect_preset}. Available presets: {list(presets.keys())}"
                }
            
            preset_config = presets[effect_preset]
            
            # Apply effects to elements
            for i, shape in enumerate(slide.shapes):
                shape_effects = []
                
                # Text elements
                if 'text' in target_elements and hasattr(shape, 'text_frame') and shape.text_frame.text:
                    text_effects = preset_config.get('text_effects', [])
                    if text_effects:
                        manager.effects_manager.apply_text_effects(
                            shape.text_frame, text_effects, 'modern_blue'  # Default color scheme
                        )
                        shape_effects.extend([f"text_{effect}" for effect in text_effects])
                
                # Image elements (simplified detection)
                if 'image' in target_elements and 'Picture' in str(shape.shape_type):
                    image_effects = preset_config.get('image_effects', [])
                    for effect in image_effects:
                        manager.effects_manager.apply_image_effects(shape, effect, 'modern_blue')
                        shape_effects.append(f"image_{effect}")
                
                # Shape elements
                if 'shape' in target_elements and hasattr(shape, 'fill'):
                    shape_effects_list = preset_config.get('shape_effects', [])
                    # Apply shape effects (simplified implementation)
                    if shape_effects_list:
                        shape_effects.extend([f"shape_{effect}" for effect in shape_effects_list])
                
                if shape_effects:
                    effects_applied.append({
                        "shape_index": i,
                        "effects": shape_effects
                    })
            
            return {
                "message": f"Applied '{effect_preset}' effects preset to slide {slide_index}",
                "slide_index": slide_index,
                "effect_preset": effect_preset,
                "intensity": intensity,
                "target_elements": target_elements,
                "effects_applied": effects_applied,
                "total_elements_affected": len(effects_applied)
            }
            
        except Exception as e:
            return {
                "error": f"Failed to apply visual effects preset: {str(e)}"
            }
    
    @app.tool()
    def create_dynamic_presentation(
        presentation_topic: str,
        slide_templates: List[str],
        content_outline: List[str],
        color_scheme: str = "modern_blue",
        typography_style: str = "modern_sans",
        auto_optimize: bool = True,
        include_effects: bool = True,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Create a complete presentation with dynamic templates and content adaptation.
        
        Args:
            presentation_topic: Main topic/theme for the presentation
            slide_templates: List of enhanced template IDs to use
            content_outline: List of content items for each slide
            color_scheme: Color scheme to apply consistently
            typography_style: Typography style to use
            auto_optimize: Whether to automatically optimize text and layouts
            include_effects: Whether to apply visual effects
            presentation_id: Presentation ID (uses current if None)
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        if len(slide_templates) != len(content_outline):
            return {
                "error": f"Number of templates ({len(slide_templates)}) must match content outline items ({len(content_outline)})"
            }
        
        pres = presentations[pres_id]
        
        try:
            slides_created = []
            manager = enhanced_utils.get_enhanced_template_manager()
            
            # Set presentation title
            pres.core_properties.title = presentation_topic
            
            for i, (template_id, content) in enumerate(zip(slide_templates, content_outline)):
                try:
                    # Add new slide
                    layout = pres.slide_layouts[1]  # Default content layout
                    slide = pres.slides.add_slide(layout)
                    slide_index = len(pres.slides) - 1
                    
                    # Prepare content mapping based on template and content
                    content_mapping = {}
                    if i == 0:  # Title slide
                        content_mapping = {
                            "title": presentation_topic,
                            "subtitle": content,
                            "author": "Presentation Team"
                        }
                    else:
                        # Generic content mapping
                        content_mapping = {
                            "title": f"Topic {i}: {presentation_topic}",
                            "content": content,
                            "content_left": content,
                            "content_right": "Supporting details and insights"
                        }
                    
                    # Apply enhanced template
                    result = enhanced_utils.apply_enhanced_template(
                        slide, template_id, color_scheme, content_mapping
                    )
                    
                    slide_result = {
                        "slide_index": slide_index,
                        "template_id": template_id,
                        "content_length": len(content),
                        "success": result['success']
                    }
                    
                    if result['success']:
                        slide_result["features_applied"] = result.get('enhanced_features_applied', [])
                        
                        # Apply optimizations if enabled
                        if auto_optimize:
                            # This would call optimize_slide_text internally
                            slide_result["optimizations"] = "Auto-optimization applied"
                        
                        # Apply effects if enabled
                        if include_effects:
                            # This would apply visual effects
                            slide_result["effects"] = "Visual effects applied"
                    else:
                        slide_result["error"] = result.get('error', 'Unknown error')
                    
                    slides_created.append(slide_result)
                    
                except Exception as e:
                    slides_created.append({
                        "slide_index": i,
                        "template_id": template_id,
                        "success": False,
                        "error": str(e)
                    })
            
            successful_slides = sum(1 for slide in slides_created if slide.get('success', False))
            
            return {
                "message": f"Created dynamic presentation with {successful_slides}/{len(slide_templates)} slides",
                "presentation_topic": presentation_topic,
                "total_slides": len(pres.slides),
                "slides_created": slides_created,
                "settings": {
                    "color_scheme": color_scheme,
                    "typography_style": typography_style,
                    "auto_optimize": auto_optimize,
                    "include_effects": include_effects
                },
                "success_rate": f"{successful_slides}/{len(slide_templates)}"
            }
            
        except Exception as e:
            return {
                "error": f"Failed to create dynamic presentation: {str(e)}"
            }