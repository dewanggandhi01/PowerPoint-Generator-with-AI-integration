from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import tempfile
import logging
import re
from typing import Dict, List, Any
from .llm_service import LLMService
from .template_analyzer import TemplateAnalyzer

logger = logging.getLogger(__name__)

class PresentationGenerator:
    def __init__(self, llm_service: LLMService, template_analyzer: TemplateAnalyzer):
        self.llm_service = llm_service
        self.template_analyzer = template_analyzer
        self.user_images = []
    
    def generate(self, text_input: str, guidance: str = "", image_paths: List[str] = None) -> str:
        logger.info("Starting presentation generation")
        self.user_images = image_paths or []
        structure = self.llm_service.analyze_text_structure(text_input, guidance)
        presentation = self._create_presentation_from_template()
        self._generate_slides_with_images(presentation, structure)
        output_path = self._save_presentation(presentation)
        logger.info(f"Presentation generated successfully: {output_path}")
        return output_path
    
    def _create_presentation_from_template(self) -> Presentation:
        try:
            return Presentation(self.template_analyzer.template_path)
        except Exception as e:
            logger.warning(f"Could not use template directly: {str(e)}")
            return Presentation()
    
    def _generate_slides_with_images(self, presentation: Presentation, structure: Dict[str, Any]):
        # Clear existing slides
        for i in reversed(range(len(presentation.slides))):
            r_id = presentation.slides._sldIdLst[i].rId
            presentation.part.drop_rel(r_id)
            del presentation.slides._sldIdLst[i]
        
        slides_data = structure.get('slides', [])
        image_index = 0
        
        for i, slide_data in enumerate(slides_data):
            self._create_slide(presentation, slide_data)
            if (self.user_images and image_index < len(self.user_images) and 
                slide_data.get('slide_type') == 'content' and (i + 1) % 3 == 0):
                self._create_image_slide(presentation, self.user_images[image_index], i + 1)
                image_index += 1

    def _create_image_slide(self, presentation: Presentation, image_path: str, slide_number: int):
        try:
            layout_index = self._find_content_layout_for_image(presentation)
            slide_layout = presentation.slide_layouts[layout_index]
            slide = presentation.slides.add_slide(slide_layout)
            slide_index = len(presentation.slides) - 1
            
            if slide.shapes.title:
                slide.shapes.title.text = f"Visual Insight {slide_number}"
                self._style_title(slide.shapes.title)
            
            self._clear_content_placeholder(slide)
            self._add_image_and_text_custom_layout(slide, image_path, slide_number)
            self._add_slide_animations(slide, slide_index)
        except Exception as e:
            logger.error(f"Error creating image slide: {str(e)}")
            self._create_template_based_image_slide(presentation, image_path, slide_number)

    def _find_content_layout_for_image(self, presentation: Presentation) -> int:
        layouts = presentation.slide_layouts
        for i, layout in enumerate(layouts):
            if any(keyword in layout.name.lower() for keyword in 
                   ['content', 'text', 'bullet', 'two content']):
                return i
        return 1 if len(layouts) > 1 else 0

    def _clear_content_placeholder(self, slide):
        try:
            for shape in slide.placeholders:
                if shape.placeholder_format.type == 2:
                    if hasattr(shape, 'text_frame'):
                        shape.text_frame.clear()
                    break
        except Exception as e:
            logger.warning(f"Could not clear content placeholder: {str(e)}")

    def _add_image_and_text_custom_layout(self, slide, image_path: str, slide_number: int):
        try:
            slide.shapes.add_picture(image_path, Inches(0.5), Inches(2.2), Inches(5.5), Inches(4))
            textbox = slide.shapes.add_textbox(Inches(6.2), Inches(2.5), Inches(3.3), Inches(3.5))
            text_frame = textbox.text_frame
            text_frame.word_wrap = True
            text_frame.auto_size = None
            
            caption_content = ["Key visual supporting our strategic analysis", "Important data visualization and insights", 
                             "Critical business intelligence demonstration", "Contextual evidence and supporting material",
                             "Strategic decision-making reference"]
            
            for i, item in enumerate(caption_content):
                p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
                p.text = f"• {item}"
                p.level = 0
                p.space_after = Pt(10)
                self._style_enhanced_content_paragraph(p, 'content')
        except Exception as e:
            logger.warning(f"Could not add image and text: {str(e)}")

    def _create_template_based_image_slide(self, presentation: Presentation, image_path: str, slide_number: int):
        try:
            layout_index = min(1, len(presentation.slide_layouts) - 1)
            slide = presentation.slides.add_slide(presentation.slide_layouts[layout_index])
            if slide.shapes.title:
                slide.shapes.title.text = f"Supporting Visual Evidence {slide_number}"
                self._style_title(slide.shapes.title)
            slide.shapes.add_picture(image_path, Inches(1.5), Inches(2.5), Inches(7), Inches(4))
        except Exception as e:
            logger.error(f"Failed to create template-based image slide: {str(e)}")

    def _generate_slides(self, presentation: Presentation, structure: Dict[str, Any]):
        self._generate_slides_with_images(presentation, structure)
    
    def _create_slide(self, presentation: Presentation, slide_data: Dict[str, Any]):
        slide_type = slide_data.get('slide_type', 'content')
        layout_index = self.template_analyzer.get_best_layout_for_slide_type(slide_type)
        
        try:
            slide = presentation.slides.add_slide(presentation.slide_layouts[layout_index])
            slide_index = len(presentation.slides) - 1
            title = slide_data.get('title', 'Slide Title')
            
            if len(title) > 60:
                title = title[:60] + "..."
                
            if slide.shapes.title:
                slide.shapes.title.text = self._clean_title(title)
                slide.shapes.title.top = Inches(0.15)
                slide.shapes.title.height = Inches(1.8)
                self._style_title(slide.shapes.title)
            else:
                self._create_manual_title(slide, title)
            
            content = slide_data.get('content', [])
            self._add_enhanced_slide_content(slide, content, slide_type)
            
            emphasis_points = slide_data.get('emphasis_points', [])
            if emphasis_points:
                self._add_emphasis_content(slide, emphasis_points)
            
            self._add_slide_animations(slide, slide_index)
            
            speaking_notes = slide_data.get('speaking_notes', '')
            if speaking_notes:
                self._add_detailed_speaker_notes(slide, speaking_notes)
            else:
                self._add_speaker_notes(slide, title, content)
        except Exception as e:
            logger.error(f"Error creating slide: {str(e)}")
            self._create_enhanced_basic_slide(presentation, slide_data)

    def _create_manual_title(self, slide, title_text: str):
        try:
            slide_width = self.template_analyzer.get_slide_dimensions()[0]
            if len(title_text) > 60:
                title_text = title_text[:60] + "..."
            
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), slide_width - Inches(1), Inches(1.8))
            title_frame = title_box.text_frame
            title_frame.text = self._clean_title(title_text)
            title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            title_frame.margin_left = title_frame.margin_right = Inches(0.2)
            title_frame.margin_top = title_frame.margin_bottom = Inches(0.15)
            
            for paragraph in title_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                paragraph.space_after = Pt(28)
                paragraph.space_before = Pt(8)
                
                for run in paragraph.runs:
                    fonts = self.template_analyzer.get_theme_fonts()
                    colors = self.template_analyzer.get_theme_colors()
                    run.font.name = fonts.get('title', 'Arial Black')
                    run.font.size = Pt(48)
                    run.font.bold = True
                    run.font.underline = True
                    self._apply_high_contrast_color(run, colors, is_title=True)
                    
                    if 'primary' in colors:
                        color_hex = colors['primary'].lstrip('#')
                        rgb = tuple(int(color_hex[i:i+2], 16) for i in (0, 2, 4))
                        run.font.color.rgb = RGBColor(*rgb)
        except Exception as e:
            logger.warning(f"Could not create manual title: {str(e)}")
    
    def _create_basic_slide(self, presentation: Presentation, slide_data: Dict[str, Any]):
        slide = presentation.slides.add_slide(presentation.slide_layouts[0])
        if slide.shapes.title:
            slide.shapes.title.text = self._clean_title(slide_data.get('title', 'Slide Title'))
        
        content = slide_data.get('content', [])
        if content:
            textbox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
            text_frame = textbox.text_frame
            for i, item in enumerate(content):
                p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
                p.text = f"• {item}"
                p.level = 0
    
    def _add_enhanced_slide_content(self, slide, content: List[str], slide_type: str):
        if not content:
            return
        
        slide_width, slide_height = self.template_analyzer.get_slide_dimensions()
        fitted_content = self._ensure_content_fits_slide(content, slide_height)
        
        # Clean and format content before adding to slide
        cleaned_content = self._clean_slide_content(fitted_content)
        
        content_placeholder = None
        for shape in slide.placeholders:
            if shape.placeholder_format.type == 2:
                content_placeholder = shape
                break
        
        if content_placeholder:
            safe_margin = Inches(0.8)
            content_placeholder.left = safe_margin
            content_placeholder.top = Inches(2.5)
            content_placeholder.width = slide_width - (safe_margin * 2)
            content_placeholder.height = slide_height - Inches(4.2)
            
            text_frame = content_placeholder.text_frame
            text_frame.clear()
            text_frame.word_wrap = True
            text_frame.auto_size = None
            text_frame.margin_top = text_frame.margin_bottom = Inches(0.3)
            text_frame.margin_left = text_frame.margin_right = Inches(0.4)
            
            for i, item in enumerate(cleaned_content):
                p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
                p.text = item
                p.level = 0
                p.space_after = Pt(10)
                self._style_enhanced_content_paragraph(p, slide_type)
                
                # Handle sub-points more intelligently
                if ':' in item and len(item) > 80 and i < 4:
                    parts = item.split(':', 1)
                    if len(parts) == 2 and len(parts[1].strip()) > 10:
                        p.text = parts[0] + ':'
                        sub_p = text_frame.add_paragraph()
                        sub_p.text = f"• {parts[1].strip()}"
                        sub_p.level = 1
                        sub_p.space_after = Pt(8)
                        self._style_enhanced_content_paragraph(sub_p, slide_type, is_sub=True)
        else:
            self._create_enhanced_content_textbox(slide, cleaned_content, slide_type)

    def _clean_slide_content(self, content: List[str]) -> List[str]:
        """Clean and format slide content for better presentation."""
        cleaned = []
        for item in content:
            if not item or len(item.strip()) < 3:
                continue
                
            # Remove encoding artifacts
            clean_item = item.replace('_x000D_', ' ').replace('_x000A_', ' ')
            clean_item = clean_item.replace('\\n', ' ').replace('\\r', '')
            
            # Clean excessive whitespace
            clean_item = ' '.join(clean_item.split())
            
            # Remove repetitive prefixes
            if clean_item.startswith('Analysis: Analysis:'):
                clean_item = clean_item.replace('Analysis: Analysis:', 'Analysis:', 1)
            
            # Ensure proper bullet formatting
            if not clean_item.startswith(('●', '•', '-')) and ':' not in clean_item:
                # Add descriptive prefix based on content
                if any(word in clean_item.lower() for word in ['market', 'revenue', 'growth']):
                    clean_item = f"Market Insight: {clean_item}"
                elif any(word in clean_item.lower() for word in ['feature', 'benefit', 'advantage']):
                    clean_item = f"Key Benefit: {clean_item}"
                elif any(word in clean_item.lower() for word in ['strategy', 'plan', 'approach']):
                    clean_item = f"Strategic Approach: {clean_item}"
                elif any(word in clean_item.lower() for word in ['result', 'outcome', 'impact']):
                    clean_item = f"Expected Outcome: {clean_item}"
                else:
                    clean_item = f"Key Point: {clean_item}"
            
            # Ensure proper capitalization and punctuation
            if not clean_item.endswith(('.', '!', '?')):
                clean_item += '.'
                
            cleaned.append(clean_item)
        
        return cleaned
    
    def _clean_title(self, title: str) -> str:
        """Clean and format slide titles"""
        # Remove encoding artifacts
        cleaned = title.replace('_x000D_', '').replace('_x000A_', '').replace('\r', '').replace('\n', ' ')
        
        # Remove markdown headers
        cleaned = re.sub(r'#+\s*', '', cleaned)
        
        # Remove repetitive patterns
        cleaned = re.sub(r'Analysis:\s*', '', cleaned, flags=re.IGNORECASE)
        
        # Clean extra whitespace
        cleaned = ' '.join(cleaned.split())
        
        # Ensure proper title case
        if cleaned:
            words = cleaned.split()
            # Capitalize first word and important words
            for i, word in enumerate(words):
                if i == 0 or len(word) > 3 or word.upper() in ['API', 'AI', 'ML', 'UI', 'UX']:
                    words[i] = word.capitalize()
            cleaned = ' '.join(words)
        
        return cleaned.strip()

    def _add_emphasis_content(self, slide, emphasis_points: List[str]):
        if not emphasis_points:
            return
            
        try:
            emphasis_box = slide.shapes.add_textbox(Inches(1), Inches(6), Inches(8), Inches(1))
            text_frame = emphasis_box.text_frame
            text_frame.word_wrap = True
            
            for i, point in enumerate(emphasis_points):
                p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
                p.text = f"💡 {point}"
                p.alignment = PP_ALIGN.CENTER
                
                for run in p.runs:
                    run.font.size = Pt(14)
                    run.font.bold = True
                    colors = self.template_analyzer.get_theme_colors()
                    bg_colors = self.template_analyzer.get_background_colors()
                    
                    if self._is_dark_background(bg_colors):
                        run.font.color.rgb = RGBColor(255, 255, 255)
                    else:
                        run.font.color.rgb = RGBColor(180, 60, 0)
        except Exception as e:
            logger.warning(f"Could not add emphasis content: {str(e)}")

    def _create_enhanced_content_textbox(self, slide, content: List[str], slide_type: str):
        slide_width, slide_height = self.template_analyzer.get_slide_dimensions()
        safe_margin = Inches(0.8)
        
        # Clean content before adding to textbox
        cleaned_content = self._clean_slide_content(content)
        
        textbox = slide.shapes.add_textbox(safe_margin, Inches(2.5), 
                                         slide_width - (safe_margin * 2), 
                                         slide_height - Inches(4.2))
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.auto_size = None
        text_frame.margin_top = text_frame.margin_bottom = Inches(0.3)
        text_frame.margin_left = text_frame.margin_right = Inches(0.4)
        
        for i, item in enumerate(cleaned_content):
            p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
            
            # Use better bullet symbols based on slide type
            if slide_type == 'comparison':
                bullet_symbol = "▶" if i % 2 == 0 else "◀"
            elif slide_type == 'section':
                bullet_symbol = "◆"
            else:
                bullet_symbol = "●"
            
            # Only add bullet if content doesn't already have formatting
            if not item.startswith(('●', '•', '-', '▶', '◀', '◆')):
                p.text = f"{bullet_symbol} {item}"
            else:
                p.text = item
                
            p.level = 0
            p.space_after = Pt(12)
            self._style_enhanced_content_paragraph(p, slide_type)

    def _create_enhanced_basic_slide(self, presentation: Presentation, slide_data: Dict[str, Any]):
        slide_type = slide_data.get('slide_type', 'content')
        layouts = presentation.slide_layouts
        layout_index = min(1, len(layouts) - 1) if len(layouts) > 1 else 0
        slide = presentation.slides.add_slide(layouts[layout_index])
        
        title = slide_data.get('title', 'Slide Title')
        if slide.shapes.title:
            slide.shapes.title.text = self._clean_title(title)
            self._style_title(slide.shapes.title)
        
        content = slide_data.get('content', [])
        if content:
            self._create_enhanced_content_textbox(slide, content, slide_type)
        
        emphasis_points = slide_data.get('emphasis_points', [])
        if emphasis_points:
            self._add_emphasis_content(slide, emphasis_points)
    
    def _style_title(self, title_shape):
        try:
            fonts = self.template_analyzer.get_theme_fonts()
            colors = self.template_analyzer.get_theme_colors()
            slide_width = self.template_analyzer.get_slide_dimensions()[0]
            
            title_shape.left = Inches(0.5)
            title_shape.top = Inches(0.15)
            title_shape.width = slide_width - Inches(1)
            title_shape.height = Inches(1.8)
            
            title_shape.text_frame.margin_bottom = title_shape.text_frame.margin_top = Inches(0.15)
            title_shape.text_frame.margin_left = title_shape.text_frame.margin_right = Inches(0.2)
            title_shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            for paragraph in title_shape.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.font.name = fonts.get('title', 'Arial Black')
                    run.font.size = Pt(48)
                    run.font.bold = True
                    run.font.underline = True
                    self._apply_high_contrast_color(run, colors, is_title=True)
                paragraph.space_after = Pt(28)
                paragraph.space_before = Pt(8)
        except Exception as e:
            logger.warning(f"Error styling title: {str(e)}")
    
    def _style_enhanced_content_paragraph(self, paragraph, slide_type: str, is_sub: bool = False):
        try:
            fonts = self.template_analyzer.get_theme_fonts()
            colors = self.template_analyzer.get_theme_colors()
            paragraph.alignment = PP_ALIGN.JUSTIFY
            
            for run in paragraph.runs:
                run.font.name = fonts.get('body', 'Calibri')
                if is_sub:
                    run.font.size = Pt(20)
                    run.font.italic = True
                elif slide_type == 'title':
                    run.font.size = Pt(28)
                    run.font.bold = True
                elif slide_type == 'section':
                    run.font.size = Pt(24)
                    run.font.bold = True
                else:
                    run.font.size = Pt(22)
                self._apply_high_contrast_color(run, colors, is_title=(slide_type in ['title', 'section']))
        except Exception as e:
            logger.warning(f"Error styling paragraph: {str(e)}")

    def _apply_high_contrast_color(self, run, colors, is_title=False):
        try:
            bg_colors = self.template_analyzer.get_background_colors()
            
            if is_title:
                if self._is_dark_background(bg_colors):
                    run.font.color.rgb = RGBColor(255, 255, 255)
                else:
                    run.font.color.rgb = RGBColor(20, 20, 20)
            else:
                if self._is_dark_background(bg_colors):
                    run.font.color.rgb = RGBColor(255, 255, 255)
                else:
                    run.font.color.rgb = RGBColor(40, 40, 40)
                    
            if 'primary' in colors and self._is_color_dark_enough(colors['primary']):
                color_hex = colors['primary'].lstrip('#')
                rgb = tuple(int(color_hex[i:i+2], 16) for i in (0, 2, 4))
                luminance = (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255
                if luminance < 0.6:
                    run.font.color.rgb = RGBColor(*rgb)
        except Exception as e:
            logger.warning(f"Error applying high contrast color: {str(e)}")
            run.font.color.rgb = RGBColor(30, 30, 30)

    def _is_dark_background(self, bg_colors):
        try:
            if not bg_colors:
                return False
            bg_color = bg_colors.get('primary', '#FFFFFF')
            if bg_color.startswith('#'):
                bg_color = bg_color[1:]
            r, g, b = tuple(int(bg_color[i:i+2], 16) for i in (0, 2, 4))
            luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
            return luminance < 0.5
        except:
            return False

    def _is_color_dark_enough(self, color):
        try:
            if not color:
                return False
            if color.startswith('#'):
                color = color[1:]
            r, g, b = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
            luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
            return luminance < 0.6
        except:
            return False

    def _has_good_contrast(self, color1, bg_colors):
        try:
            if not bg_colors or not color1:
                return False
            bg_color = bg_colors.get('primary', '#FFFFFF')
            if color1.startswith('#'):
                color1 = color1[1:]
            if bg_color.startswith('#'):
                bg_color = bg_color[1:]
            r1, g1, b1 = tuple(int(color1[i:i+2], 16) for i in (0, 2, 4))
            r2, g2, b2 = tuple(int(bg_color[i:i+2], 16) for i in (0, 2, 4))
            diff = abs(r1 - r2) + abs(g1 - g2) + abs(b1 - b2)
            return diff > 200
        except:
            return False

    def _add_slide_animations(self, slide, slide_index):
        try:
            transitions = ['fade', 'push', 'wipe', 'split', 'reveal', 'cover', 'cut']
            transition_type = transitions[slide_index % len(transitions)]
            self._add_text_animations(slide)
        except Exception as e:
            logger.warning(f"Error adding slide animations: {str(e)}")

    def _add_text_animations(self, slide):
        try:
            for shape in slide.shapes:
                if hasattr(shape, 'text_frame') and shape.text_frame.text.strip():
                    self._simulate_entrance_effect(shape)
        except Exception as e:
            logger.warning(f"Error adding text animations: {str(e)}")

    def _simulate_entrance_effect(self, shape):
        try:
            if hasattr(shape, 'text_frame'):
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        if hasattr(run.font, 'color'):
                            pass
        except Exception as e:
            logger.warning(f"Error simulating entrance effect: {str(e)}")

    def _ensure_content_fits_slide(self, content: List[str], slide_height: float) -> List[str]:
        try:
            available_height = slide_height - Inches(4.5)
            line_height = Pt(30)
            max_lines = int(available_height / line_height)
            max_items = min(max_lines // 2, 8)
            fitted_content = content[:max_items]
            
            for i, item in enumerate(fitted_content):
                if len(item) > 180:
                    fitted_content[i] = item[:180] + "..."
            return fitted_content
        except Exception as e:
            logger.warning(f"Error fitting content to slide: {str(e)}")
            return content[:6]

    def _add_detailed_speaker_notes(self, slide, speaking_notes: str):
        try:
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = speaking_notes
        except Exception as e:
            logger.warning(f"Error adding detailed speaker notes: {str(e)}")

    def _style_content_paragraph(self, paragraph):
        try:
            fonts = self.template_analyzer.get_theme_fonts()
            paragraph.alignment = PP_ALIGN.JUSTIFY
            for run in paragraph.runs:
                run.font.name = fonts.get('body', 'Calibri')
                run.font.size = Pt(22)
                bg_colors = self.template_analyzer.get_background_colors()
                if self._is_dark_background(bg_colors):
                    run.font.color.rgb = RGBColor(255, 255, 255)
                else:
                    run.font.color.rgb = RGBColor(40, 40, 40)
        except Exception as e:
            logger.warning(f"Error styling paragraph: {str(e)}")

    def _create_content_textbox(self, slide, content: List[str]):
        slide_width, slide_height = self.template_analyzer.get_slide_dimensions()
        safe_margin = Inches(0.8)
        
        textbox = slide.shapes.add_textbox(safe_margin, Inches(2.5), 
                                         slide_width - (safe_margin * 2), 
                                         slide_height - Inches(4.2))
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.auto_size = None
        text_frame.margin_left = text_frame.margin_right = Inches(0.4)
        text_frame.margin_top = text_frame.margin_bottom = Inches(0.3)
        
        for i, item in enumerate(content):
            if len(item) > 180:
                item = item[:180] + "..."
            p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
            p.text = f"• {item}"
            p.level = 0
            p.space_after = Pt(10)
            self._style_content_paragraph(p)

    def _add_speaker_notes(self, slide, title: str, content: List[str]):
        try:
            slide_content = f"Title: {title}\nContent: {'; '.join(content)}"
            speaker_notes = self.llm_service.generate_speaker_notes(slide_content)
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = speaker_notes
        except Exception as e:
            logger.warning(f"Error adding speaker notes: {str(e)}")
    
    def _save_presentation(self, presentation: Presentation) -> str:
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False, dir=tempfile.gettempdir()) as tmp_file:
            output_path = tmp_file.name
        presentation.save(output_path)
        return output_path
