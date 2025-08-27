"""
LLM Service module for handling different LLM providers.
"""
import requests
import json
import logging
from typing import Dict, List, Any

# Import LLM libraries with try/except to handle missing packages gracefully
try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OpenAI = None
    OPENAI_AVAILABLE = False

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    anthropic = None
    ANTHROPIC_AVAILABLE = False

try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except (ImportError, Exception):
    genai = None
    GEMINI_AVAILABLE = False

logger = logging.getLogger(__name__)

class LLMService:
    """Service for interacting with different LLM providers."""
    
    def __init__(self, provider: str, api_key: str):
        self.provider = provider.lower()
        self.api_key = api_key
        self._setup_client()
    
    def _setup_client(self):
        """Setup the appropriate LLM client based on provider."""
        if self.provider == 'openai':
            if not OPENAI_AVAILABLE:
                raise ImportError("OpenAI package is required but not available")
            self.client = OpenAI(api_key=self.api_key)
        elif self.provider == 'anthropic':
            if not ANTHROPIC_AVAILABLE:
                raise ImportError("Anthropic package is required but not available")
            self.client = anthropic.Anthropic(api_key=self.api_key)
        elif self.provider == 'gemini':
            if not GEMINI_AVAILABLE:
                raise ImportError("Google GenerativeAI package is required but not available")
            genai.configure(api_key=self.api_key)
            self.client = genai.GenerativeModel('gemini-pro')
        elif self.provider == 'aipipe':
            # AIPipe.org uses a custom API endpoint
            self.client = None  # We'll use requests directly
            self.api_base = "https://aipipe.org/api/v1"
        else:
            raise ValueError(f"Unsupported LLM provider: {self.provider}")
    
    def analyze_text_structure(self, text: str, guidance: str = "") -> Dict[str, Any]:
        """
        Analyze input text and generate slide structure.
        
        Args:
            text: Input text to analyze
            guidance: Optional guidance for tone/structure
            
        Returns:
            Dictionary with slide structure and content
        """
        prompt = self._create_structure_prompt(text, guidance)
        
        try:
            response = self._make_llm_call(prompt)
            return self._parse_structure_response(response)
        except Exception as e:
            logger.error(f"Error analyzing text structure: {str(e)}")
            return self._fallback_structure(text)
    
    def generate_speaker_notes(self, slide_content: str) -> str:
        """
        Generate speaker notes for a slide.
        
        Args:
            slide_content: Content of the slide
            
        Returns:
            Generated speaker notes
        """
        prompt = f"""
        Generate concise speaker notes for this slide content. 
        Keep it professional and helpful for a presenter.
        
        Slide Content:
        {slide_content}
        
        Speaker Notes:
        """
        
        try:
            return self._make_llm_call(prompt).strip()
        except Exception as e:
            logger.error(f"Error generating speaker notes: {str(e)}")
            return "Key points to discuss based on slide content."
    
    def _create_structure_prompt(self, text: str, guidance: str) -> str:
        """Create prompt for text structure analysis."""
        base_prompt = f"""
        Analyze the following text and break it down into a comprehensive PowerPoint presentation structure.
        Create a professional presentation with rich content and proper slide distribution.
        
        Text to analyze:
        {text}
        
        {f"Guidance: {guidance}" if guidance else ""}
        
        Please provide a JSON response with the following structure:
        {{
            "title": "Compelling Presentation Title",
            "slides": [
                {{
                    "title": "Slide Title",
                    "content": ["Point 1 with details", "Point 2 with explanation", "Point 3 with context", "Point 4 with benefits", "Point 5 with examples"],
                    "slide_type": "title|section|content|comparison|conclusion",
                    "emphasis_points": ["key point to highlight", "important statistic"],
                    "speaking_notes": "Detailed talking points for this slide"
                }}
            ]
        }}
        
        CRITICAL REQUIREMENTS:
        1. Create 8-15 slides (not just 5-6)
        2. Each content slide MUST have 4-7 bullet points (not 2-3)
        3. Make bullet points substantial with details, examples, or explanations
        4. Include different slide types: title, section headers, content, comparison, conclusion
        5. Add emphasis_points for key statistics or important highlights
        6. Create logical flow with section breaks for major topics
        7. Make titles engaging and descriptive
        8. Ensure content is presentation-ready, not just summary points
        
        Example good content structure:
        - "Revenue Growth: Increased by 45% year-over-year due to new market expansion"
        - "Key Benefits: Reduced operational costs, improved efficiency, enhanced customer satisfaction"
        - "Market Analysis: Targeting millennials aged 25-40 in urban areas with disposable income above $50k"
        
        Make it professional, engaging, and visually rich for PowerPoint presentation.
        """
        
        return base_prompt
    
    def _make_llm_call(self, prompt: str) -> str:
        """Make API call to the configured LLM provider."""
        if self.provider == 'openai':
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=2000,
                temperature=0.7
            )
            return response.choices[0].message.content
        
        elif self.provider == 'anthropic':
            response = self.client.messages.create(
                model="claude-3-sonnet-20240229",
                max_tokens=2000,
                messages=[{"role": "user", "content": prompt}]
            )
            return response.content[0].text
        
        elif self.provider == 'gemini':
            response = self.client.generate_content(prompt)
            return response.text
        
        elif self.provider == 'aipipe':
            # AIPipe.org API call using requests
            headers = {
                'Authorization': f'Bearer {self.api_key}',
                'Content-Type': 'application/json'
            }
            
            payload = {
                'model': 'gpt-3.5-turbo',  # Default model, can be configured
                'messages': [{'role': 'user', 'content': prompt}],
                'max_tokens': 2000,
                'temperature': 0.7
            }
            
            response = requests.post(
                f"{self.api_base}/chat/completions",
                headers=headers,
                json=payload,
                timeout=60
            )
            
            if response.status_code == 200:
                result = response.json()
                return result['choices'][0]['message']['content']
            else:
                raise Exception(f"AIPipe API error: {response.status_code} - {response.text}")
        
        else:
            raise ValueError(f"Unsupported provider: {self.provider}")
    
    def _parse_structure_response(self, response: str) -> Dict[str, Any]:
        """Parse the LLM response into structured data."""
        try:
            # Try to extract JSON from response
            start_idx = response.find('{')
            end_idx = response.rfind('}') + 1
            
            if start_idx != -1 and end_idx != 0:
                json_str = response[start_idx:end_idx]
                return json.loads(json_str)
            else:
                raise ValueError("No JSON found in response")
        
        except (json.JSONDecodeError, ValueError) as e:
            logger.warning(f"Failed to parse LLM response as JSON: {str(e)}")
            return self._fallback_structure(response)
    
    def _fallback_structure(self, text: str) -> Dict[str, Any]:
        """Create an enhanced fallback structure when LLM parsing fails."""
        # Enhanced text splitting fallback
        sentences = [s.strip() for s in text.split('.') if s.strip() and len(s.strip()) > 10]
        paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
        
        # Use paragraphs if available, otherwise sentences
        content_chunks = paragraphs if len(paragraphs) > 2 else sentences
        
        slides = []
        
        # Title slide
        slides.append({
            "title": "Presentation Overview",
            "content": [
                "Key topics and insights to be covered",
                "Main objectives and expected outcomes", 
                "Strategic importance and relevance",
                "Action items and next steps"
            ],
            "slide_type": "title",
            "emphasis_points": ["Comprehensive analysis", "Data-driven insights"],
            "speaking_notes": "Welcome the audience and provide overview of presentation structure and key objectives."
        })
        
        # Process content in larger, more meaningful chunks
        slides_per_chunk = max(4, len(content_chunks) // 6)  # Aim for 6-8 content slides
        
        for i in range(0, len(content_chunks), slides_per_chunk):
            chunk = content_chunks[i:i + slides_per_chunk]
            if chunk:
                slide_num = len(slides)
                
                # Create section header every 3 slides
                if slide_num > 1 and (slide_num - 1) % 3 == 0:
                    slides.append({
                        "title": f"Section {(slide_num-1)//3 + 1}: Key Analysis",
                        "content": [
                            "Detailed examination of core concepts",
                            "Critical insights and findings",
                            "Strategic implications and impact"
                        ],
                        "slide_type": "section",
                        "emphasis_points": ["Strategic focus area"],
                        "speaking_notes": "Transition to new section with key focus areas."
                    })
                
                # Enhanced content processing
                processed_content = []
                for item in chunk:
                    # Make content more presentation-friendly
                    if len(item) > 150:
                        # Split long sentences into main point and details
                        parts = item.split(',', 1)
                        processed_content.append(parts[0] + " - Key insight")
                        if len(parts) > 1:
                            processed_content.append(f"Details: {parts[1].strip()}")
                    else:
                        processed_content.append(f"Analysis: {item}")
                
                # Ensure minimum content per slide
                while len(processed_content) < 4:
                    processed_content.append(f"Supporting point {len(processed_content)}: Additional context and relevance")
                
                slides.append({
                    "title": f"Analysis Point {slide_num}",
                    "content": processed_content[:6],  # Max 6 points per slide
                    "slide_type": "content",
                    "emphasis_points": [f"Key insight #{slide_num}"],
                    "speaking_notes": f"Detailed discussion of analysis point {slide_num} with supporting evidence and examples."
                })
        
        # Enhanced conclusion slide
        slides.append({
            "title": "Key Takeaways & Next Steps",
            "content": [
                "Summary of critical findings and insights",
                "Strategic recommendations and action items",
                "Implementation timeline and milestones",
                "Success metrics and evaluation criteria",
                "Future opportunities and considerations"
            ],
            "slide_type": "conclusion",
            "emphasis_points": ["Action required", "Success metrics"],
            "speaking_notes": "Summarize key points and provide clear next steps with timeline and ownership."
        })
        
        return {
            "title": "Professional Analysis & Insights",
            "slides": slides
        }
