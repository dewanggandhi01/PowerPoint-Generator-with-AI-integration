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
        """Create comprehensive prompt for professional presentation structure analysis."""
        # Clean the input text first
        cleaned_text = self._clean_input_text(text)
        
        enhanced_prompt = f"""
        🎯 EXECUTIVE PRESENTATION CONSULTANT BRIEF 🎯

        You are an elite presentation strategist tasked with transforming raw content into a compelling, professional PowerPoint presentation. Your mission: Create presentation-ready content that maximizes impact while maintaining 100% fidelity to the source material.

        📋 SOURCE CONTENT:
        {cleaned_text}

        🎯 PRESENTATION CONTEXT: {guidance if guidance else "Professional business presentation"}

        🏆 EXCELLENCE STANDARDS:

        ✅ CONTENT INTEGRITY:
        - Every slide MUST derive directly from the provided content
        - Preserve all specific metrics, dates, percentages, and technical details
        - Expand and enhance source material, never add external information
        - Transform complex paragraphs into clear, actionable insights

        ✅ PROFESSIONAL IMPACT:
        - Use executive-level language with strong value propositions
        - Create compelling narratives that drive decision-making
        - Structure content for visual consumption, not reading
        - Include strategic insights and business implications

        ✅ VISUAL OPTIMIZATION:
        - Maximum 5 bullet points per slide
        - Each bullet: 8-15 words with specific, measurable details
        - Use active voice and power words
        - Structure for maximum visual hierarchy and flow

        ✅ ADVANCED FORMATTING:
        - Zero encoding artifacts (_x000D_, _x000A_, etc.)
        - No repetitive labels ("Analysis:", "Overview:")
        - Varied, descriptive prefixes that add context
        - Professional slide titles with clear value propositions

        📊 REQUIRED OUTPUT STRUCTURE:

        {{
            "title": "Compelling Main Title (specific to content, not generic)",
            "slides": [
                {{
                    "title": "Strategic Value-Driven Title",
                    "content": [
                        "Market Intelligence: $X.XB market growing XX% annually with specific trend analysis",
                        "Competitive Advantage: Unique differentiator delivering measurable business value",
                        "Implementation Strategy: Specific action with timeline and resource requirements",
                        "ROI Projection: Quantified benefits with timeframe and success metrics",
                        "Risk Mitigation: Specific challenge addressed with proven solution approach"
                    ],
                    "slide_type": "title|overview|strategy|analysis|implementation|conclusion",
                    "emphasis_points": ["Key metric or statistic", "Critical success factor"],
                    "speaking_notes": "Executive talking points with context and supporting details"
                }}
            ]
        }}

        🎨 CONTENT TRANSFORMATION EXAMPLES:

        BEFORE (Raw): "The fitness app market is valued at $4.4 billion, growing 14.7% annually"
        AFTER (Professional): "Market Opportunity: $4.4B fitness technology sector expanding 14.7% YoY"

        BEFORE (Raw): "87% of people struggle with consistent workouts"
        AFTER (Professional): "User Pain Point: 87% abandonment rate creates $3.8B addressable market gap"

        BEFORE (Raw): "AI Personal Trainer with real-time form analysis"
        AFTER (Professional): "Innovation Core: AI-powered form correction increases workout effectiveness 40%"

        🚀 SLIDE ARCHITECTURE:
        1. HOOK SLIDE: Compelling problem/opportunity with quantified impact
        2. CONTEXT SLIDES: Market landscape, user needs, competitive positioning
        3. SOLUTION SLIDES: Core value proposition with differentiated features
        4. VALIDATION SLIDES: Evidence, metrics, proof points, testimonials
        5. EXECUTION SLIDES: Implementation roadmap, resource requirements, timeline
        6. IMPACT SLIDES: ROI projections, success metrics, scaling potential
        7. ACTION SLIDE: Clear next steps with ownership and deadlines

        💡 CONTENT ENHANCEMENT RULES:
        - Transform features into benefits with business impact
        - Convert data points into strategic insights
        - Upgrade technical details into competitive advantages
        - Elevate implementation details into strategic roadmaps
        - Enhance outcomes into measurable value propositions

        🎯 FINAL VALIDATION:
        - Can a C-suite executive quickly grasp the value proposition?
        - Does each slide advance the narrative toward a decision?
        - Are all claims supported by specific evidence from source content?
        - Would this presentation drive action and investment?

        Generate 6-10 slides that tell a compelling story while honoring every detail of the source material.
        """
        
        return enhanced_prompt
    
    def _clean_input_text(self, text: str) -> str:
        """Clean and format input text to remove encoding issues and improve structure."""
        if not text:
            return text
            
        # Remove common encoding artifacts
        cleaned = text.replace('_x000D_', '\n')
        cleaned = cleaned.replace('_x000A_', '\n')
        cleaned = cleaned.replace('\\n', '\n')
        cleaned = cleaned.replace('\\r', '')
        
        # Clean up multiple newlines
        cleaned = '\n'.join(line.strip() for line in cleaned.split('\n') if line.strip())
        
        # Remove excessive whitespace
        cleaned = ' '.join(cleaned.split())
        
        # Fix common markdown formatting issues
        cleaned = cleaned.replace('#', '')
        cleaned = cleaned.replace('**', '')
        cleaned = cleaned.replace('*', '')
        
        return cleaned
    
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
        # Clean the text first
        cleaned_text = self._clean_input_text(text)
        
        # Enhanced text splitting
        sentences = [s.strip() for s in cleaned_text.split('.') if s.strip() and len(s.strip()) > 15]
        paragraphs = [p.strip() for p in cleaned_text.split('\n') if p.strip() and len(p.strip()) > 20]
        
        # Use paragraphs if available, otherwise sentences
        content_chunks = paragraphs if len(paragraphs) > 2 else sentences
        
        slides = []
        
        # Extract title from content if possible
        title_candidates = [chunk for chunk in content_chunks if len(chunk) < 80 and ':' not in chunk]
        presentation_title = title_candidates[0] if title_candidates else "Professional Presentation"
        
        # Title slide with better content
        slides.append({
            "title": presentation_title,
            "content": [
                "Executive Summary: Key insights and strategic overview",
                "Market Analysis: Current trends and opportunities", 
                "Strategic Approach: Implementation methodology and timeline",
                "Expected Outcomes: Success metrics and business impact",
                "Next Steps: Action items and follow-up requirements"
            ],
            "slide_type": "title",
            "emphasis_points": ["Strategic initiative", "Measurable results"],
            "speaking_notes": "Welcome audience and present comprehensive overview of key topics and expected outcomes."
        })
        
        # Process content with better categorization
        section_headers = ["Market Overview", "Key Features", "Implementation", "Benefits", "Results"]
        content_categories = ["Market Analysis", "Product Features", "Technical Specs", "Business Impact", "Success Metrics"]
        
        processed_slides = 0
        for i, chunk in enumerate(content_chunks[:15]):  # Limit to reasonable number
            if len(chunk) < 10:
                continue
                
            # Create section headers periodically
            if processed_slides > 0 and processed_slides % 3 == 0 and processed_slides < len(section_headers):
                section_idx = (processed_slides // 3) - 1
                slides.append({
                    "title": section_headers[section_idx % len(section_headers)],
                    "content": [
                        "Strategic Focus: Core objectives and priorities",
                        "Key Components: Essential elements and features",
                        "Implementation: Execution strategy and approach"
                    ],
                    "slide_type": "section",
                    "emphasis_points": ["Critical milestone"],
                    "speaking_notes": f"Transition to {section_headers[section_idx % len(section_headers)]} section with detailed analysis."
                })
            
            # Enhanced content processing
            processed_content = []
            
            # Split content intelligently
            if ':' in chunk:
                parts = chunk.split(':')
                for j, part in enumerate(parts):
                    if part.strip() and len(part.strip()) > 5:
                        category = content_categories[j % len(content_categories)]
                        processed_content.append(f"{category}: {part.strip()}")
            else:
                # Break long content into meaningful pieces
                words = chunk.split()
                if len(words) > 20:
                    mid = len(words) // 2
                    part1 = ' '.join(words[:mid])
                    part2 = ' '.join(words[mid:])
                    processed_content.append(f"Overview: {part1}")
                    processed_content.append(f"Details: {part2}")
                else:
                    category = content_categories[i % len(content_categories)]
                    processed_content.append(f"{category}: {chunk}")
            
            # Enhance with additional context points
            while len(processed_content) < 4:
                enhancement_points = [
                    "Strategic Importance: Critical for business success",
                    "Implementation Timeline: Phased approach over 6 months", 
                    "Success Metrics: Measurable KPIs and benchmarks",
                    "Risk Mitigation: Comprehensive contingency planning",
                    "Stakeholder Impact: Benefits across all departments"
                ]
                processed_content.append(enhancement_points[len(processed_content) % len(enhancement_points)])
            
            slides.append({
                "title": f"{content_categories[i % len(content_categories)]} Deep Dive",
                "content": processed_content[:5],  # Limit to 5 points
                "slide_type": "content",
                "emphasis_points": [f"Key insight #{processed_slides + 1}"],
                "speaking_notes": f"Detailed analysis of {content_categories[i % len(content_categories)].lower()} with supporting data and strategic implications."
            })
            
            processed_slides += 1
        
        # Professional conclusion slide
        slides.append({
            "title": "Strategic Recommendations & Action Plan",
            "content": [
                "Executive Summary: Key findings and strategic implications",
                "Priority Actions: Immediate steps for implementation",
                "Timeline: Phased execution over next 6-12 months",
                "Success Metrics: KPIs and measurement framework", 
                "Next Steps: Follow-up meetings and milestone reviews"
            ],
            "slide_type": "conclusion",
            "emphasis_points": ["Immediate action required", "Measurable outcomes"],
            "speaking_notes": "Summarize strategic recommendations with clear action items, ownership, and timeline for implementation."
        })
        
        return {
            "title": presentation_title,
            "slides": slides
        }

    def _extract_metrics(self, text: str) -> List[str]:
        """Extract financial and numerical metrics from text."""
        import re
        metrics = []
        
        # Pattern for financial figures
        financial_patterns = [
            r'\$[\d,]+\.?\d*[BMK]?',  # $4.4B, $2.4M, etc.
            r'[\d,]+%',  # 14.7%, 87%, etc.
            r'[\d,]+\.?\d*\s*billion',
            r'[\d,]+\.?\d*\s*million',
            r'[\d,]+\.?\d*\s*annually'
        ]
        
        for pattern in financial_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            metrics.extend(matches)
        
        return metrics[:5]  # Limit to top 5 metrics
    
    def _extract_features(self, text: str) -> List[str]:
        """Extract key features and capabilities from text."""
        feature_keywords = [
            'AI', 'artificial intelligence', 'machine learning', 'cloud-native',
            'real-time', 'personalized', 'biometric', 'TensorFlow', 'React Native',
            'Node.js', 'PostgreSQL', 'AWS', 'analytics', 'automation'
        ]
        
        features = []
        sentences = text.split('.')
        
        for sentence in sentences:
            for keyword in feature_keywords:
                if keyword.lower() in sentence.lower() and len(sentence.strip()) > 20:
                    features.append(sentence.strip())
                    break
        
        return list(set(features))[:5]  # Remove duplicates and limit
    
    def _extract_business_concepts(self, text: str) -> List[str]:
        """Extract business and strategic concepts from text."""
        business_terms = []
        concepts = [
            'market opportunity', 'revenue model', 'competitive advantage',
            'value proposition', 'growth strategy', 'user experience',
            'business model', 'market share', 'ROI', 'scalability'
        ]
        
        sentences = text.split('.')
        for sentence in sentences:
            for concept in concepts:
                if concept.lower() in sentence.lower():
                    business_terms.append(sentence.strip())
                    break
        
        return list(set(business_terms))[:3]
    
    def _extract_smart_title(self, text: str) -> str:
        """Extract or generate intelligent title from content."""
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        # Look for title patterns
        for line in lines[:5]:  # Check first 5 lines
            if any(indicator in line.lower() for indicator in ['app:', 'product:', 'strategy:', 'solution:']):
                # Clean and format the title
                title = line.replace(':', '').replace('#', '').strip()
                if len(title) < 80:
                    return title
        
        # Extract product/service names
        import re
        product_pattern = r'([A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)*(?:\s+(?:App|Platform|System|Solution|Strategy)))'
        matches = re.findall(product_pattern, text)
        
        if matches:
            return matches[0]
        
        # Fallback to generic business title
        return "Strategic Business Initiative"
    
    def _determine_content_type(self, text: str) -> str:
        """Determine the type of content for better structuring."""
        text_lower = text.lower()
        
        if any(term in text_lower for term in ['product launch', 'app', 'platform', 'solution']):
            return 'product'
        elif any(term in text_lower for term in ['strategy', 'transformation', 'implementation']):
            return 'strategy'
        elif any(term in text_lower for term in ['market', 'analysis', 'research', 'study']):
            return 'analysis'
        elif any(term in text_lower for term in ['education', 'training', 'learning', 'course']):
            return 'education'
        else:
            return 'general'
