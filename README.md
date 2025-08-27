# PowerPoint Generator

**Your Text, Your Style – Auto-Generate a Presentation**

A powerful web application that transforms any bulk text, markdown, or prose into professionally formatted PowerPoint presentations while automatically preserving and applying your chosen template's visual identity, layout, and styling.

## 🌟 Overview

The PowerPoint Generator revolutionizes presentation creation by combining advanced AI text analysis with sophisticated template processing. Instead of spending hours manually creating slides, simply paste your content, upload your preferred PowerPoint template, and let AI do the heavy lifting. The application intelligently structures your content into coherent slides while maintaining your brand's visual identity.

### Why Use PowerPoint Generator?

- **Save Time**: Convert hours of manual slide creation into minutes of automated generation
- **Maintain Consistency**: Automatically apply your brand template across all slides
- **Professional Quality**: AI ensures logical flow, proper structure, and engaging content
- **Flexible Input**: Works with any text format - from research papers to meeting notes
- **Brand Preservation**: Your template's colors, fonts, and layouts are perfectly preserved

## 🚀 Features

### Core Functionality
- **Text-to-Slides Conversion**: Paste any text content and get structured slides with intelligent content organization
- **Template Style Extraction**: Upload your PowerPoint template and the app applies its complete styling including colors, fonts, layouts, and backgrounds
- **Multiple LLM Support**: Choose from OpenAI GPT-3.5, Anthropic Claude, Google Gemini, or AIPipe.org for optimal performance
- **Smart Content Analysis**: AI breaks down text into logical slide structure with proper hierarchy and flow
- **Image Reuse**: Intelligently incorporates and repositions images from your template where contextually appropriate
- **Image & Icon Auto-Suggestions**: AI identifies opportunities for visual enhancements and suggests relevant imagery

### Advanced Features
- **Adaptive Layout Selection**: Automatically chooses the best template layout for each slide type (title, content, comparison, conclusion)
- **Content Optimization**: Adjusts text length and bullet points to fit slide boundaries perfectly
- **Visual Hierarchy**: Applies proper typography, spacing, and emphasis for maximum readability
- **Multi-format Support**: Handles various input formats including markdown, plain text, and structured documents
- **Batch Processing**: Process multiple presentations with consistent styling
- **Cross-platform Compatibility**: Works on Windows, Mac, and Linux through web interface

## 🎯 How It Works

### Step-by-Step Process

1. **Content Input**: 
   - Paste your text content into the web interface (supports up to 10,000 characters)
   - Content can be any format: research papers, meeting notes, product descriptions, technical documentation
   - The system handles unstructured text and automatically identifies key topics and themes

2. **Template Upload**: 
   - Upload your PowerPoint template file (.pptx or .potx format, up to 50MB)
   - The system extracts color schemes, font families, layout configurations, and background elements
   - Preserves your brand identity and visual consistency across all generated slides

3. **AI Provider Selection**: 
   - Choose from multiple LLM providers based on your needs and preferences
   - OpenAI GPT-3.5: Excellent general-purpose text analysis and structuring
   - Anthropic Claude: Superior at maintaining context and logical flow
   - Google Gemini: Strong performance with technical and scientific content
   - AIPipe.org: Cost-effective alternative with good performance

4. **Content Processing**: 
   - AI analyzes your text and identifies main topics, subtopics, and supporting details
   - Creates logical slide hierarchy with appropriate titles and content distribution
   - Generates speaker notes with additional context and presentation tips
   - Applies your template's styling while optimizing content layout

5. **Presentation Generation**: 
   - Downloads a fully formatted PowerPoint file maintaining your template's visual identity
   - Each slide is optimized for readability with proper text sizing and spacing
   - Includes title slide, content slides, and conclusion slide as appropriate
   - Ready for immediate use or further customization

### Behind the Scenes Technology

The application employs sophisticated algorithms to ensure professional presentation quality:

- **Natural Language Processing**: Advanced text analysis identifies document structure and key themes
- **Content Segmentation**: Intelligent chunking breaks large text into digestible slide-sized portions
- **Layout Optimization**: Dynamic content fitting ensures text never overflows slide boundaries
- **Visual Hierarchy**: Automatic application of typography best practices for maximum impact

## 🛠️ Technical Implementation

### Architecture Overview

The PowerPoint Generator is built on a modern, scalable architecture that combines web technologies with advanced AI processing:

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   Web Frontend  │────│  Flask Backend   │────│  LLM Services   │
│   (HTML/CSS/JS) │    │  (Python/Flask)  │    │  (OpenAI/etc.)  │
└─────────────────┘    └──────────────────┘    └─────────────────┘
                                │
                       ┌──────────────────┐
                       │ Template Analyzer │
                       │ (python-pptx)     │
                       └──────────────────┘
```

## 🚀 Setup and Installation

### Prerequisites
- **Python 3.8 or higher** (Python 3.9+ recommended for optimal performance)
- **pip package manager** (usually included with Python installation)
- **4GB RAM minimum** (8GB recommended for processing large templates)
- **Internet connection** for LLM API access

### Quick Start Installation

#### Standard Installation
```bash
# Clone the repository
git clone https://github.com/yourusername/powerpoint-generator.git
cd powerpoint-generator

# Create virtual environment (recommended)
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

### Environment Configuration

#### Environment Variables (Optional but Recommended)
Create a `.env` file in the project root for convenient API key management:

```bash
# OpenAI Configuration
OPENAI_API_KEY=sk-your-openai-api-key-here
OPENAI_MODEL=gpt-3.5-turbo

# Anthropic Configuration  
ANTHROPIC_API_KEY=your-anthropic-api-key-here
ANTHROPIC_MODEL=claude-3-sonnet-20240229

# Google Gemini Configuration
GEMINI_API_KEY=your-gemini-api-key-here
GEMINI_MODEL=gemini-pro

# Application Settings
FLASK_ENV=development
MAX_CONTENT_LENGTH=52428800  # 50MB
UPLOAD_FOLDER=temp_uploads
```

#### API Key Setup Instructions

**OpenAI Setup:**
1. Visit [OpenAI Platform](https://platform.openai.com/)
2. Create an account or sign in
3. Navigate to API Keys section
4. Generate a new secret key
5. Add billing information (pay-per-use model)

**Anthropic Setup:**
1. Go to [Anthropic Console](https://console.anthropic.com/)
2. Sign up for an account
3. Navigate to API Keys
4. Create a new API key
5. Set up billing (credit-based system)

**Google Gemini Setup:**
1. Visit [Google AI Studio](https://makersuite.google.com/)
2. Sign in with Google account
3. Create a new project
4. Generate API key
5. Enable Gemini API access

**AIPipe.org Setup:**
1. Register at [AIPipe.org](https://aipipe.org/login)
2. Choose a subscription plan
3. Access API credentials in dashboard
4. Cost-effective alternative to direct providers


## 📝 API Endpoints

### `POST /api/generate`
Generates a PowerPoint presentation from the provided inputs.

**Form Data:**
- `text_input` (required): The text content to convert
- `guidance` (optional): Instructions for tone/structure
- `llm_provider` (required): One of 'openai', 'anthropic', 'gemini'
- `api_key` (required): API key for the chosen LLM provider
- `template_file` (required): PowerPoint template file (.pptx or .potx)

**Response:**
- Success: Returns the generated PowerPoint file for download
- Error: JSON with error message

### `GET /api/health`
Health check endpoint.

## 🎨 Supported LLM Providers

### OpenAI
- **Model**: GPT-3.5-turbo
- **API Key**: Get from [OpenAI Platform](https://platform.openai.com/)

### Anthropic
- **Model**: Claude-3-sonnet
- **API Key**: Get from [Anthropic Console](https://console.anthropic.com/)

### Google Gemini
- **Model**: Gemini-pro
- **API Key**: Get from [Google AI Studio](https://makersuite.google.com/)

### AIPipe.org
- **Model**: GPT-3.5-turbo (via AIPipe)
- **API Key**: Get from [AIPipe.org](https://aipipe.org/login)
- **Features**: Cost-effective API access to multiple models

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with [python-pptx](https://python-pptx.readthedocs.io/) for PowerPoint manipulation
- Styled with modern CSS and responsive design
- Powered by state-of-the-art language models

## 📞 Support

If you encounter any issues or have questions, please [open an issue](https://github.com/yourusername/powerpoint-generator/issues) on GitHub.

---

**Made with ❤️ for better presentations**
