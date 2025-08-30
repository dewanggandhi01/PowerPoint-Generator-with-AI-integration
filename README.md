# PowerPoint Generator

**Transform any text into professional PowerPoint presentations using AI**

A web application that converts your text content into well-structured PowerPoint slides while preserving your template's visual identity.

## 🚀 Features

- **AI-Powered Content Structuring**: Automatically organizes text into logical slides
- **Template Style Preservation**: Maintains your brand colors, fonts, and layouts
- **Multiple LLM Support**: OpenAI, Anthropic, Google Gemini, and AIPipe.org
- **Professional Output**: Clean, presentation-ready slides with proper formatting
- **Easy Upload**: Simply paste text and upload your PowerPoint template

## 🛠️ Quick Start

### Prerequisites
- Python 3.8+
- Internet connection for AI processing

### Installation
```bash
# Clone repository
git clone https://github.com/dewanggandhi01/PowerPoint-Generator-with-AI-integration.git
cd PowerPoint-Generator-with-AI-integration

# Install dependencies
pip install -r requirements.txt

# Run application
python app.py
```

### Usage
1. **Start the app**: Visit `http://localhost:5000`
2. **Add content**: Paste your text (up to 10,000 characters)
3. **Choose AI provider**: Select OpenAI, Anthropic, Gemini, or AIPipe
4. **Enter API key**: Get from your chosen provider's website
5. **Upload template**: Your PowerPoint template (.pptx file)
6. **Generate**: Download your professional presentation

## 🔑 API Keys

Get your API key from:
- **OpenAI**: [platform.openai.com](https://platform.openai.com/)
- **Anthropic**: [console.anthropic.com](https://console.anthropic.com/)
- **Google Gemini**: [makersuite.google.com](https://makersuite.google.com/)
- **AIPipe**: [aipipe.org](https://aipipe.org/) (cost-effective alternative)

## 📁 Project Structure

```
├── app.py                 # Main Flask application
├── src/
│   ├── llm_service.py     # AI provider integrations
│   ├── template_analyzer.py # Template processing
│   └── presentation_generator.py # PowerPoint creation
├── templates/
│   └── index.html         # Web interface
└── requirements.txt       # Dependencies
```

## 🎯 How It Works

1. **Text Analysis**: AI breaks down your content into main topics and subtopics
2. **Slide Structure**: Creates logical presentation flow with titles and bullet points
3. **Template Application**: Extracts and applies your template's styling
4. **Content Optimization**: Ensures text fits properly within slide boundaries
5. **Professional Output**: Generates a polished PowerPoint presentation

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

---

**Made with ❤️ for better presentations**
