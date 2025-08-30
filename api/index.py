from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import sys
import tempfile
import logging

# Add the parent directory to Python path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.presentation_generator import PresentationGenerator
from src.llm_service import LLMService
from src.template_analyzer import TemplateAnalyzer

# Get the directory where this file is located
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)

app = Flask(__name__, template_folder=os.path.join(parent_dir, 'templates'))
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@app.route('/')
def index():
    """Serve the main application page."""
    return render_template('index.html')

@app.route('/api/generate', methods=['POST'])
def generate_presentation():
    """Generate a PowerPoint presentation from user input."""
    try:
        # Get form data
        text_input = request.form.get('text_input', '').strip()
        guidance = request.form.get('guidance', '').strip()
        llm_provider = request.form.get('llm_provider')
        api_key = request.form.get('api_key', '').strip()

        # Validation
        if not text_input:
            return jsonify({'error': 'Text input is required'}), 400
        
        if len(text_input) > 10000:
            return jsonify({'error': 'Text input exceeds 10,000 character limit'}), 400
            
        if not llm_provider:
            return jsonify({'error': 'LLM provider selection is required'}), 400
            
        if not api_key:
            return jsonify({'error': 'API key is required'}), 400

        # Handle file upload
        if 'template_file' not in request.files:
            return jsonify({'error': 'Template file is required'}), 400
            
        template_file = request.files['template_file']
        if template_file.filename == '':
            return jsonify({'error': 'No template file selected'}), 400
            
        if not template_file.filename.lower().endswith(('.pptx', '.potx')):
            return jsonify({'error': 'Template must be a .pptx or .potx file'}), 400

        # Use temporary directory for Vercel
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save uploaded template
            template_filename = secure_filename(template_file.filename)
            template_path = os.path.join(temp_dir, template_filename)
            template_file.save(template_path)

            # Initialize services
            llm_service = LLMService(llm_provider, api_key)
            template_analyzer = TemplateAnalyzer(template_path)
            presentation_generator = PresentationGenerator(template_analyzer, llm_service)

            # Generate presentation structure
            logger.info(f"Generating presentation structure using {llm_provider}")
            presentation_data = llm_service.analyze_text_structure(text_input, guidance)
            
            if not presentation_data or 'slides' not in presentation_data:
                return jsonify({'error': 'Failed to generate presentation structure'}), 500

            # Generate presentation
            logger.info("Creating PowerPoint presentation")
            output_path = os.path.join(temp_dir, 'generated_presentation.pptx')
            presentation_generator.create_presentation(presentation_data, output_path)

            # Return the file
            return send_file(
                output_path,
                as_attachment=True,
                download_name='generated_presentation.pptx',
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )

    except Exception as e:
        logger.error(f"Error generating presentation: {str(e)}")
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500

@app.route('/api/health')
def health_check():
    """Health check endpoint."""
    return jsonify({'status': 'healthy', 'message': 'PowerPoint Generator API is running'})

# Export the Flask app for Vercel
# Vercel will automatically handle the WSGI interface
app = app

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
