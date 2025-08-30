from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import tempfile
import logging
from src.presentation_generator import PresentationGenerator
from src.llm_service import LLMService
from src.template_analyzer import TemplateAnalyzer

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = 'temp_uploads'

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def home():
    """Serve the home/landing page."""
    return render_template('home.html')

@app.route('/generator')
def generator():
    """Serve the main generator application page."""
    return render_template('index.html')

@app.route('/api/generate', methods=['POST'])
def generate_presentation():
    """Generate a PowerPoint presentation from user input."""
    try:
        # Get form data
        text_input = request.form.get('text_input', '').strip()
        guidance = request.form.get('guidance', '').strip()
        llm_provider = request.form.get('llm_provider', 'openai')
        api_key = request.form.get('api_key', '').strip()
        
        # Get uploaded template file
        template_file = request.files.get('template_file')
        
        # Get uploaded image files
        image_files = request.files.getlist('image_files')
        
        # Validate inputs
        if not text_input:
            return jsonify({'error': 'Text input is required'}), 400
        
        if not api_key:
            return jsonify({'error': 'API key is required'}), 400
            
        if not template_file or template_file.filename == '':
            return jsonify({'error': 'Template file is required'}), 400
        
        # Validate file type
        allowed_extensions = {'.pptx', '.potx'}
        file_ext = os.path.splitext(template_file.filename)[1].lower()
        if file_ext not in allowed_extensions:
            return jsonify({'error': 'Only .pptx and .potx files are allowed'}), 400
        
        # Save uploaded template temporarily
        template_filename = secure_filename(template_file.filename)
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_filename)
        template_file.save(template_path)
        
        # Save uploaded images temporarily
        image_paths = []
        for image_file in image_files:
            if image_file and image_file.filename:
                # Validate image file type
                allowed_image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp'}
                img_ext = os.path.splitext(image_file.filename)[1].lower()
                if img_ext in allowed_image_extensions:
                    image_filename = secure_filename(image_file.filename)
                    image_path = os.path.join(app.config['UPLOAD_FOLDER'], f"img_{image_filename}")
                    image_file.save(image_path)
                    image_paths.append(image_path)
        
        try:
            # Initialize services
            llm_service = LLMService(llm_provider, api_key)
            template_analyzer = TemplateAnalyzer(template_path)
            presentation_generator = PresentationGenerator(llm_service, template_analyzer)
            
            # Generate presentation with images
            output_path = presentation_generator.generate(text_input, guidance, image_paths)
            
            # Return the generated file
            return send_file(
                output_path,
                as_attachment=True,
                download_name=f"generated_presentation_{template_filename}",
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
        
        finally:
            # Clean up uploaded files
            if os.path.exists(template_path):
                os.remove(template_path)
            for image_path in image_paths:
                if os.path.exists(image_path):
                    os.remove(image_path)
    
    except Exception as e:
        logger.error(f"Error generating presentation: {str(e)}")
        return jsonify({'error': f'Failed to generate presentation: {str(e)}'}), 500

@app.route('/api/health')
def health_check():
    """Health check endpoint."""
    return jsonify({'status': 'healthy'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
