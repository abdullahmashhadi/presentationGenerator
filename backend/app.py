from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import google.generativeai as genai
import json
import os
import io
import uuid
import logging
import warnings
import time
from datetime import datetime
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, asdict
import requests
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from dotenv import load_dotenv
import tempfile
import time

# Suppress Google gRPC warnings for ALTS credentials
os.environ['GRPC_VERBOSITY'] = 'ERROR'
os.environ['GLOG_minloglevel'] = '2'

# Filter out specific warnings
warnings.filterwarnings("ignore", category=UserWarning, module="google.auth")

# Configure logging for serverless
logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Suppress grpc and google auth verbose logging
logging.getLogger('google.auth').setLevel(logging.WARNING)
logging.getLogger('google.auth.transport.requests').setLevel(logging.WARNING)
logging.getLogger('urllib3.connectionpool').setLevel(logging.WARNING)

# Load environment variables
load_dotenv()

app = Flask(__name__)

# Configure CORS for production - allow all origins for Vercel
CORS(app, 
     origins="*",
     methods=['GET', 'POST'], 
     allow_headers=['Content-Type', 'Authorization'],
     supports_credentials=False)

# Configuration
GEMINI_API_KEY = os.getenv("GOOGLE_API_KEY")
PEXELS_API_KEY = os.getenv("PEXELS_API_KEY")

if not GEMINI_API_KEY:
    raise ValueError("API_KEY environment variable is required")

# Configure Gemini AI with proper error handling
try:
    genai.configure(api_key=GEMINI_API_KEY)
    logger.info("✅ Gemini AI configured successfully")
except Exception as e:
    logger.error(f"❌ Failed to configure Gemini AI: {str(e)}")
    raise

# Data models
@dataclass
class SlideData:
    header: str
    content: str
    image_url: Optional[str] = None

@dataclass
class PresentationTemplate:
    id: str
    name: str
    description: str
    preview: str
    style: Dict[str, str]

@dataclass
class PresentationRequest:
    title: str
    slides: int
    template: str = "modern"
    style: str = "professional"
    tone: str = "business"
    include_images: bool = True

@dataclass
class ApiResponse:
    success: bool
    data: Optional[Any] = None
    message: Optional[str] = None
    error: Optional[str] = None

# Templates configuration
TEMPLATES = [
    PresentationTemplate(
        id="modern",
        name="Modern Business",
        description="Clean and professional design perfect for business presentations",
        preview="/templates/modern.png",
        style={
            "backgroundColor": "#ffffff",
            "textColor": "#2d3748",
            "accentColor": "#3182ce",
            "fontFamily": "Arial"
        }
    ),
    PresentationTemplate(
        id="creative",
        name="Creative Bold",
        description="Vibrant and eye-catching design for creative presentations",
        preview="/templates/creative.png",
        style={
            "backgroundColor": "#1a202c",
            "textColor": "#ffffff",
            "accentColor": "#ed64a6",
            "fontFamily": "Helvetica"
        }
    ),
    PresentationTemplate(
        id="minimal",
        name="Minimal Clean",
        description="Simple and elegant design with focus on content",
        preview="/templates/minimal.png",
        style={
            "backgroundColor": "#f7fafc",
            "textColor": "#2d3748",
            "accentColor": "#48bb78",
            "fontFamily": "Calibri"
        }
    ),
    PresentationTemplate(
        id="academic",
        name="Academic",
        description="Traditional and formal design for academic presentations",
        preview="/templates/academic.png",
        style={
            "backgroundColor": "#ffffff",
            "textColor": "#1a202c",
            "accentColor": "#2b6cb0",
            "fontFamily": "Times New Roman"
        }
    )
]

class PresentationGenerator:
    def __init__(self):
        try:
            self.model = genai.GenerativeModel('gemini-2.5-flash')
            self.presentations_cache = {}
            logger.info("✅ PresentationGenerator initialized successfully")
        except Exception as e:
            logger.error(f"❌ Failed to initialize PresentationGenerator: {str(e)}")
            raise

    def generate_content(self, request_data: PresentationRequest) -> List[SlideData]:
        """Generate slide content using Gemini AI"""
        
        # Enhanced prompt with style and tone considerations
        tone_descriptions = {
            "formal": "formal, professional language with industry-specific terminology",
            "casual": "conversational, approachable language that's easy to understand",
            "academic": "scholarly, research-focused language with proper citations structure",
            "business": "corporate, results-oriented language focusing on value propositions"
        }
        
        style_descriptions = {
            "professional": "structured, data-driven content with clear action items",
            "creative": "innovative, visually-oriented content with storytelling elements",
            "minimal": "concise, bullet-pointed content focusing on key messages",
            "colorful": "engaging, diverse content with varied formatting and examples"
        }

        prompt = f"""
        Create a {request_data.slides}-slide presentation about "{request_data.title}".
        
        Requirements:
        - Tone: {tone_descriptions.get(request_data.tone, 'professional')}
        - Style: {style_descriptions.get(request_data.style, 'structured and clear')}
        - Each slide should have a compelling header and detailed content
        - Content should include 3-4 bullet points per slide
        - Make it engaging and informative
        - Ensure logical flow between slides
        
        Return as JSON using this exact schema:
        {{
            "slides": [
                {{
                    "header": "Compelling slide title",
                    "content": "• First key point\\n• Second key point\\n• Third key point\\n• Fourth key point (if applicable)"
                }}
            ]
        }}
        """

        try:
            logger.info(f"🤖 Generating content with Gemini AI for: {request_data.title}")
            logger.info(f"📝 Prompt length: {len(prompt)} characters")
            
            response = self.model.generate_content(
                prompt,
                generation_config={
                    "response_mime_type": "application/json",
                    "temperature": 0.7,
                    "max_output_tokens": 8192
                }
            )
            
            logger.info("✅ Received response from Gemini AI")
            logger.info(f"📄 Response length: {len(response.text)} characters")
            
            response_data = json.loads(response.text)
            slides = []
            
            for slide_data in response_data.get("slides", []):
                slides.append(SlideData(
                    header=slide_data.get("header", ""),
                    content=slide_data.get("content", "")
                ))
            
            logger.info(f"✅ Generated {len(slides)} slides successfully")
            return slides
            
        except json.JSONDecodeError as e:
            logger.error(f"❌ JSON decode error: {str(e)}")
            logger.error(f"📄 Raw response: {response.text[:500]}...")
            raise Exception(f"Failed to parse AI response: {str(e)}")
        except Exception as e:
            logger.error(f"❌ Error generating content: {str(e)}")
            logger.error(f"📋 Request details: title='{request_data.title}', slides={request_data.slides}")
            raise Exception(f"Failed to generate content: {str(e)}")

    def fetch_image(self, query: str) -> Optional[str]:
        """Fetch relevant image from Pexels API"""
        if not PEXELS_API_KEY:
            return None
            
        try:
            url = f"https://api.pexels.com/v1/search"
            headers = {"Authorization": PEXELS_API_KEY}
            params = {
                "query": query,
                "per_page": 1,
                "orientation": "landscape"
            }
            
            response = requests.get(url, headers=headers, params=params, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                if data.get('photos'):
                    return data['photos'][0]['src']['large']
                    
        except Exception as e:
            logger.warning(f"Failed to fetch image for '{query}': {str(e)}")
            
        return None

    def create_presentation(self, slides: List[SlideData], title: str, template: str) -> io.BytesIO:
        """Create PowerPoint presentation with enhanced styling"""
        
        prs = Presentation()
        template_style = next((t.style for t in TEMPLATES if t.id == template), TEMPLATES[0].style)
        
        # Add title slide
        title_layout = prs.slide_layouts[0]
        title_slide = prs.slides.add_slide(title_layout)
        title_slide.shapes.title.text = title
        title_slide.shapes.placeholders[1].text = f"Generated on {datetime.now().strftime('%B %d, %Y')}"
        
        # Add content slides
        for slide_data in slides:
            slide_layout = prs.slide_layouts[6]  # Blank layout
            slide = prs.slides.add_slide(slide_layout)
            
            # Add background image if available
            if slide_data.image_url:
                try:
                    img_response = requests.get(slide_data.image_url, timeout=10)
                    if img_response.status_code == 200:
                        img_stream = io.BytesIO(img_response.content)
                        slide.shapes.add_picture(
                            img_stream, 0, 0, 
                            width=prs.slide_width, 
                            height=prs.slide_height
                        )
                except Exception as e:
                    logger.warning(f"Failed to add image: {str(e)}")
            
            # Add semi-transparent overlay
            overlay = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.5), Inches(0.5),
                prs.slide_width - Inches(1),
                prs.slide_height - Inches(1)
            )
            
            # Style the overlay
            overlay.fill.solid()
            overlay.fill.fore_color.rgb = RGBColor(255, 255, 255)
            overlay.line.fill.background()
            
            # Add transparency
            try:
                sp = overlay._element
                solidFill = sp.find(".//a:solidFill", sp.nsmap)
                srgbClr = solidFill.find(".//a:srgbClr", sp.nsmap)
                alpha = parse_xml('<a:alpha xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="40000"/>')
                srgbClr.append(alpha)
            except:
                pass
            
            # Add text content
            text_box = slide.shapes.add_textbox(
                Inches(0.7), Inches(0.7),
                prs.slide_width - Inches(1.4),
                prs.slide_height - Inches(1.4)
            )
            
            text_frame = text_box.text_frame
            text_frame.word_wrap = True
            text_frame.margin_left = Inches(0.2)
            text_frame.margin_top = Inches(0.2)
            
            # Add header
            if slide_data.header:
                p = text_frame.paragraphs[0]
                p.text = slide_data.header
                p.font.size = Pt(28)
                p.font.bold = True
                p.font.color.rgb = RGBColor(20, 20, 20)
                p.space_after = Pt(20)
            
            # Add content
            if slide_data.content:
                content_lines = slide_data.content.split('\n')
                for line in content_lines:
                    if line.strip():
                        p = text_frame.add_paragraph()
                        p.text = line.strip()
                        p.font.size = Pt(18)
                        p.font.color.rgb = RGBColor(40, 40, 40)
                        p.level = 1 if line.startswith('•') else 0
                        p.space_after = Pt(8)
        
        # Save to BytesIO
        pptx_stream = io.BytesIO()
        prs.save(pptx_stream)
        pptx_stream.seek(0)
        
        return pptx_stream

# Initialize generator
generator = PresentationGenerator()

# Routes
@app.route('/')
def index():
    """Root endpoint for debugging"""
    return jsonify({
        "message": "AI Presentation Generator API",
        "version": "2.0.0",
        "endpoints": {
            "health": "/api/health or /health",
            "templates": "/api/templates or /templates",
            "generate": "/api/generate or /generate",
            "preview": "/api/preview or /preview",
            "download": "/api/download/<id> or /download/<id>"
        },
        "gemini_configured": bool(GEMINI_API_KEY),
        "pexels_configured": bool(PEXELS_API_KEY)
    })

@app.route('/api/health', methods=['GET'])
@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify(ApiResponse(
        success=True,
        data={
            "status": "healthy",
            "version": "2.0.0",
            "timestamp": datetime.now().isoformat()
        }
    ).__dict__)

@app.route('/api/templates', methods=['GET'])
@app.route('/templates', methods=['GET'])
def get_templates():
    """Get available presentation templates"""
    try:
        logger.info("📋 Templates endpoint called")
        response_data = ApiResponse(
            success=True,
            data=[asdict(template) for template in TEMPLATES]
        ).__dict__
        logger.info(f"✅ Returning {len(TEMPLATES)} templates")
        return jsonify(response_data)
    except Exception as e:
        logger.error(f"❌ Error in get_templates: {str(e)}")
        return jsonify(ApiResponse(
            success=False,
            error=f"Failed to load templates: {str(e)}"
        ).__dict__), 500

@app.route('/api/preview', methods=['POST'])
@app.route('/preview', methods=['POST'])
def preview_slides():
    """Generate slide content preview without creating presentation"""
    try:
        data = request.get_json()
        
        if not data or not data.get('title'):
            return jsonify(ApiResponse(
                success=False,
                error="Title is required"
            ).__dict__), 400
        
        req = PresentationRequest(
            title=data['title'],
            slides=min(data.get('slides', 5), 10),  # Limit preview to 10 slides
            include_images=False  # No images for preview
        )
        
        slides = generator.generate_content(req)
        
        return jsonify(ApiResponse(
            success=True,
            data={
                "slides": [asdict(slide) for slide in slides]
            }
        ).__dict__)
        
    except Exception as e:
        logger.error(f"Preview error: {str(e)}")
        return jsonify(ApiResponse(
            success=False,
            error=str(e)
        ).__dict__), 500

@app.route('/api/generate', methods=['POST'])
@app.route('/generate', methods=['POST'])
def generate_presentation():
    """Generate complete presentation"""
    try:
        data = request.get_json()
        logger.info(f"🚀 New presentation request received")
        logger.info(f"📋 Request data: {data}")
        
        # Validate input
        if not data or not data.get('title'):
            logger.warning("❌ Request missing title")
            return jsonify(ApiResponse(
                success=False,
                error="Title is required"
            ).__dict__), 400
        
        if not (1 <= data.get('slides', 5) <= 20):
            logger.warning(f"❌ Invalid slide count: {data.get('slides', 5)}")
            return jsonify(ApiResponse(
                success=False,
                error="Number of slides must be between 1 and 20"
            ).__dict__), 400
        
        # Create request object
        req = PresentationRequest(
            title=data['title'],
            slides=data.get('slides', 5),
            template=data.get('template', 'modern'),
            style=data.get('style', 'professional'),
            tone=data.get('tone', 'business'),
            include_images=data.get('includeImages', True)
        )
        
        logger.info(f"📝 Processing: '{req.title}' with {req.slides} slides")
        
        # Generate content
        logger.info("🤖 Starting content generation...")
        slides = generator.generate_content(req)
        logger.info(f"✅ Content generation completed: {len(slides)} slides")
        
        # Fetch images if requested
        if req.include_images and PEXELS_API_KEY:
            logger.info("🖼️  Fetching images from Pexels...")
            for i, slide in enumerate(slides):
                if slide.header:
                    logger.info(f"🔍 Fetching image for slide {i+1}: '{slide.header}'")
                    slide.image_url = generator.fetch_image(slide.header)
            logger.info("✅ Image fetching completed")
        else:
            logger.info("⏭️  Skipping image fetching")
        
        # Create presentation
        logger.info("📋 Creating PowerPoint presentation...")
        pptx_stream = generator.create_presentation(slides, req.title, req.template)
        logger.info("✅ PowerPoint creation completed")
        
        # Generate unique filename
        presentation_id = str(uuid.uuid4())
        filename = f"presentation_{presentation_id}.pptx"
        
        # Cache the presentation temporarily
        generator.presentations_cache[presentation_id] = {
            'data': pptx_stream.getvalue(),
            'filename': filename,
            'created_at': time.time()
        }
        
        logger.info(f"💾 Presentation cached with ID: {presentation_id}")
        
        return jsonify(ApiResponse(
            success=True,
            data={
                "presentation_id": presentation_id,
                "slides": [asdict(slide) for slide in slides],
                "title": req.title,
                "download_url": f"/download/{presentation_id}"
            }
        ).__dict__)
        
    except Exception as e:
        logger.error(f"❌ Generation error: {str(e)}")
        import traceback
        logger.error(f"📋 Full traceback:\n{traceback.format_exc()}")
        return jsonify(ApiResponse(
            success=False,
            error=str(e)
        ).__dict__), 500

@app.route('/api/download/<presentation_id>', methods=['GET'])
@app.route('/download/<presentation_id>', methods=['GET'])
def download_presentation(presentation_id: str):
    """Download generated presentation"""
    try:
        # Check cache
        if presentation_id not in generator.presentations_cache:
            return jsonify(ApiResponse(
                success=False,
                error="Presentation not found or expired"
            ).__dict__), 404
        
        cached_data = generator.presentations_cache[presentation_id]
        
        # Check if expired (24 hours)
        if time.time() - cached_data['created_at'] > 86400:
            del generator.presentations_cache[presentation_id]
            return jsonify(ApiResponse(
                success=False,
                error="Presentation has expired"
            ).__dict__), 404
        
        # Create temporary file for download
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
            tmp_file.write(cached_data['data'])
            tmp_file.flush()
            
            return send_file(
                tmp_file.name,
                as_attachment=True,
                download_name=cached_data['filename'],
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
            
    except Exception as e:
        logger.error(f"Download error: {str(e)}")
        return jsonify(ApiResponse(
            success=False,
            error=str(e)
        ).__dict__), 500

@app.errorhandler(404)
def not_found(error):
    return jsonify(ApiResponse(
        success=False,
        error="Endpoint not found"
    ).__dict__), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify(ApiResponse(
        success=False,
        error="Internal server error"
    ).__dict__), 500

# Cleanup old presentations periodically
def cleanup_old_presentations():
    """Remove presentations older than 24 hours"""
    current_time = time.time()
    expired_ids = [
        pid for pid, data in generator.presentations_cache.items()
        if current_time - data['created_at'] > 86400
    ]
    
    for pid in expired_ids:
        del generator.presentations_cache[pid]
    
    logger.info(f"Cleaned up {len(expired_ids)} expired presentations")

if __name__ == '__main__':
    # Additional environment setup for production
    os.environ.setdefault('GRPC_VERBOSITY', 'ERROR')
    os.environ.setdefault('GLOG_minloglevel', '2')
    
    logger.info("🚀 Starting AI Presentation Generator Backend")
    logger.info(f"🔧 Debug mode: {app.debug}")
    logger.info(f"🌐 Server will be available at: http://localhost:5000")
    logger.info(f"🔑 Gemini API configured: {'✅' if GEMINI_API_KEY else '❌'}")
    logger.info(f"🖼️  Pexels API configured: {'✅' if PEXELS_API_KEY else '❌ (images disabled)'}")
    
    try:
        app.run(debug=True, host='0.0.0.0', port=5000)
    except KeyboardInterrupt:
        logger.info("🛑 Server stopped by user")
    except Exception as e:
        logger.error(f"❌ Server error: {str(e)}")
        raise
