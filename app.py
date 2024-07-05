from flask import Flask, render_template, request, send_file
import google.generativeai as genai
import json
from pptx import Presentation
from pptx.util import Pt
from dotenv import load_dotenv
import os

app = Flask(__name__)

# Load the environment variables from the .env file
load_dotenv()

# Get the API key from the environment variables
api_key = os.getenv("API_KEY")
genai.configure(api_key=api_key)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    presentation_title = request.form['title']
    number_of_slides = request.form['slides']

    # Define the prompt for generating the presentation slides.
    prompt = f"""
    generate a {number_of_slides} slide presentation for the topic {presentation_title}. Each slide should have a (header), (content). Return as JSON. Must include 3 bullet points.
    Using this JSON schema:
        {{
            "slides": [
                {{
                    "header": "string",
                    "content": "string"
                }}
            ]
        }}
    """

    # Choose a model that's appropriate for your use case.
    model = genai.GenerativeModel('gemini-1.5-flash',
                                  generation_config={"response_mime_type": "application/json"})

    # Send the request to the Gemini API.
    response = model.generate_content(prompt)

    # Parse the JSON response.
    response_data = json.loads(response.text)
    slide_data = response_data["slides"]

    # Create a PowerPoint presentation.
    prs = Presentation()

    for slide in slide_data:
        slide_layout = prs.slide_layouts[1]
        new_slide = prs.slides.add_slide(slide_layout)
        if slide['header']:
            title = new_slide.shapes.title
            title.text = slide['header']
        if slide['content']:
            shapes = new_slide.shapes
            body_shape = shapes.placeholders[1]
            tf = body_shape.text_frame
            tf.text = slide['content']
            
            # Manually set the font size and style
            for paragraph in tf.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(18)  # Set font size as needed
                    run.font.bold = True
                    run.font.name = 'Calibri'

    # Save the PowerPoint presentation.
    pptx_path = "output.pptx"
    prs.save(pptx_path)

    return send_file(pptx_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
