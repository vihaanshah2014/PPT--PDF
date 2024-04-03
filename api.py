import sys
import os
import comtypes.client
from flask import Flask, request, send_file, jsonify

app = Flask(__name__)

@app.route('/')
def home():
    return jsonify({"message": "Welcome to the PowerPoint to PDF Conversion API"})

@app.route('/convert', methods=['POST'])
def convert_ppt_to_pdf():
    # Get the uploaded file from the request
    uploaded_file = request.files['file']
    input_file_path = os.path.join(os.getcwd(), uploaded_file.filename)
    uploaded_file.save(input_file_path)

    # Generate the output file path
    output_file_path = os.path.splitext(input_file_path)[0] + '.pdf'

    # Create powerpoint application object
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

    # Set visibility to minimize
    powerpoint.Visible = 1

    # Open the powerpoint slides
    slides = powerpoint.Presentations.Open(input_file_path)

    # Save as PDF (formatType = 32)
    slides.SaveAs(output_file_path, 32)

    # Close the slide deck
    slides.Close()

    # Return the converted PDF file
    return send_file(output_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run()