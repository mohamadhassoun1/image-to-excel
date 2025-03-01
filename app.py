from flask import Flask, render_template, request, send_file
import os
from PIL import Image
import pandas as pd
import pytesseract

# Configure Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    result = None  # Variable to store the download link
    if request.method == "POST":
        # Check if a file was uploaded
        if "file" not in request.files:
            return "No file uploaded!", 400
        
        file = request.files["file"]
        if file.filename == "":
            return "No file selected!", 400

        # Save the uploaded file temporarily
        file_path = os.path.join("uploads", file.filename)
        os.makedirs("uploads", exist_ok=True)
        file.save(file_path)

        # Extract text from the image using OCR
        try:
            text = pytesseract.image_to_string(Image.open(file_path))
            lines = [line.strip() for line in text.split("\n") if line.strip()]  # Split text into lines and remove empty ones
        except Exception as e:
            return f"Error extracting text: {str(e)}", 500

        # Create a DataFrame with each line as a separate row
        df = pd.DataFrame(lines, columns=["Extracted Text"])

        # Save the DataFrame as an Excel file
        excel_path = os.path.join("output", f"{os.path.splitext(file.filename)[0]}.xlsx")
        os.makedirs("output", exist_ok=True)
        df.to_excel(excel_path, index=False)

        # Pass the download link to the template
        result = f"/download/{os.path.basename(excel_path)}"

    return render_template("index.html", result=result)

@app.route("/download/<filename>")
def download_file(filename):
    # Serve the generated Excel file for download
    return send_file(
        os.path.join("output", filename),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True)