from PIL import Image
import pytesseract

# Configure Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Path to the image file
image_path = r'C:\Users\moham\Desktop\image_to_excel\example.jpg'

# Load the image and extract text
try:
    text = pytesseract.image_to_string(Image.open(image_path))
    print("Extracted Text:")
    print(text)
except FileNotFoundError:
    print("Error: The specified image file was not found. Please check the 'image_path' variable.")
except Exception as e:
    print(f"An error occurred: {e}")