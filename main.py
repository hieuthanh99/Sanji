import json

import pytesseract
from readmrz import MrzDetector, MrzReader


def main():
    pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
    # Process image
    detector = MrzDetector()
    reader = MrzReader()

    image = detector.read("Ho chieu me.jpg")
    cropped = detector.crop_area(image)
    result = reader.process(cropped)
    print(json.dumps(result))


if __name__ == "__main__":
    main()
