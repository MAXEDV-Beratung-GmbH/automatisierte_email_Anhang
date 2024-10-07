import os
from dotenv import load_dotenv
import requests
import logging

# Load environment variables from .env file
load_dotenv()

class OCRConnection:
    def __init__(self):
        self.api_key = os.getenv('OCR_SPACE_API_KEY')

    def ocr_space_file(self, file_path):
        with open(file_path, 'rb') as f:
            r = requests.post(
                'https://api.ocr.space/parse/image',
                files={file_path: f},
                data={'apikey': self.api_key}
            )
        result = r.json()
        if result['IsErroredOnProcessing']:
            logging.error("Error in processing: " + str(result['ErrorMessage']))
            return ''
        return result.get('ParsedResults')[0].get('ParsedText', '')
