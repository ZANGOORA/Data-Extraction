import os
import sys
import subprocess

# Install required packages
required_packages = ['openai', 'PyPDF2', 'pandas', 'openpyxl']
for package in required_packages:
    try:
        __import__(package)
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Now import the required modules
import openai
from PyPDF2 import PdfReader
import pandas as pd

# ... existing code ...

def process_text_with_openai(text):
    """Process text using OpenAI API to extract structured data."""
    # Check if API key is set
    if not openai.api_key:
        raise ValueError("OpenAI API key is not set. Please set the OPENAI_API_KEY environment variable.")
    
    # ... rest of the function ...

# ... rest of the existing code ...