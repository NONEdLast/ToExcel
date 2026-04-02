import sys
import os

# Test if Python is working
print(f"Python version: {sys.version}")
print(f"Current directory: {os.getcwd()}")

# Test if we can import installed modules
try:
    import pandas as pd
    import openpyxl
    import gradio
    print("All required modules are installed and importable")
    print(f"  - pandas version: {pd.__version__}")
    print(f"  - openpyxl version: {openpyxl.__version__}")
    print(f"  - gradio version: {gradio.__version__}")
except ImportError as e:
    print(f"Error importing modules: {e}")

# Test if we can import project modules
try:
    # Add current directory to Python path
    sys.path.insert(0, os.path.dirname(__file__))
    import txt_to_excel
    import json_to_excel
    import excel_to_other
    import sqlite_to_excel
    print("All project modules are importable")
except ImportError as e:
    print(f"Error importing project modules: {e}")
    import traceback
    traceback.print_exc()