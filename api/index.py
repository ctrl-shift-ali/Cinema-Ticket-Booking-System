import sys
import os
from pathlib import Path

# Add the parent directory to the path so we can import CinemaManagementSystem
parent_dir = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(parent_dir))

# Import the Flask app
from CinemaManagementSystem import app

# Export the app for Vercel
