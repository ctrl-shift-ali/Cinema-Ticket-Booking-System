import os
import sys
from pathlib import Path

# Set Vercel environment
os.environ['VERCEL'] = '1'

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import and export the Flask app
from CinemaManagementSystem import app

# Vercel needs the app exported at module level for WSGI
# The @vercel/python runtime will use this as the WSGI application
