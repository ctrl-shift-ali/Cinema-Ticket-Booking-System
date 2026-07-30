import os

# Set Vercel environment flag
os.environ['VERCEL'] = '1'

# Import the Flask app
from CinemaManagementSystem import app

# WSGI application for Vercel
application = app

# Also export as 'app' for compatibility
__all__ = ['app', 'application']
