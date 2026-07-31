import os
import sys
from pathlib import Path

os.environ['VERCEL'] = '1'

sys.path.insert(0, str(Path(__file__).parent.parent))

from CinemaManagementSystem import app
