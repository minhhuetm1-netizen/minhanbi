import os
import sys

# Ensure root directory is on sys.path
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from server import MinhAnRequestHandler

class handler(MinhAnRequestHandler):
    """
    Vercel Serverless Function entry point:
    Inherits all API endpoints (upload, step2, step3, step4, export, progress, paginate)
    and memory-buffered file downloads (.docx, .xlsx).
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=root_dir, **kwargs)
