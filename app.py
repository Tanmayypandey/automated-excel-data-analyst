"""Application Runner for Automated Excel Data Analyst.

Run directly via:
    python app.py
or with uvicorn:
    uvicorn api:app --reload
"""

import sys
import uvicorn
from api import app

if __name__ == "__main__":
    print("=" * 60)
    print("Starting Automated Excel Data Analyst server...")
    print("Local UI: http://127.0.0.1:8000/")
    print("Health:   http://127.0.0.1:8000/health")
    print("=" * 60)
    uvicorn.run("api:app", host="127.0.0.1", port=8000, reload=True)
