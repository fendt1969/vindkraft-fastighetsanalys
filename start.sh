#!/bin/bash
# Vindkraft Fastighetsanalys — start the web app
DIR="$(cd "$(dirname "$0")" && pwd)"

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "Python 3 krävs. Installera från https://www.python.org/downloads/"
    exit 1
fi

# Install dependencies if needed
if ! python3 -c "import streamlit" 2>/dev/null; then
    echo "Installerar beroenden..."
    pip3 install -r "$DIR/requirements.txt"
fi

echo ""
echo "  Vindkraft Fastighetsanalys"
echo "  Öppna webbläsaren på http://localhost:8501"
echo ""

cd "$DIR"
python3 -m streamlit run app.py --server.headless true --server.port 8501
