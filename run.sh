#!/bin/bash

# PO Line Comparison Tool - Start Script

echo "🚀 Starting PO Line Comparison Tool..."
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed. Please install Python 3.8 or higher."
    exit 1
fi

# Check if requirements are installed
if ! python3 -c "import streamlit" &> /dev/null; then
    echo "📦 Installing dependencies..."
    pip3 install -r requirements.txt
    echo ""
fi

# Run Streamlit app
echo "✅ Starting application..."
echo "📊 The app will open in your default browser"
echo ""
echo "Press Ctrl+C to stop the server"
echo ""

streamlit run app.py

