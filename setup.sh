#!/bin/bash

# Automated setup script for macOS/Linux
# This script creates a virtual environment, installs dependencies, and sets up the project

echo "🚀 Setting up PPA Excel Processor..."

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed. Please install Python 3 first."
    exit 1
fi

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "❌ pip3 is not installed. Please install pip3 first."
    exit 1
fi

echo "✅ Python 3 and pip3 are available"

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "📦 Creating virtual environment..."
    python3 -m venv venv
    echo "✅ Virtual environment created"
else
    echo "✅ Virtual environment already exists"
fi

# Activate virtual environment
echo "🔧 Activating virtual environment..."
source venv/bin/activate

# Upgrade pip
echo "⬆️ Upgrading pip..."
pip install --upgrade pip

# Install requirements
echo "📚 Installing dependencies..."
pip install -r requirements.txt

echo ""
echo "🎉 Setup complete! Your environment is ready."
echo ""
echo "To run the app:"
echo "1. Activate the environment: source venv/bin/activate"
echo "2. Run the app: streamlit run app.py"
echo ""
echo "Or simply run: ./run_app.sh"
