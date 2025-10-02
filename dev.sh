#!/bin/bash

# AI Presentation Generator - Development Script
# This script starts both frontend and backend servers

echo "🚀 Starting AI Presentation Generator Development Environment"
echo "=========================================================="

# Function to check if a command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Check for required tools
if ! command_exists python3; then
    echo "❌ Python3 is not installed. Please install Python 3.8+"
    exit 1
fi

if ! command_exists node; then
    echo "❌ Node.js is not installed. Please install Node.js 16+"
    exit 1
fi

if ! command_exists npm; then
    echo "❌ npm is not installed. Please install npm"
    exit 1
fi

# Set up backend
echo "🔧 Setting up Backend..."
cd backend

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "📦 Creating Python virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "🔄 Activating virtual environment..."
source venv/bin/activate

# Install dependencies
echo "📦 Installing Python dependencies..."
pip install -r requirements.txt

# Check for environment variables
if [ ! -f ".env" ]; then
    echo "⚠️  Warning: .env file not found in backend/"
    echo "   Please create .env with:"
    echo "   GOOGLE_API_KEY=your_gemini_api_key"
    echo "   PEXELS_API_KEY=your_pexels_api_key"
    echo ""
    echo "   Continuing anyway..."
fi

# Start backend server in background
echo "🌐 Starting Flask backend server..."
python app.py &
BACKEND_PID=$!
echo "✅ Backend started (PID: $BACKEND_PID)"

# Go back to root directory
cd ..

# Set up frontend
echo "🎨 Setting up Frontend..."
cd frontend

# Install dependencies if node_modules doesn't exist
if [ ! -d "node_modules" ]; then
    echo "📦 Installing Node.js dependencies..."
    npm install
fi

# Create .env.local if it doesn't exist
if [ ! -f ".env.local" ]; then
    echo "🔧 Creating .env.local..."
    echo "NEXT_PUBLIC_BACKEND_URL=http://localhost:5000" > .env.local
fi

# Start frontend development server in background
echo "🌐 Starting Next.js development server..."
npm run dev &
FRONTEND_PID=$!
echo "✅ Frontend started (PID: $FRONTEND_PID)"

# Go back to root directory
cd ..

echo ""
echo "🎉 Development servers started successfully!"
echo "=========================================================="
echo "📱 Frontend: http://localhost:3000"
echo "🔧 Backend:  http://localhost:5000"
echo "📊 API Docs: http://localhost:5000/api/docs (if available)"
echo ""
echo "Press Ctrl+C to stop all servers"
echo ""

# Function to cleanup on exit
cleanup() {
    echo ""
    echo "🛑 Stopping servers..."
    if kill -0 $BACKEND_PID 2>/dev/null; then
        kill $BACKEND_PID
        echo "✅ Backend stopped"
    fi
    if kill -0 $FRONTEND_PID 2>/dev/null; then
        kill $FRONTEND_PID
        echo "✅ Frontend stopped"
    fi
    echo "👋 Development environment stopped"
    exit 0
}

# Set trap to cleanup on script exit
trap cleanup SIGINT SIGTERM

# Wait for background processes
wait