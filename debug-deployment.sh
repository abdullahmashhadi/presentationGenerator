#!/bin/bash

echo "🔍 Vercel Deployment Debugging Script"
echo "======================================"
echo

echo "📋 Current Repository Info:"
echo "Repository: $(git remote get-url origin)"
echo "Current Branch: $(git branch --show-current)"
echo "Latest Commit: $(git log -1 --oneline)"
echo

echo "📊 Recent Commits:"
git log --oneline -5
echo

echo "🔧 Vercel Configuration Check:"
if [ -f "vercel.json" ]; then
    echo "✅ vercel.json exists"
    echo "Content preview:"
    head -20 vercel.json
else
    echo "❌ vercel.json not found"
fi
echo

echo "📁 Project Structure:"
echo "Root files:"
ls -la | grep -E '\.(json|md|py)$'
echo
echo "Frontend structure:"
if [ -d "frontend" ]; then
    echo "✅ frontend/ directory exists"
    ls frontend/ | head -10
else
    echo "❌ frontend/ directory not found"
fi
echo
echo "Backend structure:"
if [ -d "backend" ]; then
    echo "✅ backend/ directory exists"
    ls backend/ | head -10
else
    echo "❌ backend/ directory not found"
fi

echo
echo "🎯 Next Steps:"
echo "1. Check your Vercel dashboard at https://vercel.com"
echo "2. Look for project: presentationGenerator"
echo "3. Go to Deployments tab"
echo "4. Manually trigger deployment if needed"
echo "5. Check environment variables are set"
