# Vercel Deployment Instructions

## ✅ Pre-Deployment Checklist

Your project is now properly configured for Vercel deployment. Here's what was fixed:

### 🔧 Configuration Changes Made:

1. **Fixed vercel.json**:
   - Corrected distDir to use `frontend/out` instead of `frontend/.next`
   - Fixed environment variable naming (`GOOGLE_API_KEY` instead of `API_KEY`)
   - Removed env variables from config (they'll be set in Vercel dashboard)

2. **Updated Next.js config**:
   - Added `distDir: 'out'` for static export compatibility
   - Configured for static deployment

3. **Fixed Flask app**:
   - Updated environment variable names to match Vercel config
   - Added Vercel handler function
   - Added missing `time` import

## 🚀 Deployment Steps

### 1. Push Changes to GitHub
```bash
git add .
git commit -m "Fix Vercel deployment configuration"
git push origin main
```

### 2. Set Environment Variables in Vercel
Go to your Vercel project dashboard and add these environment variables:

- `GOOGLE_API_KEY`: Your Google Gemini API key
- `PEXELS_API_KEY`: Your Pexels API key (optional)

### 3. Manual Redeploy
If auto-deployment doesn't work:
1. Go to your Vercel dashboard
2. Click on your project
3. Go to Deployments tab
4. Click "Redeploy" on the latest commit

## 🐛 Common Issues & Solutions

### Issue: "Build failed"
- **Solution**: Check that environment variables are set in Vercel dashboard
- **Check**: Verify `requirements.txt` exists in backend folder

### Issue: "Function timeout"
- **Solution**: The timeout is set to 60s, which should be sufficient for AI generation

### Issue: "CORS errors"
- **Solution**: The Flask app is configured for Vercel domains, should work automatically

### Issue: "404 on API routes"
- **Solution**: Ensure routes start with `/api/` (e.g., `/api/generate`)

## 📋 Project Structure Verification

Your project should have this structure:
```
presentationMakerGemini/
├── vercel.json                 ✅ Fixed
├── backend/
│   ├── app.py                 ✅ Fixed
│   └── requirements.txt       ✅ Should exist
└── frontend/
    ├── package.json           ✅ Fixed
    ├── next.config.js         ✅ Fixed
    └── src/...               ✅ Should exist
```

## 🔍 Debugging Deployment

If deployment still fails:

1. **Check Vercel Function Logs**:
   - Go to Vercel dashboard → Your project → Functions tab
   - Check logs for any errors

2. **Verify Build Logs**:
   - Go to Deployments tab
   - Click on failed deployment
   - Check build logs for specific errors

3. **Test Locally**:
   ```bash
   # Test frontend build
   cd frontend && npm run build
   
   # Test backend
   cd backend && python app.py
   ```

## 📞 Next Actions

1. **Commit and push** the fixes made
2. **Set environment variables** in Vercel dashboard
3. **Trigger redeploy** if needed
4. **Test the deployed app** once live

The deployment should now work correctly! 🎉
