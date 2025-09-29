# Excel Comparison Tool - Deployment Guide

## Quick Start

1. **Push to GitHub:**

   ```bash
   git remote add origin https://github.com/YOUR_USERNAME/excel-comparison-tool.git
   git push -u origin main
   ```

2. **Deploy on Railway:**

   - Go to [railway.app](https://railway.app)
   - Click "New Project"
   - Select "Deploy from GitHub repo"
   - Choose your `excel-comparison-tool` repository
   - Railway will automatically deploy

3. **Set Environment Variable:**

   - In Railway dashboard, go to your project
   - Go to "Variables" tab
   - Add: `MAIN_APP_URL` = `https://web-production-9e92a.up.railway.app/`

4. **Get Your URL:**

   - Railway will give you a URL like: `https://your-app.railway.app`
   - Copy this URL

5. **Update Main App:**
   - Go to your main app's Railway project
   - Add environment variable: `COMPARISON_URL` = `https://your-app.railway.app/comparison`

## Manual Deployment (Alternative)

If you prefer to upload files directly:

1. **Create a zip file:**

   ```bash
   zip -r comparison-tool.zip . -x "*.git*"
   ```

2. **Upload to Railway:**
   - Go to [railway.app](https://railway.app)
   - Click "New Project"
   - Select "Deploy from folder"
   - Upload the zip file

## Testing

After deployment, test the comparison tool:

- Visit your Railway URL
- Upload two Excel files
- Verify the comparison works
- Test the "Main App" button redirects correctly

## Troubleshooting

- **Port issues:** Railway automatically handles port configuration
- **Environment variables:** Make sure `MAIN_APP_URL` is set correctly
- **File uploads:** Check that file size limits are appropriate
