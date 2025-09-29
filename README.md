# Excel Comparison Tool

A web-based tool for comparing two Excel files and identifying differences between them.

## Features

- Upload two Excel files for comparison
- Side-by-side comparison view
- Highlight differences between files
- Export comparison results
- Clean, responsive web interface

## Local Development

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python excel_comparison.py
```

3. Open your browser and go to: http://localhost:5002

## Deployment

This app is designed to be deployed on Railway or similar platforms.

### Environment Variables

- `MAIN_APP_URL`: URL of the main Excel automation app (for navigation)

### Railway Deployment

1. Connect your GitHub repository to Railway
2. Set the environment variable `MAIN_APP_URL` to your main app's URL
3. Deploy!

## Usage

1. Upload your first Excel file (Raw Data)
2. Upload your second Excel file (Previous Data) 
3. Click "Compare Files" to see the differences
4. Use the navigation to switch between apps

## Related Apps

- [Main Excel Automation App](../excel-ai-builder/) - Upload Excel files and give natural language instructions