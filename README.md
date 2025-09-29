# Excel Comparison Tool

A Flask web application for comparing Excel files and analyzing differences between datasets.

## Features

- Upload and compare two Excel files
- Side-by-side comparison view
- Highlight differences between datasets
- Download comparison results
- Clean, responsive interface

## Local Development

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python excel_comparison.py
```

3. Open your browser to `http://localhost:5002`

## Deployment

This app is designed to be deployed on Railway. The main app will automatically link to this comparison tool.

### Railway Deployment

1. Create a new Railway project
2. Connect this repository
3. Railway will automatically detect the `Procfile` and deploy
4. Set the environment variable `MAIN_APP_URL` to your main app's URL

## Environment Variables

- `MAIN_APP_URL`: URL of the main Excel AI Builder app (default: http://localhost:5001/)

## API Endpoints

- `GET /` - Redirects to `/comparison`
- `GET /comparison` - Main comparison interface
- `POST /upload` - Upload Excel files for comparison
- `GET /download` - Download comparison results

## Requirements

- Python 3.10+
- Flask
- Pandas
- OpenPyXL
