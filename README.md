# Deal Deliverability Review Generator

**Accenture Security Practice** — Automated Deal Deliverability Review PowerPoint Generation

## What This Does

Upload a filled Excel assessment + RFP document → AI analyzes both → Download a complete 5-slide PowerPoint deck ready for leadership review.

The tool:
- Parses all 21 assessment questions across 5 dimensions
- Calculates RAG scores (Green/Amber/Red) per dimension and overall
- Reads and analyzes the RFP document
- Uses AI to generate narrative content (deal overview, justification, bullets, assumptions, next steps)
- Produces a professional PowerPoint deck matching Accenture design standards

## Quick Start (Local — 5 minutes)

### 1. Install Python
Download from [python.org](https://www.python.org/downloads/) if you don't have it. Version 3.9+ required.

### 2. Clone or download this folder
Put all files in one folder on your machine.

### 3. Install dependencies
Open a terminal in the project folder and run:
```bash
pip install -r requirements.txt
```

### 4. Get a free AI API key
**Recommended: Google Gemini (free)**
1. Go to [aistudio.google.com](https://aistudio.google.com)
2. Sign in with any Google account
3. Click "Get API Key" → "Create API Key"
4. Copy the key

### 5. Run the app
```bash
streamlit run app.py
```
The app opens in your browser at `http://localhost:8501`.

### 6. Use it
1. Select your AI provider in the sidebar
2. Paste your API key
3. Upload your filled Excel (.xlsx/.xlsm) and RFP (.docx)
4. Click "Analyze & Generate"
5. Download the PPT

## Deploy to Streamlit Cloud (Free — Public URL)

This gives you a URL like `https://your-app.streamlit.app` that anyone can access.

### 1. Push code to GitHub
Create a GitHub repository and push all files:
```bash
git init
git add .
git commit -m "Deliverability Review Generator"
git remote add origin https://github.com/YOUR_USERNAME/deliverability-review.git
git push -u origin main
```

**Important:** Do NOT push `.streamlit/secrets.toml` with real API keys. Add it to `.gitignore`.

### 2. Deploy on Streamlit Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with your GitHub account
3. Click "New app"
4. Select your repository, branch `main`, file `app.py`
5. Click "Deploy"

### 3. Add API keys
1. In your deployed app, click the ⋮ menu → "Settings"
2. Go to "Secrets"
3. Paste your API keys:
```toml
GEMINI_API_KEY = "your-key-here"
```
4. Save — the app restarts with keys configured

### 4. Share the URL
Copy the URL and share with your team. Anyone with the link can use it.

## Project Structure

```
deliverability-review-app/
├── app.py              # Main Streamlit UI
├── excel_parser.py     # Excel assessment parser + RAG scoring
├── rfp_reader.py       # RFP .docx text extraction
├── ai_engine.py        # AI provider integration (Gemini/Groq/Cohere)
├── ppt_builder.py      # PowerPoint generation with python-pptx
├── requirements.txt    # Python dependencies
├── .streamlit/
│   └── secrets.toml    # API keys (don't commit to git!)
└── README.md           # This file
```

## AI Providers Supported

| Provider | Model | Free Tier | Sign Up |
|----------|-------|-----------|---------|
| Google Gemini | Gemini 2.0 Flash | 15 req/min | [aistudio.google.com](https://aistudio.google.com) |
| Groq | Llama 3.3 70B | Generous | [console.groq.com](https://console.groq.com) |
| Cohere | Command R+ | Trial | [dashboard.cohere.com](https://dashboard.cohere.com) |

## Excel File Requirements

The Excel file must have these sheets:
- **02_Assessment**: 21 questions, Col B=Dimension, D=Question, F=Response, H=Team, I=RAG, J=Justification, K=Action
- **01_Deal_Overview**: Cell A3 = deal overview text
- **03_Risks**: Col A=Risk, B=Mitigation, C=Owner

## Built By

Raghad Altawil — Accenture Security Practice  
