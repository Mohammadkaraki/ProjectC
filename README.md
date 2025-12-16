# Slide Translator - Context-Aware RTL Translation

## 🎯 Project Overview

**Slide Translator** is an AI-powered automation tool that translates consulting PowerPoint slides from English to Arabic with automatic Right-to-Left (RTL) layout conversion. This MVP demonstrates end-to-end workflow automation for consulting environments where slides need to be translated while preserving their professional structure and visual layout.

**Case Study:** Project C - Consulting Automation Hub
**Feature:** Slide Translator (Context-Aware RTL Translation)
**Timeline:** 12-hour MVP development

### 🌐 Live Production Deployment
- **Live URL:** [creativeshowroom.site](https://creativeshowroom.site)
- **Infrastructure:** AWS EC2 with Static IP (Elastic IP)
- **Containerization:** Docker
- **CI/CD:** Automated deployment pipeline
- **Status:** ✅ Production-ready and operational

---

## ✨ Key Features

- **Context-Aware Translation:** Uses OpenAI GPT-4 to understand slide hierarchy (title, header, bullets) for accurate, professional translation
- **Automatic RTL Layout Conversion:** Flips slide layout from Left-to-Right to Right-to-Left via PowerPoint XML manipulation
- **Structure Preservation:** Maintains bullet hierarchy, formatting, and slide organization
- **Professional Output:** Generates consulting-grade Arabic slides ready for executive presentations
- **Web Interface:** User-friendly Project C homepage with drag-and-drop file upload

---

## 🏗️ System Architecture

```
┌─────────────────────────────────────────────────────┐
│              PROJECT C FRONTEND (HTML/CSS/JS)       │
│  - Feature tiles homepage                           │
│  - Drag & drop file upload                          │
│  - Progress tracking UI                             │
└────────────────────┬────────────────────────────────┘
                     │ HTTP POST /api/translate-slide
                     ▼
┌─────────────────────────────────────────────────────┐
│          FLASK API SERVER (Python)                  │
│  - File upload endpoint                             │
│  - Workflow orchestration                           │
└────────────────────┬────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────┐
│         SLIDE TRANSLATOR WORKFLOW                   │
│  Step 1: Parse Slide Structure                      │
│  Step 2: Build Context Map                          │
│  Step 3: Translate with OpenAI GPT-4               │
│  Step 4: Convert to RTL Layout (XML manipulation)   │
│  Step 5: Replace Text with Translations             │
│  Step 6: Generate Output File                       │
└─────────────────────────────────────────────────────┘
```

---

## 📁 Project Structure

```
slide-translator/
├── backend/                      # Python backend
│   ├── app.py                    # Flask API entry point
│   ├── config.py                 # Configuration management
│   ├── requirements.txt          # Python dependencies
│   ├── .env.example              # Environment variables template
│   │
│   ├── modules/                  # Core processing modules
│   │   ├── slide_parser.py       # Extract slide structure
│   │   ├── context_builder.py    # Identify element roles
│   │   ├── llm_translator.py     # OpenAI translation
│   │   ├── rtl_converter.py      # RTL layout flipping
│   │   └── text_replacer.py      # Text replacement
│   │
│   ├── workflows/
│   │   └── slide_translator.py   # Main workflow orchestrator
│   │
│   ├── utils/
│   │   ├── file_handler.py       # File I/O utilities
│   │   └── logger.py             # Logging setup
│   │
│   └── tests/
│       ├── create_sample_slide.py # Generate test slides
│       └── fixtures/
│           └── sample_slide.pptx  # Sample consulting slide
│
├── frontend/                     # Web UI
│   ├── index.html                # Project C homepage
│   ├── slide-translator.html     # Slide Translator page
│   ├── styles.css                # Styling
│   └── translator.js             # Client-side logic
│
├── tmp/                          # Temporary files
│   ├── uploads/                  # Uploaded files
│   └── outputs/                  # Translated outputs
│
├── docs/                         # Documentation
│   ├── ARCHITECTURE.md           # System architecture
│   └── WORKFLOW_EXPLANATION.md   # Detailed workflow explanation
│
└── README.md                     # This file
```

---

## 🚀 Quick Start

### Prerequisites

- **Python 3.8+**
- **OpenAI API Key** (GPT-4 access recommended)
- **Modern web browser** (Chrome, Firefox, Edge)

### Installation

1. **Clone or extract the project**
```bash
cd slide-translator
```

2. **Set up Python backend**
```bash
cd backend
pip install -r requirements.txt
```

3. **Configure environment variables**
```bash
# Create .env file from template
cp .env.example .env

# Edit .env and add your OpenAI API key
# OPENAI_API_KEY=sk-proj-your-key-here
```

4. **Start the Flask backend**
```bash
python app.py
```

Backend will run on `http://localhost:5000`

5. **Start the frontend** (in a new terminal)
```bash
cd frontend
python -m http.server 3000
```

Frontend will run on `http://localhost:3000`

6. **Open in browser**
```
http://localhost:3000
```

---

## 📖 Usage Guide

### Using the Web Interface

1. **Navigate to Project C Homepage**
   Open `http://localhost:3000` in your browser

2. **Click "Slide Translator" feature tile**
   The active feature with 🌐 icon

3. **Upload PowerPoint slide**
   - Drag & drop .pptx file onto the upload area, OR
   - Click "Choose File" to select from your computer
   - Max file size: 16 MB

4. **Start translation**
   Click "Translate Slide" button

5. **Monitor progress**
   Watch real-time progress through 5 workflow steps:
   - Parsing slide structure
   - Building context map
   - Translating with AI
   - Converting to RTL layout
   - Finalizing translation

6. **Download translated slide**
   Click "Download Translated Slide" when complete

### Using the Command Line

**Test individual modules:**
```bash
cd backend

# Test slide parser
python modules/slide_parser.py tests/fixtures/sample_slide.pptx

# Test full workflow
python workflows/slide_translator.py tests/fixtures/sample_slide.pptx output.pptx
```

**Test API endpoint:**
```bash
curl -X POST http://localhost:5000/api/translate-slide \
  -F "file=@tests/fixtures/sample_slide.pptx" \
  --output translated.pptx
```

---

## 🔧 Technical Implementation

### Core Modules

#### 1. **Slide Parser** (`slide_parser.py`)
- **Purpose:** Extract text content and structure from PowerPoint
- **Technology:** python-pptx library
- **Output:** Structured JSON with titles, headers, bullets, and positions

#### 2. **Context Builder** (`context_builder.py`)
- **Purpose:** Identify the role of each text element
- **Logic:** Classifies elements as slide_title, header, bullet_group, or supporting_text
- **Why:** Enables context-aware translation (titles translated differently than bullets)

#### 3. **LLM Translator** (`llm_translator.py`)
- **Purpose:** Translate text from English to Arabic
- **Technology:** OpenAI GPT-4 API
- **Features:**
  - Context-aware prompts for professional consulting tone
  - Batch translation for consistency
  - JSON-structured output for reliability

#### 4. **RTL Converter** (`rtl_converter.py`) ⭐ **Most Complex**
- **Purpose:** Convert slide layout from LTR to RTL
- **Challenges:** PowerPoint doesn't expose RTL via python-pptx API
- **Solution:**
  - Direct XML manipulation using lxml
  - Add `rtl="1"` attribute to paragraph properties
  - Flip text box positions: `new_left = slide_width - (old_left + width)`
  - Set text alignment to RIGHT
  - Apply Arabic-compatible fonts

#### 5. **Text Replacer** (`text_replacer.py`)
- **Purpose:** Insert translated text into RTL-converted slide
- **Features:** Preserves bullet hierarchy and slide structure

#### 6. **Workflow Orchestrator** (`slide_translator.py`)
- **Purpose:** Execute all steps in correct sequence
- **Features:** Error handling, logging, state management

---

## 🎨 Frontend Features

### Project C Homepage
- **Feature Tiles Layout:** 6 tiles showcasing automation features
- **Active Feature:** Slide Translator (functional)
- **Placeholder Features:** 5 coming-soon features demonstrating platform vision
- **Professional Design:** Gradient styling, hover effects, responsive layout

### Slide Translator Page
- **Drag & Drop Upload:** Intuitive file upload with visual feedback
- **Progress Tracking:** Real-time workflow step indication
- **Error Handling:** User-friendly error messages
- **Download Management:** One-click download of translated file

---

## 🧪 Testing

### Sample Slide Provided

A consulting slide is included at:
```
backend/tests/fixtures/sample_slide.pptx
```

**Contents:**
- Title: "Market Expansion Strategy"
- Header: "Key findings from MENA region analysis"
- 4 bullet points with consulting content
- 1 sub-bullet demonstrating hierarchy

### Test Workflow

1. **Generate fresh sample slide:**
```bash
cd backend
python tests/create_sample_slide.py
```

2. **Test via web interface:**
- Upload `sample_slide.pptx`
- Download translated output
- Open in PowerPoint to verify:
  - Text is in Arabic
  - Layout is RTL (text aligned right, shapes flipped)
  - Structure preserved

3. **Test via command line:**
```bash
python workflows/slide_translator.py tests/fixtures/sample_slide.pptx output.pptx
```

---

## 🛠️ Configuration

### Environment Variables (.env)

```bash
# OpenAI API Configuration
OPENAI_API_KEY=sk-proj-your-key-here
OPENAI_MODEL=gpt-4  # or gpt-3.5-turbo for cost savings

# Flask Configuration
FLASK_ENV=development
FLASK_DEBUG=True

# File Paths (auto-configured)
UPLOAD_FOLDER=../tmp/uploads
OUTPUT_FOLDER=../tmp/outputs

# Logging
LOG_LEVEL=INFO
```

### Cost Considerations

**OpenAI API Usage:**
- **GPT-3.5-turbo: ~$0.002 per slide (DEFAULT - fast and cost-effective)**
- GPT-4: ~$0.03 per slide (premium option for higher quality)

The application uses GPT-3.5-turbo by default. For higher quality translations, change to GPT-4 by setting `OPENAI_MODEL=gpt-4` in .env

---

## 🚢 Production Deployment

### Infrastructure Architecture

```
User → creativeshowroom.site (AWS Route 53)
  ↓
AWS Elastic IP (Static IP)
  ↓
AWS EC2 Instance
  ↓
Docker Container → Flask API + React Frontend
  ↓
OpenAI API (GPT-3.5-turbo)
```

### Deployment Features

✅ **Docker Containerization**
- Consistent environment across dev/prod
- Easy scaling and deployment
- Isolated dependencies

✅ **CI/CD Pipeline**
- Automated testing
- Automated deployment
- Version control integration

✅ **AWS Infrastructure**
- Static IP (Elastic IP) for reliability
- Custom domain: creativeshowroom.site
- Scalable EC2 instance

### Deploy with Docker

```bash
# Build Docker image
docker build -t slide-translator .

# Run container
docker run -d -p 5000:5000 \
  -e OPENAI_API_KEY=your_key_here \
  --name slide-translator \
  slide-translator

# Check logs
docker logs -f slide-translator
```

---

## 🐛 Troubleshooting

### Backend Issues

**"OPENAI_API_KEY is required"**
- Solution: Add your API key to backend/.env file

**"ModuleNotFoundError"**
- Solution: Ensure you're in backend/ directory and ran `pip install -r requirements.txt`

**"Port 5000 already in use"**
- Solution: Kill process using port 5000 or change port in app.py

### Frontend Issues

**"Failed to fetch"**
- Solution: Ensure Flask backend is running on http://localhost:5000
- Check CORS is enabled (flask-cors installed)

**"Translation failed"**
- Check backend logs in terminal
- Verify OpenAI API key is valid and has credits
- Ensure .pptx file is valid PowerPoint format

### RTL Layout Issues

**"Text not aligned right"**
- This is expected behavior in some PowerPoint versions
- Open file in Microsoft PowerPoint (not Google Slides) for proper RTL rendering

---

## 📈 Future Enhancements

Post-MVP improvements:
- [ ] Image and chart handling
- [ ] Batch processing (multiple files)
- [ ] Translation memory for consistency
- [ ] User authentication and history
- [ ] Enhanced cloud deployment (Azure, GCP)

---

## 📝 License

This project is a case study MVP for evaluation purposes.

---

## 👤 Author

**Case Study Submission:** Slide Translator Feature
**Project:** Project C - Consulting Automation Hub
**Date:** 2024

---

## 📚 Additional Documentation

- **[WORKFLOW_EXPLANATION.md](docs/WORKFLOW_EXPLANATION.md):** Detailed step-by-step workflow explanation
- **[ARCHITECTURE.md](docs/ARCHITECTURE.md):** System architecture and design decisions

---

## 🙏 Acknowledgments

- **python-pptx:** PowerPoint manipulation library
- **OpenAI GPT-4:** Context-aware translation
- **lxml:** XML parsing and manipulation
- **Flask:** REST API framework

---

**Thank you for reviewing this case study submission!** 🚀
