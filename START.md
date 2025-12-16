# Quick Start Guide - Slide Translator

## Prerequisites Check

Before starting, ensure you have:
- ✅ Python 3.8 or higher installed
- ✅ OpenAI API key (with GPT-4 or GPT-3.5-turbo access)
- ✅ Modern web browser (Chrome, Firefox, Edge)

## Step-by-Step Setup

### 1. Install Python Dependencies

Open Command Prompt or PowerShell:

```bash
cd backend
pip install -r requirements.txt
```

Wait for installation to complete (~1-2 minutes).

### 2. Configure OpenAI API Key

Create `.env` file in `backend/` folder:

```bash
cd backend
copy .env.example .env
```

Edit `backend/.env` file and add your OpenAI API key:

```
OPENAI_API_KEY=sk-proj-your-actual-key-here
```

**Where to get an API key?**
- Visit: https://platform.openai.com/api-keys
- Create new key
- Copy and paste into .env file

### 3. Start the Backend Server

In Command Prompt (in backend/ directory):

```bash
python app.py
```

You should see:
```
Starting Slide Translator API server...
 * Running on http://0.0.0.0:5000
```

**Keep this terminal window open!**

### 4. Start the Frontend Server

Open a **NEW** Command Prompt window:

```bash
cd frontend
python -m http.server 3000
```

You should see:
```
Serving HTTP on :: port 3000 ...
```

**Keep this terminal window open too!**

### 5. Open in Browser

Open your web browser and go to:

```
http://localhost:3000
```

You should see the **Project C** homepage!

---

## Using the Application

### Quick Test

1. Click on **"Slide Translator"** tile (the blue one with 🌐)

2. **Upload Test Slide:**
   - Navigate to: `backend/tests/fixtures/sample_slide.pptx`
   - Drag it onto the upload area OR click "Choose File"

3. **Click "Translate Slide"**

4. **Wait 5-10 seconds** (you'll see progress updates)

5. **Click "Download Translated Slide"** when done

6. **Open the downloaded file in Microsoft PowerPoint** to verify:
   - Text is in Arabic ✓
   - Layout is RTL (text on right side) ✓
   - Bullets are preserved ✓

---

## Troubleshooting

### "OPENAI_API_KEY is required"

**Problem:** You didn't set the API key

**Solution:**
1. Check that `backend/.env` file exists
2. Open it and verify OPENAI_API_KEY is set
3. Restart the backend server (`python app.py`)

### "Port 5000 already in use"

**Problem:** Another application is using port 5000

**Solution:**
```bash
# Windows: Find and kill process on port 5000
netstat -ano | findstr :5000
taskkill /PID <process_id> /F

# Or change port in backend/app.py:
app.run(host='0.0.0.0', port=5001)  # Change to 5001
# Then update frontend/translator.js API_URL to http://localhost:5001/api/translate-slide
```

### "Failed to fetch" Error in Browser

**Problem:** Backend server is not running or CORS issue

**Solution:**
1. Check terminal - is backend running?
2. Visit http://localhost:5000/api/health in browser
   - Should show: `{"status": "ok", ...}`
3. If not, restart backend server

### Want Higher Quality Translations?

**Default:** Application uses GPT-3.5-turbo (fast, cost-effective)

**Upgrade:** Use GPT-4 for premium quality translations:

Edit `backend/.env`:
```
OPENAI_MODEL=gpt-4
```

Restart backend server.

**Note:** GPT-4 is ~15x more expensive but provides better consulting-grade language.

### Arabic Text Not Showing Correctly

**Problem:** Viewing in Google Slides or LibreOffice

**Solution:** **Open the file in Microsoft PowerPoint** for proper RTL rendering. Google Slides and LibreOffice have limited RTL support.

---

## Stopping the Application

1. **Stop Backend:** Press `Ctrl+C` in backend terminal
2. **Stop Frontend:** Press `Ctrl+C` in frontend terminal

---

## Command-Line Usage (Advanced)

Test without the UI:

```bash
cd backend

# Test slide parser
python modules/slide_parser.py tests/fixtures/sample_slide.py

# Test full workflow
python workflows/slide_translator.py tests/fixtures/sample_slide.pptx output.pptx

# Check output.pptx in PowerPoint
```

---

## Next Steps

- Read [README.md](README.md) for detailed documentation
- See [docs/WORKFLOW_EXPLANATION.md](docs/WORKFLOW_EXPLANATION.md) for how it works
- Upload your own slides to translate!

---

## Need Help?

Check the logs in the backend terminal for error messages and debugging information.

**Happy Translating!** 🚀
