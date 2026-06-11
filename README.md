# Law School Tools

A small toolkit for studying smarter. Two tools:

1. **PowerPoint Generator** — turns your notes into a clean academic .pptx (needs the backend running)
2. **Quiz Me** — fill-in-the-blank or multiple-choice practice questions from your notes (works entirely in the browser)

## Files

- `index.html` — the website
- `api_server.py` — the FastAPI backend that builds .pptx files from `template3.pptx`
- `template3.pptx` — the PowerPoint template

## Hosting

**If you only want Quiz Me:** just put `index.html` on GitHub Pages and it works.

**If you also want the PowerPoint Generator** you need to host the Python backend somewhere that allows running Python (GitHub Pages does NOT). Easy free options:
- Render.com (free tier)
- Railway.app
- Fly.io
- Replit
- Or run it on your own computer with `python api_server.py`

Then in `index.html`, find this line:
```js
const PPT_API = '__PORT_8000__'.startsWith('__') ? 'http://localhost:8000' : '__PORT_8000__';
```
and replace it with your backend's URL, e.g.:
```js
const PPT_API = 'https://your-backend.onrender.com';
```

## Running the backend locally

```bash
pip install fastapi uvicorn python-pptx
python api_server.py
```
Then open `index.html` in your browser.
