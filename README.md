# 🧠 Personal AI Flashcard & Quiz Generator

A Streamlit web app that turns your study notes (pasted text, PDF or TXT files)
into **flashcards** and **multiple-choice quizzes** using the Google Gemini API.
It works with a free API key from Google AI Studio.

## Features

- Paste notes, or upload PDF / TXT files
- Interactive flashcards (tap to reveal the answer)
- Multiple-choice quizzes (4 options), score and explanations
- Easy / Medium / Hard difficulty and 3-20 items
- Study sets organised by subject or topic
- Input validation, error messages and loading indicators
- AI output is requested in a fixed JSON structure and checked before display
- Mobile-friendly layout

## Important limits (please read)

- There is **no database and no login**. Study sets live in Streamlit
  *session state*, so they disappear when you refresh or close the page.
- Scanned PDFs (pictures of text) cannot be read. Use PDFs with real text.
- Very long material is cut to the first 24,000 characters.
- The free Gemini tier has daily and per-minute limits. If you see a
  "quota" message, wait and try again.
- On Google's free tier, your prompts may be used to improve Google products.
  Do not paste private or sensitive notes. Check Google's current terms.
- Once deployed, anyone with the link can use the app and use **your** quota.

## Files

| File | Purpose |
|------|---------|
| `app.py` | The whole application |
| `requirements.txt` | Python packages to install |
| `.gitignore` | Stops secrets and junk files being uploaded |
| `README.md` | This guide |

## Get a free API key

1. Open https://aistudio.google.com/apikey and sign in with a Google account.
2. Tap **Create API key** and copy it.
3. Keep it private. Never put it in GitHub or in `app.py`.

## Run on your computer

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Create a file `.streamlit/secrets.toml` (it is ignored by Git):

```toml
GEMINI_API_KEY = "paste-your-key-here"
```

Then start the app:

```bash
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Put the project in a GitHub repository (never upload your key).
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Click **Create app**, choose the repository, branch `main`, main file `app.py`.
4. Open **Advanced settings**:
   - Choose Python **3.12** (or any version offered).
   - In **Secrets**, paste:
     ```toml
     GEMINI_API_KEY = "paste-your-key-here"
     ```
5. Click **Deploy**.

To change a secret later: app menu -> **Settings** -> **Secrets**.

### Optional: choose the AI model

The app tries `gemini-3.8-flash`, then `gemini-3.5-flash`, then
`gemini-3.5-flash-lite`. To force one model, add this to secrets:

```toml
GEMINI_MODEL = "gemini-3.5-flash"
```

Model names change over time. See https://ai.google.dev/gemini-api/docs/models
for the current list.

## Troubleshooting

| Message | What to do |
|---------|-----------|
| "No API key found" | Add `GEMINI_API_KEY` in Secrets, then reboot the app |
| "Google rejected the API key" | The key is wrong or deleted. Create a new one |
| "free quota is used up" | Wait a few minutes, or until tomorrow |
| "model was not found" | Set `GEMINI_MODEL` in Secrets to a current model name |
| "no readable text" in a PDF | The PDF is a scan. Use a text PDF or paste the text |
