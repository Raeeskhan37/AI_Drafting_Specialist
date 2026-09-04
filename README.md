# DraftForge — AI Document Composer

A free Streamlit app that drafts professional emails, letters, cover letters,
and reports using an LLM (Groq's Llama 3.3, or Google Gemini), lets you edit
the result, keeps a history of your drafts, and exports as **PDF, DOCX, TXT,
PNG, or JPG** with an optional letterhead (logo + company name/address).

## Step-by-step: deploy for free on Streamlit Community Cloud

This gives you a permanent public URL (e.g. `draftforge.streamlit.app`) at
zero cost.

**1. Get a free API key**
- Groq (recommended, fast + generous free tier): https://console.groq.com/keys
- Gemini (alternative): https://aistudio.google.com/apikey

**2. Create a GitHub repo**
- Go to https://github.com/new, name it (e.g. `draftforge`), keep it public, create it.

**3. Upload the three files**
- On the repo page, click "Add file → Upload files".
- Drag in `app.py`, `requirements.txt`, and `README.md`.
- Commit the changes.

  *(Or via command line instead of the web UI:)*
  ```bash
  git init
  git add app.py requirements.txt README.md
  git commit -m "Initial commit: DraftForge"
  git branch -M main
  git remote add origin https://github.com/<your-username>/draftforge.git
  git push -u origin main
  ```

**4. Deploy on Streamlit Cloud**
- Go to https://share.streamlit.io and sign in with GitHub.
- Click "New app".
- Pick your repo, branch `main`, and main file `app.py`.
- Click "Deploy". It builds for ~1–2 minutes.

**5. (Optional but recommended) Save your API key as a secret**
- On your app's page, click "⋮ → Settings → Secrets".
- Add:
  ```
  GROQ_API_KEY = "gsk_your_key_here"
  GEMINI_API_KEY = "your_key_here"
  ```
- Save. The app already reads these as defaults, so you and visitors won't
  need to paste a key in every time (anyone using your hosted app will use
  your key and quota, so only do this if that's what you want — otherwise
  leave the key field blank and let each visitor paste their own).

**6. Done** — your app is live at `https://<your-app-name>.streamlit.app`.
Any future `git push` to `main` auto-redeploys it.

## Alternative ways to run it

### Locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
Open http://localhost:8501.

### Google Colab (quick testing, no GitHub needed)
```python
!pip install -q streamlit requests python-docx fpdf2 pillow pyngrok
# upload app.py and requirements.txt to the Colab file browser first, or:
!wget -q -O app.py https://raw.githubusercontent.com/<your-username>/draftforge/main/app.py

from pyngrok import ngrok
ngrok.set_auth_token("YOUR_NGROK_TOKEN")  # free at ngrok.com
print(ngrok.connect(8501))

!streamlit run app.py &>/content/log.txt &
```
Open the printed ngrok URL.

## Features
- Choose document type (email, formal letter, business report, cover letter, custom) and tone.
- Fill in recipient, subject, and key points — the model drafts the full document.
- Edit the AI's draft directly in the browser before exporting.
- **History**: past drafts are saved locally (SQLite) and browsable/reloadable from the sidebar, tagged by the name you enter.
- **Letterhead**: add a company name, address, and logo — applied automatically to PDF, DOCX, and image exports.
- Export to PDF, Word (.docx), plain text, PNG, or JPG.

## Notes on the history feature
Drafts are stored in a local `drafts.db` SQLite file next to `app.py`. On
Streamlit Community Cloud this file persists only for the life of the running
container (it resets on redeploy or after long inactivity) — fine for
personal/demo use. For durable multi-user history in production, swap the
SQLite calls for a hosted database (e.g. Supabase, Turso, or Streamlit's
built-in SQL connector) — the `save_draft` / `get_history` / `delete_draft`
functions are isolated so this is a small, contained change.

## Extending it
- Add real login (`streamlit-authenticator`) if multiple people will use one deployment and need private history.
- Add more providers (OpenAI, Anthropic, Cohere) via another `call_*` function.
- Add more letterhead templates (colors, fonts, layouts).
