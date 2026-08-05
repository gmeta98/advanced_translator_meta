# Professional Word Translation App

Streamlit app for translating one or more `.docx` files between Albanian, Italian, English, French, German, and custom languages.

## Run

```bash
python3 -m streamlit run app.py
```

## Local Setup

Create a `.env` file with the app password and OpenAI configuration:

```bash
APP_PASSWORD=choose_a_strong_password
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=your_preferred_model_here
```

If no API key is set, the app runs in preview mode so you can test uploads and exports without sending text anywhere.

## Streamlit Community Cloud

1. Put `app.py`, `requirements.txt`, and the supporting repository files on GitHub.
2. In Streamlit Community Cloud, create an app and select `app.py` as the entrypoint.
3. Open **Advanced settings** and add these values in **Secrets**:

```toml
APP_PASSWORD = "choose_a_strong_password"
OPENAI_API_KEY = "your_api_key_here"
OPENAI_MODEL = "your_preferred_model_here"
```

Never commit `.env` or `.streamlit/secrets.toml`. The app accepts configuration from either local environment variables or Streamlit Secrets.

The repository may be public, but the deployed app should still use a strong password. Keep client Word files out of the repository; `.gitignore` excludes `.docx`, ZIP exports, logs, and local work folders as an additional safeguard.

### Public repository contents

- `app.py` - Streamlit application
- `requirements.txt` - Python dependencies installed by Streamlit
- `.env.example` - safe local configuration template with placeholders only
- `.streamlit/secrets.toml.example` - safe Cloud Secrets template with placeholders only
- `.gitignore` - prevents credentials and client documents from being committed
- `README.md` - setup and deployment instructions

## Notes

- Uploads stay in memory during the Streamlit session.
- A shared password is required before upload, translation, or API settings are shown.
- Logging out clears uploaded and translated document data from the active session.
- Upload one file to download one translated `.docx`, or upload several files to download a ZIP of translated `.docx` files.
- The translated export preserves the Word document structure, paragraph styles, run-level text formatting, tables, headers, and footers by replacing existing text nodes only.
- The app analyzes repeated text and likely repeated table-header rows before translation. Exact repeated source text is translated once and reused for speed and consistency.
- Legal, judicial, police, Guardia di Finanza, prosecution, immigration, and other professional profiles add document-specific translation guidance while keeping layout untouched.
- The app is designed for office review workflows: preview first, translate, then download the translated `.docx`.
