# GenDiploma Web (local) - Gmail API OAuth

Local web UI for creating and sending diplomas with HTML letter templates (Hebrew supported).

## Setup

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r webappGamilAPI/requirements.txt
```

If you see an error about multipart form data, install the missing dependency:

```bash
pip install python-multipart
```

If you see an error about `usedforsecurity` in OpenSSL/MD5, use ReportLab 3.x:

```bash
pip install "reportlab<4"
```

PDF-to-JPG needs Poppler:
- Ubuntu: `sudo apt-get install poppler-utils`
- Windows: install Poppler, then add `bin` to PATH

## Run

```bash
python -m uvicorn webappGamilAPI.app:app --reload --port 8001
```

Open http://127.0.0.1:8001

## Notes

- HTML supports `{{name}}` placeholder.
- Plain text letters are converted to HTML automatically for preview and sending.
- Use the Send management table to choose recipients and download or save an updated CSV with `new`/`send` changes.
- Inline logo can be referenced with `cid:logo_cid` (upload a logo in the form).
- Output files are written under `webappGamilAPI/output/`.

## Gmail OAuth setup (localhost)

1) Create a Google Cloud project, enable the Gmail API.
2) Create OAuth client credentials (Web application).
3) Add redirect URI for local testing, for example:
   - http://127.0.0.1:8001/oauth/callback
4) Download the client secret JSON and save it as:
   - `webappGamilAPI/client_secret.json`
5) Start the app and click "Connect Google".

The OAuth token is stored locally in `webappGamilAPI/.tokens/gmail_token.json`.
If you run on a different port, update the redirect URI and restart the app.
This app requests Gmail send and read-only scopes to auto-fill the sender email.

### Shared client secret

You can place the OAuth client secret anywhere and point the app to it:

```
export GMAIL_CLIENT_SECRET_PATH=/path/to/shared/client_secret.json
```

If set, it overrides `webappGamilAPI/client_secret.json`.
