# מערכת שליחת תעודות (Gmail API OAuth)

Local web UI for creating and sending diplomas with HTML or plain‑text letters (Hebrew supported).

## Requirements

- Python 3.8+
- Poppler (for PDF → JPG)

## Install (Linux)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r webappGamilAPI/requirements.txt

# Poppler
sudo apt-get install poppler-utils
```

## Install (Windows)

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r webappGamilAPI\requirements.txt
```

Poppler:
- Install Poppler for Windows and add its `bin` folder to PATH.

## Run

```bash
python -m uvicorn webappGamilAPI.app:app --reload --port 8001
```

Open: `http://127.0.0.1:8001`

## Run with Docker (Windows / Linux)

Prerequisites (Windows):
- Docker Desktop with WSL2 backend enabled

Run:

```bash
docker compose up --build
```

Stop:
- Ctrl+C, then:

```bash
docker compose down
```

Troubleshooting:
- Port 8000 already in use: change the compose mapping to `"8001:8000"`
- Check logs: `docker compose logs -f`
- PDF → JPG fails: Poppler is already included in the image

## Gmail OAuth setup (localhost)

1) Create a Google Cloud project and enable **Gmail API**.
2) Create OAuth client credentials (**Web application**).
3) Add redirect URI: `http://127.0.0.1:8001/oauth/callback`
4) Download the client secret JSON and save it as:
   - `webappGamilAPI/client_secret.json`
5) Start the app and click **Connect Google**.

Tokens are stored locally at `webappGamilAPI/.tokens/gmail_token.json`.

### Shared client secret (optional)

Set a shared location for the OAuth client secret:

```bash
export GMAIL_CLIENT_SECRET_PATH=/path/to/shared/client_secret.json
```

## Notes

- CSV expected headers: `email,first,last,new,send`
- Sends only when `new=1` and `send` is empty.
- Output files are written under `webappGamilAPI/output/`.
