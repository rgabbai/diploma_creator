from __future__ import annotations

import base64
import csv
import html as html_lib
import io
import json
import os
import re
import time
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO
from pathlib import Path
from typing import List, Dict, Tuple, Optional

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.requests import Request
from google.auth.transport.requests import Request as GoogleRequest
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
try:
    from pdf2image import convert_from_path
except Exception:  # noqa: BLE001
    convert_from_path = None
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import PyPDF2


APP_DIR = Path(__file__).resolve().parent
REPO_DIR = APP_DIR.parent
OUTPUT_DIR = APP_DIR / "output"
FONT_PATH = REPO_DIR / "Noto_Sans_Hebrew,Noto_Serif_Hebrew" / "Noto_Sans_Hebrew" / "static" / "NotoSansHebrew_Condensed-Bold.ttf"
FONT_PREF_ENV = "DIPLOMA_FONT"
FONT_CANDIDATES = [
    ("Noto", [FONT_PATH]),
    ("Arial", [Path("C:/Windows/Fonts/arial.ttf"), Path("/usr/share/fonts/truetype/msttcorefonts/Arial.ttf")]),
    ("David", [Path("C:/Windows/Fonts/david.ttf"), Path("C:/Windows/Fonts/davidr.ttf")]),
]
ACTIVE_FONT_NAME: str | None = None
ACTIVE_FONT_PATH: Path | None = None
DEFAULT_SUBJECT = "תעודת סיום קורס"
ADMIN_NOTIFY_EMAIL = "rony.gabbai@gmail.com"
CLIENT_SECRETS = APP_DIR / "client_secret.json"
CLIENT_SECRETS_ENV = "GMAIL_CLIENT_SECRET_PATH"
TOKEN_DIR = APP_DIR / ".tokens"
TOKEN_FILE = TOKEN_DIR / "gmail_token.json"
DISABLE_JPG_ENV = "DIPLOMA_DISABLE_JPG"
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
]
OAUTH_STATE: str | None = None
OAUTH_REDIRECT_URI: str | None = None

app = FastAPI()
app.mount("/static", StaticFiles(directory=APP_DIR / "static"), name="static")
app.mount("/output", StaticFiles(directory=OUTPUT_DIR), name="output")
templates = Jinja2Templates(directory=APP_DIR / "templates")


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


def register_font_once() -> None:
    global ACTIVE_FONT_NAME, ACTIVE_FONT_PATH
    if ACTIVE_FONT_NAME and ACTIVE_FONT_NAME in pdfmetrics.getRegisteredFontNames():
        return

    for font_name, paths in FONT_CANDIDATES:
        if font_name in pdfmetrics.getRegisteredFontNames():
            if ACTIVE_FONT_NAME is None:
                ACTIVE_FONT_NAME = font_name
            continue
        for path in paths:
            if path.exists():
                pdfmetrics.registerFont(TTFont(font_name, str(path)))
                if ACTIVE_FONT_NAME is None:
                    ACTIVE_FONT_NAME = font_name
                    ACTIVE_FONT_PATH = path
                break

    preferred = os.environ.get(FONT_PREF_ENV, "").strip()
    if preferred and preferred in pdfmetrics.getRegisteredFontNames():
        ACTIVE_FONT_NAME = preferred
        for font_name, paths in FONT_CANDIDATES:
            if font_name == preferred:
                for path in paths:
                    if path.exists():
                        ACTIVE_FONT_PATH = path
                        break
                break

    if ACTIVE_FONT_NAME is None:
        raise FileNotFoundError(f"No supported font files found. Checked: {FONT_CANDIDATES}")


def get_active_font() -> str:
    register_font_once()
    return ACTIVE_FONT_NAME or "Helvetica"


def get_active_font_path() -> Optional[Path]:
    register_font_once()
    return ACTIVE_FONT_PATH


def is_hebrew(text: str) -> bool:
    return any("\u0590" <= char <= "\u05FF" for char in text)


def make_overlay_pdf(student_name: str, x_offset: int = 0, y_offset: int = 0) -> BytesIO:
    register_font_once()
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    width, height = letter

    display_name = student_name[::-1] if is_hebrew(student_name) else student_name
    name_y_offset = 220

    can.setFont(get_active_font(), 42)
    can.drawCentredString(width / 2.0 + x_offset, height - name_y_offset + y_offset, display_name)
    can.save()

    packet.seek(0)
    return packet


def generate_diploma_pdf(
    pdf_template_path: Path,
    student_name: str,
    output_dir: Path,
    name_x_offset: int = 0,
    name_y_offset: int = 0,
) -> Path:
    overlay_packet = make_overlay_pdf(student_name, x_offset=name_x_offset, y_offset=name_y_offset)
    template_pdf = PyPDF2.PdfReader(open(pdf_template_path, "rb"))
    existing_page = template_pdf.pages[0]

    new_pdf = PyPDF2.PdfReader(overlay_packet)
    overlay_page = new_pdf.pages[0]
    existing_page.merge_page(overlay_page)

    output = PyPDF2.PdfWriter()
    output.add_page(existing_page)

    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_filename = output_dir / f"{student_name.replace(' ', '_')}_diploma.pdf"
    with open(pdf_filename, "wb") as output_stream:
        output.write(output_stream)

    return pdf_filename


def convert_pdf_to_jpg(pdf_filename: Path) -> Optional[Path]:
    if os.environ.get(DISABLE_JPG_ENV, "1").strip() in {"1", "true", "yes"}:
        return None
    if convert_from_path is None:
        return None
    images = convert_from_path(str(pdf_filename))
    jpg_filename = pdf_filename.with_suffix(".jpg")
    images[0].save(jpg_filename, "JPEG")
    return jpg_filename


def generate_diploma_jpg(
    jpg_template_path: Path,
    student_name: str,
    output_dir: Path,
    name_x_offset: int = 0,
    name_y_offset: int = 0,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    image = Image.open(jpg_template_path).convert("RGB")
    draw = ImageDraw.Draw(image)

    width, height = image.size
    base_x = letter[0] / 2.0
    base_y = letter[1] - 220

    scale_x = width / letter[0]
    scale_y = height / letter[1]
    absolute_x = (base_x + name_x_offset) * scale_x
    absolute_y = height - ((base_y + name_y_offset) * scale_y)

    display_name = student_name
    font_size = max(10, int(42 * scale_y))
    font_path = get_active_font_path()
    if font_path and font_path.exists():
        font = ImageFont.truetype(str(font_path), font_size)
    else:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), display_name, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    draw.text((absolute_x - text_width / 2, absolute_y - text_height / 2), display_name, fill=(0, 0, 0), font=font)

    jpg_filename = output_dir / f"{student_name.replace(' ', '_')}_diploma.jpg"
    image.save(jpg_filename, "JPEG", quality=95)
    return jpg_filename


def build_message(
    student_name: str,
    student_email: str,
    html_content: str,
    subject: str,
    from_email: str,
    pdf_filename: Optional[Path],
    jpg_filename: Optional[Path],
    logo_bytes: Optional[bytes] = None,
    logo_filename: Optional[str] = None,
) -> MIMEMultipart:
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject or DEFAULT_SUBJECT
    if from_email:
        msg["From"] = from_email
    msg["To"] = student_email

    html_content = html_content.replace("{{name}}", student_name)
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(html_content, "html", "utf-8"))
    msg.attach(alt)

    if logo_bytes:
        image = MIMEImage(logo_bytes, name=logo_filename or "logo")
        image.add_header("Content-ID", "<logo_cid>")
        msg.attach(image)

    if pdf_filename and pdf_filename.exists():
        with open(pdf_filename, "rb") as f:
            pdf_data = f.read()
            pdf_attachment = MIMEApplication(pdf_data, _subtype="pdf")
            pdf_attachment.add_header("Content-Disposition", "attachment", filename=pdf_filename.name)
            msg.attach(pdf_attachment)

    if jpg_filename and jpg_filename.exists():
        with open(jpg_filename, "rb") as f:
            jpg_data = f.read()
            jpg_attachment = MIMEApplication(jpg_data, _subtype="jpeg")
            jpg_attachment.add_header("Content-Disposition", "attachment", filename=jpg_filename.name)
            msg.attach(jpg_attachment)

    return msg


def build_text_message(subject: str, body: str, from_email: str, to_email: str) -> MIMEMultipart:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    if from_email:
        msg["From"] = from_email
    msg["To"] = to_email
    msg.attach(MIMEText(body, "plain", "utf-8"))
    return msg


def send_batch_notification(credentials: Credentials, from_email: str, total_sent: int) -> None:
    subject = "Batch send report"
    body = f"From email: {from_email}\nTotal sent: {total_sent}\n"
    msg = build_text_message(subject, body, from_email, ADMIN_NOTIFY_EMAIL)
    send_email_gmail(msg, credentials)


def save_credentials(credentials: Credentials) -> None:
    TOKEN_DIR.mkdir(parents=True, exist_ok=True)
    TOKEN_FILE.write_text(credentials.to_json(), encoding="utf-8")


def load_credentials() -> Credentials | None:
    if not TOKEN_FILE.exists():
        return None
    credentials = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if credentials.expired and credentials.refresh_token:
        try:
            credentials.refresh(GoogleRequest())
            save_credentials(credentials)
        except Exception:
            return None
    return credentials if credentials.valid else None


def get_client_secret_path() -> Path:
    env_path = os.environ.get(CLIENT_SECRETS_ENV, "").strip()
    if env_path:
        return Path(env_path).expanduser()
    return CLIENT_SECRETS


def get_flow(redirect_uri: str) -> Flow:
    client_secret_path = get_client_secret_path()
    if not client_secret_path.exists():
        raise FileNotFoundError(
            "Missing client_secret.json. Provide webappGamilAPI/client_secret.json or set GMAIL_CLIENT_SECRET_PATH."
        )
    return Flow.from_client_secrets_file(
        str(client_secret_path),
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )


def send_email_gmail(
    message: MIMEMultipart,
    credentials: Credentials,
) -> None:
    service = build("gmail", "v1", credentials=credentials, cache_discovery=False)
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    service.users().messages().send(userId="me", body={"raw": raw}).execute()


def get_gmail_address(credentials: Credentials) -> str:
    service = build("gmail", "v1", credentials=credentials, cache_discovery=False)
    profile = service.users().getProfile(userId="me").execute()
    return profile.get("emailAddress", "")


def text_to_html(text: str) -> str:
    cleaned = text.replace("\r\n", "\n").strip()
    if not cleaned:
        return ""
    paragraphs = re.split(r"\n\\s*\n", cleaned)
    rendered = []
    for para in paragraphs:
        lines = "<br>".join(html_lib.escape(line) for line in para.split("\n"))
        rendered.append(f"<p style=\"line-height: 1.6;\">{lines}</p>")
    body = "".join(rendered)
    return (
        '<!DOCTYPE html><html lang="he"><head><meta charset="UTF-8"></head>'
        '<body style="direction: rtl; text-align: right; font-family: Arial, sans-serif; color: #333; margin: 20px;">'
        f"{body}</body></html>"
    )


def inject_logo_cid(html_content: str, cid: str = "logo_cid") -> str:
    if f"cid:{cid}" in html_content:
        return html_content
    img_tag = f'<img src="cid:{cid}" style="max-width: 220px; height: auto; display: block; margin-bottom: 12px;" />'
    match = re.search(r"</body>", html_content, re.IGNORECASE)
    if match:
        insert_at = match.start()
        return html_content[:insert_at] + img_tag + html_content[insert_at:]
    return html_content + img_tag


def decode_csv_content(content: bytes) -> str:
    for encoding in ("utf-8-sig", "ascii", "latin-1"):
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    return content.decode("utf-8", errors="replace")


def normalize_student_row(row: Dict[str, str], fieldnames: List[str]) -> Dict[str, str]:
    name_field = (row.get("name") or "").strip()
    first = (row.get("first") or "").strip()
    last = (row.get("last") or "").strip()
    email = (row.get("email") or "").strip() or (row.get("mail") or "").strip()
    new = (row.get("new") or "").strip()
    send = (row.get("send") or "").strip()

    has_name_field = "name" in fieldnames
    full_name = name_field if has_name_field and name_field else f"{first} {last}".strip()

    return {
        "first": first,
        "last": last,
        "email": email,
        "new": new,
        "send": send,
        "name": full_name,
    }


def parse_csv(content: bytes) -> List[Dict[str, str]]:
    decoded = decode_csv_content(content)
    reader = csv.DictReader(io.StringIO(decoded, newline=""))
    rows = []
    fieldnames = reader.fieldnames or []
    for row in reader:
        rows.append(normalize_student_row(row, fieldnames))
    return rows


def parse_csv_with_headers(content: bytes) -> Tuple[List[str], List[Dict[str, str]]]:
    decoded = decode_csv_content(content)
    reader = csv.DictReader(io.StringIO(decoded, newline=""))
    fieldnames = reader.fieldnames or []
    if "new" not in fieldnames:
        fieldnames.append("new")
    if "send" not in fieldnames:
        fieldnames.append("send")
    rows = []
    for row in reader:
        raw = {key: (row.get(key) or "").strip() for key in fieldnames}
        normalized = normalize_student_row(raw, fieldnames)
        rows.append({
            "raw": raw,
            "name": normalized["name"],
            "email": normalized["email"],
            "new": normalized["new"],
            "send": normalized["send"],
        })
    return fieldnames, rows

def save_upload(file: UploadFile, dest_dir: Path) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    file_path = dest_dir / file.filename
    content = file.file.read()
    file_path.write_bytes(content)
    return file_path


def save_bytes(filename: str, data: bytes, dest_dir: Path) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    safe_name = Path(filename).name or "upload.bin"
    file_path = dest_dir / safe_name
    file_path.write_bytes(data)
    return file_path


@app.get("/oauth/start")
def oauth_start(request: Request):
    redirect_uri = os.environ.get("GMAIL_OAUTH_REDIRECT_URI") or f"{request.base_url}oauth/callback"
    try:
        flow = get_flow(redirect_uri)
    except FileNotFoundError as exc:
        return HTMLResponse(str(exc), status_code=400)
    auth_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    global OAUTH_STATE, OAUTH_REDIRECT_URI
    OAUTH_STATE = state
    OAUTH_REDIRECT_URI = redirect_uri
    return RedirectResponse(auth_url)


@app.get("/oauth/callback", response_class=HTMLResponse)
def oauth_callback(code: str = "", state: str = ""):
    if not code:
        return HTMLResponse("Missing OAuth code.", status_code=400)
    if OAUTH_STATE and state and state != OAUTH_STATE:
        return HTMLResponse("Invalid OAuth state. Please try again.", status_code=400)

    redirect_uri = OAUTH_REDIRECT_URI or os.environ.get("GMAIL_OAUTH_REDIRECT_URI") or ""
    if not redirect_uri:
        return HTMLResponse("Missing redirect URI. Set GMAIL_OAUTH_REDIRECT_URI.", status_code=400)

    try:
        flow = get_flow(redirect_uri)
    except FileNotFoundError as exc:
        return HTMLResponse(str(exc), status_code=400)

    flow.fetch_token(code=code)
    save_credentials(flow.credentials)
    return HTMLResponse("Gmail connected. You can close this tab.", status_code=200)


@app.get("/oauth/status")
def oauth_status():
    credentials = load_credentials()
    email = ""
    if credentials:
        try:
            email = get_gmail_address(credentials)
        except Exception:
            email = ""
    return {"ok": bool(credentials), "email": email}


@app.get("/oauth/check-setup")
def oauth_check_setup():
    client_secret_path = get_client_secret_path()
    return {
        "ok": client_secret_path.exists(),
        "client_secret_path": str(client_secret_path),
    }


@app.post("/oauth/logout")
def oauth_logout():
    if TOKEN_FILE.exists():
        TOKEN_FILE.unlink()
    return {"ok": True}


@app.post("/api/test-send")
def test_send(
    pdf_template: Optional[UploadFile] = File(None),
    jpg_template: Optional[UploadFile] = File(None),
    logo_file: Optional[UploadFile] = File(None),
    html_content: str = Form(""),
    text_content: str = Form(""),
    letter_format: str = Form("html"),
    subject: str = Form(""),
    test_name: str = Form("אמיר"),
    test_email: str = Form(""),
    from_email: str = Form(""),
    name_x_offset: int = Form(0),
    name_y_offset: int = Form(0),
):
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = OUTPUT_DIR / f"test-{run_id}"

    pdf_path = None
    if pdf_template and pdf_template.filename:
        pdf_path = save_upload(pdf_template, run_dir)
    jpg_template_path = None
    if jpg_template and jpg_template.filename:
        jpg_template_path = save_upload(jpg_template, run_dir)

    credentials = load_credentials()
    if not credentials:
        return JSONResponse({"ok": False, "error": "Not connected to Gmail. Click Connect Google first."}, status_code=400)
    if not from_email:
        try:
            from_email = get_gmail_address(credentials)
        except Exception:
            from_email = ""

    if not test_email:
        return JSONResponse({"ok": False, "error": "Missing test email."}, status_code=400)
    if not from_email:
        return JSONResponse({"ok": False, "error": "Missing from email."}, status_code=400)

    if letter_format == "text":
        html_content = text_to_html(text_content)
    if not html_content:
        return JSONResponse({"ok": False, "error": "Letter content is empty."}, status_code=400)

    pdf_output = None
    if pdf_path:
        pdf_output = generate_diploma_pdf(
            pdf_path,
            test_name,
            run_dir,
            name_x_offset=name_x_offset,
            name_y_offset=name_y_offset,
        )
    jpg_output = None
    if jpg_template_path:
        jpg_output = generate_diploma_jpg(
            jpg_template_path,
            test_name,
            run_dir,
            name_x_offset=name_x_offset,
            name_y_offset=name_y_offset,
        )

    logo_bytes = None
    logo_filename = None
    if logo_file and logo_file.filename:
        logo_bytes = logo_file.file.read()
        logo_filename = Path(logo_file.filename).name

    if logo_bytes:
        html_content = inject_logo_cid(html_content)

    message = build_message(
        student_name=test_name,
        student_email=test_email,
        html_content=html_content,
        subject=subject or DEFAULT_SUBJECT,
        from_email=from_email,
        pdf_filename=pdf_output,
        jpg_filename=jpg_output,
        logo_bytes=logo_bytes,
        logo_filename=logo_filename,
    )
    send_email_gmail(message, credentials)

    return {"ok": True, "message": f"Test email sent to {test_email}", "output_dir": str(run_dir)}


@app.post("/api/preview")
def preview_csv(csv_file: UploadFile = File(...)):
    csv_bytes = csv_file.file.read()
    headers, students = parse_csv_with_headers(csv_bytes)
    preview = []
    for idx, student in enumerate(students):
        eligible = student["new"] == "1" and student["send"] == ""
        preview.append({
            "index": idx,
            "name": student["name"],
            "email": student["email"],
            "new": student["new"],
            "send": student["send"],
            "eligible": eligible,
            "raw": student["raw"],
        })
    return {"ok": True, "headers": headers, "rows": preview}


@app.post("/api/preview-pdf")
def preview_pdf(
    pdf_template: Optional[UploadFile] = File(None),
    jpg_template: Optional[UploadFile] = File(None),
    test_name: str = Form(""),
    name_x_offset: int = Form(0),
    name_y_offset: int = Form(0),
):
    if not test_name.strip():
        return JSONResponse({"ok": False, "error": "Missing test name."}, status_code=400)

    run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = OUTPUT_DIR / f"preview-{run_id}"

    pdf_path = None
    if pdf_template and pdf_template.filename:
        pdf_path = save_upload(pdf_template, run_dir)
    jpg_template_path = None
    jpg_received = bool(jpg_template and jpg_template.filename)
    if jpg_received:
        jpg_template_path = save_upload(jpg_template, run_dir)

    pdf_url = None
    if pdf_path:
        pdf_output = generate_diploma_pdf(
            pdf_path,
            test_name.strip(),
            run_dir,
            name_x_offset=name_x_offset,
            name_y_offset=name_y_offset,
        )
        pdf_url = f"/output/{pdf_output.parent.name}/{pdf_output.name}"
    jpg_url = None
    if jpg_template_path:
        jpg_output = generate_diploma_jpg(
            jpg_template_path,
            test_name.strip(),
            run_dir,
            name_x_offset=name_x_offset,
            name_y_offset=name_y_offset,
        )
        jpg_url = f"/output/{jpg_output.parent.name}/{jpg_output.name}"
    return {
        "ok": True,
        "pdf_url": pdf_url,
        "jpg_url": jpg_url,
        "jpg_received": jpg_received,
    }


@app.post("/api/send")
def send_batch(
    csv_file: UploadFile = File(...),
    pdf_template: Optional[UploadFile] = File(None),
    jpg_template: Optional[UploadFile] = File(None),
    logo_file: Optional[UploadFile] = File(None),
    html_content: str = Form(""),
    text_content: str = Form(""),
    letter_format: str = Form("html"),
    subject: str = Form(""),
    selected_indices: str = Form(""),
    from_email: str = Form(""),
    name_x_offset: int = Form(0),
    name_y_offset: int = Form(0),
):
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = OUTPUT_DIR / f"batch-{run_id}"

    if not pdf_template and not jpg_template:
        return JSONResponse({"ok": False, "error": "Missing PDF or JPG template."}, status_code=400)

    pdf_path = None
    if pdf_template and pdf_template.filename:
        pdf_path = save_upload(pdf_template, run_dir)
    jpg_template_path = None
    if jpg_template and jpg_template.filename:
        jpg_template_path = save_upload(jpg_template, run_dir)
    csv_bytes = csv_file.file.read()
    students = parse_csv(csv_bytes)

    credentials = load_credentials()
    if not credentials:
        return JSONResponse({"ok": False, "error": "Not connected to Gmail. Click Connect Google first."}, status_code=400)

    if not from_email:
        try:
            from_email = get_gmail_address(credentials)
        except Exception:
            from_email = ""

    if not from_email:
        return JSONResponse({"ok": False, "error": "Missing from email."}, status_code=400)

    if letter_format == "text":
        html_content = text_to_html(text_content)
    if not html_content:
        return JSONResponse({"ok": False, "error": "Letter content is empty."}, status_code=400)

    selected = set()
    if selected_indices.strip():
        for chunk in selected_indices.split(","):
            chunk = chunk.strip()
            if chunk.isdigit():
                selected.add(int(chunk))

    logo_bytes = None
    logo_filename = None
    if logo_file and logo_file.filename:
        logo_bytes = logo_file.file.read()
        logo_filename = Path(logo_file.filename).name

    if logo_bytes:
        html_content = inject_logo_cid(html_content)

    sent = []
    sent_indices = []
    skipped = []
    errors = []
    log_lines = []

    for idx, student in enumerate(students):
        if selected and idx not in selected:
            skipped.append({"name": student["name"], "reason": "Not selected"})
            if student["email"]:
                log_lines.append(f"skip:{student['email']}")
            continue
        if not student["email"]:
            skipped.append({"name": student["name"], "reason": "Missing email"})
            log_lines.append(f"skip:{student['name']}")
            continue
        if student["new"] != "1" or student["send"] != "":
            skipped.append({"name": student["name"], "reason": "Not eligible or already sent"})
            log_lines.append(f"skip:{student['email']}")
            continue

        try:
            pdf_output = None
            if pdf_path:
                pdf_output = generate_diploma_pdf(
                    pdf_path,
                    student["name"],
                    run_dir,
                    name_x_offset=name_x_offset,
                    name_y_offset=name_y_offset,
                )
            jpg_output = None
            if jpg_template_path:
                jpg_output = generate_diploma_jpg(
                    jpg_template_path,
                    student["name"],
                    run_dir,
                    name_x_offset=name_x_offset,
                    name_y_offset=name_y_offset,
                )
            message = build_message(
                student_name=student["name"],
                student_email=student["email"],
                html_content=html_content,
                subject=subject or DEFAULT_SUBJECT,
                from_email=from_email,
                pdf_filename=pdf_output,
                jpg_filename=jpg_output,
                logo_bytes=logo_bytes,
                logo_filename=logo_filename,
            )
            send_email_gmail(message, credentials)
            sent.append({"name": student["name"], "email": student["email"]})
            sent_indices.append(idx)
            log_lines.append(f"send:{student['email']}")
        except Exception as exc:  # noqa: BLE001
            errors.append({"name": student["name"], "error": str(exc)})

    try:
        send_batch_notification(credentials, from_email, len(sent))
    except Exception as exc:  # noqa: BLE001
        log_lines.append(f"notify_error:{str(exc)}")

    return {
        "ok": True,
        "output_dir": str(run_dir),
        "sent": sent,
        "sent_indices": sent_indices,
        "skipped": skipped,
        "errors": errors,
        "log_lines": log_lines,
        "summary": {
            "sent": len(sent),
            "skipped": len(skipped),
        },
    }


@app.post("/api/send-stream")
async def send_batch_stream(
    csv_file: UploadFile = File(...),
    pdf_template: Optional[UploadFile] = File(None),
    jpg_template: Optional[UploadFile] = File(None),
    logo_file: Optional[UploadFile] = File(None),
    html_content: str = Form(""),
    text_content: str = Form(""),
    letter_format: str = Form("html"),
    subject: str = Form(""),
    selected_indices: str = Form(""),
    from_email: str = Form(""),
    name_x_offset: int = Form(0),
    name_y_offset: int = Form(0),
):
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = OUTPUT_DIR / f"batch-{run_id}"

    csv_bytes = await csv_file.read()
    if not pdf_template and not jpg_template:
        return JSONResponse({"ok": False, "error": "Missing PDF or JPG template."}, status_code=400)

    pdf_bytes = await pdf_template.read() if pdf_template else None
    jpg_bytes = await jpg_template.read() if jpg_template and jpg_template.filename else None
    logo_bytes = await logo_file.read() if logo_file and logo_file.filename else None
    logo_filename = Path(logo_file.filename).name if logo_file and logo_file.filename else None

    credentials = load_credentials()
    if not credentials:
        return JSONResponse({"ok": False, "error": "Not connected to Gmail. Click Connect Google first."}, status_code=400)

    if not from_email:
        try:
            from_email = get_gmail_address(credentials)
        except Exception:
            from_email = ""

    if not from_email:
        return JSONResponse({"ok": False, "error": "Missing from email."}, status_code=400)

    if letter_format == "text":
        html_content = text_to_html(text_content)
    if not html_content:
        return JSONResponse({"ok": False, "error": "Letter content is empty."}, status_code=400)

    if logo_bytes:
        html_content = inject_logo_cid(html_content)

    pdf_path = None
    if pdf_template and pdf_bytes is not None:
        pdf_path = save_bytes(pdf_template.filename or "template.pdf", pdf_bytes, run_dir)
    jpg_template_path = None
    if jpg_bytes and jpg_template and jpg_template.filename:
        jpg_template_path = save_bytes(jpg_template.filename or "template.jpg", jpg_bytes, run_dir)
    students = parse_csv(csv_bytes)

    selected = set()
    if selected_indices.strip():
        for chunk in selected_indices.split(","):
            chunk = chunk.strip()
            if chunk.isdigit():
                selected.add(int(chunk))

    sent = []
    sent_indices = []
    skipped = []
    errors = []

    async def event_stream():
        for idx, student in enumerate(students):
            if selected and idx not in selected:
                skipped.append({"name": student["name"], "reason": "Not selected"})
                if student["email"]:
                    yield f"skip:{student['email']}\n"
                continue
            if not student["email"]:
                skipped.append({"name": student["name"], "reason": "Missing email"})
                yield f"skip:{student['name']}\n"
                continue
            if student["new"] != "1" or student["send"] != "":
                skipped.append({"name": student["name"], "reason": "Not eligible or already sent"})
                yield f"skip:{student['email']}\n"
                continue

            try:
                pdf_output = None
                if pdf_path:
                    pdf_output = generate_diploma_pdf(
                        pdf_path,
                        student["name"],
                        run_dir,
                        name_x_offset=name_x_offset,
                        name_y_offset=name_y_offset,
                    )
                jpg_output = None
                if jpg_template_path:
                    jpg_output = generate_diploma_jpg(
                        jpg_template_path,
                        student["name"],
                        run_dir,
                        name_x_offset=name_x_offset,
                        name_y_offset=name_y_offset,
                    )
                message = build_message(
                    student_name=student["name"],
                    student_email=student["email"],
                    html_content=html_content,
                    subject=subject or DEFAULT_SUBJECT,
                    from_email=from_email,
                    pdf_filename=pdf_output,
                    jpg_filename=jpg_output,
                    logo_bytes=logo_bytes,
                    logo_filename=logo_filename,
                )
                send_email_gmail(message, credentials)
                sent.append({"name": student["name"], "email": student["email"]})
                sent_indices.append(idx)
                yield f"send:{student['email']}\n"
            except Exception as exc:  # noqa: BLE001
                errors.append({"name": student["name"], "error": str(exc)})
                yield f"error:{student['email']}:{str(exc)}\n"

        summary = {"sent": len(sent), "skipped": len(skipped)}
        try:
            send_batch_notification(credentials, from_email, summary["sent"])
        except Exception as exc:  # noqa: BLE001
            yield f"notify_error:{str(exc)}\n"
        yield f"summary: sent={summary['sent']} skipped={summary['skipped']}\n"
        payload = {
            "sent_indices": sent_indices,
            "sent": sent,
            "skipped": skipped,
            "errors": errors,
            "summary": summary,
        }
        yield f"json:{json.dumps(payload)}\n"

    return StreamingResponse(event_stream(), media_type="text/plain")


@app.post("/api/save-csv")
def save_csv(
    csv_content: str = Form(...),
    filename: str = Form("updated_list.csv"),
    target_path: str = Form(""),
):
    if target_path:
        dest = Path(target_path)
    else:
        safe_name = Path(filename).name or "updated_list.csv"
        dest = OUTPUT_DIR / safe_name
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(csv_content, encoding="utf-8")
    return {"ok": True, "saved_to": str(dest)}
