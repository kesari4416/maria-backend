from dotenv import load_dotenv
from pathlib import Path

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

import os
import uuid
import logging
import bcrypt
import jwt
import smtplib
from datetime import datetime, timezone, timedelta
from email.message import EmailMessage
from typing import Optional, List

from fastapi import FastAPI, APIRouter, HTTPException, Depends, UploadFile, File, Form, Request
from fastapi.responses import StreamingResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from starlette.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
from pydantic import BaseModel, EmailStr, Field
import io
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

# ---------- MongoDB ----------
mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

# ---------- Logging ----------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ---------- App ----------
app = FastAPI(title="Maria Glass & Plywood API")
api_router = APIRouter(prefix="/api")
security = HTTPBearer(auto_error=False)

JWT_ALGORITHM = "HS256"
JWT_EXPIRY_HOURS = 24

# ---------- Helpers ----------
def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

def verify_password(plain: str, hashed: str) -> bool:
    return bcrypt.checkpw(plain.encode("utf-8"), hashed.encode("utf-8"))

def create_token(user_id: str, email: str, role: str) -> str:
    payload = {
        "sub": user_id,
        "email": email,
        "role": role,
        "exp": datetime.now(timezone.utc) + timedelta(hours=JWT_EXPIRY_HOURS),
        "iat": datetime.now(timezone.utc),
    }
    return jwt.encode(payload, os.environ["JWT_SECRET"], algorithm=JWT_ALGORITHM)

async def get_current_user(credentials: Optional[HTTPAuthorizationCredentials] = Depends(security)) -> dict:
    if credentials is None:
        raise HTTPException(status_code=401, detail="Not authenticated")
    try:
        payload = jwt.decode(credentials.credentials, os.environ["JWT_SECRET"], algorithms=[JWT_ALGORITHM])
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token expired")
    except jwt.InvalidTokenError:
        raise HTTPException(status_code=401, detail="Invalid token")

    user = await db.users.find_one({"id": payload["sub"]})
    if not user:
        raise HTTPException(status_code=401, detail="User not found")
    return {"id": user["id"], "email": user["email"], "name": user.get("name", ""), "role": user["role"]}

def require_admin(user: dict = Depends(get_current_user)) -> dict:
    if user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return user

# ---------- Email ----------
def send_submission_email(submission: dict, photo_bytes: Optional[bytes] = None, photo_filename: Optional[str] = None) -> bool:
    host = os.environ["SMTP_HOST"]
    port = int(os.environ["SMTP_PORT"])
    username = os.environ["SMTP_USERNAME"]
    password = os.environ["SMTP_APP_PASSWORD"]
    to_email = os.environ["COMPANY_EMAIL"]
    company = os.environ.get("COMPANY_NAME", "Maria Glass & Plywood")

    msg = EmailMessage()
    msg["Subject"] = f"[{company}] New Field Report — {submission['client_name']}"
    msg["From"] = username
    msg["To"] = to_email

    g = submission.get("geo") or {}
    maps_link = ""
    if g.get("latitude") is not None and g.get("longitude") is not None:
        maps_link = f"https://www.google.com/maps?q={g['latitude']},{g['longitude']}"

    html = f"""
    <div style="font-family:Arial,sans-serif;max-width:640px;margin:auto">
      <h2 style="color:#15803D">New Field Worker Submission</h2>
      <p><strong>Submitted by:</strong> {submission.get('worker_name','')} ({submission.get('worker_email','')})</p>
      <p><strong>Submitted at:</strong> {submission.get('created_at','')}</p>
      <hr/>
      <h3>Client Details</h3>
      <table style="border-collapse:collapse;width:100%">
        <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Client Name</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission['client_name']}</td></tr>
        <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Company</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission.get('client_company','')}</td></tr>
        <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Mobile</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission['client_mobile']}</td></tr>
        <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Email</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission.get('client_email','')}</td></tr>
      </table>
      <h3>Construction Site</h3>
      <table style="border-collapse:collapse;width:100%">
        <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Address / Notes</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission.get('site_address','')}</td></tr>
        <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>GPS Coordinates</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{g.get('latitude','-')}, {g.get('longitude','-')}</td></tr>
        <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Google Maps</strong></td><td style="padding:6px;border:1px solid #e5e5e5"><a href="{maps_link}">{maps_link}</a></td></tr>
        <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Notes</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission.get('notes','')}</td></tr>
      </table>
      <p style="margin-top:24px;color:#57534E;font-size:12px">Auto-sent from {company} field portal.</p>
    </div>
    """
    msg.set_content("New Field Worker Submission. View HTML email for details.")
    msg.add_alternative(html, subtype="html")

    if photo_bytes and photo_filename:
        msg.add_attachment(photo_bytes, maintype="image", subtype="jpeg", filename=photo_filename)

    try:
        with smtplib.SMTP(host, port, timeout=20) as server:
            server.starttls()
            server.login(username, password)
            server.send_message(msg)
        return True
    except Exception as e:
        logger.exception(f"Email sending failed: {e}")
        return False

# ---------- Models ----------
class LoginRequest(BaseModel):
    email: EmailStr
    password: str

class CreateWorkerRequest(BaseModel):
    name: str
    email: EmailStr
    password: str = Field(min_length=4)

class ContactRequest(BaseModel):
    name: str
    email: EmailStr
    phone: Optional[str] = ""
    subject: Optional[str] = ""
    message: str

class UserOut(BaseModel):
    id: str
    name: str
    email: EmailStr
    role: str
    created_at: str

# ---------- Auth Endpoints ----------
@api_router.post("/auth/login")
async def login(payload: LoginRequest):
    email = payload.email.lower().strip()
    user = await db.users.find_one({"email": email})
    if not user or not verify_password(payload.password, user["password_hash"]):
        raise HTTPException(status_code=401, detail="Invalid email or password")
    token = create_token(user["id"], user["email"], user["role"])
    return {
        "token": token,
        "user": {"id": user["id"], "name": user.get("name", ""), "email": user["email"], "role": user["role"]},
    }

@api_router.get("/auth/me")
async def me(user: dict = Depends(get_current_user)):
    return user

# ---------- Admin: Field Worker Management ----------
@api_router.post("/admin/workers", response_model=UserOut)
async def create_worker(payload: CreateWorkerRequest, _: dict = Depends(require_admin)):
    email = payload.email.lower().strip()
    if await db.users.find_one({"email": email}):
        raise HTTPException(status_code=400, detail="Email already exists")
    uid = str(uuid.uuid4())
    now = datetime.now(timezone.utc).isoformat()
    doc = {
        "id": uid,
        "name": payload.name.strip(),
        "email": email,
        "password_hash": hash_password(payload.password),
        "role": "field_worker",
        "created_at": now,
    }
    await db.users.insert_one(doc)
    return UserOut(id=uid, name=doc["name"], email=email, role="field_worker", created_at=now)

@api_router.get("/admin/workers", response_model=List[UserOut])
async def list_workers(_: dict = Depends(require_admin)):
    workers = await db.users.find({"role": "field_worker"}, {"_id": 0, "password_hash": 0}).sort("created_at", -1).to_list(1000)
    return [UserOut(**w) for w in workers]

@api_router.delete("/admin/workers/{worker_id}")
async def delete_worker(worker_id: str, _: dict = Depends(require_admin)):
    res = await db.users.delete_one({"id": worker_id, "role": "field_worker"})
    if res.deleted_count == 0:
        raise HTTPException(status_code=404, detail="Worker not found")
    return {"ok": True}

# ---------- Admin: View Submissions ----------
@api_router.get("/admin/submissions")
async def list_submissions(_: dict = Depends(require_admin)):
    items = await db.submissions.find({}, {"_id": 0, "photo_data": 0}).sort("created_at", -1).to_list(1000)
    return items

# ---------- Admin: Export Submissions ----------
EXPORT_COLUMNS = [
    ("Submitted At", "created_at"),
    ("Worker Name", "worker_name"),
    ("Worker Email", "worker_email"),
    ("Client Name", "client_name"),
    ("Client Company", "client_company"),
    ("Client Mobile", "client_mobile"),
    ("Client Email", "client_email"),
    ("Site Address", "site_address"),
    ("Latitude", None),
    ("Longitude", None),
    ("Maps Link", None),
    ("Notes", "notes"),
    ("Email Sent", "email_sent"),
]

def _row_value(sub: dict, key: Optional[str], label: str) -> str:
    if key is not None:
        v = sub.get(key, "")
        return "" if v is None else str(v)
    g = sub.get("geo") or {}
    if label == "Latitude":
        return str(g.get("latitude", "") or "")
    if label == "Longitude":
        return str(g.get("longitude", "") or "")
    if label == "Maps Link":
        if g.get("latitude") and g.get("longitude"):
            return f"https://www.google.com/maps?q={g['latitude']},{g['longitude']}"
        return ""
    return ""

@api_router.get("/admin/submissions/export")
async def export_submissions(format: str = "excel", _: dict = Depends(require_admin)):
    items = await db.submissions.find({}, {"_id": 0, "photo_data": 0}).sort("created_at", -1).to_list(10000)
    fmt = format.lower()
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M")

    if fmt in ("excel", "xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Submissions"
        headers = [c[0] for c in EXPORT_COLUMNS]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True, color="FFFFFF")
            cell.fill = cell.fill.copy(fgColor="15803D", patternType="solid")
        for sub in items:
            ws.append([_row_value(sub, key, label) for label, key in EXPORT_COLUMNS])
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 50)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="maria_submissions_{stamp}.xlsx"'},
        )

    if fmt == "pdf":
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=28, bottomMargin=24)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("title", parent=styles["Title"], textColor=colors.HexColor("#15803D"), alignment=0)
        meta_style = ParagraphStyle("meta", parent=styles["Normal"], textColor=colors.HexColor("#57534E"), fontSize=9)
        cell_style = ParagraphStyle("cell", parent=styles["Normal"], fontSize=8, leading=10)
        pdf_columns = [
            ("Date", "created_at"),
            ("Client", "client_name"),
            ("Company", "client_company"),
            ("Mobile", "client_mobile"),
            ("Email", "client_email"),
            ("Site", "site_address"),
            ("GPS", None),
            ("Worker", "worker_name"),
        ]
        data = [[Paragraph(f"<b>{label}</b>", cell_style) for label, _ in pdf_columns]]
        for sub in items:
            row = []
            for label, key in pdf_columns:
                if label == "GPS":
                    g = sub.get("geo") or {}
                    val = f"{g.get('latitude','')}, {g.get('longitude','')}" if g.get("latitude") else "-"
                elif label == "Date":
                    val = (sub.get("created_at") or "")[:19].replace("T", " ")
                else:
                    val = str(sub.get(key, "") or "-")
                row.append(Paragraph(val, cell_style))
            data.append(row)

        table = Table(data, repeatRows=1, colWidths=[68, 90, 80, 70, 110, 140, 90, 90])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#15803D")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#E7E5E4")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFAF9")]),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story = [
            Paragraph("Maria Glass & Plywood — Field Submissions", title_style),
            Paragraph(f"Exported on {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')} · {len(items)} record(s)", meta_style),
            Spacer(1, 12),
            table,
        ]
        doc.build(story)
        buf.seek(0)
        return StreamingResponse(
            buf,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="maria_submissions_{stamp}.pdf"'},
        )

    raise HTTPException(status_code=400, detail="format must be 'excel' or 'pdf'")

# ---------- Field Worker: Submit ----------
@api_router.post("/field/submit")
async def submit_field_report(
    client_name: str = Form(...),
    client_company: str = Form(""),
    client_mobile: str = Form(...),
    client_email: str = Form(""),
    site_address: str = Form(""),
    latitude: Optional[str] = Form(None),
    longitude: Optional[str] = Form(None),
    notes: str = Form(""),
    photo: Optional[UploadFile] = File(None),
    user: dict = Depends(get_current_user),
):
    if user["role"] != "field_worker":
        raise HTTPException(status_code=403, detail="Only field workers can submit reports")

    photo_bytes = None
    photo_filename = None
    if photo is not None:
        photo_bytes = await photo.read()
        photo_filename = photo.filename or "site-photo.jpg"

    sub_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc).isoformat()
    geo = {}
    try:
        if latitude not in (None, ""):
            geo["latitude"] = float(latitude)
        if longitude not in (None, ""):
            geo["longitude"] = float(longitude)
    except ValueError:
        geo = {}

    submission = {
        "id": sub_id,
        "worker_id": user["id"],
        "worker_name": user["name"],
        "worker_email": user["email"],
        "client_name": client_name.strip(),
        "client_company": client_company.strip(),
        "client_mobile": client_mobile.strip(),
        "client_email": client_email.strip(),
        "site_address": site_address.strip(),
        "geo": geo,
        "notes": notes.strip(),
        "photo_filename": photo_filename,
        "created_at": now,
    }

    await db.submissions.insert_one({**submission})

    email_ok = send_submission_email(submission, photo_bytes=photo_bytes, photo_filename=photo_filename)
    await db.submissions.update_one({"id": sub_id}, {"$set": {"email_sent": email_ok}})

    return {"id": sub_id, "email_sent": email_ok, "message": "Report submitted successfully"}

# ---------- Public: Contact Form ----------
@api_router.post("/contact")
async def contact(payload: ContactRequest):
    host = os.environ["SMTP_HOST"]
    port = int(os.environ["SMTP_PORT"])
    username = os.environ["SMTP_USERNAME"]
    password = os.environ["SMTP_APP_PASSWORD"]
    to_email = os.environ["COMPANY_EMAIL"]
    company = os.environ.get("COMPANY_NAME", "Maria Glass & Plywood")

    msg = EmailMessage()
    msg["Subject"] = f"[{company} Website] {payload.subject or 'Contact Enquiry'}"
    msg["From"] = username
    msg["To"] = to_email
    msg["Reply-To"] = payload.email

    html = f"""
    <div style="font-family:Arial,sans-serif">
      <h3 style="color:#15803D">New website enquiry</h3>
      <p><strong>Name:</strong> {payload.name}</p>
      <p><strong>Email:</strong> {payload.email}</p>
      <p><strong>Phone:</strong> {payload.phone or '-'}</p>
      <p><strong>Subject:</strong> {payload.subject or '-'}</p>
      <p><strong>Message:</strong><br/>{payload.message}</p>
    </div>
    """
    msg.set_content(payload.message)
    msg.add_alternative(html, subtype="html")

    now = datetime.now(timezone.utc).isoformat()
    await db.contacts.insert_one({
        "id": str(uuid.uuid4()),
        "name": payload.name,
        "email": payload.email,
        "phone": payload.phone,
        "subject": payload.subject,
        "message": payload.message,
        "created_at": now,
    })

    try:
        with smtplib.SMTP(host, port, timeout=20) as server:
            server.starttls()
            server.login(username, password)
            server.send_message(msg)
        return {"ok": True}
    except Exception as e:
        logger.exception(f"Contact email failed: {e}")
        return {"ok": False, "message": "Stored, but email failed"}

@api_router.get("/")
async def root():
    return {"message": "Maria Glass & Plywood API"}

# ---------- Startup ----------
@app.on_event("startup")
async def startup_event():
    await db.users.create_index("email", unique=True)
    await db.submissions.create_index("created_at")
    admin_email = os.environ["ADMIN_EMAIL"].lower().strip()
    admin_password = os.environ["ADMIN_PASSWORD"]
    existing = await db.users.find_one({"email": admin_email})
    if not existing:
        await db.users.insert_one({
            "id": str(uuid.uuid4()),
            "name": "Admin",
            "email": admin_email,
            "password_hash": hash_password(admin_password),
            "role": "admin",
            "created_at": datetime.now(timezone.utc).isoformat(),
        })
        logger.info(f"Seeded admin: {admin_email}")
    else:
        if not verify_password(admin_password, existing["password_hash"]):
            await db.users.update_one({"email": admin_email}, {"$set": {"password_hash": hash_password(admin_password)}})
            logger.info("Admin password updated from env")

@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()

app.include_router(api_router)

app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=os.environ.get('CORS_ORIGINS', '*').split(','),
    allow_methods=["*"],
    allow_headers=["*"],
)
