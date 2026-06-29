from dotenv import load_dotenv
from pathlib import Path

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

import os
import uuid
import math
import logging
import bcrypt
import jwt
import smtplib
from datetime import datetime, timezone, timedelta
from email.message import EmailMessage
from typing import Optional, List

from fastapi import FastAPI, APIRouter, HTTPException, Depends, UploadFile, File, Form, Request, BackgroundTasks
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
def send_submission_email(submission: dict, photos: Optional[List[tuple]] = None, timeline: Optional[List[dict]] = None) -> bool:
    """photos: list of (filename, bytes) tuples. timeline: full visit history (oldest first)."""
    host = os.environ["SMTP_HOST"]
    port = int(os.environ["SMTP_PORT"])
    username = os.environ["SMTP_USERNAME"]
    password = os.environ["SMTP_APP_PASSWORD"]
    to_email = os.environ["COMPANY_EMAIL"]
    company = os.environ.get("COMPANY_NAME", "Maria Glass & Plywood")

    visit_num = submission.get("visit_number", 1)
    status = submission.get("status", "Site Visited")
    visit_label = f"Visit #{visit_num}" if visit_num > 1 else "Initial Visit"
    subject_prefix = f"Follow-up #{visit_num}" if visit_num > 1 else "New Field Report"

    msg = EmailMessage()
    msg["Subject"] = f"[{company}] {subject_prefix} — {submission['client_name']} · Status: {status}"
    msg["From"] = username
    msg["To"] = to_email

    g = submission.get("geo") or {}
    maps_link = ""
    if g.get("latitude") is not None and g.get("longitude") is not None:
        maps_link = f"https://www.google.com/maps?q={g['latitude']},{g['longitude']}"

    photos = photos or []
    photo_count_label = f"{len(photos)} photo(s) attached" if photos else "No photos attached"

    timeline_html = ""
    if timeline and len(timeline) > 1:
        rows = ""
        for v in timeline:
            vg = v.get("geo") or {}
            vmap = f"<a href='https://www.google.com/maps?q={vg.get('latitude','')},{vg.get('longitude','')}'>map</a>" if vg.get("latitude") else "-"
            rows += (
                f"<tr>"
                f"<td style='padding:6px;border:1px solid #e5e5e5'>#{v.get('visit_number', 1)}</td>"
                f"<td style='padding:6px;border:1px solid #e5e5e5'>{(v.get('created_at') or '')[:19].replace('T', ' ')}</td>"
                f"<td style='padding:6px;border:1px solid #e5e5e5'>{v.get('status', '-')}</td>"
                f"<td style='padding:6px;border:1px solid #e5e5e5'>{v.get('worker_name', '-')}</td>"
                f"<td style='padding:6px;border:1px solid #e5e5e5'>{v.get('photo_count', 0)} photo(s)</td>"
                f"<td style='padding:6px;border:1px solid #e5e5e5'>{vmap}</td>"
                f"</tr>"
            )
        timeline_html = f"""
        <h3 style="margin-top:24px">Client Visit Timeline ({len(timeline)} visits)</h3>
        <table style="border-collapse:collapse;width:100%;font-size:13px">
          <thead style="background:#f5f5f4">
            <tr>
              <th style='padding:6px;border:1px solid #e5e5e5;text-align:left'>Visit</th>
              <th style='padding:6px;border:1px solid #e5e5e5;text-align:left'>When</th>
              <th style='padding:6px;border:1px solid #e5e5e5;text-align:left'>Status</th>
              <th style='padding:6px;border:1px solid #e5e5e5;text-align:left'>Worker</th>
              <th style='padding:6px;border:1px solid #e5e5e5;text-align:left'>Photos</th>
              <th style='padding:6px;border:1px solid #e5e5e5;text-align:left'>Map</th>
            </tr>
          </thead>
          <tbody>{rows}</tbody>
        </table>
        """

    html = f"""
    <div style="font-family:Arial,sans-serif;max-width:680px;margin:auto">
      <div style="background:#15803D;color:#fff;padding:14px 18px;border-radius:8px 8px 0 0">
        <div style="font-size:11px;letter-spacing:0.2em;text-transform:uppercase;opacity:0.85">{visit_label} · {status}</div>
        <h2 style="margin:4px 0 0">{submission['client_name']}</h2>
      </div>
      <div style="border:1px solid #e5e5e5;border-top:0;padding:18px;border-radius:0 0 8px 8px">
        <p><strong>Submitted by:</strong> {submission.get('worker_name','')} ({submission.get('worker_email','')})</p>
        <p><strong>Submitted at:</strong> {submission.get('created_at','')}</p>
        <h3>Client Details</h3>
        <table style="border-collapse:collapse;width:100%">
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Client Name</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission['client_name']}</td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Role</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission.get('client_role','') or '-'}</td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Company</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission.get('client_company','')}</td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Mobile</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission['client_mobile']}</td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Email</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission.get('client_email','')}</td></tr>
        </table>
        <h3>This Visit</h3>
        <table style="border-collapse:collapse;width:100%">
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Visit Number</strong></td><td style="padding:6px;border:1px solid #e5e5e5">#{visit_num}</td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Status</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{status}</td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Location</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission.get('location','') or '-'}</td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Site Address</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission.get('site_address','')}</td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>GPS Coordinates</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{g.get('latitude','-')}, {g.get('longitude','-')}</td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Google Maps</strong></td><td style="padding:6px;border:1px solid #e5e5e5"><a href="{maps_link}">{maps_link}</a></td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Notes</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{submission.get('notes','')}</td></tr>
          <tr><td style="padding:6px;border:1px solid #e5e5e5"><strong>Photos</strong></td><td style="padding:6px;border:1px solid #e5e5e5">{photo_count_label}</td></tr>
        </table>
        {timeline_html}
        <p style="margin-top:24px;color:#57534E;font-size:12px">Auto-sent from {company} field portal.</p>
      </div>
    </div>
    """
    msg.set_content(f"{subject_prefix} — {submission['client_name']} · Status: {status}. View HTML email for details.")
    msg.add_alternative(html, subtype="html")

    for fname, fbytes in photos:
        if not fbytes:
            continue
        msg.add_attachment(fbytes, maintype="image", subtype="jpeg", filename=fname or "site-photo.jpg")

    try:
        with smtplib.SMTP(host, port, timeout=30) as server:
            server.starttls()
            server.login(username, password)
            server.send_message(msg)
        return True
    except Exception as e:
        logger.exception(f"Email sending failed: {e}")
        return False

# ---------- Models ----------
ALLOWED_STATUSES = [
    "Site Visited",
    "Materials Delivered",
    "Work in Progress",
    "Completed",
    "On Hold",
    "Cancelled",
]
CLIENT_ROLES = ["Engineer", "Plumber", "Electrician", "Mastri", "Other"]
LOCATIONS = ["Nagercoil", "Monday Market", "Valliyoor", "Thisayanvilai"]
FOLLOWUP_RADIUS_METERS = int(os.environ.get("FOLLOWUP_RADIUS_METERS", "300"))

def haversine_m(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Great-circle distance between two GPS points in metres."""
    r = 6371000.0
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return r * c

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
    ("Visit #", "visit_number"),
    ("Status", "status"),
    ("Worker Name", "worker_name"),
    ("Worker Email", "worker_email"),
    ("Client Name", "client_name"),
    ("Client Role", "client_role"),
    ("Client Company", "client_company"),
    ("Client Mobile", "client_mobile"),
    ("Client Email", "client_email"),
    ("Location", "location"),
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
            ("Role", "client_role"),
            ("Location", "location"),
            ("Company", "client_company"),
            ("Mobile", "client_mobile"),
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

        table = Table(data, repeatRows=1, colWidths=[60, 80, 55, 70, 70, 65, 110, 65, 75])
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

# ---------- Field Worker: helpers ----------
async def _read_photos(uploads) -> List[tuple]:
    photos = []
    for idx, ph in enumerate(uploads, start=1):
        if ph is None:
            continue
        data = await ph.read()
        if not data:
            continue
        photos.append((ph.filename or f"site-photo-{idx}.jpg", data))
    return photos

def _parse_geo(latitude, longitude) -> dict:
    geo = {}
    try:
        if latitude not in (None, ""):
            geo["latitude"] = float(latitude)
        if longitude not in (None, ""):
            geo["longitude"] = float(longitude)
    except ValueError:
        return {}
    return geo

async def _build_timeline(original_id: str) -> List[dict]:
    visits = await db.submissions.find(
        {"$or": [{"id": original_id}, {"original_submission_id": original_id}]},
        {"_id": 0, "photo_data": 0},
    ).sort("visit_number", 1).to_list(1000)
    return visits

def _send_and_flag(sub_id: str, submission: dict, photos: List[tuple], timeline: Optional[List[dict]] = None):
    """Sync function safe to run as a FastAPI BackgroundTask (runs in a worker thread)."""
    try:
        ok = send_submission_email(submission, photos=photos, timeline=timeline)
    except Exception as e:
        logger.exception(f"Background email crashed: {e}")
        ok = False
    # Update the DB flag synchronously using a fresh client (motor must run in the event loop, so use pymongo here).
    try:
        from pymongo import MongoClient
        sync_client = MongoClient(os.environ["MONGO_URL"])
        sync_client[os.environ["DB_NAME"]].submissions.update_one({"id": sub_id}, {"$set": {"email_sent": ok}})
        sync_client.close()
    except Exception as e:
        logger.exception(f"Failed to update email_sent flag: {e}")

# ---------- Field Worker: Submit (initial visit) ----------
@api_router.post("/field/submit")
async def submit_field_report(
    background_tasks: BackgroundTasks,
    client_name: str = Form(...),
    client_company: str = Form(""),
    client_mobile: str = Form(...),
    client_email: str = Form(""),
    client_role: str = Form(""),
    location: str = Form(""),
    site_address: str = Form(""),
    latitude: Optional[str] = Form(None),
    longitude: Optional[str] = Form(None),
    notes: str = Form(""),
    status: str = Form("Site Visited"),
    photo1: Optional[UploadFile] = File(None),
    photo2: Optional[UploadFile] = File(None),
    photo3: Optional[UploadFile] = File(None),
    user: dict = Depends(get_current_user),
):
    if user["role"] != "field_worker":
        raise HTTPException(status_code=403, detail="Only field workers can submit reports")

    if status not in ALLOWED_STATUSES:
        status = "Site Visited"
    # Allow free-text roles (when worker picked "Other" and typed a custom value).
    # Cap length to keep storage / exports clean.
    client_role = (client_role or "").strip()[:50]
    location = (location or "").strip()[:50]

    photos = await _read_photos([photo1, photo2, photo3])
    sub_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc).isoformat()
    geo = _parse_geo(latitude, longitude)

    submission = {
        "id": sub_id,
        "original_submission_id": sub_id,
        "parent_submission_id": None,
        "visit_number": 1,
        "status": status,
        "worker_id": user["id"],
        "worker_name": user["name"],
        "worker_email": user["email"],
        "client_name": client_name.strip(),
        "client_company": client_company.strip(),
        "client_mobile": client_mobile.strip(),
        "client_email": client_email.strip(),
        "client_role": client_role,
        "location": location,
        "site_address": site_address.strip(),
        "geo": geo,
        "notes": notes.strip(),
        "photo_filenames": [p[0] for p in photos],
        "photo_count": len(photos),
        "created_at": now,
    }
    await db.submissions.insert_one({**submission})
    background_tasks.add_task(_send_and_flag, sub_id, submission, photos)

    return {"id": sub_id, "visit_number": 1, "email_queued": True, "photo_count": len(photos), "message": "Report submitted successfully"}

# ---------- Field Worker: Follow-up visit ----------
@api_router.post("/field/follow-up/{submission_id}")
async def submit_follow_up(
    submission_id: str,
    background_tasks: BackgroundTasks,
    status: str = Form(...),
    notes: str = Form(""),
    latitude: Optional[str] = Form(None),
    longitude: Optional[str] = Form(None),
    photo1: Optional[UploadFile] = File(None),
    photo2: Optional[UploadFile] = File(None),
    photo3: Optional[UploadFile] = File(None),
    user: dict = Depends(get_current_user),
):
    if user["role"] != "field_worker":
        raise HTTPException(status_code=403, detail="Only field workers can submit follow-ups")
    if status not in ALLOWED_STATUSES:
        raise HTTPException(status_code=400, detail=f"status must be one of: {', '.join(ALLOWED_STATUSES)}")

    parent = await db.submissions.find_one({"id": submission_id}, {"_id": 0, "photo_data": 0})
    if not parent:
        raise HTTPException(status_code=404, detail="Original submission not found")

    original_id = parent.get("original_submission_id") or parent["id"]
    original = await db.submissions.find_one({"id": original_id}, {"_id": 0, "photo_data": 0})
    if not original:
        original = parent

    async def _log_rejected(reason: str, attempted_geo: dict, distance_m):
        await db.rejected_attempts.insert_one({
            "id": str(uuid.uuid4()),
            "worker_id": user["id"],
            "worker_name": user["name"],
            "worker_email": user["email"],
            "parent_submission_id": submission_id,
            "original_submission_id": original_id,
            "client_name": original.get("client_name", ""),
            "client_mobile": original.get("client_mobile", ""),
            "original_geo": original.get("geo") or {},
            "attempted_geo": attempted_geo or {},
            "distance_m": round(distance_m, 1) if distance_m is not None else None,
            "radius_m": FOLLOWUP_RADIUS_METERS,
            "reason": reason,
            "attempted_status": status,
            "created_at": datetime.now(timezone.utc).isoformat(),
        })

    # GPS guard: must be within FOLLOWUP_RADIUS_METERS of the original visit (if original has GPS)
    new_geo = _parse_geo(latitude, longitude)
    orig_geo = original.get("geo") or {}
    distance_m = None
    if orig_geo.get("latitude") is not None and orig_geo.get("longitude") is not None:
        if not (new_geo.get("latitude") is not None and new_geo.get("longitude") is not None):
            await _log_rejected("missing_gps", new_geo, None)
            raise HTTPException(
                status_code=400,
                detail=f"GPS location is required for follow-up visits (must be within {FOLLOWUP_RADIUS_METERS}m of the original site).",
            )
        distance_m = haversine_m(
            orig_geo["latitude"], orig_geo["longitude"],
            new_geo["latitude"], new_geo["longitude"],
        )
        if distance_m > FOLLOWUP_RADIUS_METERS:
            await _log_rejected("out_of_range", new_geo, distance_m)
            raise HTTPException(
                status_code=400,
                detail=f"You are {int(distance_m)}m from the original site. Follow-ups must be within {FOLLOWUP_RADIUS_METERS}m.",
            )

    # Next visit number
    last = await db.submissions.find(
        {"$or": [{"id": original_id}, {"original_submission_id": original_id}]},
        {"visit_number": 1, "_id": 0},
    ).sort("visit_number", -1).to_list(1)
    next_visit = (last[0]["visit_number"] if last else 1) + 1

    photos = await _read_photos([photo1, photo2, photo3])
    sub_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc).isoformat()

    submission = {
        "id": sub_id,
        "original_submission_id": original_id,
        "parent_submission_id": submission_id,
        "visit_number": next_visit,
        "status": status,
        "worker_id": user["id"],
        "worker_name": user["name"],
        "worker_email": user["email"],
        # client snapshot copied from the original visit (read-only on UI)
        "client_name": original["client_name"],
        "client_company": original.get("client_company", ""),
        "client_mobile": original["client_mobile"],
        "client_email": original.get("client_email", ""),
        "client_role": original.get("client_role", ""),
        "location": original.get("location", ""),
        "site_address": original.get("site_address", ""),
        "geo": new_geo or orig_geo,
        "distance_from_original_m": round(distance_m, 1) if distance_m is not None else None,
        "notes": notes.strip(),
        "photo_filenames": [p[0] for p in photos],
        "photo_count": len(photos),
        "created_at": now,
    }
    await db.submissions.insert_one({**submission})

    timeline = await _build_timeline(original_id)
    background_tasks.add_task(_send_and_flag, sub_id, submission, photos, timeline)

    return {
        "id": sub_id,
        "visit_number": next_visit,
        "email_queued": True,
        "photo_count": len(photos),
        "distance_from_original_m": submission["distance_from_original_m"],
        "message": "Follow-up submitted successfully",
    }

# ---------- Field Worker: My reports (root visits only, grouped) ----------
@api_router.get("/field/my-reports")
async def my_reports(user: dict = Depends(get_current_user)):
    if user["role"] != "field_worker":
        raise HTTPException(status_code=403, detail="Field worker access required")
    # All originals submitted by this worker OR any chain where this worker contributed
    own_originals = await db.submissions.find(
        {"worker_id": user["id"], "visit_number": 1},
        {"_id": 0, "photo_data": 0},
    ).sort("created_at", -1).to_list(500)
    # Attach last visit info per chain
    for o in own_originals:
        chain = await db.submissions.find(
            {"$or": [{"id": o["id"]}, {"original_submission_id": o["id"]}]},
            {"_id": 0, "visit_number": 1, "status": 1, "created_at": 1},
        ).sort("visit_number", -1).to_list(50)
        if chain:
            o["latest_visit_number"] = chain[0]["visit_number"]
            o["latest_status"] = chain[0].get("status", o.get("status"))
            o["latest_visit_at"] = chain[0]["created_at"]
            o["total_visits"] = len(chain)
        else:
            o["latest_visit_number"] = 1
            o["latest_status"] = o.get("status", "Site Visited")
            o["latest_visit_at"] = o["created_at"]
            o["total_visits"] = 1
    return own_originals

# ---------- Admin / Field: Timeline for a chain ----------
@api_router.get("/submissions/{submission_id}/timeline")
async def submission_timeline(submission_id: str, user: dict = Depends(get_current_user)):
    parent = await db.submissions.find_one({"id": submission_id}, {"_id": 0, "photo_data": 0})
    if not parent:
        raise HTTPException(status_code=404, detail="Submission not found")
    original_id = parent.get("original_submission_id") or parent["id"]
    # Field workers can only see their own chains
    if user["role"] == "field_worker":
        if parent.get("worker_id") != user["id"]:
            other = await db.submissions.find_one({"original_submission_id": original_id, "worker_id": user["id"]}, {"_id": 0})
            if not other:
                raise HTTPException(status_code=403, detail="Not your submission")
    timeline = await _build_timeline(original_id)
    return {"original_id": original_id, "visits": timeline, "total_visits": len(timeline)}

# ---------- Admin: Rejected follow-up attempts ----------
@api_router.get("/admin/rejected-attempts")
async def list_rejected(_: dict = Depends(require_admin)):
    items = await db.rejected_attempts.find({}, {"_id": 0}).sort("created_at", -1).to_list(1000)
    return items

# ---------- Statuses metadata ----------
@api_router.get("/field/statuses")
async def list_statuses():
    return {"statuses": ALLOWED_STATUSES, "client_roles": CLIENT_ROLES, "locations": LOCATIONS, "followup_radius_m": FOLLOWUP_RADIUS_METERS}

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
    await db.submissions.create_index("original_submission_id")
    await db.submissions.create_index("worker_id")
    await db.rejected_attempts.create_index("created_at")

    # Backfill old submissions to the new visit-tracking schema (idempotent)
    await db.submissions.update_many(
        {"visit_number": {"$exists": False}},
        {"$set": {"visit_number": 1, "status": "Site Visited", "parent_submission_id": None}},
    )
    async for doc in db.submissions.find({"original_submission_id": {"$exists": False}}, {"_id": 0, "id": 1}):
        await db.submissions.update_one({"id": doc["id"]}, {"$set": {"original_submission_id": doc["id"]}})

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