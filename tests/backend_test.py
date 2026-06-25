"""Backend API tests for Maria Glass & Plywood."""
import os
import io
import uuid
import pytest
import requests

BASE_URL = os.environ.get("REACT_APP_BACKEND_URL", "https://maria-ply.preview.emergentagent.com").rstrip("/")
API = f"{BASE_URL}/api"

ADMIN_EMAIL = "admin@mariaglass.com"
ADMIN_PASSWORD = "Admin@123"

WORKER_EMAIL = f"TEST_worker_{uuid.uuid4().hex[:6]}@mariaglass.com"
WORKER_PASSWORD = "Worker@123"
WORKER_NAME = "TEST Worker"


@pytest.fixture(scope="session")
def admin_token():
    r = requests.post(f"{API}/auth/login", json={"email": ADMIN_EMAIL, "password": ADMIN_PASSWORD}, timeout=20)
    assert r.status_code == 200, f"Admin login failed: {r.status_code} {r.text}"
    data = r.json()
    assert "token" in data and data["user"]["role"] == "admin"
    return data["token"]


@pytest.fixture(scope="session")
def admin_headers(admin_token):
    return {"Authorization": f"Bearer {admin_token}"}


@pytest.fixture(scope="session")
def worker_info(admin_headers):
    # Create worker via admin
    r = requests.post(f"{API}/admin/workers", headers=admin_headers,
                      json={"name": WORKER_NAME, "email": WORKER_EMAIL, "password": WORKER_PASSWORD}, timeout=20)
    assert r.status_code == 200, f"Worker create failed: {r.status_code} {r.text}"
    worker = r.json()
    # Login as worker
    lr = requests.post(f"{API}/auth/login", json={"email": WORKER_EMAIL, "password": WORKER_PASSWORD}, timeout=20)
    assert lr.status_code == 200, f"Worker login failed: {lr.text}"
    token = lr.json()["token"]
    yield {"worker": worker, "token": token, "headers": {"Authorization": f"Bearer {token}"}}
    # Cleanup
    requests.delete(f"{API}/admin/workers/{worker['id']}", headers=admin_headers, timeout=20)


# --- Auth tests ---
class TestAuth:
    def test_admin_login_success(self, admin_token):
        assert isinstance(admin_token, str) and len(admin_token) > 20

    def test_admin_login_wrong_password(self):
        r = requests.post(f"{API}/auth/login", json={"email": ADMIN_EMAIL, "password": "WRONG"}, timeout=20)
        assert r.status_code == 401

    def test_login_nonexistent_user(self):
        r = requests.post(f"{API}/auth/login", json={"email": "nope@nope.com", "password": "x"}, timeout=20)
        assert r.status_code == 401

    def test_me_endpoint(self, admin_headers):
        r = requests.get(f"{API}/auth/me", headers=admin_headers, timeout=20)
        assert r.status_code == 200
        data = r.json()
        assert data["email"] == ADMIN_EMAIL
        assert data["role"] == "admin"

    def test_me_no_token(self):
        r = requests.get(f"{API}/auth/me", timeout=20)
        assert r.status_code == 401


# --- Admin worker management ---
class TestWorkerManagement:
    def test_create_worker_and_list(self, admin_headers, worker_info):
        worker = worker_info["worker"]
        assert worker["email"] == WORKER_EMAIL.lower()
        assert worker["role"] == "field_worker"
        assert "id" in worker
        # List
        r = requests.get(f"{API}/admin/workers", headers=admin_headers, timeout=20)
        assert r.status_code == 200
        ids = [w["id"] for w in r.json()]
        assert worker["id"] in ids

    def test_duplicate_worker_email(self, admin_headers, worker_info):
        r = requests.post(f"{API}/admin/workers", headers=admin_headers,
                          json={"name": "Dup", "email": WORKER_EMAIL, "password": "x123"}, timeout=20)
        assert r.status_code == 400

    def test_non_admin_cannot_list_workers(self, worker_info):
        r = requests.get(f"{API}/admin/workers", headers=worker_info["headers"], timeout=20)
        assert r.status_code == 403

    def test_non_admin_cannot_create_workers(self, worker_info):
        r = requests.post(f"{API}/admin/workers", headers=worker_info["headers"],
                          json={"name": "X", "email": "TEST_x@x.com", "password": "x123"}, timeout=20)
        assert r.status_code == 403


# --- Worker auth + field submit ---
class TestFieldSubmit:
    def test_worker_login_role(self, worker_info):
        # Already logged in - validate role via /me
        r = requests.get(f"{API}/auth/me", headers=worker_info["headers"], timeout=20)
        assert r.status_code == 200
        assert r.json()["role"] == "field_worker"

    def test_submit_without_photo(self, worker_info, admin_headers):
        data = {
            "client_name": "TEST Client A",
            "client_company": "TEST Co",
            "client_mobile": "9999900000",
            "client_email": "test_a@x.com",
            "site_address": "TEST Site Nagercoil",
            "latitude": "8.1780",
            "longitude": "77.4326",
            "notes": "Initial visit",
        }
        r = requests.post(f"{API}/field/submit", headers=worker_info["headers"], data=data, timeout=60)
        assert r.status_code == 200, r.text
        body = r.json()
        assert "id" in body and "email_sent" in body
        # Verify in admin submissions
        sub_id = body["id"]
        ar = requests.get(f"{API}/admin/submissions", headers=admin_headers, timeout=20)
        assert ar.status_code == 200
        match = [s for s in ar.json() if s["id"] == sub_id]
        assert len(match) == 1
        s = match[0]
        assert s["client_name"] == "TEST Client A"
        assert s["client_mobile"] == "9999900000"
        assert s["geo"]["latitude"] == 8.1780
        assert s["geo"]["longitude"] == 77.4326
        assert s["worker_name"] == WORKER_NAME

    def test_submit_with_photo(self, worker_info):
        # Minimal 1x1 JPEG bytes (valid header sufficient)
        fake_jpg = (b"\xff\xd8\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00"
                    b"\xff\xdb\x00C\x00" + b"\x08" * 64 + b"\xff\xd9")
        data = {
            "client_name": "TEST Client B",
            "client_company": "TEST Photo Co",
            "client_mobile": "8888800000",
            "client_email": "",
            "site_address": "TEST Photo Site",
            "latitude": "",
            "longitude": "",
            "notes": "with photo",
        }
        files = {"photo": ("site.jpg", io.BytesIO(fake_jpg), "image/jpeg")}
        r = requests.post(f"{API}/field/submit", headers=worker_info["headers"], data=data, files=files, timeout=60)
        assert r.status_code == 200, r.text
        body = r.json()
        assert "id" in body
        # email_sent may be True or False; either is OK

    def test_admin_cannot_submit(self, admin_headers):
        data = {"client_name": "X", "client_mobile": "0"}
        r = requests.post(f"{API}/field/submit", headers=admin_headers, data=data, timeout=20)
        assert r.status_code == 403


# --- Contact form ---
class TestContact:
    def test_contact_submit(self):
        r = requests.post(f"{API}/contact", json={
            "name": "TEST Contact",
            "email": "test_contact@x.com",
            "phone": "9000000000",
            "subject": "Enquiry",
            "message": "Hello, this is a test enquiry."
        }, timeout=60)
        assert r.status_code == 200, r.text
        body = r.json()
        assert "ok" in body  # may be True or False depending on SMTP


# --- Worker deletion ---
class TestWorkerDelete:
    def test_delete_worker(self, admin_headers):
        # Create a temp worker just to delete
        email = f"TEST_del_{uuid.uuid4().hex[:6]}@x.com"
        cr = requests.post(f"{API}/admin/workers", headers=admin_headers,
                           json={"name": "TEST Del", "email": email, "password": "abc1"}, timeout=20)
        assert cr.status_code == 200
        wid = cr.json()["id"]
        dr = requests.delete(f"{API}/admin/workers/{wid}", headers=admin_headers, timeout=20)
        assert dr.status_code == 200
        assert dr.json().get("ok") is True
        # Confirm not in list
        lr = requests.get(f"{API}/admin/workers", headers=admin_headers, timeout=20)
        ids = [w["id"] for w in lr.json()]
        assert wid not in ids

    def test_delete_nonexistent_worker(self, admin_headers):
        r = requests.delete(f"{API}/admin/workers/{uuid.uuid4()}", headers=admin_headers, timeout=20)
        assert r.status_code == 404
