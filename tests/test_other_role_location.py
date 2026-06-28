"""Tests for BUG: 'Other' role/location free-text submission.
Verifies backend accepts arbitrary free-text role/location, persists them,
and follow-ups inherit the original values.
"""
import os
import uuid
import time
import requests
import pytest

BASE_URL = os.environ["REACT_APP_BACKEND_URL"].rstrip("/")
API = f"{BASE_URL}/api"

ADMIN_EMAIL = "admin@mariaglass.com"
ADMIN_PASSWORD = "Admin@123"


@pytest.fixture(scope="module")
def admin_token():
    r = requests.post(f"{API}/auth/login", json={"email": ADMIN_EMAIL, "password": ADMIN_PASSWORD})
    assert r.status_code == 200, r.text
    return r.json()["token"]


@pytest.fixture(scope="module")
def worker(admin_token):
    """Create a fresh field worker for the test and clean up after."""
    suffix = uuid.uuid4().hex[:6]
    email = f"TEST_bug_{suffix}@mariaglass.com"
    password = "Worker@123"
    create = requests.post(
        f"{API}/admin/workers",
        json={"name": "TEST Bug Worker", "email": email, "password": password},
        headers={"Authorization": f"Bearer {admin_token}"},
    )
    assert create.status_code == 200, create.text
    worker_id = create.json()["id"]

    login = requests.post(f"{API}/auth/login", json={"email": email, "password": password})
    assert login.status_code == 200, login.text
    token = login.json()["token"]

    yield {"id": worker_id, "email": email, "token": token}

    # cleanup
    requests.delete(
        f"{API}/admin/workers/{worker_id}",
        headers={"Authorization": f"Bearer {admin_token}"},
    )


def _submit(token, **fields):
    data = {
        "client_name": fields.get("client_name", "TEST Bug Client"),
        "client_company": fields.get("client_company", ""),
        "client_mobile": fields.get("client_mobile", "9999999999"),
        "client_email": fields.get("client_email", ""),
        "client_role": fields.get("client_role", ""),
        "location": fields.get("location", ""),
        "site_address": fields.get("site_address", ""),
        "latitude": fields.get("latitude", ""),
        "longitude": fields.get("longitude", ""),
        "notes": fields.get("notes", ""),
        "status": fields.get("status", "Site Visited"),
    }
    return requests.post(
        f"{API}/field/submit",
        data=data,
        headers={"Authorization": f"Bearer {token}"},
    )


def _find_submission(admin_token, sub_id):
    r = requests.get(f"{API}/admin/submissions", headers={"Authorization": f"Bearer {admin_token}"})
    assert r.status_code == 200
    for s in r.json():
        if s["id"] == sub_id:
            return s
    return None


# BUG-1: Other role -> custom typed role persisted
def test_other_role_custom_text_persisted(worker, admin_token):
    r = _submit(worker["token"], client_name="TEST_Carpenter_Client", client_role="Carpenter", location="Nagercoil")
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["visit_number"] == 1
    sub = _find_submission(admin_token, body["id"])
    assert sub is not None
    assert sub["client_role"] == "Carpenter"
    assert sub["location"] == "Nagercoil"


# BUG-2: Other location -> custom typed location persisted
def test_other_location_custom_text_persisted(worker, admin_token):
    r = _submit(worker["token"], client_name="TEST_Marthandam_Client", client_role="Engineer", location="Marthandam")
    assert r.status_code == 200, r.text
    body = r.json()
    sub = _find_submission(admin_token, body["id"])
    assert sub is not None
    assert sub["client_role"] == "Engineer"
    assert sub["location"] == "Marthandam"


# EDGE-1: Empty role/location still accepted
def test_blank_role_and_location_accepted(worker, admin_token):
    r = _submit(worker["token"], client_name="TEST_Blank_Client", client_role="", location="")
    assert r.status_code == 200, r.text
    sub = _find_submission(admin_token, r.json()["id"])
    assert sub is not None
    assert sub["client_role"] == ""
    assert sub["location"] == ""


# REGRESSION-1 & 2: Predefined values still work
def test_predefined_role_location_unchanged(worker, admin_token):
    r = _submit(worker["token"], client_name="TEST_Predef_Client", client_role="Engineer", location="Tirunelveli")
    assert r.status_code == 200, r.text
    sub = _find_submission(admin_token, r.json()["id"])
    assert sub["client_role"] == "Engineer"
    assert sub["location"] == "Tirunelveli"


# 50-char cap enforced
def test_role_location_capped_at_50_chars(worker, admin_token):
    long_role = "X" * 80
    long_loc = "Y" * 80
    r = _submit(worker["token"], client_name="TEST_LongCap_Client", client_role=long_role, location=long_loc)
    assert r.status_code == 200, r.text
    sub = _find_submission(admin_token, r.json()["id"])
    assert len(sub["client_role"]) == 50
    assert len(sub["location"]) == 50


# FOLLOWUP: Custom role/location inherited by follow-up
def test_followup_inherits_custom_role_and_location(worker, admin_token):
    # Initial with custom role + location and GPS
    r = _submit(
        worker["token"],
        client_name="TEST_Followup_Client",
        client_role="Carpenter",
        location="Marthandam",
        latitude="8.179000",
        longitude="77.430000",
    )
    assert r.status_code == 200, r.text
    original_id = r.json()["id"]

    # Follow-up within 300m
    fu = requests.post(
        f"{API}/field/follow-up/{original_id}",
        data={
            "status": "Materials Delivered",
            "notes": "TEST follow-up inheritance",
            "latitude": "8.179050",
            "longitude": "77.430000",
        },
        headers={"Authorization": f"Bearer {worker['token']}"},
    )
    assert fu.status_code == 200, fu.text
    fu_id = fu.json()["id"]
    assert fu.json()["visit_number"] == 2

    sub = _find_submission(admin_token, fu_id)
    assert sub is not None
    assert sub["client_role"] == "Carpenter", f"Follow-up did not inherit role: {sub['client_role']}"
    assert sub["location"] == "Marthandam", f"Follow-up did not inherit location: {sub['location']}"


# Email sent flag set (will be True/False depending on SMTP; just ensure key exists & returned)
def test_email_sent_flag_present(worker):
    r = _submit(worker["token"], client_name="TEST_Email_Client", client_role="Painter", location="Kanyakumari")
    assert r.status_code == 200, r.text
    body = r.json()
    assert "email_sent" in body
    assert isinstance(body["email_sent"], bool)
