"""Tests for BUG: Client Role missing from HQ email body.

Approach: monkey-patch smtplib.SMTP inside backend.server to capture the
EmailMessage object and assert on the rendered HTML payload.
"""
import os
import sys
import importlib
import pytest

# Make the backend module importable.
sys.path.insert(0, "/app/backend")


@pytest.fixture(scope="module")
def server_module():
    # Force a fresh import so env vars are loaded.
    if "server" in sys.modules:
        del sys.modules["server"]
    return importlib.import_module("server")


class _FakeSMTP:
    """Captures the sent EmailMessage; replaces smtplib.SMTP context manager."""
    captured = []

    def __init__(self, host, port, timeout=30):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def starttls(self):
        pass

    def login(self, u, p):
        pass

    def send_message(self, msg):
        _FakeSMTP.captured.append(msg)


@pytest.fixture
def capture_email(server_module, monkeypatch):
    _FakeSMTP.captured = []
    monkeypatch.setattr(server_module.smtplib, "SMTP", _FakeSMTP)
    yield _FakeSMTP


def _get_html(msg):
    """Extract the text/html alternative as a string from an EmailMessage."""
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            return part.get_content()
    # fall back to full string
    return msg.as_string()


def _base_sub(**overrides):
    sub = {
        "id": "test-id",
        "visit_number": 1,
        "status": "Site Visited",
        "worker_name": "Test Worker",
        "worker_email": "worker@test.com",
        "client_name": "Ravi Kumar",
        "client_company": "Kumar Constructions",
        "client_mobile": "9876543210",
        "client_email": "ravi@test.com",
        "client_role": "",
        "location": "",
        "site_address": "12 Main Street",
        "notes": "Initial visit",
        "geo": {"latitude": 8.18, "longitude": 77.43},
        "created_at": "2026-01-15T10:00:00",
    }
    sub.update(overrides)
    return sub


# BUG: Role row present with custom 'Carpenter' value
def test_email_html_contains_carpenter_role(server_module, capture_email):
    ok = server_module.send_submission_email(_base_sub(client_role="Carpenter", location="Kanyakumari"))
    assert ok is True
    assert len(capture_email.captured) == 1
    html = _get_html(capture_email.captured[0])
    assert "<strong>Role</strong>" in html, "Role label missing from email HTML"
    assert "Carpenter" in html, "Custom role value 'Carpenter' missing from email HTML"


# REGRESSION: Predefined role 'Engineer' appears
def test_email_html_contains_engineer_role(server_module, capture_email):
    server_module.send_submission_email(_base_sub(client_role="Engineer", location="Tirunelveli"))
    html = _get_html(capture_email.captured[0])
    assert "<strong>Role</strong>" in html
    assert "Engineer" in html


# REGRESSION: Empty role -> label shown with '-'
def test_email_html_empty_role_shows_dash(server_module, capture_email):
    server_module.send_submission_email(_base_sub(client_role=""))
    html = _get_html(capture_email.captured[0])
    assert "<strong>Role</strong>" in html
    # Find the Role row and confirm it renders '-'
    # The cell is: <td ...><strong>Role</strong></td><td ...>-</td>
    role_idx = html.find("<strong>Role</strong>")
    assert role_idx != -1
    after = html[role_idx:role_idx + 400]
    assert ">-<" in after, f"Expected '-' for empty role, got: {after}"


# REGRESSION: custom location 'Kanyakumari' under This Visit table
def test_email_html_contains_custom_location(server_module, capture_email):
    server_module.send_submission_email(_base_sub(client_role="Engineer", location="Kanyakumari"))
    html = _get_html(capture_email.captured[0])
    assert "<strong>Location</strong>" in html
    # Confirm 'Kanyakumari' appears after the Location label
    loc_idx = html.find("<strong>Location</strong>")
    assert loc_idx != -1
    after = html[loc_idx:loc_idx + 400]
    assert "Kanyakumari" in after


# FOLLOWUP: visit_number=2 email also carries Role
def test_followup_email_contains_role(server_module, capture_email):
    server_module.send_submission_email(
        _base_sub(visit_number=2, status="Materials Delivered", client_role="Carpenter", location="Marthandam"),
        timeline=[
            _base_sub(client_role="Carpenter", location="Marthandam"),
            _base_sub(visit_number=2, status="Materials Delivered", client_role="Carpenter", location="Marthandam"),
        ],
    )
    html = _get_html(capture_email.captured[0])
    assert "<strong>Role</strong>" in html
    assert "Carpenter" in html
    # Timeline section should appear for >1 visit
    assert "Client Visit Timeline" in html


# ALL-TOGETHER: full structural check
def test_email_html_full_structure(server_module, capture_email):
    server_module.send_submission_email(_base_sub(client_role="Plumber", location="Nagercoil"))
    html = _get_html(capture_email.captured[0])
    # Client Details table rows
    for label in ["Client Name", "Role", "Company", "Mobile", "Email"]:
        assert f"<strong>{label}</strong>" in html, f"Client Details row '{label}' missing"
    # This Visit table rows
    for label in ["Visit Number", "Status", "Location", "Site Address"]:
        assert f"<strong>{label}</strong>" in html, f"This Visit row '{label}' missing"
    # Ordering: Role appears AFTER Client Name and BEFORE Company
    assert html.index("<strong>Client Name</strong>") < html.index("<strong>Role</strong>") < html.index("<strong>Company</strong>")
