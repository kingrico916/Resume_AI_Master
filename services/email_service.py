"""
services.email_service
======================
Send emails via Microsoft Graph API and log every attempt to the database.
"""

import os
import sys
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from data.email_templates import render_template, get_template
from data.database import log_email

OUTLOOK_USER        = os.getenv("OUTLOOK_USER", "")
GRAPH_TENANT_ID     = os.getenv("GRAPH_TENANT_ID", "")
GRAPH_CLIENT_ID     = os.getenv("GRAPH_CLIENT_ID", "")
GRAPH_CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET", "")


def _get_graph_token():
    url = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     GRAPH_CLIENT_ID,
        "client_secret": GRAPH_CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def _send_via_graph(to_addr, to_name, reply_to, subject, body):
    token = _get_graph_token()
    message = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": [{"emailAddress": {"address": to_addr, "name": to_name}}],
    }
    if reply_to:
        message["replyTo"] = [{"emailAddress": {"address": reply_to}}]
    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{OUTLOOK_USER}/sendMail",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"message": message, "saveToSentItems": True},
        timeout=30,
    )
    resp.raise_for_status()


def send_email(vsc_name: str, vsc_email: str,
               to_name: str, to_email: str,
               template_key: str, variables: dict) -> dict:
    """
    Render template, send via Microsoft Graph API, log result.
    """
    tmpl = get_template(template_key)
    if not tmpl:
        return {"success": False, "error": f"Unknown template: {template_key}"}

    rendered = render_template(template_key, variables)
    subject  = rendered["subject"]
    body     = rendered["body"]

    status        = "sent"
    error_message = None

    try:
        if not all([OUTLOOK_USER, GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET]):
            raise ValueError("Graph API credentials not configured — set GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET")
        _send_via_graph(
            to_addr  = to_email,
            to_name  = to_name,
            reply_to = vsc_email or OUTLOOK_USER,
            subject  = subject,
            body     = body,
        )
    except Exception as e:
        status        = "failed"
        error_message = str(e)

    log_id = log_email(
        vsc_name      = vsc_name,
        vsc_email     = vsc_email,
        to_name       = to_name,
        to_email      = to_email,
        template_key  = template_key,
        template_label= tmpl["label"],
        subject       = subject,
        body          = body,
        status        = status,
        error_message = error_message,
    )

    if status == "sent":
        return {"success": True, "log_id": log_id}
    else:
        return {"success": False, "log_id": log_id, "error": error_message}


def preview_email(template_key: str, variables: dict) -> dict:
    """Return rendered subject and body without sending."""
    tmpl = get_template(template_key)
    if not tmpl:
        return {"error": f"Unknown template: {template_key}"}
    rendered = render_template(template_key, variables)
    return {
        "subject": rendered["subject"],
        "body":    rendered["body"],
        "label":   tmpl["label"],
        "when":    tmpl["when"],
    }
