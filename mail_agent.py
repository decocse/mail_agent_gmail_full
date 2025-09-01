#!/usr/bin/env python3
"""
Mail Agent with Web UI (Gmail version)

- Connects to Gmail via IMAP
- Reads only new incoming emails (ignores old unread mails)
- Searches Stack Overflow for relevant answers
- Sends reply via Gmail SMTP
- Logs responses into a web dashboard (Flask)

Run:
    python mail_agent_ui.py
Then open:
    http://localhost:5000
"""

import email
import imaplib
import os
import re
import smtplib
import sqlite3
import threading
import time
import traceback
from contextlib import contextmanager
from dataclasses import dataclass
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import parseaddr, formataddr, make_msgid
from typing import Optional, Tuple

import requests
from flask import Flask, render_template_string


# ------------------------------
# Config
# ------------------------------

@dataclass
class Config:
    imap_host: str
    imap_user: str
    imap_pass: str
    imap_folder: str = "INBOX"
    smtp_host: Optional[str] = None
    smtp_port: int = 587
    smtp_user: Optional[str] = None
    smtp_pass: Optional[str] = None
    from_name: str = "StackOverflow Mail Agent"
    api_timeout: int = 30
    poll_interval_sec: int = 15


def load_env_file(path=".env"):
    if not os.path.exists(path):
        return
    with open(path) as f:
        for line in f:
            if not line.strip() or line.strip().startswith("#") or "=" not in line:
                continue
            k, v = line.strip().split("=", 1)
            os.environ.setdefault(k, v.strip())


def build_config() -> Config:
    load_env_file(".env")
    return Config(
        imap_host=os.environ["IMAP_HOST"],
        imap_user=os.environ["IMAP_USER"],
        imap_pass=os.environ["IMAP_PASS"],
        imap_folder=os.environ.get("IMAP_FOLDER", "INBOX"),
        smtp_host=os.environ.get("SMTP_HOST", "smtp.gmail.com"),
        smtp_port=int(os.environ.get("SMTP_PORT", "587")),
        smtp_user=os.environ.get("SMTP_USER", os.environ["IMAP_USER"]),
        smtp_pass=os.environ.get("SMTP_PASS", os.environ["IMAP_PASS"]),
        from_name=os.environ.get("FROM_NAME", "StackOverflow Mail Agent"),
        api_timeout=int(os.environ.get("API_TIMEOUT", "30")),
        poll_interval_sec=int(os.environ.get("POLL_INTERVAL_SEC", "15")),
    )


# ------------------------------
# Helpers
# ------------------------------

def strip_html(html: str) -> str:
    html = re.sub(r"(?is)<(script|style).*?>.*?</\1>", "", html)
    html = re.sub(r"(?i)<br\s*/?>", "\n", html)
    html = re.sub(r"(?i)</p>", "\n\n", html)
    text = re.sub(r"(?s)<.*?>", "", html)
    return text.strip()


def decode_mime_header(value: Optional[str]) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def get_body_text(msg: email.message.Message) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="replace")
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                return strip_html(part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="replace"))
    else:
        return msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8", errors="replace")
    return ""


def fetch_stackoverflow_answer(query: str, timeout: int = 20) -> tuple[bool, str]:
    try:
        search_url = "https://api.stackexchange.com/2.3/search/advanced"
        params = {
            "order": "desc",
            "sort": "relevance",
            "q": query,
            "site": "stackoverflow",
            "accepted": True,
            "answers": 1,
        }
        r = requests.get(search_url, params=params, timeout=timeout)
        items = r.json().get("items", [])
        if not items:
            return False, "No relevant results found on Stack Overflow."
        qid = items[0]["question_id"]
        ans_url = f"https://api.stackexchange.com/2.3/questions/{qid}/answers"
        ans_params = {"order": "desc", "sort": "votes", "site": "stackoverflow", "filter": "withbody"}
        r2 = requests.get(ans_url, params=ans_params, timeout=timeout)
        answers = r2.json().get("items", [])
        if not answers:
            return False, "No answers found."
        return True, strip_html(answers[0]["body"])[:3000]
    except Exception as e:
        return False, f"Error: {e}"


def build_reply(cfg: Config, original: email.message.Message, reply_text: str, to_addr: str) -> EmailMessage:
    msg = EmailMessage()
    subject = decode_mime_header(original.get("Subject"))
    in_reply_to = original.get("Message-ID") or make_msgid()

    msg["From"] = formataddr((cfg.from_name, cfg.smtp_user))
    msg["To"] = to_addr
    msg["Subject"] = f"Re: {subject}" if subject and not subject.lower().startswith("re:") else subject or "Re: (no subject)"
    msg["In-Reply-To"] = in_reply_to
    msg["References"] = in_reply_to
    msg.set_content(reply_text)
    return msg


def smtp_send(cfg: Config, msg: EmailMessage):
    with smtplib.SMTP(cfg.smtp_host, cfg.smtp_port) as s:
        s.starttls()
        s.login(cfg.smtp_user, cfg.smtp_pass)
        s.send_message(msg)


# ------------------------------
# IMAP handling (new mails only)
# ------------------------------

def get_last_uid(imap, folder="INBOX") -> int:
    imap.select(folder)
    typ, data = imap.search(None, "ALL")
    if typ != "OK" or not data[0]:
        return 0
    latest_id = data[0].split()[-1]
    typ, uid_data = imap.uid("fetch", latest_id, "(UID)")
    if typ != "OK" or not uid_data or not isinstance(uid_data[0], tuple):
        return 0
    m = re.search(r"UID (\\d+)", uid_data[0][1].decode())
    return int(m.group(1)) if m else 0


# ------------------------------
# Web UI
# ------------------------------

app = Flask(__name__)
mail_logs = []


def add_log(uid, frm, subject, status, message):
    mail_logs.append({
        "uid": uid,
        "from": frm,
        "subject": subject,
        "status": status,
        "message": message
    })
    if len(mail_logs) > 100:
        mail_logs.pop(0)


@app.route("/")
def index():
    return render_template_string("""
    <html>
    <head>
        <title>Mail Agent Dashboard</title>
        <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #ddd; padding: 6px; font-size: 14px; }
            th { background: #f0f0f0; }
        </style>
    </head>
    <body>
        <h1>📬 Mail Agent Dashboard</h1>
        <table>
            <tr><th>UID</th><th>From</th><th>Subject</th><th>Status</th><th>Message</th></tr>
            {% for log in logs %}
            <tr>
                <td>{{log.uid}}</td>
                <td>{{log.from}}</td>
                <td>{{log.subject}}</td>
                <td>{{log.status}}</td>
                <td>{{log.message}}</td>
            </tr>
            {% endfor %}
        </table>
    </body>
    </html>
    """, logs=reversed(mail_logs))


# ------------------------------
# Agent loop
# ------------------------------

def agent_loop(cfg: Config):
    with imaplib.IMAP4_SSL(cfg.imap_host) as imap:
        imap.login(cfg.imap_user, cfg.imap_pass)
        last_uid = get_last_uid(imap, cfg.imap_folder)
        add_log("-", "-", "-", "INFO", f"Starting after UID={last_uid}")

    while True:
        try:
            with imaplib.IMAP4_SSL(cfg.imap_host) as imap:
                imap.login(cfg.imap_user, cfg.imap_pass)
                imap.select(cfg.imap_folder)
                typ, data = imap.uid("search", None, f"UID {last_uid+1}:*")
                if typ != "OK":
                    time.sleep(cfg.poll_interval_sec)
                    continue
                ids = data[0].split()
                for uid in ids:
                    last_uid = int(uid)
                    typ, body_data = imap.uid("fetch", uid, "(RFC822)")
                    if typ != "OK" or not body_data or not isinstance(body_data[0], tuple):
                        continue
                    raw = body_data[0][1]
                    msg = email.message_from_bytes(raw)
                    frm = parseaddr(msg.get("From"))[1]
                    subject = decode_mime_header(msg.get("Subject"))
                    body = get_body_text(msg)

                    ok, reply = fetch_stackoverflow_answer(body, cfg.api_timeout)
                    if not ok:
                        reply = f"No answer found. ({reply})"

                    try:
                        reply_msg = build_reply(cfg, msg, reply, frm)
                        smtp_send(cfg, reply_msg)
                        add_log(uid.decode(), frm, subject, "SENT", reply[:200] + "...")
                    except Exception as e:
                        add_log(uid.decode(), frm, subject, "ERROR", str(e))
        except Exception as e:
            add_log("-", "-", "-", "ERROR", f"Loop failed: {e}\n{traceback.format_exc()}")
        time.sleep(cfg.poll_interval_sec)


# ------------------------------
# Main
# ------------------------------

if __name__ == "__main__":
    cfg = build_config()
    t = threading.Thread(target=agent_loop, args=(cfg,), daemon=True)
    t.start()
    app.run(host="0.0.0.0", port=5000)
