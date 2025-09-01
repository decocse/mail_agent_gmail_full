#!/usr/bin/env python3
"""
Mail Agent: Gmail Primary inbox → auto-reply with Stack Overflow answers.

Features:
- Only new incoming mails (since today)
- Skips Gmail Social & Promotions
- Skips non-programming mails
- Uses email body as query (subject fallback)
- Cleans queries to avoid API 400 errors
- Adds 1s throttle + API key support to avoid 429 rate limits
- Caches repeated queries
- Waits/backoff if rate limited
- Preserves code formatting in replies
"""

import email
import imaplib
import os
import re
import smtplib
import sys
import time
import traceback
from contextlib import contextmanager
from dataclasses import dataclass
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import parseaddr, formataddr, make_msgid
from typing import Optional, Tuple
from datetime import datetime

import requests


def log(msg: str) -> None:
    print(msg, flush=True)


def warn(msg: str) -> None:
    print(f"[WARN] {msg}", flush=True)


def err(msg: str) -> None:
    print(f"[ERROR] {msg}", file=sys.stderr, flush=True)


def load_env_file_if_present(path: str = ".env") -> None:
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                os.environ.setdefault(k, v)
        log(f"Loaded environment variables from {path}")
    except Exception as e:
        warn(f"Failed to load .env: {e}")


def strip_html(html: str) -> str:
    # Preserve <pre><code> blocks as fenced code blocks
    html = re.sub(
        r"(?is)<pre><code>(.*?)</code></pre>",
        lambda m: "\n```java\n" + m.group(1).strip() + "\n```\n",
        html,
    )

    # Inline <code> → backticks
    html = re.sub(
        r"(?is)<code>(.*?)</code>",
        lambda m: "`" + m.group(1).strip() + "`",
        html,
    )

    # Remove <script> and <style> blocks
    html = re.sub(r"(?is)<(script|style).*?>.*?</\1>", "", html)

    # Replace <br> and </p> with line breaks
    html = re.sub(r"(?i)<br\s*/?>", "\n", html)
    html = re.sub(r"(?i)</p>", "\n\n", html)

    # Remove remaining tags
    text = re.sub(r"(?s)<.*?>", "", html)

    # Decode HTML entities
    text = (
        text.replace("&nbsp;", " ")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", "\"")
        .replace("&#39;", "'")
    )

    # Normalize whitespace
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text)

    return text.strip()


def decode_mime_header(value: Optional[str]) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def get_part_text(msg: email.message.Message) -> Tuple[str, str]:
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            if ctype == "text/plain":
                charset = part.get_content_charset() or "utf-8"
                try:
                    return part.get_payload(decode=True).decode(
                        charset, errors="replace"
                    ), ctype
                except Exception:
                    continue
        for part in msg.walk():
            ctype = part.get_content_type()
            if ctype == "text/html":
                charset = part.get_content_charset() or "utf-8"
                try:
                    html = part.get_payload(decode=True).decode(
                        charset, errors="replace"
                    )
                    return strip_html(html), ctype
                except Exception:
                    continue
        return "", "text/plain"
    else:
        ctype = msg.get_content_type()
        payload = msg.get_payload(decode=True)
        if payload:
            try:
                return payload.decode(
                    msg.get_content_charset() or "utf-8", errors="replace"
                ), ctype
            except Exception:
                pass
        return str(msg.get_payload()), ctype


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
    from_name: str = "Mail Agent"
    api_timeout: int = 60
    poll_interval_sec: int = 30
    reply_subject_prefix: str = "Re: "
    dry_run: bool = False


def build_config() -> Config:
    load_env_file_if_present(".env")
    cfg = Config(
        imap_host=os.environ.get("IMAP_HOST", ""),
        imap_user=os.environ.get("IMAP_USER", ""),
        imap_pass=os.environ.get("IMAP_PASS", ""),
        imap_folder=os.environ.get("IMAP_FOLDER", "INBOX"),
        smtp_host=os.environ.get("SMTP_HOST", "smtp.gmail.com"),
        smtp_port=int(os.environ.get("SMTP_PORT", "587")),
        smtp_user=os.environ.get("SMTP_USER", os.environ.get("IMAP_USER", None)),
        smtp_pass=os.environ.get("SMTP_PASS", os.environ.get("IMAP_PASS", None)),
        from_name=os.environ.get("FROM_NAME", "Mail Agent"),
        api_timeout=int(os.environ.get("API_TIMEOUT", "60")),
        poll_interval_sec=int(os.environ.get("POLL_INTERVAL_SEC", "30")),
        reply_subject_prefix=os.environ.get("REPLY_SUBJECT_PREFIX", "Re: "),
        dry_run=os.environ.get("DRY_RUN", "false").lower() == "true",
    )
    required = {
        "IMAP_HOST": cfg.imap_host,
        "IMAP_USER": cfg.imap_user,
        "IMAP_PASS": cfg.imap_pass,
    }
    missing = [k for k, v in required.items() if not v]
    if missing:
        raise SystemExit(
            f"Missing required env vars: {', '.join(missing)}. See .env.example."
        )
    return cfg


@contextmanager
def imap_session(host: str, user: str, password: str):
    m = imaplib.IMAP4_SSL(host)
    try:
        m.login(user, password)
        yield m
    finally:
        try:
            m.logout()
        except Exception:
            pass


# ---------------------------
# Stack Overflow API lookup with cache + backoff
# ---------------------------

CACHE = {}


def fetch_stackoverflow_answer(query: str, timeout: int = 30) -> tuple[bool, str]:
    try:
        query_key = query.lower().strip()
        if query_key in CACHE:
            return True, CACHE[query_key]

        time.sleep(1)  # ✅ throttle to avoid 429
        search_url = "https://api.stackexchange.com/2.3/search/advanced"
        params = {
            "order": "desc",
            "sort": "relevance",
            "q": query,
            "site": "stackoverflow",
            "answers": 1,
        }
        key = os.environ.get("STACKOVERFLOW_KEY")
        if key:
            params["key"] = key

        r = requests.get(search_url, params=params, timeout=timeout)
        if r.status_code == 429:
            log("[RATE LIMIT] Hit 429, sleeping 60s...")
            time.sleep(60)
            return False, "Rate limited by Stack Overflow (429)."
        if not r.ok:
            return False, f"Search failed: {r.status_code}"

        items = r.json().get("items", [])
        if not items:
            return False, "No relevant results found on Stack Overflow."

        qid = items[0]["question_id"]
        ans_url = f"https://api.stackexchange.com/2.3/questions/{qid}/answers"
        ans_params = {
            "order": "desc",
            "sort": "votes",
            "site": "stackoverflow",
            "filter": "withbody",
        }
        if key:
            ans_params["key"] = key

        r2 = requests.get(ans_url, params=ans_params, timeout=timeout)
        if r2.status_code == 429:
            log("[RATE LIMIT] Hit 429 on answer fetch, sleeping 60s...")
            time.sleep(60)
            return False, "Rate limited on answer fetch (429)."
        if not r2.ok:
            return False, f"Answer fetch failed: {r2.status_code}"

        answers = r2.json().get("items", [])
        if not answers:
            return False, "No answers found for the top question."

        answer_text = strip_html(answers[0]["body"])[:4000]
        CACHE[query_key] = answer_text
        return True, answer_text

    except Exception as e:
        return False, f"Error fetching from StackOverflow: {e}"


# ---------------------------
# Reply helpers
# ---------------------------

def build_reply_message(
    cfg: Config, original: email.message.Message, reply_text: str, to_addr: str
) -> EmailMessage:
    msg = EmailMessage()
    orig_subject = decode_mime_header(original.get("Subject"))
    in_reply_to = original.get("Message-ID") or make_msgid()
    references = original.get_all("References", []) or []
    if in_reply_to and in_reply_to not in references:
        references.append(in_reply_to)

    from_disp = formataddr((cfg.from_name, cfg.smtp_user or cfg.imap_user))
    msg["From"] = from_disp
    msg["To"] = to_addr
    msg["Subject"] = (
        (cfg.reply_subject_prefix + orig_subject)
        if orig_subject
        and not orig_subject.lower().startswith(cfg.reply_subject_prefix.lower())
        else orig_subject
        or "Re: (no subject)"
    )
    msg["In-Reply-To"] = in_reply_to
    if references:
        msg["References"] = " ".join(references)

    msg.set_content(reply_text)
    return msg


def smtp_send(cfg: Config, message: EmailMessage) -> None:
    if cfg.dry_run:
        log(
            f"[DRY RUN] Would send email to {message['To']} with subject '{message['Subject']}'"
        )
        return

    with smtplib.SMTP(cfg.smtp_host, cfg.smtp_port) as s:
        s.starttls()
        s.login(cfg.smtp_user or cfg.imap_user, cfg.smtp_pass or cfg.imap_pass)
        s.send_message(message)


# ---------------------------
# Filters + Cleaning
# ---------------------------

def is_programming_related(text: str) -> bool:
    keywords = [
        "python",
        "java",
        "c++",
        "c#",
        "javascript",
        "typescript",
        "error",
        "bug",
        "stack",
        "traceback",
        "function",
        "class",
        "method",
        "variable",
        "thread",
        "sql",
        "database",
        "code",
        "compile",
        "runtime",
        "exception",
        "algorithm",
    ]
    text_lower = text.lower()
    return any(kw in text_lower for kw in keywords)


def clean_query(subject: str, body: str) -> str:
    # ✅ Prefer body text over subject
    if len(body.strip()) > 10:
        text = body[:200]
    else:
        text = f"{subject} {body[:200]}"

    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"[^a-zA-Z0-9\s\+\-\.\,\?\!]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# ---------------------------
# Process each mail
# ---------------------------

def process_message(cfg: Config, raw: bytes, uid: str) -> None:
    msg = email.message_from_bytes(raw)
    from_addr = parseaddr(msg.get("From"))[1]
    subject = decode_mime_header(msg.get("Subject"))
    body_text, ctype = get_part_text(msg)

    log(f"Processing UID={uid} From={from_addr} Subject={subject!r}")

    if not is_programming_related(subject + " " + body_text):
        log(f"Skipping UID={uid} (not programming related).")
        return

    query = clean_query(subject, body_text)

    ok, api_reply = fetch_stackoverflow_answer(query, cfg.api_timeout)

    if not ok:
        err(f"StackOverflow lookup failed for UID={uid}: {api_reply}")
        return

    reply_body = api_reply.strip()
    if not reply_body:
        log(f"No useful answer found for UID={uid}, skipping reply.")
        return

    to_addr = from_addr or (cfg.smtp_user or cfg.imap_user)
    reply_msg = build_reply_message(cfg, msg, reply_body, to_addr)
    smtp_send(cfg, reply_msg)
    log(f"Replied to {to_addr} for UID={uid}")


# ---------------------------
# Gmail poll loop
# ---------------------------

def poll_loop(cfg: Config) -> None:
    log("Starting poll loop (Primary only, new mails since today)...")
    today = datetime.now().strftime("%d-%b-%Y")

    while True:
        try:
            with imap_session(cfg.imap_host, cfg.imap_user, cfg.imap_pass) as m:
                m.select(cfg.imap_folder)

                typ, data = m.search(None, f'(SINCE "{today}")')
                if typ != "OK":
                    time.sleep(cfg.poll_interval_sec)
                    continue

                ids = data[0].split()
                for msg_id in ids:
                    typ_labels, label_data = m.fetch(msg_id, "(X-GM-LABELS)")
                    labels = (
                        label_data[0][1].decode()
                        if label_data and isinstance(label_data[0], tuple)
                        else ""
                    )
                    if "Social" in labels or "Promotions" in labels:
                        log(f"Skipping id={msg_id} (non-Primary, labels={labels})")
                        continue

                    typ_body, body_data = m.fetch(msg_id, "(RFC822)")
                    if (
                        typ_body != "OK"
                        or not body_data
                        or not isinstance(body_data[0], tuple)
                    ):
                        continue
                    raw = body_data[0][1]

                    try:
                        process_message(cfg, raw, msg_id.decode())
                    except Exception as e:
                        err(
                            f"Processing error for id={msg_id}: {e}\n{traceback.format_exc()}"
                        )
        except KeyboardInterrupt:
            log("Received KeyboardInterrupt, exiting.")
            break
        except Exception as e:
            err(f"Top-level loop error: {e}\n{traceback.format_exc()}")
        time.sleep(cfg.poll_interval_sec)


def main():
    cfg = build_config()
    log("Configuration loaded.")
    log(
        f"IMAP host: {cfg.imap_host}, user: {cfg.imap_user}, folder: {cfg.imap_folder}"
    )
    log(
        f"SMTP host: {cfg.smtp_host}, user: {cfg.smtp_user or cfg.imap_user}, port: {cfg.smtp_port}"
    )
    log(f"Poll interval: {cfg.poll_interval_sec}s")
    poll_loop(cfg)


if __name__ == "__main__":
    main()
