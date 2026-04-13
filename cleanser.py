#!/usr/bin/env python3
"""
Email Cleanser — Dry-Run Scanner (v1)
======================================
Scans your Microsoft 365 Inbox via Microsoft Graph and classifies
messages into: Newsletter, Receipt, Quarantine, or Untouched.

Produces a Markdown report and optional CSV for review.
No messages are moved or deleted in v1.

Usage:
    python cleanser.py                  # full dry-run
    python cleanser.py --limit 500      # scan only 500 messages (for testing)
    python cleanser.py --help
"""

import argparse
import csv
import io
import json
import os
import sys
import time
import webbrowser
import yaml

# Force UTF-8 output on Windows so emoji in print() don't crash
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from enum import Enum
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

# ── Third-party imports (installed via requirements.txt) ─────
try:
    import msal
    import requests
except ImportError:
    print(
        "Missing dependencies. Please run:\n"
        "  pip install msal requests pyyaml\n"
        "Then re-run this script."
    )
    sys.exit(1)


# ============================================================
# Constants
# ============================================================
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.ReadWrite", "User.Read"]  # User.Read needed to identify the connected account
TOKEN_CACHE_FILE = Path(__file__).parent / ".token_cache.json"
CONFIG_FILE = Path(__file__).parent / "config.yaml"
REPORT_DIR = Path(__file__).parent / "reports"


class Classification(str, Enum):
    NEWSLETTER = "Newsletter"
    RECEIPT = "Receipt"
    QUARANTINE = "Quarantine"
    UNTOUCHED = "Untouched"


@dataclass
class MessageRecord:
    message_id: str
    subject: str
    sender_name: str
    sender_email: str
    sender_domain: str
    received_datetime: datetime
    is_read: bool
    conversation_id: str
    has_unsubscribe_signal: bool = False
    matched_newsletter_rule: bool = False
    matched_receipt_rule: bool = False
    matched_quarantine_rule: bool = False
    keep_subscription: bool = False  # quarantined but sender was recently engaged
    classification: Classification = Classification.UNTOUCHED
    rule_triggers: list = field(default_factory=list)
    body_snippet: str = ""


# ============================================================
# Config loader
# ============================================================
def load_config(path: Path = CONFIG_FILE) -> dict:
    if not path.exists():
        print(f"ERROR: Config file not found at {path}")
        sys.exit(1)
    with open(path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    # Normalise domains to lowercase
    cfg["newsletter_domains"] = [
        d.lower().strip() for d in cfg.get("newsletter_domains", [])
    ]
    cfg["receipt_sender_domains"] = [
        d.lower().strip() for d in cfg.get("receipt_sender_domains", [])
    ]
    cfg["receipt_keywords"] = [
        k.lower().strip() for k in cfg.get("receipt_keywords", [])
    ]
    cfg["unsubscribe_signals"] = [
        s.lower().strip() for s in cfg.get("unsubscribe_signals", [])
    ]
    cfg["excluded_sender_domains"] = [
        d.lower().strip() for d in (cfg.get("excluded_sender_domains") or [])
    ]
    cfg["excluded_sender_addresses"] = [
        a.lower().strip() for a in (cfg.get("excluded_sender_addresses") or [])
    ]
    return cfg


# ============================================================
# Authentication (Device Code Flow — no web server needed)
# ============================================================
class GraphAuth:
    """
    Handles Microsoft Graph authentication using the Device Code flow.
    This is ideal for personal/desktop scripts — you don't need a web server.
    
    You must register an app in Azure AD first (see README).
    """

    # Default: Microsoft's well-known "mobile/desktop" client ID for
    # personal accounts. Replace with YOUR app registration's client ID.
    CLIENT_ID = "YOUR_CLIENT_ID_HERE"
    AUTHORITY = "https://login.microsoftonline.com/organizations"

    def __init__(self):
        self._load_client_id()
        self.cache = msal.SerializableTokenCache()
        if TOKEN_CACHE_FILE.exists():
            self.cache.deserialize(TOKEN_CACHE_FILE.read_text())
        self.app = msal.PublicClientApplication(
            self.CLIENT_ID,
            authority=self.AUTHORITY,
            token_cache=self.cache,
        )

    def _load_client_id(self):
        """Load client ID from environment or .env file."""
        # 1. Check environment variable
        env_id = os.environ.get("EMAIL_CLEANSER_CLIENT_ID")
        if env_id and env_id != "your-client-id-goes-here":
            self.CLIENT_ID = env_id
            print(f"  Client ID loaded from environment variable")
            return

        # 2. Check .env file in same directory as this script
        env_file = Path(__file__).parent / ".env"
        if env_file.exists():
            for line in env_file.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, _, value = line.partition("=")
                    key = key.strip()
                    value = value.strip().strip('"').strip("'")
                    if key == "EMAIL_CLEANSER_CLIENT_ID" and value and value != "your-client-id-goes-here":
                        self.CLIENT_ID = value
                        print(f"  Client ID loaded from .env file")
                        return
            print(f"  WARNING: .env file found at {env_file} but no valid CLIENT_ID in it")
        else:
            print(f"  WARNING: No .env file found at {env_file}")

        if self.CLIENT_ID == "YOUR_CLIENT_ID_HERE":
            print(
                "\nERROR: No Azure AD Client ID configured.\n"
                "Create a .env file in the same folder as cleanser.py with:\n"
                "  EMAIL_CLEANSER_CLIENT_ID=your-actual-client-id\n"
                "\nSee README.md for setup instructions."
            )
            sys.exit(1)

    def _save_cache(self):
        if self.cache.has_state_changed:
            TOKEN_CACHE_FILE.write_text(self.cache.serialize())

    def clear_cache(self):
        """Delete the cached token, forcing fresh authentication on next run."""
        if TOKEN_CACHE_FILE.exists():
            TOKEN_CACHE_FILE.unlink()
            print("  Token cache cleared.")
        # Reset the in-memory cache too
        self.cache = msal.SerializableTokenCache()
        self.app = msal.PublicClientApplication(
            self.CLIENT_ID,
            authority=self.AUTHORITY,
            token_cache=self.cache,
        )

    def get_token(self) -> str:
        accounts = self.app.get_accounts()
        result = None
        if accounts:
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result:
            flow = self.app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                print(f"Auth error: {json.dumps(flow, indent=2)}")
                sys.exit(1)
            print("\n" + "=" * 60)
            print("SIGN IN REQUIRED")
            print("=" * 60)
            print(f"1. Open:  {flow['verification_uri']}")
            print(f"2. Enter: {flow['user_code']}")
            print("=" * 60 + "\n")
            # Try to open the browser automatically
            try:
                webbrowser.open(flow["verification_uri"])
            except Exception:
                pass
            result = self.app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            print(f"Auth failed: {result.get('error_description', result)}")
            sys.exit(1)
        self._save_cache()
        return result["access_token"]


# ============================================================
# Microsoft Graph client
# ============================================================
class GraphClient:
    """Thin wrapper around Graph API calls for mail operations."""

    PAGE_SIZE = 250  # messages per request (max 1000, 250 is safe)

    def __init__(self, auth: "GraphAuth"):
        self.auth = auth
        self.session = requests.Session()
        self._apply_token(auth.get_token())

    def _apply_token(self, token: str):
        self.session.headers.update({"Authorization": f"Bearer {token}"})
        self.session.headers.setdefault("Content-Type", "application/json")

    def _refresh_token(self):
        """Silently refresh the access token (called on 401)."""
        print("  🔄 Token expired — refreshing …")
        # Force MSAL to fetch a new token (bypass silent cache)
        accounts = self.auth.app.get_accounts()
        if accounts:
            result = self.auth.app.acquire_token_silent(SCOPES, account=accounts[0], force_refresh=True)
            if result and "access_token" in result:
                self.auth._save_cache()
                self._apply_token(result["access_token"])
                return
        # Fallback: full device flow (shouldn't normally happen)
        self._apply_token(self.auth.get_token())

    def _get(self, url: str, params: dict = None) -> dict:
        max_retries = 5
        base_delay = 2
        auth_refreshed = False

        for attempt in range(max_retries):
            try:
                resp = self.session.get(url, params=params, timeout=90)

                if resp.status_code == 429:
                    retry_after = resp.headers.get("Retry-After")
                    if retry_after and retry_after.isdigit():
                        delay = int(retry_after)
                    else:
                        delay = base_delay * (2 ** attempt)
                    print(f"  ⏳ Throttled — waiting {delay}s …")
                    time.sleep(delay)
                    continue

                if resp.status_code == 401:
                    if not auth_refreshed:
                        print("  🔄 Token expired — refreshing …")
                        self._refresh_token()
                        auth_refreshed = True
                        continue
                    resp.raise_for_status()

                resp.raise_for_status()
                return resp.json()

            except (requests.exceptions.ReadTimeout, requests.exceptions.ConnectionError) as e:
                if attempt < max_retries - 1:
                    delay = base_delay * (2 ** attempt)
                    print(f"  ⚠️ Request failed ({type(e).__name__}) — retrying in {delay}s …")
                    time.sleep(delay)
                    continue
                raise

        raise RuntimeError(f"GET request failed after {max_retries} attempts: {url}")

    def _post(self, url: str, data: dict) -> dict:
        max_retries = 5
        base_delay = 2
        auth_refreshed = False

        for attempt in range(max_retries):
            try:
                resp = self.session.post(url, json=data, timeout=90)

                if resp.status_code == 429:
                    retry_after = resp.headers.get("Retry-After")
                    if retry_after and retry_after.isdigit():
                        delay = int(retry_after)
                    else:
                        delay = base_delay * (2 ** attempt)
                    print(f"  ⏳ Throttled — waiting {delay}s …")
                    time.sleep(delay)
                    continue

                if resp.status_code == 401:
                    if not auth_refreshed:
                        print("  🔄 Token expired — refreshing …")
                        self._refresh_token()
                        auth_refreshed = True
                        continue
                    resp.raise_for_status()

                resp.raise_for_status()
                return resp.json()

            except (requests.exceptions.ReadTimeout, requests.exceptions.ConnectionError) as e:
                if attempt < max_retries - 1:
                    delay = base_delay * (2 ** attempt)
                    print(f"  ⚠️ Request failed ({type(e).__name__}) — retrying in {delay}s …")
                    time.sleep(delay)
                    continue
                raise

        raise RuntimeError(f"POST request failed after {max_retries} attempts: {url}")

    def get_or_create_folder(self, display_name: str) -> str:
        """Return folder ID by display name, creating it if it doesn't exist."""
        data = self._get(f"{GRAPH_BASE}/me/mailFolders", params={"$top": 100})
        for folder in data.get("value", []):
            if folder.get("displayName", "").lower() == display_name.lower():
                return folder["id"]
        print(f"  📁 Creating folder '{display_name}' …")
        result = self._post(f"{GRAPH_BASE}/me/mailFolders", {"displayName": display_name})
        return result["id"]

    def move_message(self, message_id: str, folder_id: str) -> None:
        """Move a message to the specified folder."""
        self._post(
            f"{GRAPH_BASE}/me/messages/{message_id}/move",
            {"destinationId": folder_id},
        )

    def get_inbox_folder_id(self) -> str:
        data = self._get(f"{GRAPH_BASE}/me/mailFolders/Inbox")
        return data["id"]

    def iter_inbox_pages(
        self,
        limit: Optional[int] = None,
        resume_after: Optional[str] = None,
    ):
        """
        Yield one page of Inbox messages at a time (newest-first).
        resume_after: ISO datetime string — if set, only fetches messages older than this.
        """
        fields = (
            "id,subject,sender,from,receivedDateTime,isRead,"
            "conversationId,internetMessageHeaders,bodyPreview"
        )

        url = f"{GRAPH_BASE}/me/mailFolders/Inbox/messages"
        params = {
            "$select": fields,
            "$top": min(self.PAGE_SIZE, limit) if limit else self.PAGE_SIZE,
            "$orderby": "receivedDateTime desc",
        }
        if resume_after:
            params["$filter"] = f"receivedDateTime lt '{resume_after}'"

        fetched = 0
        page = 1
        while url:
            print(f"  📬 Fetching page {page} …", end=" ", flush=True)
            try:
                data = self._get(url, params)
            except Exception as e:
                print(f"\n  ⚠️ Error fetching page {page}: {e}")
                print(f"  Stopping inbox fetch early — {fetched:,} messages fetched so far.")
                return
            batch = data.get("value", [])
            if limit:
                batch = batch[:limit - fetched]
            print(f"got {len(batch)}")
            yield batch
            fetched += len(batch)
            if limit and fetched >= limit:
                return
            url = data.get("@odata.nextLink")
            params = None  # nextLink includes params
            page += 1

    def get_inbox_messages(
        self,
        limit: Optional[int] = None,
    ) -> list[dict]:
        """Fetch all Inbox messages at once. Used by dry-run mode."""
        all_messages = []
        for page in self.iter_inbox_pages(limit=limit):
            all_messages.extend(page)
        return all_messages

    def get_sent_conversation_ids(self) -> set[str]:
        """
        Fetch conversation IDs where you have sent/replied.
        Used for thread-protection: never quarantine threads you've engaged with.
        """
        print("  📤 Building protected-thread index from Sent Items …")
        url = f"{GRAPH_BASE}/me/mailFolders/SentItems/messages"
        params = {
            "$select": "conversationId",
            "$top": 500,
        }
        conv_ids = set()
        page = 1
        while url:
            print(f"     Page {page} …", end=" ", flush=True)
            data = self._get(url, params)
            batch = data.get("value", [])
            for msg in batch:
                cid = msg.get("conversationId")
                if cid:
                    conv_ids.add(cid)
            print(f"({len(conv_ids)} unique threads)")
            url = data.get("@odata.nextLink")
            params = None
            page += 1
        return conv_ids


# ============================================================
# Classifier
# ============================================================
class EmailClassifier:
    def __init__(self, config: dict, protected_threads: set[str], recently_engaged_domains: set[str] = None):
        self.cfg = config
        self.protected_threads = protected_threads
        self.recently_engaged_domains = recently_engaged_domains or set()
        self.cutoff = datetime.now(timezone.utc) - timedelta(
            days=config.get("quarantine_age_days", 30)
        )

    @staticmethod
    def build_engaged_domains(raw_messages: list[dict], config: dict) -> set[str]:
        """
        Pre-pass over all inbox messages to find sender domains the user has
        recently engaged with (read a newsletter-style email within the window).
        These senders will be flagged keep_subscription=True when quarantined.
        """
        window = timedelta(days=config.get("engagement_window_days", 60))
        cutoff = datetime.now(timezone.utc) - window
        signals = config.get("unsubscribe_signals", [])
        engaged = set()
        for raw in raw_messages:
            if not raw.get("isRead"):
                continue
            received = datetime.fromisoformat(
                raw["receivedDateTime"].replace("Z", "+00:00")
            )
            if received < cutoff:
                continue
            headers = raw.get("internetMessageHeaders", []) or []
            has_signal = any(
                h.get("name", "").lower() == "list-unsubscribe" for h in headers
            )
            if not has_signal:
                preview = (raw.get("bodyPreview", "") or "").lower()
                has_signal = any(s in preview for s in signals)
            if has_signal:
                sender_obj = (
                    (raw.get("sender", {}) or {}).get("emailAddress", {})
                    or (raw.get("from", {}) or {}).get("emailAddress", {})
                    or {}
                )
                email = (sender_obj.get("address", "") or "").lower()
                domain = email.split("@")[-1] if "@" in email else ""
                if domain:
                    engaged.add(domain)
        return engaged

    def _extract_sender(self, raw: dict) -> tuple[str, str, str]:
        sender_obj = (
            raw.get("sender", {}).get("emailAddress", {})
            or raw.get("from", {}).get("emailAddress", {})
        )
        name = sender_obj.get("name", "")
        email = (sender_obj.get("address", "") or "").lower()
        domain = email.split("@")[-1] if "@" in email else ""
        return name, email, domain

    def _has_unsubscribe_header(self, raw: dict) -> bool:
        headers = raw.get("internetMessageHeaders", []) or []
        for h in headers:
            if h.get("name", "").lower() == "list-unsubscribe":
                return True
        return False

    def _has_unsubscribe_body(self, raw: dict) -> bool:
        preview = (raw.get("bodyPreview", "") or "").lower()
        for signal in self.cfg["unsubscribe_signals"]:
            if signal in preview:
                return True
        return False

    def classify(self, raw: dict) -> MessageRecord:
        name, email, domain = self._extract_sender(raw)
        received = datetime.fromisoformat(
            raw["receivedDateTime"].replace("Z", "+00:00")
        )
        is_read = raw.get("isRead", False)
        conv_id = raw.get("conversationId", "")
        subject = raw.get("subject", "") or "(no subject)"
        subject_lower = subject.lower()

        rec = MessageRecord(
            message_id=raw["id"],
            subject=subject,
            sender_name=name,
            sender_email=email,
            sender_domain=domain,
            received_datetime=received,
            is_read=is_read,
            conversation_id=conv_id,
            body_snippet=(raw.get("bodyPreview", "") or "")[:120],
        )

        # ── 1) Newsletter check ────────────────────────────
        if domain in self.cfg["newsletter_domains"]:
            rec.matched_newsletter_rule = True
            rec.keep_subscription = True
            rec.classification = Classification.NEWSLETTER
            rec.rule_triggers.append(f"newsletter_domain={domain}")
            return rec

        # ── 2) Receipt check ───────────────────────────────
        # Check sender domain
        if domain in self.cfg["receipt_sender_domains"]:
            for kw in self.cfg["receipt_keywords"]:
                if kw in subject_lower:
                    rec.matched_receipt_rule = True
                    rec.classification = Classification.RECEIPT
                    rec.rule_triggers.append(
                        f"receipt_domain={domain}, keyword='{kw}'"
                    )
                    return rec

        # Check subject keywords alone
        for kw in self.cfg["receipt_keywords"]:
            if kw in subject_lower:
                rec.matched_receipt_rule = True
                rec.classification = Classification.RECEIPT
                rec.rule_triggers.append(f"receipt_keyword='{kw}'")
                return rec

        # ── 3) Quarantine check ────────────────────────────
        unsub_header = self._has_unsubscribe_header(raw)
        unsub_body = self._has_unsubscribe_body(raw)
        rec.has_unsubscribe_signal = unsub_header or unsub_body

        is_old = received < self.cutoff
        is_protected = conv_id in self.protected_threads
        is_excluded_domain = domain in self.cfg["excluded_sender_domains"]
        is_excluded_addr = email in self.cfg["excluded_sender_addresses"]

        if (
            not is_read
            and is_old
            and rec.has_unsubscribe_signal
            and not rec.matched_receipt_rule
            and not is_protected
            and not is_excluded_domain
            and not is_excluded_addr
        ):
            rec.matched_quarantine_rule = True
            rec.classification = Classification.QUARANTINE
            triggers = []
            if unsub_header:
                triggers.append("List-Unsubscribe header")
            if unsub_body:
                triggers.append("unsubscribe in body")
            triggers.append(f"unread + older than {self.cfg['quarantine_age_days']}d")
            if domain in self.recently_engaged_domains:
                rec.keep_subscription = True
                triggers.append(f"keep_subscription (read within {self.cfg.get('engagement_window_days', 60)}d)")
            rec.rule_triggers = triggers
            return rec

        # ── 4) Untouched ──────────────────────────────────
        rec.classification = Classification.UNTOUCHED
        return rec


# ============================================================
# Report generator
# ============================================================
class ReportGenerator:
    def __init__(self, config: dict, records: list[MessageRecord]):
        self.cfg = config
        self.records = records
        self.sample_size = config.get("report", {}).get("sample_size", 20)
        self.include_snippet = config.get("report", {}).get(
            "include_body_snippet", False
        )
        self.buckets: dict[Classification, list[MessageRecord]] = defaultdict(list)
        for r in records:
            self.buckets[r.classification].append(r)

    def _bucket_sample(self, cls: Classification) -> list[MessageRecord]:
        return self.buckets[cls][: self.sample_size]

    def generate_markdown(self, path: Path):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        total = len(self.records)
        lines = [
            f"# Email Cleanser — Dry Run Report",
            f"**Generated:** {ts}  ",
            f"**Total messages scanned:** {total:,}",
            "",
            "## Summary",
            "",
            "| Bucket | Count | % |",
            "|--------|------:|---:|",
        ]
        for cls in Classification:
            cnt = len(self.buckets[cls])
            pct = (cnt / total * 100) if total else 0
            lines.append(f"| {cls.value} | {cnt:,} | {pct:.1f}% |")

        # Samples per bucket
        for cls in Classification:
            items = self._bucket_sample(cls)
            if not items:
                continue
            lines.append("")
            lines.append(f"## {cls.value} — sample (top {self.sample_size})")
            lines.append("")
            is_quarantine = cls == Classification.QUARANTINE
            if is_quarantine:
                header = "| # | Sender | Subject | Date | Keep Sub? | Triggers |"
                sep = "|---|--------|---------|------|:---------:|----------|"
            else:
                header = "| # | Sender | Subject | Date | Triggers |"
                sep = "|---|--------|---------|------|----------|"
            if self.include_snippet:
                header = header.rstrip("|") + " Snippet |"
                sep = sep.rstrip("|") + "---------|"
            lines.append(header)
            lines.append(sep)
            for i, rec in enumerate(items, 1):
                sender = self._esc(rec.sender_email)
                subj = self._esc(rec.subject[:80])
                dt = rec.received_datetime.strftime("%Y-%m-%d")
                triggers = self._esc("; ".join(rec.rule_triggers))
                if is_quarantine:
                    keep = "✓" if rec.keep_subscription else ""
                    row = f"| {i} | {sender} | {subj} | {dt} | {keep} | {triggers} |"
                else:
                    row = f"| {i} | {sender} | {subj} | {dt} | {triggers} |"
                if self.include_snippet:
                    row = row.rstrip("|") + f" {self._esc(rec.body_snippet)} |"
                lines.append(row)

        # Top senders in quarantine
        q = self.buckets[Classification.QUARANTINE]
        if q:
            domain_counts: dict[str, int] = defaultdict(int)
            keep_counts: dict[str, int] = defaultdict(int)
            for r in q:
                domain_counts[r.sender_domain] += 1
                if r.keep_subscription:
                    keep_counts[r.sender_domain] += 1
            top = sorted(domain_counts.items(), key=lambda x: -x[1])[:30]
            lines.append("")
            lines.append("## Quarantine — top sender domains")
            lines.append("")
            lines.append("| Domain | Count | Keep Sub? |")
            lines.append("|--------|------:|:---------:|")
            for dom, cnt in top:
                keep = "✓" if dom in keep_counts else ""
                lines.append(f"| {dom} | {cnt:,} | {keep} |")

            # Keep-subscription summary
            keep_domains = sorted(keep_counts.items(), key=lambda x: -x[1])
            if keep_domains:
                lines.append("")
                lines.append("## Quarantine — keep-subscription domains")
                lines.append("")
                lines.append(
                    "_These senders have quarantined emails but you've recently read "
                    "their newsletters — the subscription will be preserved in v2._"
                )
                lines.append("")
                lines.append("| Domain | Quarantined emails |")
                lines.append("|--------|-------------------:|")
                for dom, cnt in keep_domains:
                    lines.append(f"| {dom} | {cnt:,} |")

        path.write_text("\n".join(lines), encoding="utf-8")
        print(f"  📝 Markdown report: {path}")

    def generate_csv(self, path: Path):
        fieldnames = [
            "classification",
            "sender_email",
            "sender_domain",
            "subject",
            "received_date",
            "is_read",
            "has_unsubscribe",
            "keep_subscription",
            "rule_triggers",
        ]
        if self.include_snippet:
            fieldnames.append("body_snippet")

        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for rec in self.records:
                row = {
                    "classification": rec.classification.value,
                    "sender_email": rec.sender_email,
                    "sender_domain": rec.sender_domain,
                    "subject": rec.subject,
                    "received_date": rec.received_datetime.strftime(
                        "%Y-%m-%d %H:%M"
                    ),
                    "is_read": rec.is_read,
                    "has_unsubscribe": rec.has_unsubscribe_signal,
                    "keep_subscription": rec.keep_subscription,
                    "rule_triggers": "; ".join(rec.rule_triggers),
                }
                if self.include_snippet:
                    row["body_snippet"] = rec.body_snippet
                writer.writerow(row)
        print(f"  📊 CSV report:      {path}")

    @staticmethod
    def _esc(text: str) -> str:
        return text.replace("|", "\\|").replace("\n", " ").replace("\r", "")


# ============================================================
# Execution engine (v2)
# ============================================================
class ExecutionEngine:
    """Moves classified emails to their destination folders inline as they are classified."""

    ACTIONABLE = {Classification.NEWSLETTER, Classification.QUARANTINE, Classification.RECEIPT}

    def __init__(self, client: GraphClient, config: dict):
        self.client = client
        self.cfg = config
        self.folder_ids: dict[Classification, str] = {}
        self.results: dict = defaultdict(lambda: {"moved": 0, "failed": 0})

    def prepare(self):
        """Resolve destination folder IDs once upfront, creating folders if needed."""
        folders_cfg = self.cfg.get("folders", {})
        folder_map = {
            Classification.NEWSLETTER: folders_cfg.get("newsletters", "Newsletters"),
            Classification.QUARANTINE: folders_cfg.get("quarantine", "CLEANSE_REVIEW"),
            Classification.RECEIPT: folders_cfg.get("receipts", "Receipts"),
        }
        print("📁 Resolving destination folders …")
        for cls, name in folder_map.items():
            self.folder_ids[cls] = self.client.get_or_create_folder(name)
            print(f"  {cls.value:<12} → {name}")
        print()

    def execute_one(self, rec: MessageRecord):
        """Move a single message if it is actionable. Tracks results."""
        if rec.classification not in self.ACTIONABLE:
            return
        try:
            self.client.move_message(rec.message_id, self.folder_ids[rec.classification])
            self.results[rec.classification]["moved"] += 1
        except Exception:
            self.results[rec.classification]["failed"] += 1

    def print_summary(self):
        """Print moved/failed counts per classification."""
        print("\n" + "=" * 50)
        print("EXECUTION COMPLETE")
        print("=" * 50)
        if not self.results:
            print("  Nothing was moved.")
        for cls, counts in self.results.items():
            print(f"  {cls.value:<12}  moved: {counts['moved']:,}  failed: {counts['failed']:,}")
        print("=" * 50)


# ============================================================
# Main entry point
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="Email Cleanser — Dry-run inbox scanner"
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Limit how many messages to scan (useful for testing)",
    )
    parser.add_argument(
        "--config",
        type=str,
        default=str(CONFIG_FILE),
        help="Path to config YAML file",
    )
    parser.add_argument(
        "--reauth",
        action="store_true",
        help="Clear cached credentials and sign in fresh (use this to switch accounts)",
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Execute moves — actually move emails to destination folders (default: dry-run only)",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Ignore any saved progress and start from the beginning (use with --execute)",
    )
    args = parser.parse_args()

    # ── Load config ────────────────────────────────────────
    config = load_config(Path(args.config))
    print("✅ Config loaded")

    # ── Authenticate ───────────────────────────────────────
    print("\n🔑 Authenticating with Microsoft Graph …")
    auth = GraphAuth()
    if args.reauth:
        auth.clear_cache()
        print("  Re-authentication requested — you will be prompted to sign in.")
    auth.get_token()  # ensure initial token is cached
    client = GraphClient(auth)

    # Verify and display which account we're connected to
    try:
        me = client._get(f"{GRAPH_BASE}/me", params={"$select": "displayName,mail,userPrincipalName"})
        email = me.get("mail") or me.get("userPrincipalName") or "unknown"
        name = me.get("displayName") or ""
        print(f"✅ Authenticated as: {name} <{email}>")
        print("   If this is the wrong account, re-run with --reauth to switch.\n")
    except Exception:
        print("✅ Authenticated\n")

    # ── Build protected thread index ───────────────────────
    protected = client.get_sent_conversation_ids()
    print(f"✅ Protected threads: {len(protected):,}\n")

    # ── Load progress (execute mode only) ─────────────────────
    progress_file = Path(__file__).parent / "progress.json"
    resume_after: Optional[str] = None
    engaged_domains: set[str] = set()
    total_processed = 0

    if args.execute:
        if args.fresh and progress_file.exists():
            progress_file.unlink()
            print("🗑️  Progress file cleared — starting fresh.\n")
        elif not args.fresh and progress_file.exists():
            try:
                prog = json.loads(progress_file.read_text(encoding="utf-8"))
                resume_after = prog.get("resume_after")
                engaged_domains = set(prog.get("engaged_domains", []))
                total_processed = prog.get("processed_count", 0)
                print(f"↩️  Resuming — {total_processed:,} already processed, skipping emails newer than {resume_after}\n")
            except Exception:
                print("  ⚠️ Could not read progress file — starting fresh.\n")

    # ── Setup engine and classifier ────────────────────────────
    records = []
    engine = None
    if args.execute:
        engine = ExecutionEngine(client, config)
        engine.prepare()

    classifier = EmailClassifier(config, protected, engaged_domains)

    # ── Per-page fetch + classify (+ inline move) ──────────────
    print("🔍 Classifying and moving messages …" if args.execute else "🔍 Fetching and classifying messages …")
    classify_errors = 0

    try:
        any_pages = False
        for page_msgs in client.iter_inbox_pages(
            limit=args.limit,
            resume_after=resume_after,
        ):
            any_pages = True
            # Update engagement index from this page before classifying it
            engaged_domains |= EmailClassifier.build_engaged_domains(page_msgs, config)

            for msg in page_msgs:
                try:
                    rec = classifier.classify(msg)
                    records.append(rec)
                    if engine:
                        engine.execute_one(rec)
                except Exception as e:
                    classify_errors += 1
                    if classify_errors <= 3:
                        print(f"  ⚠️ Classification error (skipping message): {e}")
                    elif classify_errors == 4:
                        print("  ⚠️ Further classification errors suppressed …")

            total_processed += len(page_msgs)
            print(f"  Total processed: {total_processed:,}", flush=True)

            # Save progress after each page (execute mode only)
            if args.execute and page_msgs:
                oldest = min(page_msgs, key=lambda m: m["receivedDateTime"])
                progress_file.write_text(json.dumps({
                    "resume_after": oldest["receivedDateTime"],
                    "engaged_domains": list(engaged_domains),
                    "processed_count": total_processed,
                }, indent=2), encoding="utf-8")

        if not any_pages:
            print("No messages found. Exiting.")
            return

        if classify_errors:
            print(f"  ⚠️ {classify_errors} message(s) skipped due to classification errors")
        print("✅ Done\n")

        # Clean up progress file on successful completion
        if args.execute and progress_file.exists():
            progress_file.unlink()

    except Exception as e:
        print(f"\n⚠️ Processing stopped early: {e}")
        if not records:
            print("  No messages classified yet. Exiting.")
            return
        print(f"  Proceeding with {len(records):,} already-processed messages …\n")

    if args.execute and engine:
        engine.print_summary()

    # ── Generate reports ───────────────────────────────────
    REPORT_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    formats = config.get("report", {}).get("output_formats", ["markdown"])

    reporter = ReportGenerator(config, records)

    prefix = "execute" if args.execute else "dryrun"
    if "markdown" in formats:
        md_path = REPORT_DIR / f"{prefix}_{timestamp}.md"
        reporter.generate_markdown(md_path)

    if "csv" in formats:
        csv_path = REPORT_DIR / f"{prefix}_{timestamp}.csv"
        reporter.generate_csv(csv_path)

    # ── Print quick summary to console ─────────────────────
    print("\n" + "=" * 50)
    print("EXECUTION SUMMARY" if args.execute else "DRY RUN SUMMARY")
    print("=" * 50)
    for cls in Classification:
        cnt = len(reporter.buckets[cls])
        pct = (cnt / len(records) * 100) if records else 0
        icon = {
            Classification.NEWSLETTER: "📰",
            Classification.RECEIPT: "🧾",
            Classification.QUARANTINE: "🗑️ ",
            Classification.UNTOUCHED: "✉️ ",
        }[cls]
        print(f"  {icon} {cls.value:<12} {cnt:>7,}  ({pct:.1f}%)")
    print("=" * 50)
    print("Done! Review the reports above, then adjust config.yaml and re-run.\n")


if __name__ == "__main__":
    main()
