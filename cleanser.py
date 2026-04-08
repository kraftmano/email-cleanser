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

    def __init__(self, token: str):
        self.token = token
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            }
        )

    def _get(self, url: str, params: dict = None) -> dict:
        resp = self.session.get(url, params=params)
        if resp.status_code == 429:
            retry_after = int(resp.headers.get("Retry-After", 5))
            print(f"  ⏳ Throttled — waiting {retry_after}s …")
            time.sleep(retry_after)
            return self._get(url, params)
        resp.raise_for_status()
        return resp.json()

    def get_inbox_folder_id(self) -> str:
        data = self._get(f"{GRAPH_BASE}/me/mailFolders/Inbox")
        return data["id"]

    def get_inbox_messages(
        self,
        limit: Optional[int] = None,
        include_body: bool = False,
    ) -> list[dict]:
        """
        Fetch all Inbox messages (paginated).
        Returns list of raw Graph message dicts.
        """
        fields = (
            "id,subject,sender,from,receivedDateTime,isRead,"
            "conversationId,internetMessageHeaders"
        )
        if include_body:
            fields += ",bodyPreview"

        url = f"{GRAPH_BASE}/me/mailFolders/Inbox/messages"
        params = {
            "$select": fields,
            "$top": min(self.PAGE_SIZE, limit) if limit else self.PAGE_SIZE,
            "$orderby": "receivedDateTime desc",
        }

        all_messages = []
        page = 1
        while url:
            print(f"  📬 Fetching page {page} …", end=" ", flush=True)
            data = self._get(url, params)
            batch = data.get("value", [])
            all_messages.extend(batch)
            print(f"got {len(batch)} messages (total: {len(all_messages)})")

            if limit and len(all_messages) >= limit:
                all_messages = all_messages[:limit]
                break

            url = data.get("@odata.nextLink")
            params = None  # nextLink includes params
            page += 1

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
    def __init__(self, config: dict, protected_threads: set[str]):
        self.cfg = config
        self.protected_threads = protected_threads
        self.cutoff = datetime.now(timezone.utc) - timedelta(
            days=config.get("quarantine_age_days", 30)
        )

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
            header = "| # | Sender | Subject | Date | Triggers |"
            sep = "|---|--------|---------|------|----------|"
            if self.include_snippet:
                header = "| # | Sender | Subject | Date | Triggers | Snippet |"
                sep = "|---|--------|---------|------|----------|---------|"
            lines.append(header)
            lines.append(sep)
            for i, rec in enumerate(items, 1):
                sender = self._esc(rec.sender_email)
                subj = self._esc(rec.subject[:80])
                dt = rec.received_datetime.strftime("%Y-%m-%d")
                triggers = self._esc("; ".join(rec.rule_triggers))
                row = f"| {i} | {sender} | {subj} | {dt} | {triggers} |"
                if self.include_snippet:
                    row = f"| {i} | {sender} | {subj} | {dt} | {triggers} | {self._esc(rec.body_snippet)} |"
                lines.append(row)

        # Top senders in quarantine
        q = self.buckets[Classification.QUARANTINE]
        if q:
            domain_counts: dict[str, int] = defaultdict(int)
            for r in q:
                domain_counts[r.sender_domain] += 1
            top = sorted(domain_counts.items(), key=lambda x: -x[1])[:30]
            lines.append("")
            lines.append("## Quarantine — top sender domains")
            lines.append("")
            lines.append("| Domain | Count |")
            lines.append("|--------|------:|")
            for dom, cnt in top:
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
    token = auth.get_token()
    client = GraphClient(token)

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

    # ── Fetch Inbox messages ───────────────────────────────
    include_body = config.get("report", {}).get("include_body_snippet", False)
    print("📥 Fetching Inbox messages …")
    raw_messages = client.get_inbox_messages(
        limit=args.limit,
        include_body=include_body,
    )
    print(f"✅ Fetched {len(raw_messages):,} messages\n")

    if not raw_messages:
        print("No messages found. Exiting.")
        return

    # ── Classify ───────────────────────────────────────────
    print("🔍 Classifying messages …")
    classifier = EmailClassifier(config, protected)
    records = [classifier.classify(msg) for msg in raw_messages]
    print("✅ Classification complete\n")

    # ── Generate reports ───────────────────────────────────
    REPORT_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    formats = config.get("report", {}).get("output_formats", ["markdown"])

    reporter = ReportGenerator(config, records)

    if "markdown" in formats:
        md_path = REPORT_DIR / f"dryrun_{timestamp}.md"
        reporter.generate_markdown(md_path)

    if "csv" in formats:
        csv_path = REPORT_DIR / f"dryrun_{timestamp}.csv"
        reporter.generate_csv(csv_path)

    # ── Print quick summary to console ─────────────────────
    print("\n" + "=" * 50)
    print("DRY RUN SUMMARY")
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
