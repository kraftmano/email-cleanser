# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Full dry-run scan
python cleanser.py

# Scan a limited batch (for testing)
python cleanser.py --limit 200

# Force re-authentication (switch accounts)
python cleanser.py --reauth

# Debug auth / inspect connected account
python test.py
```

## Architecture

This is a single-file Python script (`cleanser.py`) that connects to Microsoft 365 via the Microsoft Graph API to classify Inbox emails — no messages are moved or deleted in v1.

**Auth flow:** `GraphAuth` uses MSAL device code flow. The Azure AD Client ID is loaded from `EMAIL_CLEANSER_CLIENT_ID` env var or a `.env` file. Tokens are cached in `.token_cache.json`.

**Pipeline:**
1. `GraphAuth.get_token()` → OAuth token via device code
2. `GraphClient.get_sent_conversation_ids()` → builds a set of conversation IDs the user has replied to (thread protection)
3. `GraphClient.get_inbox_messages()` → paginated fetch of inbox (250 per page, handles 429 throttling with auto-retry)
4. `EmailClassifier.classify()` → classifies each message in priority order: Newsletter → Receipt → Quarantine → Untouched
5. `ReportGenerator` → writes timestamped `.md` and/or `.csv` to `reports/`

**Classification priority** (first match wins):
1. **Newsletter**: sender domain in `newsletter_domains` list
2. **Receipt**: subject matches `receipt_keywords` (optionally combined with `receipt_sender_domains`)
3. **Quarantine**: unread + older than `quarantine_age_days` + has unsubscribe signal (header or body) + not a receipt + not a protected thread + not in exclusion lists
4. **Untouched**: everything else

**Configuration** (`config.yaml`): all classification rules live here — domain lists, keywords, age thresholds, exclusions, folder names, report settings. Edit this file to tune classification; no code changes needed.

**Planned v2**: execution mode to actually move messages to `Newsletters`, `Receipts`, and `CLEANSE_REVIEW` folders.
