# Email Cleanser

A personal, rerunnable inbox scanner for Microsoft 365. Classifies emails into Newsletters, Receipts, Quarantine (unwanted marketing), or Untouched — then produces a report so you can review before any messages are moved.

**v1 is dry-run only.** Nothing is moved or deleted.

---

## Setup Instructions

### Step 1 — Install Python

You need Python 3.10+ on Windows. Check with:

```
python --version
```

If not installed, download from https://www.python.org/downloads/ and ensure "Add to PATH" is checked during install.

### Step 2 — Download the project

Put the entire `email-cleanser` folder somewhere on your machine (e.g. `C:\Users\YOU\email-cleanser`).

### Step 3 — Install dependencies

Open a terminal in the project folder and run:

```
pip install -r requirements.txt
```

### Step 4 — Register an Azure AD App (one-time)

This gives the script permission to read your mailbox via Microsoft Graph.

1. Go to https://portal.azure.com → **Azure Active Directory** → **App registrations** → **New registration**
2. Fill in:
   - **Name:** `Email Cleanser` (anything you like)
   - **Supported account types:** "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI:** select **Public client/native (mobile & desktop)** and enter `https://login.microsoftonline.com/common/oauth2/nativeclient`
3. Click **Register**
4. On the app's overview page, copy the **Application (client) ID** — this is your Client ID
5. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
   - Search for and add: `Mail.ReadWrite`
6. Click **Grant admin consent** if you have admin access, or just proceed (you'll consent on first login)

### Step 5 — Configure your Client ID

Copy `.env.example` to `.env` and paste your Client ID:

```
copy .env.example .env
```

Edit `.env`:

```
EMAIL_CLEANSER_CLIENT_ID=a1b2c3d4-e5f6-7890-abcd-ef1234567890
```

### Step 6 — Run the dry-run

```
python cleanser.py
```

On first run it will prompt you to sign in via your browser using a device code. After that, the token is cached locally.

To test with a small batch first:

```
python cleanser.py --limit 200
```

### Step 7 — Review the report

Reports are saved in the `reports/` folder:
- `dryrun_YYYYMMDD_HHMMSS.md` — Markdown summary with sample tables
- `dryrun_YYYYMMDD_HHMMSS.csv` — Full CSV of all classified messages (open in Excel to filter/sort)

### Step 8 — Tune and repeat

Edit `config.yaml` to:
- Add/remove newsletter domains
- Adjust receipt keywords or trusted merchant domains
- Add sender domains or addresses to the exclusion list
- Change the quarantine age threshold

Then re-run `python cleanser.py` and check the new report.

---

## Project Files

| File | Purpose |
|------|---------|
| `cleanser.py` | Main script — authenticates, fetches, classifies, reports |
| `config.yaml` | All classification rules (edit this to tune) |
| `.env` | Your Azure Client ID (not committed to git) |
| `.env.example` | Template for `.env` |
| `requirements.txt` | Python dependencies |
| `reports/` | Generated reports (gitignored) |
| `.token_cache.json` | Cached auth token (gitignored) |

---

## How Classification Works

| Priority | Bucket | Condition |
|----------|--------|-----------|
| 1 | **Newsletter** | Sender domain is in `newsletter_domains` list |
| 2 | **Receipt** | Subject matches a `receipt_keywords` entry |
| 3 | **Quarantine** | Unread AND older than 30 days AND has unsubscribe signal AND not a receipt AND not a thread you've replied to AND not excluded |
| 4 | **Untouched** | Everything else stays in Inbox |

Thread protection: the script scans your Sent Items to find every conversation you've participated in. Those threads are never quarantined.

---

## Troubleshooting

**"Missing dependencies"** → Run `pip install -r requirements.txt`

**"No Azure AD Client ID configured"** → Create `.env` file with your Client ID (see Step 5)

**Auth errors** → Delete `.token_cache.json` and re-run to sign in fresh

**Throttling (429 errors)** → The script auto-retries. If persistent, wait a few minutes and try `--limit 500`

**Large inbox is slow** → Use `--limit` for testing, then run full scan overnight. The Graph API paginated fetch handles 115k+ messages but it takes time.
