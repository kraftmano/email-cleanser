# Email Cleanser — Master Plan

## 1) Overview and objectives
### What this is
A personal, rerunnable “email cleanser” that scans only the Outlook Inbox (Microsoft 365) and:
- Moves newsletters into a dedicated folder
- Moves receipts into a dedicated folder
- Quarantines unwanted marketing/sales emails into a review folder (no deletion in v1)

### Why it exists
- Reduce inbox clutter at scale (115k emails)
- Make your inbox usable again without manually triaging thousands of messages
- Provide a safe workflow where you can validate the logic before any irreversible actions

### Success criteria
- After a run, the Inbox contains mostly human/important messages
- Newsletters are consolidated in `Newsletters`
- Receipts are consolidated in `Receipts`
- Unwanted marketing/sales is quarantined in `CLEANSE_REVIEW`
- No important threads you’ve engaged with are moved to quarantine

---

## 2) Core features and functionality
### A) Authentication and access (Microsoft Graph)
- Sign in to Microsoft Graph with least-privilege permissions needed to read and move messages in your mailbox.
- Operate only on the Inbox folder.

### B) Folder management
Ensure the following folders exist (create if missing):
- `CLEANSE_REVIEW`
- `Newsletters`
- `Receipts`

### C) Classification rules (conceptual)
The cleanser categorizes messages into one of these outcomes:

#### 1) Newsletters → move to `Newsletters` (never delete)
Detection approach:
- Sender domain belongs to a “newsletter providers” list (starting with Substack; extensible).
- Optional future: maintain a “keep list” to exempt specific newsletter senders from any additional processing.

Outcome:
- Move matching messages to `Newsletters` regardless of age or read/unread status.

#### 2) Receipts → move to `Receipts` (never delete)
Detection approach:
- Subject and/or sender patterns consistent with receipts and invoices.
- Keyword heuristics (e.g., “receipt”, “invoice”, “order”, “payment”, “tax invoice”, “confirmation”).
- Optional future: maintain a “trusted merchants” sender allowlist.

Outcome:
- Move matching messages to `Receipts`.

#### 3) Unwanted marketing/sales → move to `CLEANSE_REVIEW`
Eligibility (all must be true):
- Message is unread
- Message is older than 1 month
- Message contains an “unsubscribe” signal (body and/or headers)
- Message is NOT classified as a receipt
- Message is NOT part of a thread you have replied to (thread protection)

Outcome:
- Move matching messages to `CLEANSE_REVIEW`

#### 4) Everything else → leave in Inbox
If a message does not satisfy any rule, it remains untouched.

---

## 3) Operating modes and user workflow
### Mode: Level 2 Dry Run (v1 requirement)
The script will not move or delete anything. It will only produce a report.

Dry run output should include:
- Total messages scanned
- Counts per classification bucket:
  - Newsletter candidates
  - Receipt candidates
  - Quarantine candidates
  - Untouched
- A sample set per bucket (e.g., top 20 subjects/senders)
- A summary of the rule triggers that caused classification (to help debug false positives)

### Suggested workflow
1. Run dry run and review the report
2. Adjust:
   - Newsletter provider domains list
   - Receipt keywords and trusted senders
   - Any additional exclusions
3. Repeat until results look correct
4. (Future) Enable execution mode to actually move messages

---

## 4) High-level technical stack recommendations (no implementation detail)
### Language/runtime
- Python on Windows

### APIs
- Microsoft Graph API for:
  - Reading messages from Inbox
  - Accessing message metadata (read/unread, received date, conversation/thread identifiers)
  - (Future) moving messages to folders

### Configuration
- A simple config file (e.g., YAML/JSON) containing:
  - Newsletter provider domains (initially include Substack)
  - Receipt keywords
  - Optional allow/deny lists for senders/domains
  - Date thresholds (1 month marketing/sales; newsletters are always moved but included for future flexibility)

### Reporting
- Generate a local report file (Markdown or HTML) for easy review.
- Optionally also output a CSV for filtering/sorting in Excel.

My recommended stack choices:
- Microsoft Graph (future-proof, robust mailbox access)
- Config-file driven rules (easy iteration without editing logic)
- Markdown/HTML + CSV reporting (fast review, easy auditing)

---

## 5) Conceptual data model
### Entities (conceptual)
#### Message (read-only representation)
- message_id
- subject
- sender_name
- sender_email
- sender_domain
- received_datetime
- is_read
- conversation_id (thread identifier)
- has_unsubscribe_signal (boolean)
- matched_receipt_rule (boolean)
- matched_newsletter_rule (boolean)
- matched_quarantine_rule (boolean)
- final_classification (enum: NEWSLETTER, RECEIPT, QUARANTINE, UNTOUCHED)

#### Thread protection index
- conversation_id → protected (true/false)
- “protected” means: you have replied somewhere in that conversation thread

#### Rule configuration
- newsletter_domains: list
- receipt_keywords: list
- receipt_sender_allowlist (optional)
- quarantine_unsubscribe_signals: list/patterns
- quarantine_age_days: numeric (≈ 30)
- exclusions: domains/senders/keywords (optional)

---

## 6) Security considerations
### Least privilege and access scope
- Request only permissions needed for read access (v1) and later moving messages (v2).
- Restrict logic to Inbox folder to minimize impact radius.

### Secrets handling
- Do not hardcode client secrets or tokens in files committed to git.
- Store credentials/tokens locally using OS-appropriate secure storage where possible.
- Ensure report outputs do not leak sensitive message bodies unnecessarily; default to subject/sender/metadata only.

### Safety rails
- Start with dry run only (no side effects).
- When execution is added:
  - Add a maximum-moves-per-run limit (e.g., 2,000) to prevent runaway actions
  - Add “protected thread” safeguard as a hard rule
  - Add strong receipt protections (keywords + allowlist)
  - Log every moved message ID for rollback auditing

---

## 7) Potential challenges and solutions
### Challenge: Performance scanning 115,000 Inbox emails
- Solution:
  - Use server-side filtering where possible (date ranges, unread status)
  - Paginate and stream results rather than loading everything at once
  - Cache intermediate indexes (e.g., protected thread IDs) to avoid repeated work

### Challenge: False positives from “unsubscribe”
- Solution:
  - Combine “unsubscribe” with bulk-sender heuristics and your explicit exclusions
  - Always exclude receipts/invoices and protected threads
  - Use dry-run sampling to refine receipt keywords and sender allowlists

### Challenge: “Replied-to thread” detection nuances
- Solution:
  - Treat thread protection as conservative: if any reply exists in the conversation, protect the entire conversation from quarantine

### Challenge: Folder naming and localization differences
- Solution:
  - Identify folders by ID after discovery rather than relying solely on display names
  - Store folder IDs locally once resolved

---

## 8) Future expansion possibilities
### V2: Execution mode
- Add a flag to actually move messages after dry-run confidence.
- Keep `CLEANSE_REVIEW` as default “delete staging”.

### V3: Optional auto-delete of quarantine
- Add a second script/action:
  - Delete everything in `CLEANSE_REVIEW` older than X days
- Only after months of confidence.

### Smarter categorization
- Add additional buckets (e.g., “Social”, “Promotions”, “Job alerts”)
- Add a “VIP / humans” allowlist of domains or contacts that should never be moved

### Scheduling / convenience
- Add a simple “run whenever” launcher:
  - Double-clickable shortcut
  - Or a Windows Task Scheduler option (manual/weekly)

### UI improvements
- A simple local dashboard (HTML report with filters)
- “Approve list” export to confirm which messages would be moved

---

## 9) Open questions to confirm before implementation
1. Newsletter providers list beyond Substack: which domains should be included initially?
2. Receipt keyword list: any merchants/senders you want guaranteed to go to `Receipts`?
3. Should “Newsletters” move include Inbox-only messages, or also sweep any subfolders later (v2)?
4. Do you want the dry-run report to include small body snippets for better accuracy, or metadata only?
