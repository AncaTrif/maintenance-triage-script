# Maintenance Triage Automation

Apps Script bound to the shared Maintenance Google Sheet. Watches for portal maintenance-request
emails, triages new Triage rows with Claude, alerts Slack on emergencies/low-confidence calls, and
logs approved tickets to the Tracker tab. No new hosting, no database — just Sheets + Gmail +
one Claude API call + one Slack webhook.

## Files

- `appsscript.json` — Apps Script manifest (timezone, OAuth scopes).
- `Code.gs` — the entire script.
- `triage_guide.txt` — starting content for the triage guide; the live copy lives in a Google Doc
  (see setup below), this file is just for reference/version history.

## One-time setup (on the shared company Google account)

1. **Bind the script.** Open the Maintenance Sheet on the shared account → Extensions → Apps Script.
   Paste `appsscript.json` content into the manifest (View → Show manifest file) and `Code.gs` into
   the script file. Save.
   - Alternative: use [`clasp`](https://github.com/google/clasp) logged into the shared account —
     `clasp push` from this folder.
2. **Create the triage guide Doc.** In the same shared Drive, create a Google Doc, paste in the
   contents of `triage_guide.txt` (edit as needed — this is meant to be tuned over time), and copy
   its Doc ID from the URL.
3. **Reload the Sheet** so the "Triage Tools" custom menu appears.
4. **Run setup**: Triage Tools → "Run setup (create Triage/Config tabs)". Authorize the script when
   prompted (must be authorized as the shared account, not a personal one).
5. **Fill in the Config tab**:
   - "Gmail label to watch" — defaults to `Maintenance-Portal`.
   - "Triage guide Google Doc ID" — paste the Doc ID from step 2.
6. **Set secrets**: Triage Tools → "Set API keys / webhook". You'll be prompted for:
   - Claude API key
   - Slack incoming webhook URL
   - Dana's Slack member ID (optional — without it, alerts post as plain `@Dana` text, which
     won't actually ping her; get her ID from her Slack profile → "Copy member ID")
   These are stored in Script Properties, never in the sheet or in code.
7. **Create the Gmail filter**: in Gmail on the shared account, create a filter matching the portal's
   notification sender/subject pattern (`New Maintenance Request – Unit ... – ...`) and apply the
   `Maintenance-Portal` label (or whatever you set in step 5) to matching mail.
8. **Install the trigger**: Triage Tools → "Install 15-minute trigger". This scans Gmail and
   processes new Triage rows every 15 minutes going forward.

## How it works

- **Gmail scan** (`scanGmailForNewTickets`): searches `label:Maintenance-Portal` minus an internal
  `Triage-Processed` label, parses Unit / Tenant / Description out of each email body, appends a
  Triage row with Source = Portal, then labels the thread processed so it's never re-read. Blank
  Description or Category selected in the source email is handled without failing.
- **Phone-ins**: front desk adds rows to Triage manually (Unit #, Tenant Name, Tenant Text,
  Source = Phone). No automation needed for this path.
- **Claude triage** (`processTriageRows`): for any row with a blank Suggested Category, sends
  *only* the Tenant Text field to Claude, along with the triage guide (from the Doc) and the
  Vendors sheet as reference — tenant name and unit number are never sent. Writes back Suggested
  Category / Priority / Vendor / Confidence / Confidence Reason.
- **Slack alerts** (`postSlackAlert`): if Priority = Emergency, or Confidence is Low or Medium,
  posts to the configured webhook, tagging Dana.
- **Approval → Tracker** (`onEdit`): a simple edit trigger fires when Status is set to "Approved" on
  a Triage row. It appends exactly one row to Tracker (matched by Tracker's own column headers, so
  Tracker's structure is never assumed or altered), using whatever is currently in Suggested Vendor
  — including any manual override Dana typed in before approving. "Skip" and "Waiting on Tenant
  Callback" do nothing to Tracker. No vendor email is ever sent automatically — this only logs the
  ticket; sending stays manual.

## Notes for whoever maintains this next

- Everything lives in one file (`Code.gs`) on purpose — there's no framework, no build step, just
  functions. Read top to bottom.
- Secrets (API key, webhook, Dana's Slack ID) are in Script Properties
  (Project Settings → Script Properties in the Apps Script editor), not in this repo.
- The triage guide is a Google Doc, not code — edit it directly to change triage behavior without
  touching the script.
- If a Triage row's Suggested Category stays blank after a scan, check Stackdriver/Apps Script
  execution logs — it usually means the Claude API call failed or the API key isn't set; it will
  retry automatically on the next 15-minute run.

## Known issue to revisit

- Approving a Triage row writes `Status = "Sent to Vendor"` into Tracker, but nothing was actually
  sent — no vendor email goes out automatically (by design, per the original spec). That label
  describes intent ("approved, ready for dispatch"), not a completed action, and reads as
  misleading at a glance. Left as-is for now since the exact string was part of the original spec
  and Tracker may already have a data-validation dropdown built around it — changing the literal
  text should be a deliberate decision, not a silent edit. If it's ever changed, it's this one line
  in `Code.gs`, inside `onEdit()`:
  ```javascript
  'Status': 'Sent to Vendor'
  ```
  and any matching dropdown validation on Tracker's own Status column would need updating too.
