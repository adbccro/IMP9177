# IMP9177 — Document Review Gate: Power Automate + Adaptive Card
# Claude Code Session Spec
# Generated: 2026-04-27

## Mission
Build a 21 CFR Part 11 compliant document review gate for DCO approvals.
Replace the current Power Automate approval email (which bypasses the portal
document-open gate) with an Adaptive Card email that:
  1. Lists every document in the DCO
  2. Shows live open/not-open status per document (refreshes on card reopen)
  3. Locks the Approve button until all documents have been opened in SharePoint
  4. Captures document-open evidence organically from the M365 Unified Audit Log

## License confirmed
- Microsoft 365 E5 Information Protection and Governance — UAL full audit events
- Exchange Online Plan 1 — Adaptive Cards + Action.Execute in OWA/Outlook
- Project Plan 3 + 5 — Power Automate Premium connectors

## SharePoint context
Site:     https://adbccro.sharepoint.com/sites/IMP9177
ClientId: ba48ac81-6f23-43bd-9797-ec2866071102
Drive ID: b!TjkbwMYC5EiRJVzpwdOMb7D2R8-AVVZEpk9Rx3u59Y8oTtzCk9OzQZtdxpr5MeqU

## Key SP lists
QMS_DCOs           — DCO records (DCO_Phase, DCO_Docs, DCO_Title, DCO_Originator)
QMS_DCOApprovals   — Approver rows (Sig_ApproverEmail, Appr_Status, Appr_DCOID)
QMS_RoutingHistory — Audit trail (AL_EventType, AL_DCOID, AL_Actor, AL_Timestamp,
                                  AL_DocID, AL_Hash, AL_PrevHash)
QMS_Config         — Key/value config (key='anthropic_api_key', etc.)

## Document zones
Published: Shared Documents/Published/QMS/Documents
           Shared Documents/Published/QMS/Forms
           Shared Documents/Published/QMS/Quality Manual
These are the URLs approvers open. FileAccessed events fire against these paths.

## What to build — 3 flows + 1 Adaptive Card JSON

### Flow 1: UAL Audit Subscriber (IMP9177 — UAL FileAccessed Listener)
Purpose: Subscribe to the Office 365 Management Activity API and write
         DocumentOpened events to QMS_RoutingHistory when an approver opens
         a DCO document in SharePoint.

Trigger:  Recurrence — every 5 minutes
Steps:
  1. HTTP GET to start/continue subscription:
     https://manage.office.com/api/v1.0/{tenantId}/activity/feed/subscriptions/start
     contentType: Audit.SharePoint
     Auth: OAuth with resource=https://manage.office.com
  2. HTTP GET to list available content:
     https://manage.office.com/api/v1.0/{tenantId}/activity/feed/subscriptions/content
     ?contentType=Audit.SharePoint&startTime=...&endTime=...
  3. For each content blob URL returned:
     HTTP GET the blob — parse array of audit records
  4. Filter records where:
     - Operation = "FileAccessed"
     - SiteUrl contains "IMP9177"
     - ObjectId (file path) contains "/Published/QMS/"
     - UserId = any active approver email in QMS_DCOApprovals
  5. For each matching record:
     - Extract: UserId (actor), ObjectId (file path), CreationTime
     - Derive DocID from ObjectId (filename without extension, match against DCO_Docs)
     - Find matching DCO: query QMS_DCOApprovals for approver email, get Appr_DCOID
     - Check QMS_RoutingHistory — skip if DocumentOpened already exists for
       this (DCOID + DocID + ActorEmail) combination
     - If new: POST to QMS_RoutingHistory:
         AL_EventType: "DocumentOpened"
         AL_DCOID: <derived DCO ID>
         AL_DocID: <derived doc ID>
         AL_Actor: <UserId>
         AL_Timestamp: <CreationTime>
         AL_Source: "UAL_FileAccessed"
         Title: "DocumentOpened-<DCOID>-<DocID>-<actor initials>"
         AL_Hash: SHA256(AL_EventType + AL_DCOID + AL_DocID + AL_Actor + AL_Timestamp)
         AL_PrevHash: <last hash in QMS_RoutingHistory for this DCOID>

Notes:
  - The UAL has ~15 min latency. Flow runs every 5 min but events may be
    15-20 min behind real time. This is acceptable — gate check happens
    at approve-click time, not continuously.
  - Use the Flow's connection to manage.office.com (HTTP with Azure AD auth,
    audience = https://manage.office.com). Requires admin consent for
    ActivityFeed.Read application permission on the AAD app registration
    (ba48ac81-6f23-43bd-9797-ec2866071102).
  - Store last-processed timestamp in QMS_Config (key='ual_last_processed')
    to avoid reprocessing old events on each poll.

### Flow 2: DCO Review Email (IMP9177 — DCO Review Notification)
Purpose: When a DCO is submitted, send an Adaptive Card email to all
         Required approvers listing every document with direct SP links.

Trigger: SharePoint — When an item is created or modified
         List: QMS_DCOs
         Condition: DCO_Phase equals "Submitted"
         (Add condition to skip if already sent: check QMS_RoutingHistory
          for ReviewEmailSent event for this DCO)

Steps:
  1. Get DCO details: Id, DCO_Title, DCO_Docs, DCO_Originator, Title (DCO_ID)
  2. Parse DCO_Docs (comma-separated doc IDs) into array
  3. For each doc ID, resolve SharePoint file URL:
     - Query Shared Documents/Published/QMS/ for file matching doc ID
     - Build direct URL: https://adbccro.sharepoint.com/sites/IMP9177/
       Shared Documents/Published/QMS/Documents/<filename>
     - If not found in Documents, check Forms, then Quality Manual
  4. Get approvers: query QMS_DCOApprovals where Appr_DCOID = DCO ID
     and Appr_Type = "Required" — collect Sig_ApproverEmail values
  5. Build Adaptive Card JSON (see adaptive_card_template.json)
     - Inject: DCO_Title, DCO_ID, originator, due date (+5 business days)
     - Inject: doc list with SP URLs, count
     - Refresh URL: https://<flow-http-trigger-url>/gate-status?dcoId=<ID>
  6. Send email via Office 365 Outlook — Send email with options
     - To: each Required approver
     - Subject: [IMP9177] DCO <ID> — Document review required
     - Body: Adaptive Card JSON (set Importance: High)
  7. Write ReviewEmailSent to QMS_RoutingHistory:
     AL_EventType: "ReviewEmailSent"
     AL_DCOID: <DCO ID>
     AL_Actor: "System"

### Flow 3: Approval Gate Endpoint (IMP9177 — Approval Gate HTTP)
Purpose: Two endpoints called by the Adaptive Card:
  A. GET /gate-status?dcoId=X  — returns current open/not-open per doc (card refresh)
  B. POST /approve             — verifies gate, PATCHes DCO, writes audit record

Trigger: When an HTTP request is received (Premium — requires Plan 3/5)

Endpoint A — GET /gate-status:
  1. Parse dcoId from query string
  2. Get DCO record: DCO_Docs list
  3. Query QMS_RoutingHistory for all DocumentOpened events for this dcoId
     grouped by AL_DocID + AL_Actor (the requesting approver)
  4. Build response JSON:
     {
       "dcoId": "DCO-0001",
       "dcoTitle": "Week 1 QMS Package",
       "totalDocs": 15,
       "openedDocs": 3,
       "docs": [
         { "docId": "QM-001", "name": "Quality Manual Rev B",
           "opened": true, "openedAt": "2026-04-28T09:42:00Z",
           "url": "https://adbccro.sharepoint.com/..." },
         { "docId": "SOP-QMS-001", "name": "Management Responsibility",
           "opened": false, "openedAt": null,
           "url": "https://adbccro.sharepoint.com/..." }
       ],
       "gateOpen": false,
       "approverEmail": "tinaqwork@gmail.com"
     }
  5. Return 200 with JSON — Adaptive Card uses this to render checkmarks

Endpoint B — POST /approve:
  Request body: { "dcoId": "DCO-0001", "approverEmail": "...",
                  "action": "approve"|"reject", "reason": "" }
  1. Verify gate: query QMS_RoutingHistory — count DocumentOpened events
     for this dcoId + approverEmail. Must equal DCO_Docs count.
  2. If gate not satisfied:
     Return 400: { "error": "Not all documents reviewed",
                   "opened": N, "required": M }
     Adaptive Card shows error message, Approve stays locked.
  3. If gate satisfied and action = "approve":
     a. PATCH QMS_DCOs: DCO_Phase = "Approved", DCO_ApprovedDate = now
     b. PATCH QMS_DCOApprovals: Appr_Status = "Approved" for this approver
     c. POST QMS_RoutingHistory:
          AL_EventType: "DCOApproved"
          AL_DCOID: <dcoId>
          AL_Actor: <approverEmail>
          AL_Note: "Approved via Adaptive Card email — all documents reviewed"
          AL_Hash: SHA256(...)
          AL_PrevHash: <last hash>
     d. Return 200: { "success": true, "message": "DCO approved" }
  4. If action = "reject":
     a. Require non-empty reason in request body
     b. PATCH QMS_DCOs: DCO_Phase = "Rejected"
     c. Write DCORejected audit event with reason
     d. Send rejection notification to DCO originator
     e. Return 200: { "success": true }

## Adaptive Card JSON
See adaptive_card_template.json in this folder.
Key features:
  - refresh action pointing to Flow 3 Endpoint A
  - ColumnSet for each document: checkmark icon + doc name + open link
  - ProgressBar showing N of M opened (TextBlock with color logic)
  - Approve/Reject Action.Execute buttons pointing to Flow 3 Endpoint B
  - Approve button disabled (isEnabled: false) when gateOpen = false
    (use Adaptive Card templating: "isEnabled": "${gateOpen}")

## AAD App Registration changes needed
The existing app (ba48ac81-6f23-43bd-9797-ec2866071102) needs one new
API permission for the UAL subscriber Flow:
  API: Office 365 Management APIs
  Permission: ActivityFeed.Read (Application permission)
  Admin consent: required

Add this in Azure Portal → App registrations → IMP9177-PnP-Script
→ API permissions → Add a permission → APIs my org uses →
  "Office 365 Management APIs" → Application permissions → ActivityFeed.Read
→ Grant admin consent

## Jira tickets to create after build
Create these tickets and mark Done as each component ships:
  IMP9177-29 — AAD app permission: ActivityFeed.Read + admin consent
  IMP9177-30 — Flow 1: UAL FileAccessed listener → QMS_RoutingHistory
  IMP9177-31 — Flow 2: DCO review Adaptive Card email
  IMP9177-32 — Flow 3: Approval gate HTTP endpoint (A + B)
  IMP9177-33 — Adaptive Card JSON template + refresh/approve wiring
  IMP9177-34 — Remove old Power Automate approval action from existing DCO flow
  IMP9177-35 — Smoke test: end-to-end DCO submit → email → open docs → approve

## Existing DCO flow to modify
Flow name: "IMP9177 — DCO Approval Routing"
Current behavior: sends M365 Approval action email → approver clicks
  Approve in Outlook → flow advances DCO_Phase
Required change: REMOVE the "Start and wait for an approval" action.
Replace with: Send notification email only (no approve button) pointing
  to the portal. The actual approval now comes from Flow 3 Endpoint B.
This is the critical change that closes the compliance gap.

## Build order
1. AAD permission (manual — admin must consent in Azure Portal)
2. Flow 3 (gate endpoint) — build first so the card has a real URL to point to
3. Adaptive Card JSON — test against Flow 3 manually
4. Flow 2 (email sender) — wire in the card + Flow 3 URL
5. Flow 1 (UAL listener) — requires ActivityFeed.Read consent to be done
6. Modify existing DCO flow — remove approval action, add notify-only email
7. End-to-end smoke test
8. Create and close Jira tickets

## Testing notes
- UAL latency is 15-20 min. For smoke test: open a doc, wait 20 min,
  check QMS_RoutingHistory for DocumentOpened event, then test approve.
- Alternatively: manually POST a DocumentOpened event to QMS_RoutingHistory
  for the test DCO + your email to simulate the UAL write, then test the
  gate endpoint immediately.
- Adaptive Card refresh: open the email, open a doc in SP, wait 20 min,
  reopen the email — checkmark should appear. Or POST manually and reopen.
- Flow 3 HTTP trigger URL will be generated when Flow is saved —
  copy it into the Adaptive Card JSON refresh.url before deploying Flow 2.
