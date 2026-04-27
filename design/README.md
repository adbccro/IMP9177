# Handoff: 3H Pharmaceuticals QMS Portal (IMP9177)

> **Audience:** Developer / engineering team picking up the IMP9177 build, working in Claude Code or any IDE.
> **Status:** Designs approved (Apr 2026). Ready to implement.
> **Project:** 3H Pharmaceuticals LLC × ADB Consulting & CRO Inc. — internal compliance software for the 21 CFR Part 111 / FSMA quality program.
> **Source repo to extend:** [`adbccro/IMP9177`](https://github.com/adbccro/IMP9177) (existing SPFx scaffold).

---

## 1 · Overview

This package contains everything needed to build the IMP9177 QMS Portal — a SharePoint-hosted application that manages the document-control lifecycle (Change Requests → Change Orders → signed/effective documents) for a regulated pharmaceutical operation. The system is a 21 CFR Part 11 controlled application, which has three non-negotiable consequences for the build:

1. **Append-only audit trails.** The `Signatures` and `ActivityLog` lists must be cryptographically chained and physically prevent edit/delete, even by site collection admins.
2. **Re-authentication on every binding action.** Sign, reject, delegate — each requires fresh password (or step-up MFA) per Part 11 §11.200.
3. **Validation package.** URS / FRS / IQ / OQ / PQ documents accompany the code. Don't treat this as "ship and iterate."

Two product surfaces:

- **QMS Portal** — interactive multi-tab work surface used daily by Tina Qin (3H QA Approver), Liu Hong (3H Production), QD Yang (3H QA Operations), and Andre Butler (ADB PM). DM Sans + DM Mono. Cool navy palette.
- **Executive PM Tracker** — read-only dashboard for the steering committee. Same data, different visual register: Libre Baskerville serif headings, sparser layout, designed for at-a-glance status.

Both surfaces share **one design system** (one type stack, one color set, one component library) — they read as the same product.

---

## 2 · About the design files in this bundle

The HTML files in `designs/` are **design references**, not production code. They were authored as static prototypes to lock down look, structure, copy, and behavior. **Do not ship them directly.** Your task is to recreate these designs faithfully inside the target codebase's environment using its established patterns.

**Recommended target stack:**
- **SharePoint Framework (SPFx) 1.18+** — extends the existing `adbccro/IMP9177` scaffold. Don't fork; build on top.
- **React 18 + TypeScript** — already in the SPFx scaffold.
- **Fluent UI v9 (`@fluentui/react-components`)** — for form controls, modals, focus trapping, accessibility primitives. The visual treatment from these designs sits *on top of* Fluent UI primitives — keep accessibility behavior, restyle to match.
- **Microsoft Graph + SP REST** — data layer. See `sharepoint-schema.md` for the full list/column structure.
- **MSAL** — auth, including step-up for Part 11 re-authentication.

If you decide on a different stack, you must still satisfy:
- WCAG 2.1 AA (regulated workplace requirement)
- 21 CFR Part 11 audit-trail integrity
- Reading from SharePoint lists per `sharepoint-schema.md`

---

## 3 · Fidelity

**Hi-fi.** Every color, type ramp, spacing value, and interaction in the prototypes is final. Recreate pixel-for-pixel inside Fluent UI, with the design tokens in §6 below as the source of truth.

The only fidelity ambiguities, called out explicitly:

- **Logos.** No real 3H Pharmaceuticals or ADB logo was found in the source repo. The prototypes use a text + "QMS" badge. Replace with real lockups when supplied; the layout reserves space for a 32px-tall mark.
- **Redline preview** — `dco-detail.html` shows a hand-styled redline. Production redline rendering needs a real diff source. Two options:
  1. Word Online iframe pointing at the DOCX in `/Drafts`, with `?embed=true` and Track Changes on.
  2. Pre-rendered redline PDF generated on draft upload (via Aspose.Words, OpenXML SDK, or Office Online conversion) stored in `DCODocuments.RedlineFileRef`.
  Pick #2 for performance; the iframe approach has authentication friction.
- **Signature hash** — the design shows truncated hashes ("`7c14...d9f1`"). The full hash chain spec is in `sharepoint-schema.md` §11.

---

## 4 · Screens / views

The bundle contains three HTML files. Each one packages multiple states — open them in a browser and scroll to see all variants.

### 4.1 `designs/01-calm-pair.html`

Two screens stacked vertically, separated by a horizontal divider. Both read at 1280–1480px design width.

#### Screen 1 — QMS Portal home

| Region | Layout | Notes |
|---|---|---|
| **Header bar** | Full width · `padding: 22px 56px 20px` · `border-bottom: 1px solid #e0e6ed` · white | Left: "QMS" mono badge (8×14px padding, pill, navy on light blue), product title in serif, sub-line in 13px sans grey. Right: 3 status chips (mono, 11px). |
| **Nav strip** | `padding: 0 56px` · 9 tabs · 16px vertical · 2px bottom-border accent on active | Active = navy text + navy underline. Inactive = slate text. Hover = blue text. |
| **Main wrapper** | `max-width: 1480px` · `padding: 48px 56px 96px` · centered | |
| **Eyebrow + greeting** | "Apr 11, 2026 · 14:03" mono eyebrow, `Good afternoon, Andre.` in 30px serif, lede beneath | One sentence quantifying the day. Tone is direct, not chirpy. |
| **KPI strip (6 cells)** | `display: grid; grid-template-columns: repeat(6, 1fr)` · top + bottom 1px borders · cells separated by 1px right-borders | Each cell: mono uppercase label (10.5px), 38px serif numeral (color varies by KPI), 12px slate sublabel. Numerals use color: navy default, red for Past Due, amber for Due Soon, green for Signed, purple for Awaiting Training. |
| **My buckets row** | 3 cards · `display: grid; grid-template-columns: repeat(3, 1fr); gap: 24px` | Each: mono eyebrow + 42px serif count, serif h4 title, sans body, foot row with mono meta. Hover: lift `translateY(-1px)` + navy border. |
| **DCO register** | Full-width table | Filter row (pill tabs + search input) above. Mono uppercase column headers. Row: mono ID (navy), serif title with sans 12px meta below, mono uppercase type, status with colored dot, target date (red if late). Hover row: `background: #f0f4f8` + 12px left padding shift on first cell. |
| **Two-column footer** | `grid-template-columns: 1.5fr 1fr` · `gap: 64px` | Effective documents (left), PM Publish queue (right). Same table styling as register. |

#### Screen 2 — Executive PM Tracker

Same chrome (header, nav, main wrapper). Data-shaped differently:

| Region | Notes |
|---|---|
| Eyebrow + greeting | "Status as of Apr 11" → "M-1 close-out in eleven days." Tone is briefing, not greeting. |
| KPI strip | 6 cells: Phase / Complete / Days to M-1 / Past Due / Budget Used / Meetings. |
| Milestone schedule | Each row is a 4-col grid (`90px 1fr 130px 100px`): mono ID, serif description with sans sub-line, 4px progress bar with mono % + status caption, mono target date. Bar color: green (default), red (past due), amber (due soon). |
| Meeting timeline | Two-column rows (`80px 1fr`): mono date stamp, serif h5 title, sans body. The "Next" row tints amber. |
| Critical open items | Same table treatment as the DCO register. |
| Budget + Decisions log | 2-col split. Budget rows are label/value with serif numerals. Decisions log uses the same timeline pattern. |

#### Tweaks panel (toolbar toggle)

Both screens are tweakable via a floating panel. Defaults:

```jsonc
{
  "density": "calm",        // "calm" | "compact"
  "serifTitles": true,
  "showBuckets": true,
  "denseTable": false
}
```

In production, Tweaks is a designer/PM tool, not a user feature. **Do not ship it** — it exists in the prototype for stakeholder review only. Read the CSS overrides for `body.compact`, `body.no-serif-titles`, etc. to understand which density variants 3H reviewed and approved.

### 4.2 `designs/02-dco-detail.html`

A single DCO record (DCO-0001) in `Past due — awaiting your signature` state, viewed by Andre Butler.

| Region | Layout | Notes |
|---|---|---|
| Header + nav | Same as calm pair | |
| Breadcrumb | `padding: 18px 56px 0` · mono uppercase | Change Orders › Open › DCO-0001 |
| Document header | 2-col flex · left: ID + 36px serif title + 4 meta fields; right: status pill + 2 CTAs | CTAs: `Reject & return` (white/red), `Sign & approve` (navy filled). |
| Blocking callout | White, 3px red left border, 4px radius | "This DCO is on the critical path…" |
| 2-col layout | `grid-template-columns: 1fr 360px; gap: 48px` | Main column + sidebar. |
| Change summary | Serif body (18px), 2-col field grid below | Justification / Risk / Supersedes / Effective on. |
| Training X/Y bar | White card · mono label, 22px serif counter ("4 / 7"), 6px progress bar, mono foot with outstanding names | Bar color: amber while < Y, green at Y/Y. |
| Documents in package | 5-col grid rows (`90px 1fr 100px 100px 120px`) | ID / serif title / rev / signed-X-of-Y / open link. Active row: blue tint background + 3px blue left border, shifted left 14px. |
| Embedded preview | White card, slate-tinted toolbar, body styled as physical document (Times, 13.5px, 1.6 leading, dashed-bordered tables) | Redline view: red strikethrough on deletions, green highlight on insertions. View buttons: Redline / Clean / Side-by-side / Download PDF. |
| Activity log | 2-col rows (`110px 1fr`) · dashed bottom borders | Each: mono timestamp, sentence with mono `actor`, colored `verb` (green=sign, red=reject, blue=submit, amber=return, purple=training), optional inline blockquote, mono meta line (sig hash · IP · session). **Most recent first. No future projections.** |
| Sidebar (right) | 4 stacked cards · 22×24px padding · `margin-bottom: 18px` | Required signatories / Audit metadata / Predicate references / Delegate signature CTA card (light blue tint). |

#### Sign modal

Triggered by `Sign & approve` CTA. Centered over a navy/55% backdrop with 4px blur. 580px max width.

- **Header** — mono eyebrow ("Electronic signature · 21 CFR Part 11"), serif title.
- **Body** —
  - Plain explanation (1 paragraph).
  - **Attestation block** — slate-1 background, 3px navy left border, serif italic copy verbatim from §11.50/11.70/11.200.
  - **Reason for signature** select — required. Options: Approval / Review / Authorship / Responsibility (per Part 11 §11.200(b)).
  - **Comment** textarea — optional, attached to the sig record.
  - **Username** input — readonly, prefilled from session.
  - **Password** input — required, type=password. Helper text cites §11.200(a)(1)(ii).
- **Footer** — mono meta on left, `Cancel` ghost + `Sign now` primary on right.

#### Reject modal

Same shape, different copy and gates:

- Title: "Reject DCO-0001"
- Body explains: returns to originator, reverts status to Draft, **voids any signatures captured on this revision**, requires resubmission.
- **Reason category** select — required. Options: ContentError / Scope / Reference / Format / Other.
- **Detailed comment** textarea — required, **min 20 characters**, helper text says "Returned to originator immediately on submit."
- Same readonly username + password.
- Submit button is `btn-danger` styled (white with red border).

### 4.3 `designs/03-dco-flow-variants.html`

Eight states of the same DCO record. Each section is self-contained and labeled with a navigable anchor at the top of the page.

| # | State | When it shows |
|---|---|---|
| 01 | **Viewer / not a signatory** | User has no role on this DCO. Action surface collapses to a read-only callout. CTA: `Subscribe to updates`. |
| 02 | **You signed — waiting on others** | Right after the user signs. Shows their signature timestamp + hash, names remaining signatories, exposes `Remind signatories` and `Revoke my signature` (with attestation). |
| 03 | **All signatures captured — gated on training** | Status = AwaitingTraining. X/Y counter visible, names outstanding users, shows last reminder date, exposes `Send reminders` and `Request waiver (QA Approver only)`. |
| 04 | **Returned to originator** | After a reject. Shows quoted reason, names rejector + timestamp, says all sigs voided, links to originator view. |
| 05 | **Effective — post-publish** | DCO closed. Header is green. Lists published documents, links to signature manifest PDF, full activity log. The page is now an audit record. Hash chain status visible. |
| 06 | **Empty register** | Filtered view returns zero. States the filter explicitly, offers "Show all" / "Show late" / "Originate new DCO". |
| 07 | **Originator draft → submit** | Tina's view. 4-step progress strip (Link CR / Describe / Attach drafts / Review & submit). Form on left (title, type, risk, target date, routing policy, draft uploads, signatories, training scope), Readiness checklist on right (live checkmarks/warnings). Autosave indicator in foot. `Continue → Review` disabled until all checks ok. |
| 08 | **Delegate signature modal** | From the Delegate card on DCO detail. Scope picker (3 options: this DCO only / all DCOs until date / role-bound). Date range required (no open-ended). Justification (min 20 chars). Auto-revoke conditions stated. Same Part 11 attestation + re-auth. |

---

## 5 · Interactions & behavior

### 5.1 State machines

The DCO state machine is the central model. From `sharepoint-schema.md` §2:

```
Draft ─submit──► Submitted ──open──► InReview ─┬─ all sigs + all training ──► Effective
                       ▲                       ├─ reject ─► Returned ──reopen──► Submitted
                       │                       └─ withdraw ─► Withdrawn
                       └─── (delegated route changes recorded in ActivityLog) ───
```

**Transition rules:**
- `Submitted → InReview` is automatic on first signatory open.
- `InReview → AwaitingTraining` happens when last required sig is captured and any training assignment is incomplete.
- `* → Effective` requires `count(Signatures.Sign) === count(RequiredSignatories)` AND `count(Certifications) === count(Assignments)`.
- `Returned` voids prior signatures **on this revision only**. The void is itself a row in `Signatures` (Action=`RevokeOnRejection`).
- `Effective` is terminal. The record locks. Permission to edit collapses for everyone including QA Approver.

### 5.2 Signing (Part 11)

The flow:

1. User clicks `Sign & approve` on the DCO detail.
2. Modal opens. User picks a Reason code, optionally adds a comment.
3. User re-enters password (or completes MFA step-up if tenant requires).
4. **Server-side** (NOT client):
   - Validate credentials against AAD.
   - Compute `SignatureHash = SHA-256(RecordId || ActorUPN || Verb || RelatedRecordId || ISO8601(now) || PreviousHash)`.
   - Insert row into `Signatures` (atomic, transactional).
   - Insert row into `ActivityLog` referencing the new `Signatures.RecordId`.
   - Recompute `DCO.Status` per state machine.
   - Return success to client.
5. Client refreshes the DCO record from the server (do not optimistic-update — Part 11 forbids showing a sig that hasn't actually persisted).

**Rejection** is the same flow with `Action=Reject`, `RejectionCategory` set, `Comment` required.

### 5.3 Animations and transitions

Quiet. Nothing exuberant.

| Element | Property | Value |
|---|---|---|
| Buttons / links | `color`, `background`, `border` | `transition: all 0.15s` |
| Card hover | `transform`, `box-shadow`, `border-color` | `transition: all 0.18s` |
| Row hover | `background` | `transition: background 0.12s` |
| Filter tab focus | `border-bottom-color` | `transition: border-color 0.15s` |
| Past-due dot pulse | `opacity` 0 → 1 → 0 | `animation: pulse 1.6s infinite` |
| Modal entry | None per current design — appears immediately. **Add `opacity 0→1, scale 0.96→1, 200ms ease-out`** in production. |

### 5.4 Hover / focus / loading / error / empty

- **Hover.** Cards lift 1px + change border to navy. Table rows tint to slate-1. Filter tabs change text color to navy.
- **Focus.** Inputs get navy border. Use Fluent UI's default 2px focus ring on interactive elements that aren't form fields.
- **Loading.** Not designed in this round. Use Fluent UI `Spinner` components inline; for table loads, use 6 skeleton rows with the `slate-1` background pulsing. Don't block the page on load — render the chrome first, fill data progressively.
- **Error.** Inline beneath the failing field, red text 12px sans, no icon. For page-level errors, use the same callout pattern as the blocking callout but with a red top-border accent and a `Retry` CTA.
- **Empty.** See variant 06. Always: state the filter explicitly + offer the most likely next action.

### 5.5 Responsive behavior

- **≥1100px** — full 2-column layout on DCO detail.
- **<1100px** — 2-column collapses to 1-column. Sidebar moves below main content.
- **<900px (tablet)** — KPI strip wraps to 3-up. Filter row wraps. Tables scroll horizontally inside their container.
- **<600px (phone)** — **NOT DESIGNED.** Tina has asked about mobile for travel weeks; flag this as a separate design scope before attempting it. Do not auto-ship a phone view.

---

## 6 · Design tokens

Source of truth: `designs/colors_and_type.css`. Reproduced here for convenience.

### 6.1 Colors

```css
/* Brand */
--n:  #0a3259;  /* navy — h1, primary text, dark CTAs */
--b:  #0f4c81;  /* blue — primary action, links, active tab */
--b2: #1a6bb5;  /* lighter blue — accents */
--b1: #e3f2fd;  /* tint — borders, badges */
--b0: #f0f7ff;  /* wash — hover backgrounds */

/* Status */
--r:  #c62828; --r1: #fde8e8;  /* red — past due, reject, critical */
--a:  #e65100; --a1: #fff3e0;  /* amber — warning, due soon, training */
--g:  #2e7d32; --g1: #e8f5e9;  /* green — effective, signed, success */
--p:  #7b1fa2; --p1: #f3e5f5;  /* purple — awaiting training */

/* Neutrals */
--s0: #fafbfc;  /* page background */
--s1: #f0f4f8;  /* row hover, subtle fills */
--s2: #e0e6ed;  /* borders */
--s4: #8b97a8;  /* disabled, tertiary text */
--s5: #5a6a7e;  /* secondary text, labels */
--s7: #1a2332;  /* body text */
--w:  #ffffff;
```

### 6.2 Typography

Three families. Loaded from Google Fonts.

```css
--serif: 'Libre Baskerville', Georgia, serif;
--sans:  'DM Sans', system-ui, sans-serif;
--mono:  'DM Mono', ui-monospace, monospace;
```

Type scale (most-used → least):

| Class | Family | Size | Weight | Letter | Use |
|---|---|---|---|---|---|
| h1 page title | serif | 26px / 1.15 | 700 | -.4 | Header product title |
| greeting | serif | 30px / 1.2 | 700 | -.4 | Main h2 ("Good afternoon, Andre.") |
| section h3 | serif | 19px / 1.3 | 700 | -.2 | Section headings |
| body large | serif | 18px / 1.55 | 400 |  | Change summary |
| KPI numeral | serif | 38px / 1.0 | 700 | -.5 | KPI cell values |
| bucket count | serif | 42px / 1.0 | 700 | -.6 | Bucket card big numbers |
| body | sans | 15px / 1.55 | 400 |  | Default |
| field value | sans | 14px / 1.5 | 500 |  | Form output, meta |
| meta | sans | 13px / 1.5 | 400 |  | Sublines beneath titles |
| eyebrow | mono | 11px / 1.0 | 500 | 2px | Section eyebrows, dates |
| label | mono | 10.5px / 1.0 | 500 | 1.4–1.6px | Form labels, table headers, caps treatment |
| ID code | mono | 12px / 1.0 | 500 | .3 | DCO-0001 etc. |
| chip / button | sans | 13.5px / 1.0 | 500 | .1 | All buttons |

Weight rules: serif uses only 400 and 700 (italic is reserved for attestation copy and quoted comments). Sans uses 300/400/500/600/700. Mono uses 400/500.

### 6.3 Spacing

4-based scale. Use these values exclusively.

```
4 · 6 · 8 · 10 · 12 · 14 · 16 · 18 · 20 · 22 · 24 · 26 · 28 · 32 · 36 · 40 · 48 · 56 · 64 · 96
```

Page horizontal padding is `56px`. Card internal padding is `22px 24px` to `28px 26px`. Form field vertical padding is `11px`.

### 6.4 Radii

```
4px  — default for cards, buttons, form inputs
6px  — modals
8px  — large cards (hero blocks)
12px — feature cards (rare)
999px — pills, status chips, avatar circles
```

### 6.5 Borders & elevation

- Default border: `1px solid #e0e6ed` (`--s2`).
- Active/focus border: `1px solid #0f4c81` (`--b`).
- Status callout borders: `3px solid {color}` on the **left edge only** (not all sides).

Elevation is restrained. Three levels:

```
none           — flat. Most cards use this.
hover          — box-shadow: 0 4px 14px rgba(15, 76, 129, 0.12). Card hover only.
modal          — box-shadow: 0 30px 80px rgba(0, 0, 0, 0.35). Backdrop: rgba(10, 50, 89, 0.55) + 4px blur.
```

No drop shadows under buttons, form fields, table rows, or chips. The look is "document," not "Material."

### 6.6 Status indicator (the dot)

Used in tables, sidebars, anywhere status appears inline. Single visual primitive:

```
display: inline-flex; align-items: center; gap: 8px; font-size: 12.5px;
::before { width: 7px; height: 7px; border-radius: 50%; background: {status-color}; }
```

The `pdue` (past-due) variant adds the `pulse` animation. Everything else is static.

---

## 7 · State management

### 7.1 What needs to live in client state

- **Current user** (UPN, role, groups) — fetched once at app load via Graph `/me`.
- **Active DCO record** when on detail page — fetched fresh on mount, refreshed after every Part 11 action.
- **Filter selections** on the register — URL-synced (querystring) so links are shareable.
- **Modal open/close state** — local component state.
- **Tweaks panel state** — for the prototype only. Do not ship.

### 7.2 What must NOT live in client state

- Signature records. Always read from server.
- Activity log entries. Always read from server.
- Computed status (`IsPastDue`, `SignatureProgress`, `TrainingProgress`) — server-computed; treat as authoritative.
- The hash chain. Never expose `PreviousHash` or recompute on client.

### 7.3 Recommended approach

- **TanStack Query** (or SWR) for server state. Set `staleTime: 30s` for register lists, `staleTime: 0` for active DCO detail (always refetch on mount).
- **Zustand** or **React Context** for current user + UI state.
- **No global signature cache.** After every sign/reject, invalidate the affected DCO query and refetch.
- **Optimistic updates: forbidden** for any Part 11 action. The user must see the persisted server state, not an optimistic preview.

### 7.4 Data fetching outline

```
GET /portal/dashboard         → KPIs, my buckets, register slice (server-paginated)
GET /dcos?status=...&phase=...→ register filter
GET /dcos/{id}                → full DCO detail incl. linked docs, sigs, log
POST /dcos/{id}/signatures    → sign (with re-auth payload)
POST /dcos/{id}/rejections    → reject (with re-auth payload)
POST /dcos/{id}/delegations   → delegate (with re-auth payload)
GET /pm/tracker               → milestones, meetings, budget, decisions
```

In SPFx these are wrappers around SP REST + a custom Azure Function (or SP-hosted API) that owns the hash-chain logic. **Do not write to `Signatures` or `ActivityLog` from client REST calls.** The custom API is the only allowed writer.

---

## 8 · Assets

| Asset | Status | Notes |
|---|---|---|
| 3H Pharmaceuticals logo | **Missing** | Designs use text + "QMS" badge. Get from 3H marketing before go-live. Reserve a 32px-tall slot. |
| ADB Consulting logo | **Missing** | Footer signature is text-only. Same as above. |
| Iconography | **None** | The product uses no icons in v1. Resist the urge to add them — Tina specifically validated that "the absence of icons reduces visual noise on a regulated screen." If a user need emerges later, propose Heroicons (outline) at 16px, slate-5 color. |
| Fonts | **Loaded from Google Fonts** | DM Sans, DM Mono, Libre Baskerville. For SPFx, self-host them via the SP CDN to avoid third-party font requests on the regulated tenant. The tenant's CSP may block Google Fonts entirely — confirm with 3H IT. |
| Imagery | **None** | The product is text + tables. Do not add stock photos or illustrations. |

---

## 9 · Files in this bundle

```
design_handoff_qms_portal/
├── README.md                          ← this file
├── sharepoint-schema.md               ← full data model — read this first
├── design-system-readme.md            ← original design-system docs (context)
└── designs/
    ├── 01-calm-pair.html              ← Portal home + Executive PM Tracker
    ├── 02-dco-detail.html             ← DCO record with sign/reject modals
    ├── 03-dco-flow-variants.html      ← 8 states (empty, signed, returned, etc.)
    └── colors_and_type.css            ← design tokens
```

Open the HTML files in a browser. Scroll. The prototypes are static — clicking the `Sign` and `Reject` buttons in `02-dco-detail.html` opens working modals (escape to close). The flow-variants file has anchor links at the top for fast navigation.

---

## 10 · Build order (recommended)

Eight working weeks if engineering is full-time. Each sprint ends with a stakeholder demo.

| Sprint | Scope | Demo to |
|---|---|---|
| **S1** | Provision SharePoint lists per `sharepoint-schema.md`. Build read-only Portal home (header, nav, KPIs, register, effective-docs panel). No interactions yet. | Andre + Tina |
| **S2** | DCO detail page. Documents-in-package list. Embedded redline preview (start with iframe, replace with pre-rendered PDF). Activity log read path. | Andre + Tina + Liu Hong |
| **S3** | Sign modal with Part 11 re-auth. Server-side hash-chain writer. Signature table + activity log writes. Reject path. | Andre + 3H QA, validation dry-run |
| **S4** | Training assignments + certification tracking. X/Y gate enforcement on `Effective` transition. Reminder send. | All stakeholders |
| **S5** | PM Publish queue. Draft → Official transition. Version stamping (Supersedes, SupersededBy, EffectiveDate atomicity). | Andre + Tina |
| **S6** | PM Tracker dashboard (Phase 2A milestone view). Read-only — no writes from this surface. | Steering committee |
| **S7** | UAT, validation execution (IQ/OQ/PQ), bug fixes. Hash chain integrity test under load. | 3H QA + ADB |
| **S8** | Production cutover, training rollout (the system itself is a controlled SOP), hypercare. | Everyone |

---

## 11 · Open questions for engineering before sprint 1

These need answers from 3H IT, 3H QA, or the validation lead before the first line of code:

1. **MFA / step-up.** Is 3H's AAD tenant configured for Conditional Access step-up on a sign event? If yes, password re-auth alone is insufficient and the modal needs to launch the MFA challenge. If no, password alone is acceptable per Part 11 — but document the rationale in the FRS.
2. **CSP and font hosting.** Does 3H's tenant CSP allow `fonts.googleapis.com` and `fonts.gstatic.com`? If not, self-host all three families via the SP CDN.
3. **Redline rendering.** Word Online iframe or pre-rendered PDF? See §3 for trade-offs.
4. **Auto-disposition.** Does QA want auto-archive at `Documents.RetentionUntil`, or manual review with a notification 30 days prior?
5. **Mobile.** Is a phone form factor required for travel weeks? If yes, scope a separate design round.
6. **External reviewers.** Will B2B guest accounts ever need access (e.g. FDA inspector)? Affects auth design.

Get these answered before kickoff. Don't guess.

---

## 12 · Validation reminders (do not skip)

This is the part that tempts every team to cut corners. Don't.

- **URS / FRS** must be authored before sprint 1 ends, and they must trace every requirement to a clause in `sharepoint-schema.md` and a screen in `designs/`.
- **IQ / OQ / PQ** protocols must be drafted by sprint 4 and executed in sprint 7. Run them in the production tenant before cutover.
- **Audit-trail integrity test** — every release runs a script that verifies the hash chain end-to-end. A mismatch is a Severity 1 incident, not a bug ticket.
- **Signature manifest export** — generate a PDF of the full signature chain for any DCO on demand. The export itself goes through the audit log.
- **The system is a controlled SOP.** Every user signs the SOP for the new system before they get a license. Plan training rollout into sprint 8.

The FDA inspector will ask for the validation package on day one. 3H is the regulated party — they own the consequence. Make their job easy.

---

**Questions while building?** The design author left a trail of decisions in the original design-system README (`design-system-readme.md` in this bundle). When in doubt, the canonical answers are: less, not more · text and tables · no icons · serif for emphasis, sans for work, mono for IDs · audit-first.
