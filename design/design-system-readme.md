# 3H Pharmaceuticals QMS Design System

A design system for the **IMP9177 SharePoint portal suite** — a set of internal SPFx web parts built by **ADB Consulting & CRO Inc.** for **3H Pharmaceuticals LLC** to manage their FDA-regulated Quality Management System (21 CFR Part 111 / FSMA dietary supplement compliance).

This is **enterprise compliance software** — read by QA managers, document controllers, regulatory consultants, and 3H operations staff during audits and inspections. Aesthetic priorities are **density, scannability, and seriousness** over playfulness or marketing polish.

---

## Sources

This design system was built by reading the source repository the user attached:

- **GitHub:** `adbccro/IMP9177` (branch: `main`)
- **CLAUDE.md** in the repo identifies the project: *"IMP9177 — 3H Pharmaceuticals LLC QMS Portal"*
- **SP site:** `https://adbccro.sharepoint.com/sites/IMP9177` (not accessible from here)
- No Figma link, brand guide, or external asset library was provided. Everything in this system was reverse-engineered from the inline CSS and HTML in the SPFx web part TypeScript files.

The repo contains six SPFx web parts. We treat each as a **product surface** of the same QMS system:

| Web part | What it is | Audience |
|---|---|---|
| `qmsPortalWebPart` | The main multi-tab portal: DCOs, CRs, Records, Training, Publish, Documents, Admin, Config | All portal users |
| `imp9177Dashboard` | Read-only executive dashboards (PM, MRS, RAID, DOCREPO) — toggled via property pane | Project leads, exec stakeholders |
| `mrsFormsWebPart` | Master Record Sheet — document gap tracker | Doc controllers |
| `pmFormsWebPart` | Project management forms | Project leads |
| `raidFormsWebPart` | Risks / Actions / Issues / Decisions register | Project leads |
| `supFormsWebPart` | Supplier forms | QA / supply chain |

Two distinct visual treatments coexist in the codebase. Both are documented and supported here as **modes** of the same brand:

- **Portal mode** (`qmsPortalWebPart`, `mrsFormsWebPart`, `supFormsWebPart`) — DM Sans + DM Mono, navy gradient header, 12–14 px text, denser tables, more functional.
- **Dashboard mode** (`imp9177Dashboard`) — Libre Baskerville (serif display) + Source Sans 3, slightly lighter palette (`#0C2D5E` / `#1E56A0`), serif numerals on KPIs. Used for read-only executive views.

When in doubt, **Portal mode is the default** — it's the one users live in.

---

## Index

| File | Purpose |
|---|---|
| `README.md` | This file — context, content fundamentals, visual foundations, iconography |
| `SKILL.md` | Agent-skill front-matter for use in Claude Code or other agent runners |
| `colors_and_type.css` | All design tokens (CSS custom properties) for both Portal and Dashboard modes, plus semantic typography classes |
| `assets/` | Logos, welcome illustrations, raw imagery copied from the source repo |
| `fonts/` | (empty — fonts loaded from Google Fonts CDN; see Visual Foundations) |
| `preview/` | Small HTML cards used to populate the Design System review tab |
| `ui_kits/portal/` | UI kit recreating the QMS Portal (main multi-tab portal) |
| `ui_kits/dashboard/` | UI kit recreating the executive Dashboard (PM tracker layout) |
| `slides/` | *(none — no slide template was provided in the source)* |

---

## Content Fundamentals

### Voice

**Compliance-first, plain-spoken, no marketing.** This is software for people who sign legally-binding documents. Copy is short, declarative, and assumes the reader knows the regulatory acronyms. There is no "we", no "let's", no exclamation marks, no welcoming prose. Buttons and labels are commands or nouns: *Submit*, *Reject DCO*, *Cancel Submission → Return to Draft*, *Approver Setup*.

### Tone qualities

- **Direct** — *"Documents Needing My Signature"*, not *"You have documents to review"*.
- **Audit-ready** — every action implies a record will be written. *"Reason Required"* appears in cancellation and rejection modal titles.
- **Specific over generic** — *"DCOs awaiting your approval signature"* (not *"Items"*); *"Drafts ready to push to Published zone"* (not *"Pending"*).
- **Quietly urgent** — red is reserved for things that block compliance. The same string can be calm (*"4 active"*) or alarming (*"PAST DUE — Tina + Liu Hong must sign"*) depending on state.

### Grammar and casing

- **Title Case** for screen names, panel headers, modal titles: *Change Order Register*, *PM Publish Queue*, *System Configuration*.
- **UPPERCASE + letter-spacing 0.6–0.8 px** for KPI labels, table column headers, section eyebrows: *MY PENDING ACTIONS*, *PHASE*, *DCO #*.
- **Sentence case** for KPI sub-text and helper copy: *"Requiring your input"*, *"Within 7 days"*, *"Locked on active DCO"*.
- **Em-dashes** (`—` or rendered as `&mdash;` / `--`) separate clauses inside a single line: *"Reject DCO — Reason Required"*, *"3H — Tina Qin"*.
- **Middle dot** (`·` / `&middot;`) joins inline metadata: *"IMP9177 · 3H Pharmaceuticals LLC · 21 CFR Part 111 / FSMA · Phase 2A"*.
- **Voice:** mostly *I* and *My* in the user-facing dashboard buckets (*"My Pending Actions"*, *"Documents I Originated"*, *"Change Orders I'm Involved With"*) — the user is the subject of their work. System-level chrome is third-person/object: *"Document Repository"*, *"Approver Setup"*.

### Domain vocabulary (use as-is)

`DCO` (Document Change Order), `CR` (Change Request), `Approver`, `Originator`, `Signatory`, `Zone` (Draft / Published / Official), `Phase` (Draft / Submitted / In Review / Implemented / Awaiting Training / Effective), `Rev` (revision), `Sig ID` (signature ID), `21 CFR Part 111`, `FSMA`, `SOP`, `FM` (Form), `QM` (Quality Manual), `FPS`, `MRS` (Master Record Sheet), `RAID` (Risks/Actions/Issues/Decisions).

### Examples lifted verbatim from the source

- Modal titles: *"Cancel Submission — Reason Required"*, *"Reject DCO — Reason Required"*
- Cancel categories: *"Documents need revision"*, *"Wrong documents assigned"*, *"Submitted in error"*
- Reject categories: *"Incomplete documentation"*, *"Missing approver sign-off"*, *"Regulatory non-compliance"*, *"Training not completed"*
- KPI sub-text: *"Past due date"*, *"Within 7 days"*, *"Signed & filed"*, *"In routing"*
- Bucket descriptions: *"DCOs where I'm an approver or originator"*, *"Based on your role assignments"*, *"Drafts ready to push to Published zone"*
- Empty/loading: *"Loading document matrix from SharePoint…"*, *"Loading config…"*
- Footer: *"IMP9177 QMS Portal · ADB Consulting & CRO Inc. · 21 CFR Part 111 / FSMA · Read-only view — data live from SharePoint"*
- Real production strings: *"W1 Sign-off PAST DUE — Tina + Liu Hong must sign"*, *"Melatonin Powder — No CoA on file (active production)"*, *"Orkin pest addendum — URGENT before inspection"*

### Emoji as iconography

The codebase uses **single-glyph color emoji** as functional icons throughout the UI — they are not decorative. They appear in nav tabs (🏠 📋 🔄 📁 🎓 🚀 📄 🏛 ⚙️), bucket headers (📚 📋 🔄 ✍️ 📄 🎓 🚀), action labels (🔴 Late, ⚠️ Due Soon, 💾 Save, ⟳ Refresh, ↩ Reject, 🚀 Publish Selected), and zone indicators (📝 Draft / ✅ Published / 🏛️ Official). When recreating screens, use **the same emoji from the source files** — don't substitute SVG icon sets.

---

## Visual Foundations

### Color palette

Two named-and-numbered palettes coexist. Use **Portal palette** by default; **Dashboard palette** for read-only executive views.

**Portal palette** (from `qmsPortalWebPart`)

| Token | Hex | Use |
|---|---|---|
| `--n` | `#0a3259` | Navy / primary dark — header gradient start, headings, KPI numerals |
| `--b` | `#0f4c81` | Blue — primary action buttons, links, focus rings |
| `--b2` | `#1a6bb5` | Blue 2 — header gradient end, button hover |
| `--b1` | `#e3f2fd` | Blue 100 — soft tint for "info" pills, hover backgrounds |
| `--b0` | `#f0f7ff` | Blue 50 — table row hover, panel header gradient end |
| `--r` | `#c62828` | Red — destructive / late / blocked |
| `--r1` | `#fde8e8` | Red tint — alert background, late pills |
| `--a` | `#e65100` | Amber — warning / pending |
| `--a1` | `#fff3e0` | Amber tint |
| `--g` | `#2e7d32` | Green — success / signed / effective |
| `--g1` | `#e8f5e9` | Green tint |
| `--s0` | `#f8fafc` | Page background |
| `--s1` | `#f0f4f8` | Subtle row separator, chip background |
| `--s2` | `#e0e6ed` | Borders |
| `--s5` | `#5a6a7e` | Muted body text, sub-labels |
| `--s7` | `#1a2332` | Body text |
| `--w` | `#fff` | Card / panel surface |

**Dashboard palette** (from `imp9177Dashboard`) is similar but colder and slightly more saturated: `#0C2D5E` / `#1E56A0` / `#3B82D4` / `#DBEAFE` / `#EFF6FF`, with Tailwind-style red `#DC2626`, amber `#D97706`, green `#059669`. See `colors_and_type.css` for the full set under `[data-mode="dashboard"]`.

**Status purple** (`#7b1fa2` / `#f3e5f5`) is reserved for "Awaiting Training" — a distinct phase that isn't success, isn't pending, isn't error.

### Typography

| | Portal mode | Dashboard mode |
|---|---|---|
| **Sans (body, UI)** | DM Sans 300/400/500/600/700 | Source Sans 3 300/400/500/600/700 |
| **Display (headings, KPI numerals)** | DM Sans 700 | **Libre Baskerville 400/700** (serif) |
| **Mono (numerals, IDs, timestamps)** | DM Mono 400/500 | system monospace |
| **Body size** | 14 px (default) | 14 px |
| **Densest readable size** | 9–10 px (table headers, eyebrows) | 9–10 px |
| **KPI numeral** | 28–34 px DM Mono 700 | 28–30 px Libre Baskerville 700 |

Both fonts load from **Google Fonts CDN** at runtime — no font files are bundled in the repo. We follow the same convention here (no `fonts/` payload). If working offline, fetch them via `fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600;700` and `family=Libre+Baskerville:wght@400;700&family=Source+Sans+3:wght@300;400;500;600;700`.

**Font substitution flag:** the source loads DM Sans / DM Mono / Libre Baskerville / Source Sans 3 directly from Google Fonts; no local TTF/WOFF was provided. If the user wants an **offline-safe bundle**, they should provide the font files or confirm we may continue to depend on the CDN.

### Spacing and rhythm

The system uses **tight, intentional spacing** typical of compliance dashboards.

- Page outer padding: **20–28 px** horizontal, **20 px** top
- Card outer gap: **10–16 px** (12 px most common in the dashboard, 14 px in the portal)
- Card inner padding: **12–18 px** vertical / **14–22 px** horizontal
- Form field gap: **13–14 px** between rows; label sits 4 px above input
- Table cell padding: **8–9 px vertical / 12 px horizontal**
- Header height: **56 px** (portal) / **60 px** (dashboard)
- KPI strip gap: **10–12 px**

There is no formal 4 / 8 / 12 / 16 / 24 step scale — the codebase uses pixel values directly. We expose `--space-1` through `--space-6` in `colors_and_type.css` as a recommendation for new work (4 / 8 / 12 / 16 / 20 / 24).

### Borders, radii, and elevation

- **Card radius:** 8 px (panels, KPIs, modals)
- **Pill radius:** 12 px (filter pills, status pills)
- **Filter button radius:** 18–20 px (full pill)
- **Small chip / badge radius:** 4 px (zone badges, type badges, late badges)
- **Borders:** 1 px solid `--s2` (`#e0e6ed`) for all card edges and table dividers.
- **Accent borders:** `border-top: 3px solid <status color>` on KPI cards (`--b2` default, `--r` / `--a` / `--g` for status).
- **Shadows:** very subtle, layered.
  - Card resting: `0 1px 3px rgba(0,0,0,.06)`
  - Card raised on hover: `0 4px 14px rgba(15,76,129,.12)` (translateY -2 px)
  - Header: `0 2px 8px rgba(0,0,0,…)` for the navy bar
  - Modal: `0 20px 60px rgba(0,0,0,.3)`

### Backgrounds

- **Page:** flat `#f8fafc`. No textures, no patterns, no full-bleed imagery.
- **Header:** **diagonal navy gradient** `linear-gradient(135deg, #0a3259 0%, #0f4c81 60%, #1a6bb5 100%)` — this is the single signature visual element of the brand. Reuse it on any product header.
- **Panel headers:** subtle horizontal gradient `linear-gradient(to right, var(--b0), var(--w))`.
- **MRS variant:** the navy gradient gets a decorative **40 % opacity white circle** offset top-right (`::before`, 200 × 200 px, blur via `border-radius:50%`) — a quiet flourish, optional.
- **No** purple gradients. **No** glassmorphism. **No** illustrated full-bleed scenes.

### Animation and interaction

- **Transitions:** `transition: all .15s` on filter buttons, table rows, cards. `.1s` on row-hover backgrounds. **Never longer than 250 ms.**
- **Easing:** the codebase uses the browser default (linear/ease) — keep it. No custom cubic-béziers, no bounce, no spring.
- **Loading spinners:** `@keyframes spin` — 0.8 s linear rotate. 16 × 16 px, 2 px border, top color `--b`.
- **KPI loading:** `@keyframes mrsPulse` — 1.2 s ease-in-out opacity pulse on placeholder text.
- **Hover (cards):** raise 2 px with `transform: translateY(-2px)` + softer-larger blue-tinted shadow + `border-color: var(--b)`.
- **Hover (rows):** `background: var(--b0)`.
- **Hover (filter pills):** flip to filled blue (`background: var(--b)`, `color: #fff`).
- **Hover (text links):** add underline; on dark links color shifts to `--linkHovered: #014446` (Fluent default).
- **Press states:** the codebase doesn't add explicit press styling — buttons rely on the browser default `:active`. Don't invent one.
- **Focus:** inputs change `border-color: var(--b)` on `:focus`. No outline/ring is drawn — keep this convention; if accessibility requires more, add `outline: 2px solid var(--b)` with a 1 px offset.

### Layout rules

- Header is **fixed-height (56–60 px)**, full-width, with the gradient running edge-to-edge.
- Below the header sits a **navy nav strip** (`--n` background) with horizontally-scrolling tabs.
- Body is `padding: 20px 28px` with no max-width on the portal, `max-width: 1600px; margin: auto` on the dashboard.
- KPI rows are `display: flex` with `flex: 1; min-width: 120px` (portal) or CSS grid `repeat(N, 1fr)` (dashboard). Both wrap at smaller breakpoints.
- Tables are `width: 100%`, sit inside `.tcard` (rounded 8 px container), and never have visible outer borders on the table itself — the card supplies them.
- Modals: `max-width: 660–780 px`, centered, with sticky header and footer; backdrop is `rgba(0,0,0,.45)`.

### Transparency and blur

- Header text has `rgba(255,255,255,.5–.6)` for sub-titles, `rgba(255,255,255,.45)` for timestamps, `rgba(255,255,255,.85)` for chips on the gradient.
- Header buttons are `rgba(255,255,255,.15)` background with `rgba(255,255,255,.3)` border — translucent but not blurred.
- **No `backdrop-filter: blur()` anywhere.** Don't add it.

### Imagery vibe

- The repo ships only **two welcome illustrations** (the SPFx default Microsoft 365 Patterns & Practices clipart) — they are not part of the brand and are not used in any product surface in the codebase. Treat the system as **illustration-free** until 3H Pharmaceuticals provides real photography or commissioned art.
- For placeholders, use **flat colored rectangles** with the `--s1`/`--s2` tones, or the navy gradient.

---

## Iconography

**The QMS Portal uses native color emoji as its primary icon system.** This is unusual but intentional — emoji ship in every modern OS, render perfectly inside SharePoint without webfont licensing, and need no asset pipeline. Replicate this in any new product surface.

### Inventory of emoji used in the source

| Emoji | Meaning in this system |
|---|---|
| 🏠 | Dashboard / home tab |
| 📋 | Change Orders (DCO) — also "ready / queue" |
| 🔄 | Change Requests (CR) |
| 📁 | Records |
| 🎓 | Training |
| 🚀 | PM Publish — also "publish action" |
| 📄 | Documents / Document Repository |
| 🏛 / 🏛️ | Administration / Official zone |
| ⚙️ | Configuration |
| 📝 | Draft zone |
| ✅ | Published zone, signed, completed |
| 📚 | Effective / official documents bucket |
| ✍️ | Awaiting signature |
| 🔴 | Late / overdue filter |
| ⚠️ | Warning / due-soon / alert prefix |
| ↩ | Reject |
| 💾 | Save |
| ⟳ | Refresh (Unicode arrow, not emoji) |
| × | Modal close (Unicode multiplication sign) |
| 🔍 | Search / "under review" panel |
| 👤 / 👥 / 🛡 | Employees / Approvers / Roles & Permissions |
| · | Inline metadata separator (U+00B7 middle dot) |

When you need an icon that isn't already in this set, **prefer Unicode/emoji** that follows the same visual register before reaching for an icon font.

### Fallback: Lucide

If a design absolutely requires a flat monochrome stroke icon (e.g. inside a tight 12 px button where emoji would render too detailed), use **[Lucide](https://lucide.dev/)** at 1.5 px stroke weight, sized 14 × 14 px, color `var(--s5)` for muted or `var(--b)` for active. **Flag this as a substitution in your design notes** — the source codebase doesn't use Lucide; introducing it is a deliberate extension of the system.

### What we do *not* use

- No custom SVG icon set
- No icon font (Font Awesome, Material Icons, Fluent Icons)
- No PNG icon sprites
- No hand-drawn illustrations

### Logos and brand marks

The repo contains **no 3H Pharmaceuticals logo, no ADB Consulting & CRO logo, and no IMP9177 lockup**. The product identifies itself in text only:

- Header lockup: **"QMS"** badge (uppercase, white, on translucent rounded rect) + **"IMP9177 · QMS Portal"** title + sub-line *"3H Pharmaceuticals LLC · 21 CFR Part 111 / FSMA · Document Control & DCO Routing"*.
- Footer signature: *"IMP9177 QMS Portal · ADB Consulting & CRO Inc. · 21 CFR Part 111 / FSMA · Read-only view — data live from SharePoint"*.

We replicate this text-only lockup in `assets/logo-lockup.html` for use in slides and docs. **Flag for user:** if 3H or ADB has a real logo, please attach it — the system currently has none.

---

## Caveats / known gaps

- **No real logo** for 3H Pharmaceuticals or ADB Consulting & CRO. The header is text + colored badge.
- **No font files bundled.** We rely on Google Fonts CDN, matching the source.
- **No slide template.** Decks were not part of the source; we did not invent one.
- **Welcome illustrations** in `assets/welcome-light.png` and `assets/welcome-dark.png` are the SPFx default Microsoft 365 Patterns & Practices clipart, not 3H brand assets — keep them out of any production-facing design.
- **Emoji rendering varies** across Windows, macOS, iOS, and SharePoint Online. The source accepts this; if pixel-perfect cross-platform consistency matters, switch to Lucide and re-document.
