# IMP9177 QMS — SharePoint List Schema Specification

**Project:** IMP9177 — 3H Pharmaceuticals LLC × ADB Consulting & CRO Inc.
**Audience:** SPFx engineering, validation lead (3H QA), Andre Butler (PM)
**Status:** v1.0 draft · Apr 11, 2026
**Source:** Derived from approved Calm Pair UI (`ui_kits_v2/calm-pair.html`) and DCO detail (`ui_kits_v2/dco-detail.html`)

> This document defines the SharePoint list structure that the QMS Portal and PM Tracker UIs read from and write to. **Every field on the UI maps to a column in this spec.** If a field isn't here, the UI cannot render it; if a list isn't here, the UI cannot transition state.
>
> Implementation must be paired with a validation package (URS/FRS/IQ/OQ/PQ) before go-live, per 21 CFR Part 11.

---

## 0 · Conventions

- **List naming:** PascalCase (`ChangeOrders`), no spaces, plural where the list represents records.
- **Column naming:** PascalCase (`TargetEffectiveDate`); display name may differ but internal name is canonical.
- **IDs:** SharePoint integer ID is internal only. Every list with a user-facing identifier carries a separate `RecordId` field (`DCO-0001`, `CR-0001`) generated server-side, monotonic, never reused.
- **Audit columns** (every list): `Created`, `Modified`, `Author`, `Editor` (built-in) + `AuditTrailHash` (custom, see §11).
- **Item-level permissions** are NOT used for record-level access. Use a `Visibility` column + a single SP group per role; enforce via security trim in the read API and validate writes server-side.
- **Lookups** carry the SP integer ID; resolve to `RecordId` for display.
- **Soft-delete only.** Add `IsDeleted` (bool) + `DeletedBy` + `DeletedAt`. Hard delete is forbidden under Part 11.
- **Indexes:** mark required indexes per list. SP list view threshold (5,000) is real — pre-index every column that filters/sorts in the UI.

---

## 1 · `ChangeRequests`

Origin record. A CR can fan out into one or more DCOs, but the UI assumes 1:1 except where explicitly noted.

| Column | Type | Req | Indexed | Notes |
|---|---|---|---|---|
| `RecordId` | Single line | ✓ | ✓ | `CR-NNNN`, monotonic |
| `Title` | Single line | ✓ | | Short title; max 200 chars |
| `Summary` | Multi-line (rich) | ✓ | | The "what & why" (renders as Change Summary on UI) |
| `Justification` | Multi-line (plain) | ✓ | | Predicate citations, audit findings, etc. |
| `RiskClass` | Choice | ✓ | ✓ | `Low`, `Moderate`, `High`, `Critical` |
| `Originator` | Person/Group | ✓ | ✓ | Single user |
| `OriginatorOrg` | Choice | ✓ | | `3H`, `ADB` |
| `LinkedIssues` | Lookup → Issues | | | Multi |
| `LinkedAuditFindings` | Single line | | | Free-text identifier (`2025-Q4-03`) |
| `Status` | Choice | ✓ | ✓ | `Draft`, `Submitted`, `Approved`, `Rejected`, `Withdrawn`, `Closed` |
| `LinkedDCOs` | Lookup → ChangeOrders | | | Multi (CR can spawn multiple DCOs) |
| `Phase` | Choice | ✓ | ✓ | `1`, `2A`, `2B`, `3`, `4` |
| `Visibility` | Choice | ✓ | | `Internal`, `Both` (3H + ADB) |
| `IsDeleted`, `DeletedBy`, `DeletedAt` | (audit) | | | |

**Indexes required:** `RecordId`, `Status`, `Phase`, `RiskClass`, `Originator`.

---

## 2 · `ChangeOrders` (DCOs)

The core work item. Drives routing, signatures, training, effective-date logic.

| Column | Type | Req | Indexed | Notes |
|---|---|---|---|---|
| `RecordId` | Single line | ✓ | ✓ | `DCO-NNNN`, monotonic |
| `Title` | Single line | ✓ | | Renders as h1 on detail page |
| `Type` | Choice | ✓ | ✓ | `SOP`, `FM` (Form), `QM` (Quality Manual), `FPS` (Finished-Product Spec), `WI` (Work Instruction), `Package` (multi-doc) |
| `LinkedCR` | Lookup → ChangeRequests | ✓ | ✓ | Single |
| `Originator` | Person/Group | ✓ | ✓ | |
| `OriginatorOrg` | Choice | ✓ | | `3H`, `ADB` |
| `Summary` | Multi-line (rich) | ✓ | | Inherits from CR initially; editable on DCO |
| `RiskClass` | Choice | ✓ | ✓ | Inherits from CR; can be elevated |
| `Status` | Choice | ✓ | ✓ | `Draft`, `Submitted`, `InReview`, `Returned`, `AwaitingTraining`, `Effective`, `Withdrawn`, `Rejected` |
| `SubmittedDate` | Date | | ✓ | Set on Draft → Submitted transition |
| `TargetEffectiveDate` | Date | ✓ | ✓ | Drives "Past due" calculation in UI |
| `EffectiveDate` | Date | | ✓ | Set when all sigs + all training certs captured |
| `Phase` | Choice | ✓ | ✓ | Inherits from CR |
| `Milestone` | Lookup → Milestones | | | Multi (a DCO can serve multiple milestones) |
| `RequiredSignatories` | Person/Group | ✓ | | Multi; the canonical approver list at submit time |
| `RoutingPolicy` | Choice | ✓ | | `Parallel` (any order), `Sequential` (ordered), `RoleQuorum` |
| `Documents` | Lookup → DCODocuments | ✓ | | Multi; the drafts in the package |
| `IsBlocking` | Bool | | ✓ | UI shows blocking callout when true |
| `BlockingReason` | Multi-line (plain) | | | Free text rendered in callout |
| `LinkedIssues` | Lookup → Issues | | | Multi |
| `RejectionCount` | Number | | | Increments on each Reject & Return |
| `Visibility` | Choice | ✓ | | `Internal`, `Both` |
| `IsDeleted`, `DeletedBy`, `DeletedAt` | (audit) | | | |

**State machine:**

```
Draft ─submit──► Submitted ──open──► InReview ─┬─ all sigs + all training ──► Effective
                       ▲                       ├─ reject ─► Returned ──reopen──► Submitted
                       │                       └─ withdraw ─► Withdrawn
                       └─── (delegated route changes recorded in ActivityLog) ───
```

**Computed in API (not stored):**
- `IsPastDue` = `Status ∈ {Submitted, InReview, AwaitingTraining}` AND `TargetEffectiveDate < today`
- `DaysLate` = today − `TargetEffectiveDate`
- `SignatureProgress` = `count(Signatures where DCO=this and Status=Captured) / count(RequiredSignatories)`
- `TrainingProgress` = `count(TrainingCertifications where DCO=this and Status=Certified) / count(TrainingAssignments where DCO=this)`

**Indexes required:** `RecordId`, `Status`, `Phase`, `Type`, `TargetEffectiveDate`, `Originator`, `LinkedCR`.

---

## 3 · `DCODocuments`

Junction list — each row is *one draft document* attached to a DCO. The redline preview, per-document signature count, and effective-on-publish stamping all live here.

| Column | Type | Req | Indexed | Notes |
|---|---|---|---|---|
| `RecordId` | Single line | ✓ | ✓ | `DCO-0001:SOP-PM-021` (composite) |
| `DCO` | Lookup → ChangeOrders | ✓ | ✓ | |
| `DocumentId` | Single line | ✓ | ✓ | `SOP-PM-021`, `FM-014`, etc. |
| `Title` | Single line | ✓ | | Renders in document list |
| `DocType` | Choice | ✓ | ✓ | `SOP`, `FM`, `QM`, `FPS`, `WI` |
| `FromRev` | Single line | | | `r3` (null if new doc) |
| `ToRev` | Single line | ✓ | | `r4` |
| `IsNew` | Bool | | | True if no prior revision |
| `DraftFileRef` | URL | ✓ | | Path in `/Drafts` document library |
| `RedlineFileRef` | URL | | | Pre-rendered redline (DOCX or PDF) |
| `Pages` | Number | | | For UI footer "8 pages" |
| `RequiredSignatories` | Person/Group | ✓ | | Multi; per-document override of DCO list |
| `Description` | Multi-line (plain) | | | The `<small>` line under doc title in UI |
| `EffectivePath` | URL | | | Path in `/Official` once published |
| `SupersedesRef` | Single line | | | `SOP-PM-021 r3` (audit only) |

**Indexes required:** `DCO`, `DocumentId`.

---

## 4 · `Signatures`

The Part 11 audit table. **Append-only.** Never updated, never deleted. One row per signature event (success OR failure).

| Column | Type | Req | Indexed | Notes |
|---|---|---|---|---|
| `RecordId` | Single line | ✓ | ✓ | `SIG-NNNNNN`, monotonic |
| `DCO` | Lookup → ChangeOrders | ✓ | ✓ | |
| `Document` | Lookup → DCODocuments | ✓ | ✓ | The specific draft signed |
| `Signatory` | Person/Group | ✓ | ✓ | Single user |
| `SignatoryRoleAtSign` | Single line | ✓ | | Snapshot of user's role at time of signature |
| `Action` | Choice | ✓ | ✓ | `Sign`, `Reject`, `Delegate`, `RevokeOnRejection` |
| `ReasonCode` | Choice | ✓ | | Per Part 11 §11.200: `Approval`, `Review`, `Authorship`, `Responsibility` |
| `Comment` | Multi-line (plain) | | | Optional for sign, **required** for reject |
| `RejectionCategory` | Choice | | | Required when `Action=Reject`. `ContentError`, `Scope`, `Reference`, `Format`, `Other` |
| `Timestamp` | Date+Time | ✓ | ✓ | UTC, server-stamped, never client |
| `IPAddress` | Single line | ✓ | | Client IP at sign time |
| `UserAgent` | Single line | ✓ | | Browser fingerprint |
| `SessionDuration` | Number (sec) | | | Time from auth to sign |
| `ReAuthMethod` | Choice | ✓ | | `Password`, `MFA`, `SSO+Step-up` |
| `ReAuthSuccess` | Bool | ✓ | | False rows are also stored (failed attempts) |
| `SignatureHash` | Single line | ✓ | | SHA-256 of (User\|Action\|DocId\|Rev\|Timestamp\|PrevHash) |
| `PreviousHash` | Single line | ✓ | | Prior row's `SignatureHash` — chain integrity |
| `DelegationFromUser` | Person/Group | | | Set when this sig was authorized by a delegation |
| `DelegationRecordRef` | Lookup → Delegations | | | |

**Critical:**
- This list has **no Edit and no Delete permissions** for any user, including site admins. Enforce at SP level + custom event receiver. Validation must prove this.
- The hash chain is verified periodically by a scheduled job; broken chain triggers a Severity 1 alert.

**Indexes required:** `DCO`, `Document`, `Signatory`, `Timestamp`, `Action`.

---

## 5 · `ActivityLog`

Wider than `Signatures` — captures every state change, comment, system event. Drives the "Routing & activity log" section on the DCO detail page.

| Column | Type | Req | Indexed | Notes |
|---|---|---|---|---|
| `RecordId` | Single line | ✓ | ✓ | `LOG-NNNNNNN` |
| `DCO` | Lookup → ChangeOrders | ✓ | ✓ | |
| `Document` | Lookup → DCODocuments | | ✓ | Optional — log can be DCO-level |
| `Actor` | Person/Group | ✓ | ✓ | `System` for automated events (use a service account) |
| `Verb` | Choice | ✓ | ✓ | `created`, `submitted`, `signed`, `rejected`, `returned`, `commented`, `delegated`, `assigned-training`, `certified-training`, `published`, `flagged-overdue`, `auto-escalated`, `withdrew`, `linked` |
| `Summary` | Multi-line (plain) | ✓ | | UI rendering text |
| `Quote` | Multi-line (plain) | | | Inline blockquote when actor left a comment |
| `Metadata` | Multi-line (plain) | | | "Sig hash 7c14...d9f1 · 192.168.4.221 · session 28m" |
| `Timestamp` | Date+Time | ✓ | ✓ | Server UTC |
| `RelatedSignatureRef` | Lookup → Signatures | | | Backlink to sig event |
| `IsSystem` | Bool | ✓ | ✓ | True when `Actor=System` |

**Append-only**, same enforcement rules as `Signatures`.

**Indexes required:** `DCO`, `Timestamp`, `Verb`.

---

## 6 · `TrainingAssignments` & `TrainingCertifications`

Two lists. Assignments are *who needs to be trained on what*; Certifications are *who has completed it*. The X/Y bar in the UI is `count(certs) / count(assignments)`.

### `TrainingAssignments`

| Column | Type | Req | Indexed | Notes |
|---|---|---|---|---|
| `RecordId` | Single line | ✓ | ✓ | `TRN-NNNN` |
| `DCO` | Lookup → ChangeOrders | ✓ | ✓ | |
| `Document` | Lookup → DCODocuments | | ✓ | Optional — training can be DCO-wide |
| `AssignedUser` | Person/Group | ✓ | ✓ | |
| `AssignedRole` | Single line | | | Snapshot at assign time |
| `AssignedBy` | Person/Group | ✓ | | |
| `AssignedAt` | Date+Time | ✓ | | |
| `DueDate` | Date | | ✓ | |
| `Status` | Choice | ✓ | ✓ | `Pending`, `Notified`, `InProgress`, `Certified`, `Overdue`, `Waived` |
| `WaiverJustification` | Multi-line | | | Required when `Status=Waived`; written by QA Approver |

### `TrainingCertifications`

| Column | Type | Req | Indexed | Notes |
|---|---|---|---|---|
| `RecordId` | Single line | ✓ | ✓ | `CERT-NNNNN` |
| `Assignment` | Lookup → TrainingAssignments | ✓ | ✓ | Single |
| `User` | Person/Group | ✓ | ✓ | Must match Assignment.AssignedUser |
| `Method` | Choice | ✓ | | `Self-attest`, `Supervisor-attest`, `Quiz`, `Live-session` |
| `Score` | Number | | | Required if `Method=Quiz` |
| `Timestamp` | Date+Time | ✓ | ✓ | |
| `IPAddress` | Single line | ✓ | | |
| `SignatureRef` | Lookup → Signatures | ✓ | | Every cert generates a row in `Signatures` too |

**Effective-date gate:** `ChangeOrders.EffectiveDate` cannot be set unless `count(Certifications) ≥ count(Assignments)` for that DCO.

---

## 7 · `Delegations`

Authorizes one user to sign on another's behalf for a bounded scope.

| Column | Type | Req | Indexed | Notes |
|---|---|---|---|---|
| `RecordId` | Single line | ✓ | ✓ | `DEL-NNNN` |
| `FromUser` | Person/Group | ✓ | ✓ | The original signatory |
| `ToUser` | Person/Group | ✓ | ✓ | The delegate |
| `ScopeType` | Choice | ✓ | | `SingleDCO`, `AllDCOsUntil`, `RoleScope` |
| `ScopeDCO` | Lookup → ChangeOrders | | | Required when `ScopeType=SingleDCO` |
| `ValidFrom` | Date+Time | ✓ | | |
| `ValidUntil` | Date+Time | ✓ | ✓ | Hard cap; UI rejects open-ended delegations |
| `Justification` | Multi-line | ✓ | | Required (e.g. "Travel Apr 13–17") |
| `AuthorizedBySignatureRef` | Lookup → Signatures | ✓ | | The Part 11 sig that created this delegation |
| `RevokedAt` | Date+Time | | | Null while active |
| `RevokedBySignatureRef` | Lookup → Signatures | | | Set when revoked early |

A delegated signature on a DCO writes `Signatures.DelegationFromUser` + `DelegationRecordRef` so the audit trail names both parties.

---

## 8 · `Issues`, `Decisions`, `MeetingMinutes`, `ActionItems`

Used by the PM Tracker. Lighter schemas — these don't have Part 11 implications, but they do feed the executive dashboard.

### `Issues`

| Column | Type | Req | Notes |
|---|---|---|---|
| `RecordId` | Single line | ✓ | `ISS-NNN` |
| `Title` | Single line | ✓ | |
| `Description` | Multi-line | ✓ | |
| `Owner` | Person/Group | ✓ | |
| `Severity` | Choice | ✓ | `Low`, `High`, `Critical` |
| `Status` | Choice | ✓ | `Open`, `InProgress`, `Mitigated`, `Closed` |
| `OpenedAt` | Date | ✓ | |
| `ResolvedAt` | Date | | |
| `LinkedDCOs` | Lookup → ChangeOrders | | Multi |
| `RegulatoryCitation` | Single line | | e.g. `§111.75(a)(1)` |

### `Decisions`

| Column | Type | Req | Notes |
|---|---|---|---|
| `RecordId` | Single line | ✓ | `DEC-NNN` |
| `Title` | Single line | ✓ | |
| `Description` | Multi-line | ✓ | |
| `MadeAt` | Date | ✓ | |
| `MadeIn` | Lookup → MeetingMinutes | | |
| `RecordedBy` | Person/Group | ✓ | |

### `MeetingMinutes`

| Column | Type | Req | Notes |
|---|---|---|---|
| `RecordId` | Single line | ✓ | `MM-NNN` |
| `Title` | Single line | ✓ | |
| `HeldAt` | Date+Time | ✓ | |
| `DurationMin` | Number | | |
| `Attendees` | Person/Group | ✓ | Multi |
| `Notes` | Multi-line (rich) | ✓ | |
| `LinkedDecisions` | Lookup → Decisions | | Multi |
| `LinkedActions` | Lookup → ActionItems | | Multi |

### `ActionItems`

| Column | Type | Req | Notes |
|---|---|---|---|
| `RecordId` | Single line | ✓ | `AC-NNN` |
| `Title` | Single line | ✓ | |
| `Owner` | Person/Group | ✓ | |
| `Status` | Choice | ✓ | `Open`, `InProgress`, `Blocked`, `Done`, `Cancelled` |
| `Priority` | Choice | ✓ | `Low`, `Medium`, `High`, `Critical` |
| `DueDate` | Date | | |
| `LinkedMeeting` | Lookup → MeetingMinutes | | |
| `LinkedDCO` | Lookup → ChangeOrders | | |

---

## 9 · `Milestones` & `BudgetEntries`

PM Tracker only.

### `Milestones`

| Column | Type | Req | Notes |
|---|---|---|---|
| `RecordId` | Single line | ✓ | `M-NN`, `W1-SO`, etc. |
| `Title` | Single line | ✓ | |
| `Description` | Multi-line | | |
| `Phase` | Choice | ✓ | |
| `TargetDate` | Date | ✓ | |
| `ActualDate` | Date | | |
| `Status` | Choice | ✓ | `NotStarted`, `InProgress`, `Complete`, `PastDue` |
| `Progress` | Number (0–100) | | Manual on PM Tracker; auto-derived where possible |
| `LinkedDCOs` | Lookup → ChangeOrders | | Multi |

### `BudgetEntries`

| Column | Type | Req | Notes |
|---|---|---|---|
| `RecordId` | Single line | ✓ | `INV-NNNN` or `PSO-NNN` |
| `Type` | Choice | ✓ | `PSO`, `Invoice`, `ChangeOrder` |
| `Phase` | Choice | ✓ | |
| `Amount` | Currency | ✓ | |
| `Status` | Choice | ✓ | `Authorized`, `Pending`, `Paid`, `Cancelled` |
| `IssuedAt` | Date | ✓ | |
| `PaidAt` | Date | | |
| `Reference` | Single line | | External ref number |

---

## 10 · `Documents` (the canonical document library)

The end state. Drafts live in `/Drafts`; published live in `/Official`. Both are SharePoint document libraries with these metadata columns.

| Column | Type | Req | Notes |
|---|---|---|---|
| `DocumentId` | Single line | ✓ | `SOP-PM-021` |
| `Title` | Single line | ✓ | |
| `DocType` | Choice | ✓ | |
| `Revision` | Single line | ✓ | `r4` |
| `Status` | Choice | ✓ | `Draft`, `Effective`, `Superseded`, `Withdrawn` |
| `EffectiveDate` | Date | | |
| `SupersededDate` | Date | | |
| `SupersededByRev` | Single line | | `r5` |
| `OwningDCO` | Lookup → ChangeOrders | ✓ | The DCO that made this revision effective |
| `Owner` | Person/Group | ✓ | |
| `RetentionUntil` | Date | ✓ | Per QM retention policy — minimum 7 years for production records |

**Publishing:** the PM Publish action moves the file from `/Drafts` to `/Official`, stamps `EffectiveDate`, and writes the prior revision's `Status=Superseded` + `SupersededDate` + `SupersededByRev` in a single transaction. Failure on any step rolls back all three.

---

## 11 · Audit & integrity (cross-cutting)

### Hash chain

Every row in `Signatures` and `ActivityLog` carries `SignatureHash` = SHA-256 of:

```
RecordId || ActorUPN || Verb || RelatedRecordId || Timestamp(ISO-8601) || PreviousHash
```

`PreviousHash` is the prior row's hash, ordered by `Timestamp` ascending within the list. The first row in each list has `PreviousHash = '0000...'`. A scheduled job verifies the chain hourly; mismatches alert the QA Approver and lock the affected DCO until reviewed.

### Append-only enforcement

- SP permissions: `Signatures` and `ActivityLog` have **no Edit, no Delete** for any user including site collection admins.
- Custom remote event receivers (`ItemUpdating`, `ItemDeleting`) on these lists throw `403 Forbidden` regardless of permission.
- All write paths go through a server-side API that adds the row + computes the hash atomically. Direct SP REST writes to these lists are blocked at the firewall.

### Time

All timestamps are **server-generated UTC**. Client times are recorded but never authoritative. Validation must prove this.

---

## 12 · Permissions matrix

| Role | CR | DCO (own) | DCO (others) | Sign | Reject | Delegate | Publish | View Audit |
|---|---|---|---|---|---|---|---|---|
| Originator (any) | C/R/U | R/U | R | — | — | — | — | R |
| Reviewer (assigned) | R | R | R | ✓ (assigned) | ✓ (assigned) | ✓ (own) | — | R |
| QA Approver (Tina) | R | R/U | R/U | ✓ | ✓ | ✓ | ✓ | R |
| QA Operations (QD Yang) | R | R | R | ✓ (assigned) | ✓ (assigned) | ✓ (own) | — | R |
| Project Manager (Andre) | R | R | R | ✓ (assigned) | ✓ (assigned) | ✓ (own) | ✓ | R |
| Production (Liu Hong) | R | R | R | ✓ (assigned) | ✓ (assigned) | ✓ (own) | — | R |
| 3H Executive | R (own) | R | R | — | — | — | — | R (own) |
| ADB consultant | R | R | R | ✓ (assigned) | ✓ (assigned) | ✓ (own) | — | R |

C=Create, R=Read, U=Update, ✓=Action allowed.

**Item-level access** is enforced via `Visibility=Internal` on records 3H wants to keep private from ADB, and vice versa. Default is `Both`.

---

## 13 · Migration & seeding

For Phase 2 cutover:

1. **Seed reference data first** — phases, choice columns, role groups.
2. **Backfill closed CRs/DCOs** as `Status=Closed` with their original `Created` dates preserved (set via SP REST `Created` override during migration window only).
3. **Backfill effective documents** into `/Official` with their original `EffectiveDate`. The hash chain in `Signatures` starts fresh from cutover; legacy signatures are imported as a single "migration attestation" row signed by Andre + Tina.
4. **Live records (open CRs/DCOs)** migrate last, in a window where the legacy system is locked. Re-route through the new UI before unlocking.

---

## 14 · Open questions for engineering review

1. **MFA/step-up** — is 3H's AAD tenant configured for Conditional Access step-up on a sign event? If not, is password re-auth alone Part 11–defensible? (Talk to 3H IT.)
2. **Redline rendering** — DOCX compare via Word Online iframe, or pre-render to PDF on submit? Affects `RedlineFileRef` semantics.
3. **Retention auto-disposition** — does QA want auto-archive on `RetentionUntil` reach, or manual review?
4. **External reviewers** — does the DCO routing ever need to include a non-tenant user (e.g. an FDA inspector for audit)? If so, B2B guest accounts + scope.
5. **Mobile** — calm pair is desktop-first. Phone form factor needed for travel weeks? If yes, the DCO detail page needs a mobile redesign (currently 2-column at ≥1100px).

---

**Sign-off needed before build starts:**

- [ ] Andre Butler (PM, ADB) — overall scope
- [ ] Tina Qin (QA Approver, 3H) — permissions matrix, Part 11 enforcement
- [ ] 3H IT — MFA/step-up answer (Q1 above)
- [ ] Validation lead — IQ/OQ/PQ protocols can map to this schema

Once signed, this document becomes the input to the SPFx provisioning script and the validation traceability matrix.
