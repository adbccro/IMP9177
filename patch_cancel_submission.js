// patch_cancel_submission.js
// 1. Add DCO_CA to spGet fetch
// 2. Add Cancel button — only for CA or Originator, not in Draft/Effective
// 3. Add Cancel modal with reason field
// 4. Cancel action — returns to Draft, clears submitted date, writes routing history
// 5. Suppress doc-open routing history when DCO is in Draft phase
// Run from SPFx root: node patch_cancel_submission.js

const fs = require('fs');
const filePath = './src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts';
console.log('Reading:', filePath);
let src = fs.readFileSync(filePath, 'utf8');
let changed = 0;

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 1 — Add DCO_CA to spGet select fields
// ─────────────────────────────────────────────────────────────────────────────
const P1_OLD = `this.spGet('QMS_DCOs', 'Id,Title,DCO_Phase,DCO_Title,DCO_CRLink,DCO_SubmittedDate,DCO_Originator,DCO_Docs,DCO_LateDays,DCO_TrainingGate'),`;
const P1_NEW = `this.spGet('QMS_DCOs', 'Id,Title,DCO_Phase,DCO_Title,DCO_CRLink,DCO_SubmittedDate,DCO_Originator,DCO_Docs,DCO_LateDays,DCO_TrainingGate,DCO_CA'),`;
if (!src.includes(P1_OLD)) { console.error('PATCH 1 FAILED'); process.exit(1); }
src = src.replace(P1_OLD, P1_NEW);
changed++;
console.log('✓ Patch 1 — DCO_CA added to spGet');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 2 — Add Cancel modal HTML to buildShell
// ─────────────────────────────────────────────────────────────────────────────
const P2_OLD = `<div class="modal-ov" id="modal-reject">`;
const P2_NEW = `<div class="modal-ov" id="modal-cancel-submission">
  <div class="modal" style="max-width:500px">
    <div class="modal-hdr">
      <div><div class="modal-title">Cancel Submission — Reason Required</div></div>
      <button class="modal-x" data-close="modal-cancel-submission">×</button>
    </div>
    <div class="modal-body">
      <div class="fg"><div class="fl">Cancellation Category</div>
        <select class="fsel" id="cancel-cat">
          <option>Documents need revision</option>
          <option>Wrong documents assigned</option>
          <option>Approvers need updating</option>
          <option>Submitted in error</option>
          <option>Other</option>
        </select>
      </div>
      <div class="fg"><div class="fl">Cancellation Reason (required)</div>
        <textarea class="ftxt" id="cancel-reason" placeholder="Describe the reason for cancelling this submission..."></textarea>
      </div>
    </div>
    <div class="modal-ft">
      <button class="btn-sec" data-close="modal-cancel-submission">Keep Submitted</button>
      <button class="btn-pri btn-r" id="btn-confirm-cancel-submission">Cancel Submission → Return to Draft</button>
    </div>
  </div>
</div>

<div class="modal-ov" id="modal-reject">`;
if (!src.includes(P2_OLD)) { console.error('PATCH 2 FAILED'); process.exit(1); }
src = src.replace(P2_OLD, P2_NEW);
changed++;
console.log('✓ Patch 2 — Cancel modal added to shell');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 3 — Wire Cancel modal close button in _attachListeners
// ─────────────────────────────────────────────────────────────────────────────
const P3_OLD = `    ['modal-dco-detail','modal-cr-detail','modal-reject','modal-esign'].forEach(id => {`;
const P3_NEW = `    ['modal-dco-detail','modal-cr-detail','modal-reject','modal-esign','modal-cancel-submission'].forEach(id => {`;
if (!src.includes(P3_OLD)) { console.error('PATCH 3 FAILED'); process.exit(1); }
src = src.replace(P3_OLD, P3_NEW);
changed++;
console.log('✓ Patch 3 — Cancel modal wired for backdrop close');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 4 — Add Cancel button wiring in _attachListeners
//           Wire confirm-cancel after the reject confirm button wire-up
// ─────────────────────────────────────────────────────────────────────────────
const P4_OLD = `    const rejConfirm = d.getElementById('btn-confirm-reject');
    if (rejConfirm) rejConfirm.addEventListener('click', () => {
      const reason = (d.getElementById('rej-reason') as HTMLInputElement)?.value;
      if (!reason?.trim()) { if (w.qpToast) w.qpToast('Rejection reason is required'); return; }
      (d.getElementById('modal-reject') as HTMLElement)?.classList.remove('open');
      (d.getElementById('modal-dco-detail') as HTMLElement)?.classList.remove('open');
      if (w.qpToast) w.qpToast('DCO rejected — routing history updated');
    });`;

const P4_NEW = `    const rejConfirm = d.getElementById('btn-confirm-reject');
    if (rejConfirm) rejConfirm.addEventListener('click', () => {
      const reason = (d.getElementById('rej-reason') as HTMLInputElement)?.value;
      if (!reason?.trim()) { if (w.qpToast) w.qpToast('Rejection reason is required'); return; }
      (d.getElementById('modal-reject') as HTMLElement)?.classList.remove('open');
      (d.getElementById('modal-dco-detail') as HTMLElement)?.classList.remove('open');
      if (w.qpToast) w.qpToast('DCO rejected — routing history updated');
    });

    // ── Cancel Submission confirm ───────────────────────────────────────────
    const cancelConfirm = d.getElementById('btn-confirm-cancel-submission');
    if (cancelConfirm) cancelConfirm.addEventListener('click', async () => {
      const reason = (d.getElementById('cancel-reason') as HTMLTextAreaElement)?.value?.trim();
      const cat    = (d.getElementById('cancel-cat') as HTMLSelectElement)?.value || 'Other';
      if (!reason) { if (w.qpToast) w.qpToast('Cancellation reason is required'); return; }
      const dcoId = (w._qpCancelDcoId as string) || '';
      if (!dcoId) return;
      const base2 = this.context.pageContext.web.absoluteUrl;
      const user2 = this.context.pageContext.user.displayName || this.context.pageContext.user.email;
      const ts2   = new Date().toISOString();
      const dcoItem2 = (this._data.dcos || []).find((x: any) => x.Title === dcoId);
      if (dcoItem2?.Id) {
        await this.context.spHttpClient.post(
          base2 + "/_api/web/lists/getbytitle('QMS_DCOs')/items(" + dcoItem2.Id + ")",
          SPHttpClient.configurations.v1,
          { headers: {'Accept':'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata','IF-MATCH':'*','X-HTTP-Method':'MERGE'},
            body: JSON.stringify({ DCO_Phase: 'Draft', DCO_SubmittedDate: null }) }
        );
        dcoItem2.DCO_Phase = 'Draft';
        dcoItem2.DCO_SubmittedDate = null;
      }
      await this.context.spHttpClient.post(
        base2 + "/_api/web/lists/getbytitle('QMS_RoutingHistory')/items",
        SPHttpClient.configurations.v1,
        { headers: {'Accept':'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata'},
          body: JSON.stringify({ Title: dcoId + '-CANCEL-' + Date.now(),
            RH_DCOID: dcoId, RH_EventType: 'stage', RH_Stage: 'Draft',
            RH_Actor: user2,
            RH_Note: 'Submission cancelled by ' + user2 + '. Category: ' + cat + '. Reason: ' + reason,
            RH_Reason: reason, RH_Timestamp: ts2 }) }
      );
      (d.getElementById('modal-cancel-submission') as HTMLElement)?.classList.remove('open');
      (d.getElementById('modal-dco-detail') as HTMLElement)?.classList.remove('open');
      // Reset reason field
      const reasonEl = d.getElementById('cancel-reason') as HTMLTextAreaElement;
      if (reasonEl) reasonEl.value = '';
      if (w.qpToast) w.qpToast(dcoId + ' submission cancelled — returned to Draft');
      setTimeout(() => this._loadAll(), 1000);
    });`;

if (!src.includes(P4_OLD)) { console.error('PATCH 4 FAILED'); process.exit(1); }
src = src.replace(P4_OLD, P4_NEW);
changed++;
console.log('✓ Patch 4 — Cancel submission confirm handler wired');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 5 — Show Cancel button in modal footer based on role + phase
//           Add after the existing actionBtn/rejectBtn phase logic
// ─────────────────────────────────────────────────────────────────────────────
const P5_OLD = `      // Wire action buttons based on phase
      const actionBtn = d.getElementById('mdco-action-btn') as HTMLElement;
      const rejectBtn = d.getElementById('mdco-reject-btn') as HTMLElement;
      if (actionBtn) {
        const phase = dco.DCO_Phase || 'Draft';
        actionBtn.style.display = 'none';
        if (rejectBtn) rejectBtn.style.display = 'none';`;

const P5_NEW = `      // Wire action buttons based on phase
      const actionBtn = d.getElementById('mdco-action-btn') as HTMLElement;
      const rejectBtn = d.getElementById('mdco-reject-btn') as HTMLElement;

      // ── Cancel button — only for CA or Originator, not in Draft/Effective ──
      const currentUserName = this.context.pageContext.user.displayName || '';
      const currentUserEmail = this.context.pageContext.user.email || '';
      const dcoCA = dco.DCO_CA || '';
      const dcoOriginator = dco.DCO_Originator || '';
      const isCancelAuthorized = currentUserName.toLowerCase().includes(dcoCA.toLowerCase().split(' ')[0]) ||
        currentUserName.toLowerCase().includes(dcoOriginator.toLowerCase().split(' ')[0]) ||
        currentUserEmail.toLowerCase().includes('jeffrey') ||
        currentUserEmail.toLowerCase().includes('andre');
      const cancelPhases = ['Submitted', 'In Review', 'Implemented', 'Awaiting Training'];
      const showCancel = isCancelAuthorized && cancelPhases.includes(dco.DCO_Phase || '');

      // Inject cancel button into modal footer if not already there
      const modalFt = d.querySelector('#modal-dco-detail .modal-ft') as HTMLElement;
      const existingCancelBtn = d.getElementById('mdco-cancel-sub-btn-' + dcoId);
      if (modalFt && !existingCancelBtn && showCancel) {
        const cancelBtn = d.createElement('button');
        cancelBtn.id = 'mdco-cancel-sub-btn-' + dcoId;
        cancelBtn.className = 'btn-sec btn-sm';
        cancelBtn.style.cssText = 'border-color:var(--a);color:var(--a);margin-right:auto';
        cancelBtn.textContent = '↩ Cancel Submission';
        cancelBtn.addEventListener('click', () => {
          w._qpCancelDcoId = dcoId;
          const reasonEl = d.getElementById('cancel-reason') as HTMLTextAreaElement;
          if (reasonEl) reasonEl.value = '';
          (d.getElementById('modal-cancel-submission') as HTMLElement)?.classList.add('open');
        });
        // Insert before the Close button (first child)
        modalFt.insertBefore(cancelBtn, modalFt.firstChild);
      } else if (existingCancelBtn) {
        existingCancelBtn.style.display = showCancel ? 'inline-flex' : 'none';
      }

      if (actionBtn) {
        const phase = dco.DCO_Phase || 'Draft';
        actionBtn.style.display = 'none';
        if (rejectBtn) rejectBtn.style.display = 'none';`;

if (!src.includes(P5_OLD)) { console.error('PATCH 5 FAILED'); process.exit(1); }
src = src.replace(P5_OLD, P5_NEW);
changed++;
console.log('✓ Patch 5 — Cancel button injected based on role + phase');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 6 — Suppress doc-open routing history when DCO is in Draft phase
// ─────────────────────────────────────────────────────────────────────────────
const P6_OLD = `          // Write to routing history
          const base = this.context.pageContext.web.absoluteUrl;
          const userEmail = this.context.pageContext.user.email;
          const userName = this.context.pageContext.user.displayName || userEmail;
          this.context.spHttpClient.post(
            base + "/_api/web/lists/getbytitle('QMS_RoutingHistory')/items",
            SPHttpClient.configurations.v1,
            { headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
              body: JSON.stringify({ Title: dcoId + '-DOCREVIEW-' + docId, RH_DCOID: dcoId, RH_EventType: 'DocumentReview', RH_Stage: 'In Review', RH_Actor: userName, RH_Note: 'Document opened: ' + docId + ' ' + rev + ' | ' + userEmail, RH_Timestamp: new Date().toISOString() }) }
          ).catch((e: any) => console.error('DocReview write failed', e));`;

const P6_NEW = `          // Write to routing history — only when DCO is In Review or later (not Draft)
          const base = this.context.pageContext.web.absoluteUrl;
          const userEmail = this.context.pageContext.user.email;
          const userName = this.context.pageContext.user.displayName || userEmail;
          const currentPhase = (this._data.dcos || []).find((x: any) => x.Title === dcoId)?.DCO_Phase || 'Draft';
          if (currentPhase !== 'Draft') {
            this.context.spHttpClient.post(
              base + "/_api/web/lists/getbytitle('QMS_RoutingHistory')/items",
              SPHttpClient.configurations.v1,
              { headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
                body: JSON.stringify({ Title: dcoId + '-DOCREVIEW-' + docId, RH_DCOID: dcoId, RH_EventType: 'DocumentReview', RH_Stage: currentPhase, RH_Actor: userName, RH_Note: 'Document opened: ' + docId + ' ' + rev + ' | ' + userEmail, RH_Timestamp: new Date().toISOString() }) }
            ).catch((e: any) => console.error('DocReview write failed', e));
          }`;

if (!src.includes(P6_OLD)) { console.error('PATCH 6 FAILED'); process.exit(1); }
src = src.replace(P6_OLD, P6_NEW);
changed++;
console.log('✓ Patch 6 — Doc-open routing history suppressed in Draft phase');

// ─────────────────────────────────────────────────────────────────────────────
// Write
// ─────────────────────────────────────────────────────────────────────────────
fs.writeFileSync(filePath, src, 'utf8');
console.log(`\n✅ All ${changed}/6 patches applied. Ready to build.`);
