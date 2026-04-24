// patch_dco_doc_editor.js
// Adds document editor to Draft DCO modal — Documents tab
// Shows current DCO_Docs with remove buttons + searchable add from Drafts folders
// Saves to DCO_Docs field via SP REST PATCH immediately on change
// Run from SPFx root: node patch_dco_doc_editor.js

const fs = require('fs');
const filePath = './src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts';
console.log('Reading:', filePath);
let src = fs.readFileSync(filePath, 'utf8');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 1 — Inject doc editor into Documents pane for Draft phase
//           Insert after the existing isEffective block in openDCOInline
// ─────────────────────────────────────────────────────────────────────────────
const P1_OLD = `      const docReviewEl = d.querySelector('.doc-review-section');
      const gateEl = d.getElementById('sgate-' + dcoId);
      if (docsPaneEl && docReviewEl) docsPaneEl.appendChild(docReviewEl);
      if (docsPaneEl && gateEl) docsPaneEl.appendChild(gateEl);`;

const P1_NEW = `      const docReviewEl = d.querySelector('.doc-review-section');
      const gateEl = d.getElementById('sgate-' + dcoId);
      if (docsPaneEl && docReviewEl) docsPaneEl.appendChild(docReviewEl);
      if (docsPaneEl && gateEl) docsPaneEl.appendChild(gateEl);

      // ── Draft phase — document editor ──────────────────────────────────────
      if (!isEffective && docsPaneEl) {
        const _editorWrap = d.createElement('div');
        _editorWrap.id = 'dco-doc-editor-' + dcoId;
        _editorWrap.innerHTML = \`
          <div style="margin-bottom:12px">
            <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:var(--s5);margin-bottom:8px">
              Documents Assigned to \${dcoId} <span id="dco-doc-count-\${dcoId}" style="font-family:var(--mono);background:var(--b1);color:var(--b);padding:1px 7px;border-radius:10px;font-size:11px">\${displayDocs.length}</span>
            </div>
            <div id="dco-doc-list-\${dcoId}" style="border:1px solid var(--s2);border-radius:7px;overflow:hidden;margin-bottom:10px"></div>
            <div style="display:flex;gap:8px;align-items:center">
              <input id="dco-doc-search-\${dcoId}" type="text" placeholder="Type doc ID or name to add (e.g. SOP-QMS-001)..."
                style="flex:1;font-size:12px;padding:7px 10px;border:1px solid var(--s2);border-radius:6px;outline:none;font-family:var(--sans)">
              <button id="dco-doc-add-btn-\${dcoId}"
                style="font-size:12px;font-weight:600;padding:7px 14px;border-radius:6px;background:var(--b);color:#fff;border:none;cursor:pointer;white-space:nowrap">
                + Add
              </button>
            </div>
            <div id="dco-doc-suggestions-\${dcoId}" style="border:1px solid var(--s2);border-radius:6px;background:var(--w);margin-top:4px;display:none;max-height:180px;overflow-y:auto"></div>
          </div>\`;
        docsPaneEl.appendChild(_editorWrap);

        // Known documents catalogue — populated from Drafts folders
        const _allDocs: Record<string, {name:string, zone:string, rev:string}> = {
          'QM-001':      { name:'Quality Manual',                          zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-QMS-001': { name:'Management Responsibility',               zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-QMS-002': { name:'Document Control',                        zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-QMS-003': { name:'Change Control',                          zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-SUP-001': { name:'Supplier Qualification',                  zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-SUP-002': { name:'Receiving Inspection',                    zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-FS-001':  { name:'Allergen Control',                        zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-FS-002':  { name:'Equipment Cleaning',                      zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-FS-003':  { name:'Facility Sanitation',                     zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-FS-004':  { name:'Environmental Monitoring',                zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-PC-001':  { name:'Pest Sighting Response',                  zone:'Documents/Drafts', rev:'Rev A' },
          'SOP-PRD-108': { name:'Finished Product Release Management',     zone:'Documents/Drafts', rev:'Rev B' },
          'SOP-PRD-432': { name:'Finished Product Specifications',         zone:'Documents/Drafts', rev:'Rev B' },
          'SOP-FRS-549': { name:'Product Specification Sheet',             zone:'Documents/Drafts', rev:'Rev B' },
          'SOP-RCL-321': { name:'Product Recall Procedure',                zone:'Documents/Drafts', rev:'Rev B' },
          'FM-001':      { name:'Master Document Log',                     zone:'Forms/Drafts', rev:'Rev A' },
          'FM-002':      { name:'Change Request Form',                     zone:'Forms/Drafts', rev:'Rev A' },
          'FM-003':      { name:'Document Change Order Form',              zone:'Forms/Drafts', rev:'Rev A' },
          'FM-004':      { name:'Approved Supplier List',                  zone:'Forms/Drafts', rev:'Rev A' },
          'FM-005':      { name:'Receiving Log',                           zone:'Forms/Drafts', rev:'Rev A' },
          'FM-006':      { name:'Raw Material Specification Sheet',        zone:'Forms/Drafts', rev:'Rev A' },
          'FM-007':      { name:'Material Hold Label',                     zone:'Forms/Drafts', rev:'Rev A' },
          'FM-008':      { name:'Supplier CoA Requirements Checklist',     zone:'Forms/Drafts', rev:'Rev A' },
          'FM-027':      { name:'QU/QS Designation Record',                zone:'Forms/Drafts', rev:'Rev A' },
          'FM-030':      { name:'Finished Product Spec Sheet Template',    zone:'Forms/Drafts', rev:'Rev A' },
          'FM-ALG':      { name:'Allergen Status Record',                  zone:'Forms/Drafts', rev:'Rev A' },
          'FPS-001':     { name:'Lychee VD3 Gummy Finished Product Spec',  zone:'Documents/Drafts', rev:'Rev A' },
        };

        // Live doc list — re-renders on every change
        const _renderDocList = (currentDocs: string[]) => {
          const listEl = d.getElementById('dco-doc-list-' + dcoId);
          const countEl = d.getElementById('dco-doc-count-' + dcoId);
          if (!listEl) return;
          if (countEl) countEl.textContent = String(currentDocs.length);
          if (!currentDocs.length) {
            listEl.innerHTML = '<div style="padding:14px;text-align:center;color:var(--s5);font-size:12px">No documents assigned — use the search below to add</div>';
            return;
          }
          listEl.innerHTML = currentDocs.map((docId: string, i: number) => {
            const info = _allDocs[docId] || { name: docId, zone: '—', rev: '—' };
            return \`<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;border-bottom:1px solid var(--s1)\${i===currentDocs.length-1?';border-bottom:none':''}" id="dco-doc-row-\${dcoId}-\${i}">
              <div style="width:22px;height:22px;border-radius:4px;background:var(--b1);display:flex;align-items:center;justify-content:center;font-size:10px;color:var(--b);font-weight:700;flex-shrink:0">📄</div>
              <div style="flex:1;min-width:0">
                <span style="font-family:var(--mono);font-size:11px;font-weight:700;color:var(--b)">\${docId}</span>
                <span style="font-size:11px;color:var(--s5);margin-left:8px">\${info.name}</span>
              </div>
              <span style="font-size:10px;font-family:var(--mono);padding:1px 6px;border-radius:3px;background:var(--b1);color:var(--b);flex-shrink:0">\${info.rev}</span>
              <span style="font-size:9px;color:var(--s5);flex-shrink:0">\${info.zone}</span>
              <button data-remove-doc="\${docId}" data-dcoid="\${dcoId}"
                style="font-size:11px;padding:2px 8px;border-radius:4px;border:1px solid var(--r1);color:var(--r);background:var(--r1);cursor:pointer;flex-shrink:0;white-space:nowrap">
                ✕ Remove
              </button>
            </div>\`;
          }).join('');
          // Wire remove buttons
          listEl.querySelectorAll('[data-remove-doc]').forEach((btn: Element) => {
            btn.addEventListener('click', async () => {
              const removeId = (btn as HTMLElement).getAttribute('data-remove-doc') || '';
              const newDocs = _currentDocIds.filter((id: string) => id !== removeId);
              await _saveDocs(newDocs);
              _currentDocIds = newDocs;
              _renderDocList(_currentDocIds);
              if (w.qpToast) w.qpToast('Removed: ' + removeId);
            });
          });
        };

        // Save DCO_Docs to SharePoint
        const _saveDocs = async (newDocs: string[]) => {
          const dcoItem2 = (this._data.dcos || []).find((x: any) => x.Title === dcoId);
          if (!dcoItem2?.Id) return;
          const base2 = this.context.pageContext.web.absoluteUrl;
          const newVal = newDocs.join(',');
          await this.context.spHttpClient.post(
            base2 + "/_api/web/lists/getbytitle('QMS_DCOs')/items(" + dcoItem2.Id + ")",
            SPHttpClient.configurations.v1,
            { headers: { 'Accept':'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata','IF-MATCH':'*','X-HTTP-Method':'MERGE' },
              body: JSON.stringify({ DCO_Docs: newVal }) }
          );
          dcoItem2.DCO_Docs = newVal;
        };

        // Current doc list — mutable
        let _currentDocIds: string[] = (dco.DCO_Docs || '').split(',').map((s: string) => s.trim()).filter(Boolean);
        _renderDocList(_currentDocIds);

        // Search / suggestions
        const _searchEl = d.getElementById('dco-doc-search-' + dcoId) as HTMLInputElement;
        const _suggestEl = d.getElementById('dco-doc-suggestions-' + dcoId);
        const _addBtn = d.getElementById('dco-doc-add-btn-' + dcoId);

        const _showSuggestions = (q: string) => {
          if (!_suggestEl) return;
          if (!q) { _suggestEl.style.display = 'none'; return; }
          const ql = q.toLowerCase();
          const matches = Object.entries(_allDocs)
            .filter(([id, info]) => !_currentDocIds.includes(id) && (id.toLowerCase().includes(ql) || info.name.toLowerCase().includes(ql)))
            .slice(0, 8);
          if (!matches.length) { _suggestEl.style.display = 'none'; return; }
          _suggestEl.style.display = 'block';
          _suggestEl.innerHTML = matches.map(([id, info]) =>
            \`<div data-suggest="\${id}" style="padding:8px 12px;cursor:pointer;display:flex;align-items:center;gap:10px;border-bottom:1px solid var(--s1)" onmouseover="this.style.background='var(--b0)'" onmouseout="this.style.background=''">
              <span style="font-family:var(--mono);font-size:11px;font-weight:700;color:var(--b);min-width:90px">\${id}</span>
              <span style="font-size:11px;color:var(--s5);flex:1">\${info.name}</span>
              <span style="font-size:10px;font-family:var(--mono);padding:1px 6px;border-radius:3px;background:var(--b1);color:var(--b)">\${info.rev}</span>
            </div>\`
          ).join('');
          _suggestEl.querySelectorAll('[data-suggest]').forEach((el: Element) => {
            el.addEventListener('click', async () => {
              const addId = (el as HTMLElement).getAttribute('data-suggest') || '';
              if (!addId || _currentDocIds.includes(addId)) return;
              _currentDocIds = [..._currentDocIds, addId];
              await _saveDocs(_currentDocIds);
              _renderDocList(_currentDocIds);
              if (_searchEl) _searchEl.value = '';
              if (_suggestEl) _suggestEl.style.display = 'none';
              if (w.qpToast) w.qpToast('Added: ' + addId);
            });
          });
        };

        if (_searchEl) {
          _searchEl.addEventListener('input', () => _showSuggestions(_searchEl.value));
          _searchEl.addEventListener('keydown', async (e: Event) => {
            const ke = e as KeyboardEvent;
            if (ke.key === 'Enter') {
              const val = _searchEl.value.trim().toUpperCase();
              if (!val) return;
              if (_currentDocIds.includes(val)) { if (w.qpToast) w.qpToast(val + ' already in list'); return; }
              if (!_allDocs[val]) { if (w.qpToast) w.qpToast('Unknown doc ID: ' + val); return; }
              _currentDocIds = [..._currentDocIds, val];
              await _saveDocs(_currentDocIds);
              _renderDocList(_currentDocIds);
              _searchEl.value = '';
              if (_suggestEl) _suggestEl.style.display = 'none';
              if (w.qpToast) w.qpToast('Added: ' + val);
            }
          });
        }

        if (_addBtn && _searchEl) {
          _addBtn.addEventListener('click', async () => {
            const val = _searchEl.value.trim().toUpperCase();
            if (!val) return;
            if (_currentDocIds.includes(val)) { if (w.qpToast) w.qpToast(val + ' already in list'); return; }
            if (!_allDocs[val]) { if (w.qpToast) w.qpToast('Unknown doc ID: ' + val + ' — use suggestions list'); return; }
            _currentDocIds = [..._currentDocIds, val];
            await _saveDocs(_currentDocIds);
            _renderDocList(_currentDocIds);
            _searchEl.value = '';
            if (_suggestEl) _suggestEl.style.display = 'none';
            if (w.qpToast) w.qpToast('Added: ' + val);
          });
        }
      }`;

if (!src.includes(P1_OLD)) { console.error('PATCH FAILED — anchor not found'); process.exit(1); }
src = src.replace(P1_OLD, P1_NEW);
console.log('✓ Patch — DCO document editor added for Draft phase');

fs.writeFileSync(filePath, src, 'utf8');
console.log('\n✅ Patch applied. Ready to build.');
