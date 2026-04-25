// patch_doc_editor_v2.js
// Upgrades the DCO document editor:
//   1. Exclude docs already assigned to another non-Effective DCO
//   2. Replace search box with click-to-open full alpha list
// Run from SPFx root: node patch_doc_editor_v2.js

const fs = require('fs');
const filePath = './src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts';
console.log('Reading:', filePath);
let src = fs.readFileSync(filePath, 'utf8');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH — Replace the search/suggestions block with locked-doc logic + alpha list
// ─────────────────────────────────────────────────────────────────────────────
const P_OLD = `        // Search / suggestions
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
        }`;

const P_NEW = `        // Build locked-doc set — docs on other non-Effective DCOs
        const _lockedDocs = new Set<string>();
        (this._data.dcos || []).forEach((otherDco: any) => {
          if (otherDco.Title === dcoId) return; // skip current DCO
          if ((otherDco.DCO_Phase || '') === 'Effective') return; // Effective = released, doc is free
          const otherDocs = (otherDco.DCO_Docs || '').split(',').map((s: string) => s.trim()).filter(Boolean);
          otherDocs.forEach((id: string) => _lockedDocs.add(id));
        });

        // Build available list — all known docs minus current + locked, sorted alpha
        const _getAvailable = (): Array<[string, {name:string,zone:string,rev:string}]> =>
          Object.entries(_allDocs)
            .filter(([id]) => !_currentDocIds.includes(id) && !_lockedDocs.has(id))
            .sort(([a], [b]) => a.localeCompare(b));

        // Replace search/button HTML with click-to-open panel
        const _searchEl = d.getElementById('dco-doc-search-' + dcoId) as HTMLInputElement;
        const _suggestEl = d.getElementById('dco-doc-suggestions-' + dcoId);
        const _addBtn = d.getElementById('dco-doc-add-btn-' + dcoId);

        // Replace the input+button row with a single "Add Document" trigger
        const _inputRow = _searchEl?.parentElement;
        if (_inputRow) {
          _inputRow.innerHTML = \`
            <button id="dco-add-trigger-\${dcoId}"
              style="width:100%;padding:9px 14px;border:1px dashed var(--b);border-radius:6px;background:var(--b0);color:var(--b);font-size:12px;font-weight:600;cursor:pointer;text-align:left;display:flex;align-items:center;gap:8px">
              <span style="font-size:16px">+</span>
              <span id="dco-add-trigger-lbl-\${dcoId}">Click to add a document...</span>
            </button>\`;
        }

        // Available docs panel — shown/hidden on trigger click
        if (_suggestEl) {
          _suggestEl.style.cssText = 'border:1px solid var(--s2);border-radius:6px;background:var(--w);margin-top:6px;display:none;max-height:280px;overflow-y:auto;box-shadow:0 4px 14px rgba(0,0,0,.1)';
        }

        const _renderAvailableList = () => {
          if (!_suggestEl) return;
          const avail = _getAvailable();
          if (!avail.length) {
            _suggestEl.innerHTML = '<div style="padding:14px;text-align:center;color:var(--s5);font-size:12px">All documents are either assigned or locked on another active DCO</div>';
            return;
          }
          // Group by zone for clarity
          const docZone: Record<string, Array<[string, {name:string,zone:string,rev:string}]>> = {};
          avail.forEach(([id, info]) => {
            if (!docZone[info.zone]) docZone[info.zone] = [];
            docZone[info.zone].push([id, info]);
          });
          _suggestEl.innerHTML = Object.entries(docZone).map(([zone, docs]) =>
            \`<div style="padding:5px 12px 3px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--s5);background:var(--s0);border-bottom:1px solid var(--s1)">\${zone}</div>\` +
            docs.map(([id, info]) =>
              \`<div data-add-doc="\${id}"
                style="padding:8px 12px;cursor:pointer;display:flex;align-items:center;gap:10px;border-bottom:1px solid var(--s1);transition:background .1s"
                onmouseover="this.style.background='var(--b0)'" onmouseout="this.style.background=''">
                <span style="font-family:var(--mono);font-size:11px;font-weight:700;color:var(--b);min-width:90px">\${id}</span>
                <span style="font-size:11px;color:var(--s7);flex:1">\${info.name}</span>
                <span style="font-size:10px;font-family:var(--mono);padding:1px 6px;border-radius:3px;background:var(--b1);color:var(--b);flex-shrink:0">\${info.rev}</span>
              </div>\`
            ).join('')
          ).join('');
          // Wire add clicks
          _suggestEl.querySelectorAll('[data-add-doc]').forEach((el: Element) => {
            el.addEventListener('click', async () => {
              const addId = (el as HTMLElement).getAttribute('data-add-doc') || '';
              if (!addId || _currentDocIds.includes(addId) || _lockedDocs.has(addId)) return;
              _currentDocIds = [..._currentDocIds, addId];
              await _saveDocs(_currentDocIds);
              _renderDocList(_currentDocIds);
              _renderAvailableList(); // refresh list to remove just-added doc
              const lbl = d.getElementById('dco-add-trigger-lbl-' + dcoId);
              const remaining = _getAvailable().length;
              if (lbl) lbl.textContent = remaining > 0 ? \`Click to add a document (\${remaining} available)\` : 'All available documents assigned';
              if (w.qpToast) w.qpToast('Added: ' + addId);
            });
          });
        };

        // Wire trigger button
        const _trigger = d.getElementById('dco-add-trigger-' + dcoId);
        let _panelOpen = false;
        if (_trigger && _suggestEl) {
          // Update label with available count
          const _availCount = _getAvailable().length;
          const _lbl = d.getElementById('dco-add-trigger-lbl-' + dcoId);
          if (_lbl) _lbl.textContent = _availCount > 0
            ? \`Click to add a document (\${_availCount} available)\`
            : 'All available documents assigned or locked on other DCOs';
          _trigger.addEventListener('click', () => {
            _panelOpen = !_panelOpen;
            if (_panelOpen) {
              _renderAvailableList();
              _suggestEl.style.display = 'block';
              _trigger.style.borderStyle = 'solid';
            } else {
              _suggestEl.style.display = 'none';
              _trigger.style.borderStyle = 'dashed';
            }
          });
        }

        // Close panel when clicking outside
        d.addEventListener('click', (e: Event) => {
          const tgt = e.target as HTMLElement;
          if (_suggestEl && !_suggestEl.contains(tgt) && tgt.id !== 'dco-add-trigger-' + dcoId && !(_trigger?.contains(tgt))) {
            _suggestEl.style.display = 'none';
            _panelOpen = false;
            if (_trigger) (_trigger as HTMLElement).style.borderStyle = 'dashed';
          }
        }, { capture: false });`;

if (!src.includes(P_OLD)) { console.error('PATCH FAILED — anchor not found'); process.exit(1); }
src = src.replace(P_OLD, P_NEW);
console.log('✓ Patch — locked-doc logic + alpha list on click');

fs.writeFileSync(filePath, src, 'utf8');
console.log('\n✅ Patch applied. Ready to build.');
