# CR Panel Injection — Manual Steps

The patch script adds all CR methods and wires _renderCR() to be called from _renderAll().

However, the CR tab HTML panel (sc-cr div content) needs to be populated at runtime because 
buildShell() is a static string defined before data loads. The patch wires this through 
_renderCR() which calls _renderAll() → sc-cr gets populated.

## What the patch does NOT auto-do:

The existing buildShell() has a stub sc-cr panel like:
  <div class="sc" id="sc-cr">
    <div class="panel-hdr">Change Requests</div>
    <!-- empty or just NewForm button -->
  </div>

The _renderCR() method writes to #cr-list-wrap which must exist inside sc-cr.

## Manual fix required if cr-list-wrap doesn't exist:

In buildShell() find the sc-cr div and replace its content with:
  <div class="sc" id="sc-cr">
    <div style="padding:20px 24px;">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;">
        <div style="display:flex;align-items:center;gap:12px;">
          <h2 style="font-size:16px;font-weight:700;color:#1e3a5f;margin:0">Change Requests</h2>
          <span id="cr-count-badge" style="background:#e5e7eb;color:#374151;font-size:11px;font-weight:700;padding:2px 8px;border-radius:10px">—</span>
        </div>
        <button onclick="window.qpOpenNewCR()" style="padding:6px 14px;background:#1e3a5f;color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">+ New CR</button>
      </div>
      <div id="cr-filter-bar" style="display:flex;gap:10px;align-items:center;margin-bottom:16px;flex-wrap:wrap;">
        <input id="cr-search" type="text" placeholder="Search CR ID or title…"
          style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:12px;width:200px"
          oninput="window._qpRenderCR && window._qpRenderCR()">
        <select id="cr-filter-status"
          style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:12px"
          onchange="window._qpRenderCR && window._qpRenderCR()">
          <option value="">All Statuses</option>
          <option>Draft</option><option>Submitted</option><option>Approved</option>
          <option>Rejected</option><option>Linked</option><option>Closed</option>
        </select>
        <select id="cr-filter-priority"
          style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:12px"
          onchange="window._qpRenderCR && window._qpRenderCR()">
          <option value="">All Priorities</option>
          <option>Low</option><option>Medium</option><option>High</option><option>Critical</option>
        </select>
      </div>
      <div id="cr-list-wrap" style="border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
        <div style="padding:24px;text-align:center;color:#9ca3af;font-size:13px">Loading…</div>
      </div>
    </div>
  </div>

## buildShell() search string to find and replace:
Search for the sc-cr div — it will look something like:

  id="sc-cr"

and its full content until the next </div></div> close. Replace the inner content with the above.
