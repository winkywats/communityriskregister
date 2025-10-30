(function () {
    'use strict';

    /* ============================== *
     *  Utilities / Security          *
     * ============================== */
    function safeUrl(u){
      try {
        const url = new URL(String(u || ''), location.origin);
        if (!/^https?:$/i.test(url.protocol)) return '#';
        return url.href;
      } catch { return '#'; }
    }
    function makeDebounced(fn, ms = 120) {
      let t; return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
    }
    const el = (id) => document.getElementById(id);
    const clamp15 = (n) => Math.max(1, Math.min(5, Number(n || 0)));
    const round1 = (x) => Math.round(x * 10) / 10;
    const normalize = (s) => (s || '').toString().toLowerCase();
    const enabledOf = (key) => (el('en_' + key)?.checked ?? false);
    const escapeHtml = (s) => String(s || '').replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));

    /* ============================== *
     *  Config & constants            *
     * ============================== */
    const CFG = Object.freeze({
      categories: [
        { key: 'individual', title: 'Individual', enabled: false, sub: [
          { id: 'p_phys',  label: 'Physical Harm' },
          { id: 'p_psych', label: 'Psychological Harm' },
          { id: 'p_fin',   label: 'Financial Harm' }
        ]},
        { key: 'societal', title: 'Societal', enabled: false, sub: [
          { id: 's_service', label: 'Service Impact' },
          { id: 's_partner', label: 'Partner Agency Impact' },
          { id: 's_econ',    label: 'Economic Cost' },
          { id: 's_rep',     label: 'Reputation/Political Impact' }
        ]},
        { key: 'responder', title: 'Emergency Responder', enabled: false, sub: [
          { id: 'er_people',  label: 'People' },
          { id: 'er_process', label: 'Process' },
          { id: 'er_plan',    label: 'Plan' }
        ]},
        { key: 'environmental', title: 'Environmental', enabled: false, sub: [
          { id: 'env_air',        label: 'Air quality' },
          { id: 'env_resources',  label: 'Use and disposal of resources' },
          { id: 'env_landwater',  label: 'Land and Water Health' },
          { id: 'env_climate',    label: 'Climate Change' }
        ]},
        { key: 'community', title: 'Community', enabled: false, sub: [
          { id: 'com_public_services', label: 'Public Services Impact' },
          { id: 'com_cni',             label: 'Critical Infrastructure Impact' },
          { id: 'com_employers',       label: 'Large Employer/Employment Impact' },
          { id: 'com_wellbeing',       label: 'Well-being and Mental Health Impact' }
        ]},
        { key: 'heritage', title: 'Heritage', enabled: false, sub: [
          { id: 'her_heritage', label: 'Heritage' }
        ]}
      ],
      likelihoodId: 'o_likelihood',
      mitigationSteps: { 1: 0.00, 2: 0.10, 3: 0.25, 4: 0.40, 5: 0.60 }
    });

    const GUIDANCE = Object.freeze({
      p_phys: ['None / negligible injury', 'Minor first aid', 'Medical attention required', 'Serious injury / hospitalization', 'Multiple serious injuries / fatalities'],
      p_psych: ['No noticeable impact', 'Temporary distress', 'Short-term clinical support needed', 'Long-term individual impact', 'Widespread or severe psychological harm'],
      p_fin:  ['No measurable cost', 'Minor personal costs', 'Significant personal/household costs', 'Severe financial hardship', 'Widespread financial harm']
    });
  
  	let SUB_REASONING = {};


    /* ============================== *
     *  Main list state               *
     * ============================== */
    let LIST_VIEW_MODE = 'grouped'; // backward-compat for grouped/all internals
    let LIST_SEARCH = '';
    const EVENT_SORT = { key: 'score', dir: 'desc' };

    // NEW primary tabs + risks mode
    let PRIMARY_TAB = 'risks';   // 'risks' | 'objectives'
    let RISKS_MODE  = 'all';     // 'all' | 'grouped'
    let HAZARDS_CACHE = [];

    /* ============================== *
     *  App state (editing, file)     *
     * ============================== */
    let editingId = null;
    let editingHazardId = null;

    let currentFileHandle = null;
    let currentFileName = null;
    let currentDriveFileId = null;
    let fileDirty = false;
    let saving = false;
    let lastSavedSnapshot = '';

    /* ============================== *
     *  Google Drive helpers          *
     * ============================== */
    const DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.file';
    const DRIVE_DISABLE_REASON = 'Set the google-drive-client-id meta tag to enable Google Drive features.';
    const DRIVE_NEEDS_FILE_REASON = 'Save to Google Drive first to create a shareable link.';
    const DRIVE_SETUP_ALERT = 'Google Drive integration is not configured. Add your Google Drive client ID to the <meta name="google-drive-client-id"> tag first.';
    const driveClientId = (document.querySelector('meta[name="google-drive-client-id"]')?.content || '').trim();
    let driveTokenClient = null;
    let driveAccessToken = null;
    let driveTokenExpiry = 0;

    function driveConfigured() { return !!driveClientId; }
    function isDriveTokenValid() { return !!driveAccessToken && Date.now() < (driveTokenExpiry - 5000); }
    function clearDriveToken() { driveAccessToken = null; driveTokenExpiry = 0; }

    function driveStatusLabel() {
      if (currentDriveFileId) return ' (Drive)';
      if (currentFileHandle) return ' (Local)';
      return '';
    }

    function warnDriveNotConfigured() {
      alert(DRIVE_SETUP_ALERT);
    }

    function updateDriveUrlParam(id) {
      const url = new URL(location.href);
      if (id) {
        const shareLink = `https://drive.google.com/file/d/${id}/view`;
        url.searchParams.set('driveFile', shareLink);
      } else {
        url.searchParams.delete('driveFile');
      }
      history.replaceState(null, document.title, url.pathname + url.search + url.hash);
    }

    function parseDriveFileId(raw) {
      if (!raw) return null;
      let input = String(raw || '').trim();
      try { input = decodeURIComponent(input); } catch (e) { /* ignore */ }
      if (!input) return null;
      const idMatch = input.match(/[-\w]{10,}/g);
      if (/^[A-Za-z0-9_-]{10,}$/.test(input)) return input;
      const urlMatch = input.match(/\/d\/([A-Za-z0-9_-]{10,})/);
      if (urlMatch) return urlMatch[1];
      const queryMatch = input.match(/[?&]id=([A-Za-z0-9_-]{10,})/);
      if (queryMatch) return queryMatch[1];
      if (idMatch) return idMatch[0];
      return null;
    }

    function ensureGoogleIdentityLoaded(timeoutMs = 6000) {
      if (window.google?.accounts?.oauth2) return Promise.resolve();
      let elapsed = 0;
      return new Promise((resolve, reject) => {
        function check() {
          if (window.google?.accounts?.oauth2) { resolve(); return; }
          elapsed += 120;
          if (elapsed >= timeoutMs) { reject(new Error('Google Identity Services script not available.')); return; }
          setTimeout(check, 120);
        }
        check();
      });
    }

    async function ensureDriveAccessToken(interactive = true) {
      if (!driveConfigured()) throw new Error('Google Drive client ID is not configured.');
      if (isDriveTokenValid()) return driveAccessToken;
      await ensureGoogleIdentityLoaded();
      const scope = DRIVE_SCOPE;
      const tokenRequestOptions = interactive ? {} : { prompt: '' };
      return new Promise((resolve, reject) => {
        try {
          if (!driveTokenClient) {
            driveTokenClient = google.accounts.oauth2.initTokenClient({
              client_id: driveClientId,
              scope,
              callback: (resp) => {
                if (resp.error) {
                  clearDriveToken();
                  reject(new Error(resp.error_description || resp.error));
                } else {
                  driveAccessToken = resp.access_token;
                  driveTokenExpiry = Date.now() + ((resp.expires_in || 3600) * 1000);
                  resolve(driveAccessToken);
                }
              }
            });
          }
          driveTokenClient.callback = (resp) => {
            if (resp.error) {
              clearDriveToken();
              reject(new Error(resp.error_description || resp.error));
            } else {
              driveAccessToken = resp.access_token;
              driveTokenExpiry = Date.now() + ((resp.expires_in || 3600) * 1000);
              resolve(driveAccessToken);
            }
          };
          driveTokenClient.requestAccessToken(tokenRequestOptions);
        } catch (err) {
          clearDriveToken();
          reject(err);
        }
      });
    }

    async function driveAuthorizedFetch(url, options = {}, { interactive = true } = {}) {
      const headers = Object.assign({}, options.headers || {});
      const attempt = async (withAuth) => {
        const opts = Object.assign({}, options, { headers: headers });
        if (withAuth) {
          const token = await ensureDriveAccessToken(interactive);
          opts.headers = Object.assign({}, headers, { Authorization: `Bearer ${token}` });
        }
        const resp = await fetch(url, opts);
        if (!withAuth && (resp.status === 401 || resp.status === 403)) {
          if (!driveConfigured()) return resp;
          return attempt(true);
        }
        return resp;
      };
      return attempt(false);
    }

    async function fetchDriveMetadata(fileId, opts = {}) {
      if (!fileId) return null;
      const url = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?fields=id,name,mimeType&supportsAllDrives=true`;
      const resp = await driveAuthorizedFetch(url, {}, opts);
      if (!resp.ok) return null;
      return resp.json().catch(() => null);
    }

    async function downloadDriveLitl(fileId, opts = {}) {
      if (!fileId) throw new Error('Missing Google Drive file ID.');
      const mediaUrl = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media&supportsAllDrives=true`;
      const resp = await driveAuthorizedFetch(mediaUrl, {}, opts);
      if (!resp.ok) {
        const text = await resp.text().catch(() => '');
        throw new Error(`Google Drive download failed (${resp.status}) ${text}`);
      }
      const text = await resp.text();
      const json = JSON.parse(text);
      const meta = await fetchDriveMetadata(fileId, opts);
      return { json, meta };
    }

    async function uploadDriveLitl(fileId, blob, name, opts = {}) {
      const headers = { 'Content-Type': blob.type || 'application/x-litl' };
      if (!fileId) throw new Error('Missing Google Drive file ID.');
      const token = await ensureDriveAccessToken(opts.interactive !== false);
      const url = `https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(fileId)}?uploadType=media&supportsAllDrives=true&fields=id,name`;
      const resp = await fetch(url, {
        method: 'PATCH',
        headers: Object.assign({}, headers, { Authorization: `Bearer ${token}` }),
        body: blob
      });
      if (!resp.ok) {
        const text = await resp.text().catch(() => '');
        throw new Error(`Google Drive save failed (${resp.status}) ${text}`);
      }
      return resp.json().catch(() => ({ id: fileId, name }));
    }

    async function createDriveLitl(blob, name, opts = {}) {
      const fileName = name || 'community_risk_register.litl';
      const token = await ensureDriveAccessToken(opts.interactive !== false);
      const boundary = 'litl-' + Math.random().toString(16).slice(2);
      const metadata = { name: fileName, mimeType: 'application/x-litl' };
      const bodyText = await blob.text();
      const body = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: application/x-litl\r\n\r\n${bodyText}\r\n--${boundary}--`;
      const resp = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,name', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': `multipart/related; boundary=${boundary}`
        },
        body
      });
      if (!resp.ok) {
        const text = await resp.text().catch(() => '');
        throw new Error(`Google Drive create failed (${resp.status}) ${text}`);
      }
      return resp.json();
    }

    function driveShareableUrl() {
      if (!currentDriveFileId) return null;
      const share = new URL(location.href);
      share.searchParams.set('driveFile', `https://drive.google.com/file/d/${currentDriveFileId}/view`);
      share.hash = '';
      return share.toString();
    }

    /* ============================== *
     *  Modal a11y helpers            *
     * ============================== */
    let lastFocusEl = null;
    function focusablesIn(root){
      return Array.from(root.querySelectorAll('a,button,input,select,textarea,[tabindex]:not([tabindex="-1"])'))
        .filter(el => !el.hasAttribute('disabled') && el.offsetParent !== null);
    }
    function openModal(modalEl, focusSelector){
      lastFocusEl = document.activeElement;
      modalEl.classList.remove('hidden'); modalEl.classList.add('flex');
      modalEl.setAttribute('aria-hidden','false');
      modalEl.dataset.trap = '1';
      const fs = focusSelector ? modalEl.querySelector(focusSelector) : focusablesIn(modalEl)[0];
      (fs || modalEl).focus();
    }
    function closeModal(modalEl){
      modalEl.classList.add('hidden'); modalEl.classList.remove('flex');
      modalEl.setAttribute('aria-hidden','true');
      delete modalEl.dataset.trap;
      if (lastFocusEl && lastFocusEl.focus) { lastFocusEl.focus(); }
    }
    document.addEventListener('keydown', (e)=>{
  const open = Array.from(document.querySelectorAll('[role="dialog"]')).find(m => m.dataset.trap === '1');
  if (!open) return;
  if (e.key === 'Escape'){
    e.preventDefault();
    const btn = open.querySelector('button[id$="CancelBtn"], [data-cancel]');
    if (btn) btn.click(); else closeModal(open);
  } else if (e.key === 'Enter' && !e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey) {
    // Pressing Enter in any field will "OK" the modal.
    // Use Shift+Enter to insert a newline in textareas/editors.
    e.preventDefault();
    const defBtn = open.querySelector('[data-default="1"]') || open.querySelector('button[id$="OkBtn"]');
    if (defBtn) defBtn.click();
  } else if (e.key === 'Tab'){
    const f = focusablesIn(open);
    if (!f.length) return;
    const i = f.indexOf(document.activeElement);
    if (e.shiftKey && (i <= 0 || document.activeElement === open)) { e.preventDefault(); f[f.length-1].focus(); }
    else if (!e.shiftKey && (i === f.length-1)) { e.preventDefault(); f[0].focus(); }
  }
});/* ============================== *
     *  Helpers (calc/format/etc)     *
     * ============================== */
    function statusDotEl() { return el('statusDot'); }
    function statusTextEl() { return el('statusText'); }
    function formatTime(d = new Date()) {
      const hh = d.getHours().toString().padStart(2, '0');
      const mm = d.getMinutes().toString().padStart(2, '0');
      return `${hh}:${mm}`;
    }
    function refreshStatus() {
      const dot = statusDotEl(); const text = statusTextEl();
      if (!dot || !text) return;
      let base = currentFileName ? `File: ${currentFileName}${driveStatusLabel()}` : 'Unsaved';
      if (saving) { dot.className = 'pulse-dot bg-sky-500'; text.textContent = `${base} • Saving…`; return; }
      if (fileDirty) { dot.className = 'inline-block w-2 h-2 rounded-full bg-amber-500'; text.textContent = `${base} • Unsaved changes`; return; }
      dot.className = 'inline-block w-2 h-2 rounded-full bg-emerald-500'; text.textContent = `${base} • Saved ${formatTime()}`;
    }
    function hasLocalHandle() { return !!(currentFileHandle && currentFileHandle.createWritable); }
    function hasDriveTarget() { return !!currentDriveFileId; }
    function canSave() { return hasDriveTarget() || hasLocalHandle(); }
    function updateFileMenuState() {
      const saveBtn = el('menuSave');
      if (saveBtn) saveBtn.disabled = !canSave();

      const driveButtons = document.querySelectorAll('#fileMenuPanel [data-requires-drive]');
      const driveAvailable = driveConfigured();
      driveButtons.forEach((btn) => {
        if (!btn) return;
        if (!btn.dataset.defaultTitle) {
          const existingTitle = btn.getAttribute('title');
          btn.dataset.defaultTitle = existingTitle || '';
        }
        const needsDriveTarget = (btn.id === 'menuShareDrive');
        const disabled = !driveAvailable || (needsDriveTarget && !hasDriveTarget());
        btn.disabled = disabled;
        if (!driveAvailable) {
          btn.setAttribute('title', DRIVE_DISABLE_REASON);
        } else if (needsDriveTarget && !hasDriveTarget()) {
          btn.setAttribute('title', DRIVE_NEEDS_FILE_REASON);
        } else {
          const readyTitle = btn.dataset.driveReadyTitle;
          const defaultTitle = btn.dataset.defaultTitle;
          if (readyTitle) btn.setAttribute('title', readyTitle);
          else if (defaultTitle) btn.setAttribute('title', defaultTitle);
          else btn.removeAttribute('title');
        }
      });

      refreshStatus();
    }

    function getDatasetSnapshotString() {
      return Promise.all([SessionStore.getAll(), SessionStore.getAllHazards(), SessionStore.getAllObjectives()]).then(([items, hazards, objectives]) => {
        items.sort((a, b) => (a.id || 0) - (b.id || 0));
        hazards.sort((a, b) => (a.id || 0) - (b.id || 0));
        objectives.sort((a, b) => (a.id || 0) - (b.id || 0));
        return JSON.stringify({ items, hazards, objectives });
      });
    }
    async function updateSavedSnapshot() { lastSavedSnapshot = await getDatasetSnapshotString(); fileDirty = false; refreshStatus(); }
    async function recomputeDirtyAgainstSnapshot() { const snap = await getDatasetSnapshotString(); fileDirty = (snap !== lastSavedSnapshot); refreshStatus(); }

    function mitigationReduction(eff) {
      const e = clamp15(eff);
      if (CFG.mitigationSteps) {
        const v = CFG.mitigationSteps[e];
        return typeof v === 'number' ? v : 0;
      }
      return ((e - 1) / 4) * 0.4;
    }
    function riskClass(score) {
      const s = Number(score || 0);
      if (s > 25) return { label: 'Extremely High', cls: 'bg-rose-100 text-rose-800' };
      if (s >= 21) return { label: 'Very High', cls: 'bg-rose-100 text-rose-800' };
      if (s >= 16) return { label: 'High', cls: 'bg-orange-100 text-orange-800' };
      if (s >= 11) return { label: 'Moderate', cls: 'bg-amber-100 text-amber-800' };
      if (s >= 6)  return { label: 'Low', cls: 'bg-yellow-100 text-yellow-800' };
      return { label: 'Very Low', cls: 'bg-emerald-100 text-emerald-800' };
    }
    function hazardNameOf(id){
      if (!id) return '(Uncategorised)';
      const h = HAZARDS_CACHE.find(x => x.id === id);
      return h ? h.title : '(Uncategorised)';
    }

    /* ============================== *
     *  Store (in-memory session)     *
     * ============================== */
    const SessionStore = {
      _items: [], _hazards: [], _objectives: [],
      _nextItemId: 1, _nextHazId: 1, _nextObjectiveId: 1,

      // Events
      add(rec) { const id = this._nextItemId++; const copy = { ...rec, id }; this._items.unshift(copy); return Promise.resolve({ id }); },
      put(rec) { if (!rec.id) return this.add(rec); const i = this._items.findIndex(x => x.id === rec.id);
        const c = JSON.parse(JSON.stringify(rec)); if (i > -1) this._items[i] = c; else this._items.unshift(c); return Promise.resolve({ id: rec.id }); },
      getAll() { return Promise.resolve(this._items.map(x => JSON.parse(JSON.stringify(x)))); },
      get(id) { const it = this._items.find(x => x.id === id); return Promise.resolve(it ? JSON.parse(JSON.stringify(it)) : null); },
      delete(id) { this._items = this._items.filter(x => x.id !== id); return Promise.resolve(); },

      // Hazards
      addHazard(title) { const id = this._nextHazId++; this._hazards.push({ id, title: String(title || 'Untitled Hazard').trim() }); return Promise.resolve({ id }); },
      putHazard(haz) { const i = this._hazards.findIndex(h => h.id === haz.id); if (i > -1) this._hazards[i] = { id: haz.id, title: String(haz.title || 'Untitled Hazard').trim() }; return Promise.resolve({ id: haz.id }); },
      getAllHazards() { return Promise.resolve(this._hazards.map(h => ({ id: h.id, title: h.title }))); },
      deleteHazard(id) { this._hazards = this._hazards.filter(h => h.id !== id); this._items = this._items.map(ev => (ev.hazardId === id ? { ...ev, hazardId: null } : ev)); return Promise.resolve(); },

      // Objectives
      addObjective(a) { const id = this._nextObjectiveId++; const copy = { id, title: String(a.title||'Untitled Objective').trim(), description: a.description||'', status: a.status||'Planned', owner: a.owner||'', color: a.color||'' , createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() }; this._objectives.push(copy); return Promise.resolve({ id }); },
      putObjective(a) { const i = this._objectives.findIndex(x => x.id === a.id); const rec = { ...this._objectives[i], ...a, updatedAt: new Date().toISOString() }; if (i > -1) this._objectives[i] = rec; else this._objectives.push(rec); return Promise.resolve({ id: rec.id }); },
      deleteObjective(id) { this._objectives = this._objectives.filter(x => x.id !== id);
        // detach mitigations that referenced this action
        this._items = this._items.map(ev => ({ ...ev, planMitigations: (ev.planMitigations||[]).map(m => ({ ...m, objectiveId: m.objectiveId === id ? null : m.objectiveId })) }));
        return Promise.resolve();
      },
      getAllObjectives() { return Promise.resolve(this._objectives.map(x => ({ ...x }))); },

      // Bulk load/set
      setAll(payload) {
        const items = Array.isArray(payload?.items) ? payload.items : [];
        const hazards = Array.isArray(payload?.hazards) ? payload.hazards : [];
        const objectives = Array.isArray(payload?.objectives) ? payload.objectives : [];

        this._items = []; this._hazards = []; this._objectives = [];
        this._nextItemId = 1; this._nextHazId = 1; this._nextObjectiveId = 1;

        hazards.forEach(h => { this._hazards.push({ id: this._nextHazId++, title: String(h.title || 'Untitled Hazard') }); });
        objectives.forEach(a => { this._objectives.push({ id: this._nextObjectiveId++, title: String(a.title || 'Untitled Objective'), description: a.description||'', status: a.status||'Planned', owner: a.owner||'', color: a.color||'', createdAt: a.createdAt||new Date().toISOString(), updatedAt: a.updatedAt||new Date().toISOString() }); });
        items.forEach(it => { this._items.push({ ...it, id: this._nextItemId++ }); });
      }
    };

    /* ============================== *
     *  Sorting / filtering helpers   *
     * ============================== */
    const sortComparators = {
      title: (a, b) => (a.title || '').localeCompare(b.title || '', undefined, { sensitivity: 'base' }),
      status: (a, b) => (a.status || '').localeCompare(b.status || '', undefined, { sensitivity: 'base' }),
      score: (a, b) => Number((a.framework7 || {}).overall || 0) - Number((b.framework7 || {}).overall || 0),
      hazard: (a, b) => hazardNameOf(a.hazardId).localeCompare(hazardNameOf(b.hazardId), undefined, { sensitivity: 'base' })
    };
    function sortEvents(rows) { const cmp = sortComparators[EVENT_SORT.key] || sortComparators.score; const out = rows.slice().sort(cmp); if (EVENT_SORT.dir === 'desc') out.reverse(); return out; }
    function sortHazardsByTitle(hazards) { return hazards.slice().sort((a,b)=>(a.title||'').localeCompare(b.title||'',undefined,{sensitivity:'base'})); }
    function sortIcon(key) { if (EVENT_SORT.key !== key) return ''; return EVENT_SORT.dir === 'asc' ? ' ▲' : ' ▼'; }
    function handleHeaderClick(key) { if (EVENT_SORT.key === key) { EVENT_SORT.dir = EVENT_SORT.dir === 'asc' ? 'desc' : 'asc'; } else { EVENT_SORT.key = key; EVENT_SORT.dir = (key === 'score' ? 'desc' : 'asc'); } resortAllTablesInPlace(); updateSortIndicators(el('hazardsAccordion')); }
    function filterItemsBySearch(items) { const q = normalize(LIST_SEARCH).trim(); if (!q) return items; return items.filter(ev => normalize(ev.title).includes(q)); }

    /* ============================== *
     *  Guidance UI helpers           *
     * ============================== */
    function renderGuidanceHtml(id, lines) {
      if (!Array.isArray(lines) || !lines.length) return '';
      const n = lines.length;
      const items = lines.slice().reverse().map((txt, idx) => {
        const step = n - idx;
        return `<li class="flex gap-2 items-start" data-step="${step}"><span class="inline-block w-5 text-right text-sky-800 font-semibold">${step}</span><span class="flex-1">${escapeHtml(txt)}</span></li>`;
      }).join('');
      return `<div id="g_${id}" class="rounded-xl bg-sky-50 text-sky-800 px-3 text-sm overflow-hidden transition-all duration-200 ease-out gpanel-collapsed border border-sky-200"><div class="font-medium mb-1">Scoring guide</div><ul class="list-none pl-0 space-y-1">${items}</ul></div>`;
    }
    function showGuide(panelId) { const d = el(panelId); if (!d) return; d.classList.remove('gpanel-collapsed'); d.classList.add('gpanel-expanded'); }
    function hideGuide(panelId) { const d = el(panelId); if (!d) return; d.classList.remove('gpanel-expanded'); d.classList.add('gpanel-collapsed'); }
    function updateGuidanceHighlight(subId) {
      const panel = el('g_' + subId); if (!panel) return;
      const input = el(subId); if (!input) return;
      const val = clamp15(input.value);
      const lis = panel.querySelectorAll('li[data-step]');
      lis.forEach(li => {
        if (Number(li.getAttribute('data-step')) === val) { li.classList.add('bg-sky-100','rounded','px-2','py-0.5','font-medium'); }
        else { li.classList.remove('bg-sky-100','rounded','px-2','py-0.5','font-medium'); }
      });
    }

    /* ============================== *
     *  Rendering: Categories (Risk)  *
     * ============================== */
    const catRoot = el('categories');
    function renderSubRow(sf) {
      const guidance = GUIDANCE?.[sf.id];
      return `<div class="subrow">
        <div class="flex items-center justify-between gap-3">
          <label class="w-56" for="${sf.id}">${sf.label}</label>
          <div class="flex items-center gap-2">
            <input id="${sf.id}" type="number" min="1" max="5" value="3"
                   class="w-20 text-center border rounded-lg py-1" />
            <button id="rsn_${sf.id}" type="button"
                    class="px-2 py-1 rounded-lg border text-sm"
                    title="Add reasoning for ${sf.label}">Reason</button>
            <span id="rsn_${sf.id}_badge"
                  class="text-emerald-700 text-xs hidden"
                  aria-live="polite">Saved</span>
          </div>
        </div>
        ${guidance ? renderGuidanceHtml(sf.id, guidance) : ''}
      </div>`;
    }
    function buildCategoryBlocks() {
      const frag = document.createDocumentFragment();
      CFG.categories.forEach(cat => {
        const box = document.createElement('div');
        box.className = 'border rounded-2xl p-4';
        box.innerHTML = `
          <div class="flex items-center justify-between mb-2">
            <h2 class="text-xl font-semibold">${cat.title}</h2>
            <div class="flex items-center gap-2">
              <span class="text-sm text-slate-600">Include</span>
              <label class="switch">
                <input id="en_${cat.key}" type="checkbox" ${cat.enabled ? 'checked' : ''}>
                <span class="slider"></span>
              </label>
            </div>
          </div>
          <div id="sec_${cat.key}_body" class="space-y-3">
            ${cat.sub.map(sf => renderSubRow(sf)).join('')}
            <div class="flex items-center justify-between gap-3 pt-2 border-t">
              <span class="w-56 font-medium">Inherent ${cat.title} Score</span>
              <output id="roll_${cat.key}" class="w-24 text-center font-semibold" aria-live="polite">3.0</output>
            </div>

            <div class="mt-3 p-3 rounded-xl bg-emerald-50 border border-emerald-200">
              <div class="flex items-center justify-between mb-2">
                <span class="font-medium text-emerald-900">Existing mitigations</span>
                <div class="flex items-center gap-2">
                  <input id="mit_${cat.key}_input" type="text" placeholder="Add a mitigation…" class="px-2 py-1 text-sm border rounded-lg w-56" />
                  <button id="mit_${cat.key}_add" class="px-2 py-1 text-sm rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white">+ Add</button>
                </div>
              </div>
              <div id="mit_${cat.key}_list" class="flex flex-wrap gap-2"></div>
            </div>

            <div class="mt-3 p-3 rounded-xl bg-rose-50 border border-rose-200">
              <div class="flex items-center justify-between mb-2">
                <span class="font-medium text-rose-900">Identified gaps</span>
                <div class="flex items-center gap-2">
                  <input id="gap_${cat.key}__input" type="text" placeholder="Add a gap…" class="px-2 py-1 text-sm border rounded-lg w-56" />
                  <button id="gap_${cat.key}_add" class="px-2 py-1 text-sm rounded-lg bg-rose-600 hover:bg-rose-700 text-white">+ Add</button>
                </div>
              </div>
              <div id="gap_${cat.key}_list" class="flex flex-wrap gap-2"></div>
            </div>

            <div class="flex items-center justify-between gap-3 mt-2">
              <label class="w-56" for="eff_${cat.key}">Mitigation effectiveness (1-5)</label>
              <input id="eff_${cat.key}" type="number" min="1" max="5" value="1" class="w-20 text-center border rounded-lg py-1" />
            </div>

            <div class="flex items-center justify-between gap-3 pt-2 border-t">
              <span class="w-56 font-medium">Current ${cat.title} Score</span>
              <output id="roll_adj_${cat.key}" class="w-24 text-center font-semibold" aria-live="polite">3.0</output>
            </div>
          </div>`;
        frag.appendChild(box);
      });

      const like = document.createElement('div');
      like.className = 'border rounded-2xl p-4';
      like.innerHTML = `
        <h2 class="text-xl font-semibold mb-2">Likelihood</h2>
        <div class="flex items-center justify-between gap-3">
          <label class="w-56" for="${CFG.likelihoodId}">Likelihood of occurrence</label>
          <input id="${CFG.likelihoodId}" type="number" min="1" max="5" value="3" class="w-20 text-center border rounded-lg py-1" />
        </div>
        <div class="flex items-center justify-between gap-3 pt-2 border-t">
          <span class="w-56 font-medium">Overall (L × Max current)</span>
          <output id="overall_out" class="w-24 text-center font-semibold" aria-live="polite">0.0</output>
        </div>`;
      frag.appendChild(like);

      catRoot.appendChild(frag);
      CFG.categories.forEach(c => toggleSectionBody(c.key, enabledOf(c.key)));
    }

    /* ============================== *
     *  Summary pills & radar         *
     * ============================== */
    const pillsRoot = el('summaryPills');
    function buildPillsDynamic(rollsAdj, perCatRisk, overall) {
      const enabledCats = CFG.categories.filter(c => enabledOf(c.key));
      let html = enabledCats.map(cat => {
        const val = Number.isFinite(perCatRisk?.[cat.key]) ? Number(perCatRisk[cat.key]) : 0;
        const rc = riskClass(val);
        return `<div class="p-3 rounded-xl bg-slate-50">
          <div class="text-slate-500">${cat.title}</div>
          <div id="pill_${cat.key}" class="text-2xl font-bold">${val.toFixed(1)}</div>
          <div class="text-xs ${rc.cls} inline-block mt-1 px-2 py-0.5 rounded-full">${rc.label}</div>
        </div>`;
      }).join('');

      const ov = Number.isFinite(overall) ? overall : 0;
      const rcOv = riskClass(ov);
      html += `<div class="p-3 rounded-xl bg-slate-50 col-span-2">
        <div class="text-slate-500">Overall (Max impact × L)</div>
        <div id="pill_overall" class="text-2xl font-bold">${ov.toFixed(1)}</div>
        <div class="text-xs ${rcOv.cls} inline-block mt-1 px-2 py-0.5 rounded-full">${rcOv.label}</div>
      </div>`;

      pillsRoot.innerHTML = html;

      const rows = enabledCats.map(c => `<tr><th scope="row">${escapeHtml(c.title)}</th><td>${(perCatRisk[c.key] ?? 0).toFixed(1)}</td></tr>`).join('');
      el('radarA11y').innerHTML = `<h4>Risk profile data</h4><table><thead><tr><th>Category</th><th>Risk (adj × L)</th></tr></thead><tbody>${rows}</tbody></table><p>Overall: ${ov.toFixed(1)}</p>`;
    }

    function drawRadar(rolls) {
      const canvas = document.getElementById('radarChart');
      const ctx = canvas.getContext('2d');
      const dpr = window.devicePixelRatio || 1;
      if (!canvas.style.width) canvas.style.width = '100%';
      if (!canvas.style.height) canvas.style.height = '300px';
      const rect = canvas.getBoundingClientRect();
      const cssW = Math.max(1, rect.width || 300);
      const cssH = Math.max(1, rect.height || 300);
      canvas.width = Math.round(cssW * dpr);
      canvas.height = Math.round(cssH * dpr);
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.clearRect(0, 0, cssW, cssH);

      const enabledCats = CFG.categories.filter(c => enabledOf(c.key));
      const labels = enabledCats.map(c => c.title);
      const values = enabledCats.map(c => rolls[c.key] ?? 0);
      if (!labels.length) { return; }
      const max = 6;

      const pad = 24;
      const cx = cssW / 2;
      const cy = cssH / 2;
      const radius = Math.min(cssW, cssH) / 2 - pad - 10;
      const steps = max;

      ctx.lineWidth = 1;
      ctx.strokeStyle = '#cbd5e1';
      for (let s = 1; s <= steps; s++) {
        const r = radius * (s / max);
        ctx.beginPath();
        for (let i = 0; i < labels.length; i++) {
          const ang = (Math.PI * 2 * i / labels.length) - Math.PI / 2;
          const x = cx + r * Math.cos(ang);
          const y = cy + r * Math.sin(ang);
          if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
        }
        ctx.closePath(); ctx.stroke();
      }

      ctx.strokeStyle = '#e2e8f0';
      for (let i = 0; i < labels.length; i++) {
        const ang = (Math.PI * 2 * i / labels.length) - Math.PI / 2;
        const x = cx + radius * Math.cos(ang);
        const y = cy + radius * Math.sin(ang);
        ctx.beginPath(); ctx.moveTo(cx, cy); ctx.lineTo(x, y); ctx.stroke();
      }

      ctx.fillStyle = '#475569';
      ctx.font = '12px system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial';
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      for (let i = 0; i < labels.length; i++) {
        const ang = (Math.PI * 2 * i / labels.length) - Math.PI / 2;
        const lx = cx + (radius + 12) * Math.cos(ang);
        const ly = cy + (radius + 12) * Math.sin(ang);
        const text = labels[i].length > 18 ? labels[i].slice(0,15)+'…' : labels[i];
        ctx.fillText(text, lx, ly);
      }

      const points = values.map((v, i) => {
        const ang = (Math.PI * 2 * i / labels.length) - Math.PI / 2;
        const r = radius * (v / max);
        return [cx + r * Math.cos(ang), cy + r * Math.sin(ang)];
      });

      ctx.beginPath();
      points.forEach(([x, y], i) => { if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y); });
      ctx.closePath();
      ctx.fillStyle = 'rgba(99,102,241,0.15)';
      ctx.strokeStyle = 'rgba(99,102,241,0.9)';
      ctx.lineWidth = 2;
      ctx.fill(); ctx.stroke();
      ctx.fillStyle = 'rgba(99,102,241,0.9)';
      points.forEach(([x, y]) => { ctx.beginPath(); ctx.arc(x, y, 3, 0, Math.PI * 2); ctx.fill(); });
    }

    /* ============================== *
     *  Calculations (max 6)          *
     * ============================== */
    function calcRoll(cat) {
      const vals = cat.sub.map(sf => clamp15(el(sf.id).value)).filter(v => Number.isFinite(v));
      if (!vals.length) return 0;
      const sorted = vals.slice().sort((a, b) => b - a);
      const top = sorted[0];
      const others = sorted.slice(1);
      const avgOthers = others.length ? (others.reduce((a,b)=>a+b,0) / others.length) : 0;
      const score = top + (avgOthers * 0.2);
      return round1(Math.min(6, score));
    }
    function calcAdjusted(cat) {
      const base = calcRoll(cat);
      const eff = clamp15(el('eff_' + cat.key).value);
      const reduction = mitigationReduction(eff);
      const adj = Math.max(0, base * (1 - reduction));
      return { base, eff, adj: round1(adj) };
    }
    function recalc() {
      const rollsAdj = {};
      CFG.categories.forEach(cat => {
        const { base, adj } = calcAdjusted(cat);
        el('roll_' + cat.key).textContent = base.toFixed(1);
        el('roll_adj_' + cat.key).textContent = adj.toFixed(1);
        if (enabledOf(cat.key)) rollsAdj[cat.key] = adj;
      });

      const like = clamp15(el(CFG.likelihoodId).value);
      const perCatRisk = {};
      Object.keys(rollsAdj).forEach(k => { perCatRisk[k] = round1(rollsAdj[k] * like); });
      const vals = Object.values(rollsAdj);
      const maxAdj = vals.length ? Math.max(...vals) : 0;
      const overall = round1(maxAdj * like);

      el('overall_out').textContent = overall.toFixed(1);
      buildPillsDynamic(rollsAdj, perCatRisk, overall);
      drawRadar(rollsAdj);
    }

    /* ============================== *
     *  Links helpers (safe URLs)     *
     * ============================== */
    function setLinks(arr) {
      const root = el('linksList');
      root.innerHTML = '';
      (arr || []).forEach((lk, i) => root.appendChild(renderLinkRow(lk, i)));
    }
    function getLinks() {
      const root = el('linksList');
      return Array.from(root.querySelectorAll('[data-index]')).map(row => {
        const title = row.querySelector('.font-medium')?.textContent ?? '';
        const desc = row.querySelector('.text-slate-600')?.textContent ?? '';
        const url = row.querySelector('a')?.getAttribute('href') ?? '';
        return { title, description: desc, url };
      });
    }
    function renderLinkRow(link, idx) {
      const row = document.createElement('div');
      row.className = 'border rounded-xl p-3 bg-white flex items-start justify-between gap-3';
      row.setAttribute('data-index', String(idx));
      const href = safeUrl(link.url || '');
      row.innerHTML = `
        <div class="min-w-0">
          <div class="font-medium truncate">${escapeHtml(link.title || '(untitled)')}</div>
          <div class="text-sm text-slate-600 line-clamp-2">${escapeHtml(link.description || '')}</div>
          <a href="${href}" target="_blank" rel="noopener noreferrer" class="text-sm text-violet-700 underline break-all">${escapeHtml(link.url || '')}</a>
        </div>
        <div class="flex gap-2 shrink-0">
          <button class="px-2 py-1 rounded-lg border text-sm" data-role="edit">Edit</button>
          <button class="px-2 py-1 rounded-lg border border-rose-300 text-rose-700 text-sm" data-role="remove">Remove</button>
        </div>`;
      row.querySelector('[data-role="edit"]').addEventListener('click', (e) => { e.stopPropagation(); openLinkModal('Edit Link', idx); markDirtySoon(); });
      row.querySelector('[data-role="remove"]').addEventListener('click', (e) => {
        e.stopPropagation();
        if (!confirm('Remove this link?')) return;
        const arr = getLinks();
        arr.splice(idx, 1);
        setLinks(arr);
        markDirtySoon();
      });
      return row;
    }
    function openLinkModal(title, idx = null) {
      el('linkModalTitle').textContent = title || 'Add Link';
      const isEdit = (idx !== null && idx !== undefined);
      const arr = getLinks();
      const existing = isEdit ? arr[idx] : { title: '', description: '', url: '' };
      el('linkTitleInput').value = existing.title || '';
      el('linkDescInput').value = existing.description || '';
      el('linkUrlInput').value = existing.url || '';
      openModal(el('linkModal'), '#linkTitleInput');

      el('linkOkBtn').onclick = () => {
        const t = el('linkTitleInput').value.trim();
        const d = el('linkDescInput').value.trim();
        const u = el('linkUrlInput').value.trim();
        const link = { title: t, description: d, url: u };
        if (isEdit) { const a = getLinks(); a[idx] = link; setLinks(a); }
        else { const a = getLinks(); a.push(link); setLinks(a); }
        closeModal(el('linkModal')); markDirtySoon();
      };
      el('linkCancelBtn').onclick = () => closeModal(el('linkModal'));
    }

    /* ============================== *
     *  Mitigations & Gaps chips      *
     * ============================== */
    function renderTagChip({ type, key, title, details }) {
      const wrap = document.createElement('span');
      const baseBg = (type === 'mit' ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800');
      const baseHover = (type === 'mit' ? 'hover:bg-emerald-200' : 'hover:bg-rose-200');
      wrap.className = `chip inline-flex items-center gap-1 px-2 py-1 rounded-full ${baseBg} text-xs`;
      wrap.setAttribute('data-type', type);
      wrap.setAttribute('data-key', key);
      wrap.setAttribute('data-title', title || '');
      wrap.setAttribute('data-details', details || '');
      wrap.innerHTML = `
        <button type="button" class="rounded px-1 ${baseHover}" title="Edit" aria-label="Edit ${type === 'mit' ? 'mitigation' : 'gap'}" data-role="edit">✎</button>
        <span title="${escapeHtml(title || '')}" data-role="text">${escapeHtml(title || '')}</span>
        <button type="button" class="rounded px-1 ${baseHover}" title="Remove" aria-label="Remove ${type === 'mit' ? 'mitigation' : 'gap'}" data-role="remove">×</button>`;
      wrap.querySelector('[data-role="edit"]').addEventListener('click', (e) => { e.stopPropagation(); openEditModal(wrap); });
      wrap.querySelector('[data-role="remove"]').addEventListener('click', (e) => {
        e.stopPropagation();
        if (confirm('Are you sure you want to remove this item?')) { wrap.remove(); recalc(); renderPlanGaps(); markDirty(); }
      });
      return wrap;
    }
    function addTag(type, key, title, details) {
      const list = (type === 'mit' ? el(`mit_${key}_list`) : el(`gap_${key}_list`));
      if (!list) return;
      const chip = renderTagChip({ type, key, title: String(title || '').trim(), details: String(details || '').trim() });
      list.appendChild(chip);
      recalc(); renderPlanGaps(); markDirtySoon();
    }
    function setTags(type, key, arr) {
      const list = (type === 'mit' ? el(`mit_${key}_list`) : el(`gap_${key}_list`));
      if (!list) return;
      list.innerHTML = '';
      (arr || []).forEach(item => {
        if (item == null) return;
        if (typeof item === 'string') addTag(type, key, item, '');
        else addTag(type, key, item.title || '', item.details || '');
      });
    }
    function collectTags(type, key) {
      const list = (type === 'mit' ? el(`mit_${key}_list`) : el(`gap_${key}_list`));
      if (!list) return [];
      return Array.from(list.querySelectorAll('span[data-type]')).map(s => ({
        title: s.getAttribute('data-title') || '',
        details: s.getAttribute('data-details') || ''
      }));
    }
    const getMitigations = (key) => collectTags('mit', key);
    const getGaps = (key) => collectTags('gap', key);
    const setMitigations = (key, arr) => setTags('mit', key, arr);
    const setGaps = (key, arr) => setTags('gap', key, arr);

    function openEditModal(chipEl) {
      const isMit = chipEl.getAttribute('data-type') === 'mit';
      el('editModalTitle').textContent = isMit ? 'Edit Mitigation' : 'Edit Gap';
      el('editTitleInput').value = chipEl.getAttribute('data-title') || '';
      el('editDetailsInput').value = chipEl.getAttribute('data-details') || '';
      openModal(el('editModal'), '#editTitleInput');
      el('editOkBtn').onclick = () => {
        const t = el('editTitleInput').value.trim();
        const d = el('editDetailsInput').value.trim();
        chipEl.setAttribute('data-title', t);
        chipEl.setAttribute('data-details', d);
        const textNode = chipEl.querySelector('[data-role="text"]');
        textNode.textContent = t || '(untitled)';
        textNode.title = t || '(untitled)';
        closeModal(el('editModal'));
        recalc(); renderPlanGaps(); markDirtySoon();
      };
      el('editCancelBtn').onclick = () => closeModal(el('editModal'));
    }

    /* ============================== *
     *  Narrative generation          *
     * ============================== */
    function genNarrative() {
      const title = (el('riskTitle').value || 'This hazardous event').trim();
      const like = clamp15(el(CFG.likelihoodId).value);
      const enabledCats = CFG.categories.filter(c => enabledOf(c.key));
      const adjs = {}; enabledCats.forEach(cat => { adjs[cat.key] = calcAdjusted(cat).adj; });
      const maxCat = enabledCats.length ? enabledCats.reduce((best, cur) => (adjs[cur.key] > adjs[best.key] ? cur : best), enabledCats[0]) : null;
      const strong = enabledCats.filter(c => adjs[c.key] >= 4.8).map(c => c.title);

      const parts = [];
      parts.push(`${title} has a likelihood of ${like}/5${maxCat ? ` with strongest current impacts in ${maxCat.title}.` : '.'}`);
      if (strong.length) parts.push('High-impact categories (after mitigation) include: ' + strong.join(', ') + '.');
      if (maxCat) {
        const topMits = getMitigations(maxCat.key).map(m => m.title).filter(Boolean);
        if (topMits.length) { parts.push(`Existing mitigations for ${maxCat.title}: ${topMits.slice(0, 3).join('; ')}${topMits.length > 3 ? ' …' : ''}.`); }
        const topGaps = getGaps(maxCat.key).map(g => g.title).filter(Boolean);
        if (topGaps.length) { parts.push(`Key gaps for ${maxCat.title}: ${topGaps.slice(0, 3).join('; ')}${topGaps.length > 3 ? ' …' : ''}.`); }
      }
      parts.push('Priorities should focus on prevention, preparedness, and mitigation aligned to the highest residual risks.');
      el('narrative').textContent = parts.join(' ');
    }

    /* ============================== *
     *  Import via URL hash (guard)   *
     * ============================== */
    function toBase64Safe(str) { try { return btoa(unescape(encodeURIComponent(str))); } catch (e) { return btoa(str); } }
    function fromBase64Safe(b64) { try { return decodeURIComponent(escape(atob(b64))); } catch (e) { return atob(b64); } }
    function importFromHash() {
      if (!location.hash || !location.hash.startsWith('#import=')) return;
      const enc = location.hash.slice('#import='.length);
      try {
        const json = fromBase64Safe(enc);
        const data = JSON.parse(json);
        const items = Array.isArray(data?.items) ? data.items : (Array.isArray(data) ? data : []);
        const hazards = Array.isArray(data?.hazards) ? data.hazards : [];
        const objectives = Array.isArray(data?.objectives) ? data.objectives : [];
        if (items.length > 10000 || hazards.length > 5000) throw new Error('Too large');
        currentDriveFileId = null;
        currentFileHandle = null;
        currentFileName = null;
        updateDriveUrlParam(null);
        SessionStore.setAll({ items, hazards, objectives });
        refreshHazardsAccordion();
        recalc();
        updateSavedSnapshot();
        updateFileMenuState();
        alert('Imported from link');
      } catch (e) {
        console.error('Hash import failed', e);
        alert('Failed to import from link');
      } finally {
        history.replaceState(null, document.title, location.pathname + location.search);
      }
    }

    async function importDriveFromQuery() {
      const params = new URLSearchParams(location.search);
      const raw = params.get('driveFile');
      if (!raw) return;
      try {
        await openDriveFile(raw, { interactive: false });
      } catch (err) {
        console.warn('Non-interactive Drive load failed', err);
        if (driveConfigured()) {
          try {
            await openDriveFile(raw, { interactive: true });
            return;
          } catch (err2) {
            console.error('Unable to load Google Drive file from URL', err2);
            alert('Failed to load Google Drive file from link: ' + (err2?.message || err2));
          }
        } else {
          alert('Failed to load Google Drive file from link: ' + (err?.message || err));
        }
      }
    }

    /* ============================== *
     *  Collect / Save current event  *
     * ============================== */
    function collectCurrent(existing) {
      const detailsHtml = el('descEditor')?.innerHTML || '';

      const inputs = {};
      CFG.categories.forEach(c =>
        c.sub.forEach(sf => inputs[sf.id] = clamp15(el(sf.id)?.value))
      );
      inputs[CFG.likelihoodId] = clamp15(el(CFG.likelihoodId)?.value);

      const rolls = {};
      const adjs = {};
      const effs = {};
      const enabled = {};
      const risks = {};

      CFG.categories.forEach(c => {
        const r = calcAdjusted(c);
        rolls[c.key]   = r.base;
        adjs[c.key]    = r.adj;
        effs[c.key]    = r.eff;
        enabled[c.key] = enabledOf(c.key);
      });

      const like = clamp15(inputs[CFG.likelihoodId]);
      Object.keys(adjs).forEach(k => { if (enabled[k]) risks[k] = round1(adjs[k] * like); });
      const currentVals = Object.entries(adjs)
        .filter(([k]) => enabled[k])
        .map(([, v]) => v);
      const overall = round1((currentVals.length ? Math.max(...currentVals) : 0) * like);

      const mitigations = {};
      CFG.categories.forEach(c => { mitigations[c.key] = getMitigations(c.key); });

      const gaps = {};
      CFG.categories.forEach(c => { gaps[c.key] = getGaps(c.key); });

      const links = getLinks();
      const planMitigationsCopy = JSON.parse(JSON.stringify(planMitigations || []));

      const hazardRaw = parseInt(el('hazardSel')?.value, 10);
      const hazardId = Number.isFinite(hazardRaw) ? hazardRaw : null;
      const title = (el('riskTitle')?.value || 'Untitled Hazardous Event').trim();
      const status = el('statusSel')?.value || 'Tolerate';
      const narrative = (el('narrative')?.textContent || '').trim();

      const baseTimestamps = existing
        ? {
            createdAt: existing.createdAt || new Date().toISOString(),
            updatedAt: new Date().toISOString()
          }
        : { createdAt: new Date().toISOString() };

			const reasoning = { ...SUB_REASONING };

      return {
        ...(editingId != null ? { id: editingId } : {}),
        hazardId,
        title,
        status,
        detailsHtml,
        links,
        planMitigations: planMitigationsCopy,
        framework7: {
          inputs,
          rolls,
          adjs,
          risks,
          effs,
          like,
          overall,
          mitigations,
          gaps,
          enabled,
          mitigationSteps: CFG.mitigationSteps
          ,reasoning 
        },
        narrative,
        ...baseTimestamps
      };
    }

    function saveCurrent() {
      const rec = collectCurrent();
      const op = (editingId != null) ? SessionStore.put(rec) : SessionStore.add(rec);
      op.then(async ({ id }) => {
        editingId = id || editingId;
        refreshHazardsAccordion();
        switchToList();
        await recomputeDirtyAgainstSnapshot();
      }).catch(err => { console.error('Save failed', err); alert('Save failed: ' + (err?.message || err)); });
    }
    function loadEvent(id) {
      SessionStore.get(id).then(r => {
        if (!r) return;
        editingId = r.id;
        populateHazardSelect(r.hazardId ?? null);
        el('riskTitle').value = r.title || '';
        el('statusSel').value = r.status || 'Tolerate';
        el('descEditor').innerHTML = r.detailsHtml || '';
        renderAllCharts(el('descEditor'));
        ensureTrailingParagraph();
        setLinks(Array.isArray(r.links) ? r.links : []);

        const f = r.framework7 || {}; const inputs = f.inputs || {}; const enabled = f.enabled || {};
        CFG.categories.forEach(c => c.sub.forEach(sf => { if (el(sf.id)) el(sf.id).value = inputs[sf.id] ?? 3; }));
        if (el(CFG.likelihoodId)) el(CFG.likelihoodId).value = inputs[CFG.likelihoodId] ?? 3;

        const effs = f.effs || {}; CFG.categories.forEach(c => { if (el('eff_' + c.key)) el('eff_' + c.key).value = effs[c.key] ?? 1; });
        const mits = (f.mitigations) || {}; CFG.categories.forEach(c => setMitigations(c.key, mits[c.key] || []));
        const gaps = (f.gaps) || {}; CFG.categories.forEach(c => setGaps(c.key, gaps[c.key] || []));
        CFG.categories.forEach(c => { const chk = el('en_' + c.key); if (chk) { chk.checked = enabled[c.key] === true; toggleSectionBody(c.key, chk.checked); } });

        planMitigations = Array.isArray(r.planMitigations) ? r.planMitigations : (Array.isArray(r.planObjectives) ? r.planObjectives : []);
        renderMitigationsList();
        renderPlanGaps();

        SUB_REASONING = { ...(f.reasoning || {}) };
        updateAllReasonBadges();

        CFG.categories.forEach(c => { c.sub.forEach(sf => { if (GUIDANCE[sf.id]) updateGuidanceHighlight(sf.id); }); });
        recalc();
        el('narrative').textContent = r.narrative || el('narrative').textContent;
        activateTab('details');
        switchToEditor();
      });
    }

    /* ============================== *
     *  Plan: gaps & mitigations      *
     * ============================== */
    let planMitigations = [];
    let editingMitigationIndex = null;

    function gapKey(catKey, title, details) {
      return `${catKey}|${(title || '').trim().toLowerCase()}|${(details || '').trim().toLowerCase()}`;
    }
    function getAllGapsFromUI() {
      const out = [];
      CFG.categories.forEach(c => {
        getGaps(c.key).forEach(g => {
          const key = gapKey(c.key, g.title, g.details);
          out.push({ key, catKey: c.key, title: g.title || '', details: g.details || '' });
        });
      });
      return out;
    }
    function getAddressedKeys(tempSelectedKeys = new Set()) {
      const keys = new Set();
      planMitigations.forEach(a => (a.gaps || []).forEach(g => keys.add(g.key)));
      tempSelectedKeys.forEach(k => keys.add(k));
      return keys;
    }

    function renderPlanGaps(tempSelectedKeys) {
      const addressedKeys = getAddressedKeys(tempSelectedKeys || new Set());
      const byGroup = new Map();
      CFG.categories.forEach(cat => {
        const gaps = getGaps(cat.key);
        if (!gaps.length) return;
        const bucket = { cat, addressed: [], unaddressed: [] };
        gaps.forEach(g => {
          const key = gapKey(cat.key, g.title, g.details);
          const chip = `<span class="pill ${addressedKeys.has(key) ? 'pill-amber' : 'pill-red'}" title="${escapeHtml(g.details || '')}">${escapeHtml(g.title || '(gap)')}</span>`;
          if (addressedKeys.has(key)) bucket.addressed.push(chip);
          else bucket.unaddressed.push(chip);
        });
        byGroup.set(cat.key, bucket);
      });

      const root = el('plan_gaps_by_group');
      if (!byGroup.size) {
        root.innerHTML = `<div class="text-sm text-slate-500">No gaps defined yet. Add gaps in the Analysis tab.</div>`;
        return;
      }

      const rows = [];
      byGroup.forEach(({ cat, addressed, unaddressed }) => {
        rows.push(`<div class="grid grid-cols-1 md:grid-cols-3 gap-3 p-3 rounded-xl bg-white border">
          <div class="md:col-span-1">
            <div class="text-sm text-slate-500 mb-1">Risk Group</div>
            <div class="font-medium">${escapeHtml(cat.title)}</div>
          </div>
          <div>
            <div class="text-sm text-slate-600 mb-1">Addressed by Mitigations</div>
            <div class="flex flex-wrap gap-2">${addressed.join('') || '<span class="text-slate-400 text-sm">None</span>'}</div>
          </div>
          <div>
            <div class="text-sm text-slate-600 mb-1">Not Addressed Yet</div>
            <div class="flex flex-wrap gap-2">${unaddressed.join('') || '<span class="text-slate-400 text-sm">None</span>'}</div>
          </div>
        </div>`);
      });

      root.innerHTML = rows.join('');
    }

    function renderMitigationsList() {
      const root = el('mitigationsList');
      if (!planMitigations.length) {
        root.innerHTML = `<div class="text-sm text-slate-500">No mitigations yet. Click <strong>+ Add Mitigation</strong> to create one.</div>`;
        return;
      }
      root.innerHTML = planMitigations.map((a, idx) => {
        const dateStr = [a.start || '', a.end || ''].filter(Boolean).join(' → ');
        const gapBadges = (a.gaps || []).map(g => `<span class="pill pill-amber">${escapeHtml(g.title || '(gap)')}</span>`).join(' ');
        return `<div class="border rounded-xl p-3 bg-white flex flex-col gap-2" data-mitigation-index="${idx}">
          <div class="flex items-center justify-between">
            <div class="font-medium">${escapeHtml(a.title || '(untitled mitigation)')}</div>
            <div class="text-xs text-slate-500">${escapeHtml(dateStr || '')}</div>
          </div>
          <div class="text-sm text-slate-700 whitespace-pre-line">${escapeHtml(a.outline || '')}</div>
          <div class="flex flex-wrap gap-2">${gapBadges}</div>
          <div class="flex gap-2 justify-end">
            <button class="px-2 py-1 rounded-lg border text-sm" data-mit="edit" data-index="${idx}">Edit</button>
            <button class="px-2 py-1 rounded-lg border border-rose-300 text-rose-700 text-sm" data-mit="del" data-index="${idx}">Delete</button>
          </div>
        </div>`;
      }).join('');
    }

    function openMitigationModal(index = null) {
      editingMitigationIndex = (index != null ? index : null);
      const allGaps = getAllGapsFromUI();

      if (index == null) {
        el('mitigationModalTitle').textContent = 'Add Mitigation';
        el('mitigationTitleInput').value = '';
        el('mitigationOutlineInput').value = '';
        el('mitigationStartInput').value = '';
        el('mitigationEndInput').value = '';
      } else {
        const a = planMitigations[index];
        el('mitigationModalTitle').textContent = 'Edit Mitigation';
        el('mitigationTitleInput').value = a.title || '';
        el('mitigationOutlineInput').value = a.outline || '';
        el('mitigationStartInput').value = a.start || '';
        el('mitigationEndInput').value = a.end || '';
      }

      const selectedKeys = new Set(index == null ? [] : (planMitigations[index].gaps || []).map(g => g.key));
      const picker = el('mitigationGapPicker');
      picker.innerHTML = allGaps.map(g => {
        const pressed = selectedKeys.has(g.key) ? 'true' : 'false';
        return `<button type="button" class="gap-pill pill pill-muted" aria-pressed="${pressed}" data-key="${g.key}" title="${escapeHtml(g.details || '')}">${escapeHtml(g.title || '(gap)')}</button>`;
      }).join('') || '<div class="text-sm text-slate-500">No gaps defined in the Analysis tab.</div>';

      function currentTempSet() {
        return new Set(Array.from(picker.querySelectorAll('.gap-pill[aria-pressed="true"]')).map(b => b.getAttribute('data-key')));
      }
      function refreshPickerColours() {
        const temp = currentTempSet();
        const addressed = getAddressedKeys(temp);
        picker.querySelectorAll('.gap-pill').forEach(btn => {
          const k = btn.getAttribute('data-key');
          const on = btn.getAttribute('aria-pressed') === 'true';
          btn.classList.remove('pill-red', 'pill-amber', 'pill-muted');
          if (!on) { btn.classList.add('pill-muted'); }
          else { if (addressed.has(k)) btn.classList.add('pill-amber'); else btn.classList.add('pill-red'); }
        });
        renderPlanGaps(temp);
      }

      picker.querySelectorAll('.gap-pill').forEach(btn => {
        btn.addEventListener('click', () => {
          const now = btn.getAttribute('aria-pressed') === 'true';
          btn.setAttribute('aria-pressed', now ? 'false' : 'true');
          refreshPickerColours();
        });
      });

      refreshPickerColours();
      openModal(el('mitigationModal'), '#mitigationTitleInput');

      el('mitigationOkBtn').onclick = () => {
        const title = el('mitigationTitleInput').value.trim();
        const outline = el('mitigationOutlineInput').value.trim();
        const start = el('mitigationStartInput').value;
        const end = el('mitigationEndInput').value;

        const temp = currentTempSet();
        const selected = Array.from(temp).map(key => {
          const g = allGaps.find(x => x.key === key);
          return { key, catKey: g?.catKey || '', title: g?.title || '', details: g?.details || '' };
        });

        const mitigationObj = { id: (editingMitigationIndex != null ? planMitigations[editingMitigationIndex].id : Date.now()), title, outline, start, end, gaps: selected, objectiveId: (editingMitigationIndex != null ? planMitigations[editingMitigationIndex].objectiveId || null : null) };

        if (editingMitigationIndex == null) planMitigations.push(mitigationObj);
        else planMitigations[editingMitigationIndex] = mitigationObj;

        renderMitigationsList();
        renderPlanGaps();
        closeModal(el('mitigationModal'));
        markDirtySoon();
      };
      el('mitigationCancelBtn').onclick = () => closeModal(el('mitigationModal'));
    }

    /* ============================== *
     *  Hazards list (accordion/all)  *
     * ============================== */
    function buildEventRowsHTML(rows) {
      return rows.map(r => {
        const f = r.framework7 || {};
        const overall = Number(f.overall ?? 0);
        const rc = riskClass(overall);
        const status = r.status || 'Tolerate';
        const badge = `<span class="inline-block text-xs ${rc.cls} px-2 py-0.5 rounded-full">${rc.label}</span>`;
        return `<tr class="border-b hover:bg-slate-50" data-id="${r.id}">
          <td class="py-2 pr-4">${escapeHtml(r.title || '')}</td>
          <td class="py-2 pr-4"><div class="flex items-center gap-2"><span class="font-semibold">${overall.toFixed(1)}</span>${badge}</div></td>
          <td class="py-2 pr-4">${escapeHtml(status)}</td>
          <td class="py-2 pr-4">
            <button data-action="edit" data-id="${r.id}" class="px-2 py-1 rounded-lg border mr-1">Edit</button>
            <button data-action="dup"  data-id="${r.id}" class="px-2 py-1 rounded-lg border mr-1">Duplicate</button>
            <button data-action="del"  data-id="${r.id}" class="px-2 py-1 rounded-lg border border-rose-300 text-rose-700">Delete</button>
          </td>
        </tr>`;
      }).join('');
    }
    function buildEventRowsHTML_All(rows) {
      return rows.map(r => {
        const f = r.framework7 || {};
        const overall = Number(f.overall ?? 0);
        const rc = riskClass(overall);
        const status = r.status || 'Tolerate';
        const badge = `<span class="inline-block text-xs ${rc.cls} px-2 py-0.5 rounded-full">${rc.label}</span>`;
        return `<tr class="border-b hover:bg-slate-50" data-id="${r.id}">
          <td class="py-2 pr-4">${escapeHtml(r.title || '')}</td>
          <td class="py-2 pr-4">${escapeHtml(hazardNameOf(r.hazardId))}</td>
          <td class="py-2 pr-4"><div class="flex items-center gap-2"><span class="font-semibold">${overall.toFixed(1)}</span>${badge}</div></td>
          <td class="py-2 pr-4">${escapeHtml(status)}</td>
          <td class="py-2 pr-4">
            <button data-action="edit" data-id="${r.id}" class="px-2 py-1 rounded-lg border mr-1">Edit</button>
            <button data-action="dup"  data-id="${r.id}" class="px-2 py-1 rounded-lg border mr-1">Duplicate</button>
            <button data-action="del"  data-id="${r.id}" class="px-2 py-1 rounded-lg border border-rose-300 text-rose-700">Delete</button>
          </td>
        </tr>`;
      }).join('');
    }
    function hazardBlock(h, rows, isUncat = false) {
      const count = rows.length;
      const subtitle = `${isUncat ? 'Uncategorised' : 'Hazard'} • ${count} event${count === 1 ? '' : 's'}`;
      const hazIdAttr = (h.id === null ? 'null' : String(h.id));
      const headerObjectives = isUncat ? '' : `<div class="flex gap-2">
        <button data-action="haz-rename" data-id="${hazIdAttr}" class="px-2 py-1 rounded-lg border">Rename</button>
        <button data-action="haz-delete" data-id="${hazIdAttr}" class="px-2 py-1 rounded-lg border border-rose-300 text-rose-700">Delete</button></div>`;
      const rowsSorted = sortEvents(rows || []);
      const tableRows = buildEventRowsHTML(rowsSorted);
      return `<details class="rounded-xl border">
        <summary class="cursor-pointer select-none flex items-center justify-between p-4">
          <div class="flex items-center gap-3">
            <svg class="chev transition-transform" width="16" height="16" viewBox="0 0 20 20"><path fill="currentColor" d="M7 5l6 5-6 5V5z"/></svg>
            <div><div class="font-semibold">${escapeHtml(h.title || 'Untitled Hazard')}</div>
            <div class="text-sm text-slate-500">${subtitle}</div></div>
          </div>${headerObjectives}
        </summary>
        <div class="p-4 pt-0">
          <div class="overflow-x-auto mt-3">
            <table class="min-w-full text-sm" data-hazard="${hazIdAttr}">
              <thead>
                <tr class="text-left">
                  <th class="py-2 pr-4"><button class="th-sort font-semibold" data-sort="title" type="button">Event Title${sortIcon('title')}</button></th>
                  <th class="py-2 pr-4"><button class="th-sort font-semibold" data-sort="score" type="button">Risk Score${sortIcon('score')}</button></th>
                  <th class="py-2 pr-4"><button class="th-sort font-semibold" data-sort="status" type="button">Status${sortIcon('status')}</button></th>
                  <th class="py-2 pr-4">Objectives</th>
                </tr>
              </thead>
              <tbody class="tbody-fade">${tableRows || ''}</tbody>
            </table>
          </div>
        </div>
      </details>`;
    }
    function updateAriaSort(root) {
      root.querySelectorAll('table thead th').forEach(th => { th.setAttribute('aria-sort', 'none'); });
      root.querySelectorAll('.th-sort').forEach(btn => {
        const key = btn.getAttribute('data-sort');
        const th = btn.closest('th'); if (!th) return;
        th.setAttribute('aria-sort', EVENT_SORT.key === key ? (EVENT_SORT.dir === 'asc' ? 'ascending' : 'descending') : 'none');
      });
    }
    function updateSortIndicators(root) {
      root.querySelectorAll('table thead .th-sort').forEach(btn => {
        const key = btn.getAttribute('data-sort');
        const base = btn.dataset.baseLabel || btn.textContent.replace(/[▲▼]\s*$/, '').trim();
        btn.dataset.baseLabel = base;
        let suffix = '';
        if (EVENT_SORT.key === key) suffix = (EVENT_SORT.dir === 'asc' ? ' ▲' : ' ▼');
        btn.textContent = base + suffix;
      });
    }
    async function renderAllEventsTable(items, hazards) {
      HAZARDS_CACHE = hazards ? hazards.slice() : HAZARDS_CACHE;
      const root = document.getElementById('hazardsAccordion');
      const rowsSorted = sortEvents(items || []);
      root.innerHTML = `<div class="overflow-x-auto">
        <table class="min-w-full text-sm" id="allEventsTable" data-hazard="all">
          <thead><tr class="text-left">
            <th class="py-2 pr-4"><button class="th-sort font-semibold" data-sort="title" type="button">Event Title${sortIcon('title')}</button></th>
            <th class="py-2 pr-4"><button class="th-sort font-semibold" data-sort="hazard" type="button">Hazard${sortIcon('hazard')}</button></th>
            <th class="py-2 pr-4"><button class="th-sort font-semibold" data-sort="score" type="button">Risk Score${sortIcon('score')}</button></th>
            <th class="py-2 pr-4"><button class="th-sort font-semibold" data-sort="status" type="button">Status${sortIcon('status')}</button></th>
            <th class="py-2 pr-4">Objectives</th>
          </tr></thead>
          <tbody class="tbody-fade">${buildEventRowsHTML_All(rowsSorted)}</tbody>
        </table></div>`;
      updateSortIndicators(root);
      updateAriaSort(root);
    }
    function renderHazardsAccordion(hazards, items) {
      const root = el('hazardsAccordion');
      if (!hazards.length && !items.length) {
        root.innerHTML = `<div class="text-slate-500">No Hazards or Hazardous Events yet. Use <strong>Add Hazard</strong> or <strong>Add Hazardous Event</strong> to begin.</div>`;
        updateAriaSort(root);
        return;
      }
      const hazardsAZ = sortHazardsByTitle(hazards);
      const byHaz = new Map();
      hazardsAZ.forEach(h => byHaz.set(h.id, []));
      const uncategorised = [];
      items.forEach(ev => { if (ev.hazardId && byHaz.has(ev.hazardId)) byHaz.get(ev.hazardId).push(ev); else uncategorised.push(ev); });
      const blocks = [];
      hazardsAZ.forEach(h => { blocks.push(hazardBlock(h, byHaz.get(h.id) || [])); });
      blocks.push(hazardBlock({ id: null, title: 'Uncategorised' }, uncategorised, true));
      root.innerHTML = blocks.join('');
      updateSortIndicators(root);
      updateAriaSort(root);
    }
    async function renderRisksView(){
      const [hazards, items] = await Promise.all([SessionStore.getAllHazards(), SessionStore.getAll()]);
      HAZARDS_CACHE = hazards.slice();
      const filtered = filterItemsBySearch(items);
      if (RISKS_MODE === 'grouped') {
        renderHazardsAccordion(hazards, filtered);
      } else {
        renderAllEventsTable(filtered, hazards);
      }
    }
    async function renderListView() {
      if (PRIMARY_TAB === 'objectives') {
        await renderObjectivesList();
        el('risksToolbar')?.classList.add('hidden');
        return;
      }
      el('risksToolbar')?.classList.remove('hidden');
      await renderRisksView();
    }
    function refreshHazardsAccordion() { renderListView(); }
    async function resortAllTablesInPlace() {
      const [hazards, items] = await Promise.all([SessionStore.getAllHazards(), SessionStore.getAll()]);
      const hazardsAZ = sortHazardsByTitle(hazards);
      const byHaz = new Map(); hazardsAZ.forEach(h => byHaz.set(h.id, []));
      const uncategorised = [];
      items.forEach(ev => { if (ev.hazardId && byHaz.has(ev.hazardId)) byHaz.get(ev.hazardId).push(ev); else uncategorised.push(ev); });
      document.querySelectorAll('#hazardsAccordion table[data-hazard]').forEach(tbl => {
        const hidAttr = tbl.getAttribute('data-hazard');
        const rows = (hidAttr === 'null') ? uncategorised : (hidAttr === 'all') ? items : (byHaz.get(Number(hidAttr)) || []);
        const sorted = sortEvents(rows);
        const tbody = tbl.querySelector('tbody'); if (!tbody) return;
        tbody.style.opacity = '0';
        setTimeout(() => { tbody.innerHTML = buildEventRowsHTML(sorted); tbody.style.opacity = '1'; }, 180);
        updateAriaSort(tbl.closest('#hazardsAccordion') || document);
      });
    }
    async function resortAllEventsInPlace() {
      const tbl = document.getElementById('allEventsTable'); if (!tbl) return;
      const tbody = tbl.querySelector('tbody');
      const [items, hazards] = await Promise.all([SessionStore.getAll(), SessionStore.getAllHazards()]);
      HAZARDS_CACHE = hazards.slice();
      const filtered = filterItemsBySearch(items);
      const sorted = sortEvents(filtered);
      tbody.style.opacity = '0';
      setTimeout(() => { tbody.innerHTML = buildEventRowsHTML_All(sorted); tbody.style.opacity = '1'; }, 180);
      updateAriaSort(tbl.closest('#hazardsAccordion') || document);
    }

    /* ============================== *
     *  Hazard select for editor      *
     * ============================== */
    function populateHazardSelect(selectedId) {
      SessionStore.getAllHazards().then(hazards => {
        const hazardsAZ = sortHazardsByTitle(hazards);
        const sel = el('hazardSel');
        const options = [`<option value="">(Uncategorised)</option>`, ...hazardsAZ.map(h => `<option value="${h.id}">${escapeHtml(h.title)}</option>`) ];
        sel.innerHTML = options.join('');
        if (selectedId != null && hazardsAZ.some(h => h.id === selectedId)) sel.value = String(selectedId);
        else if (editingHazardId != null && hazardsAZ.some(h => h.id === editingHazardId)) sel.value = String(editingHazardId);
        else sel.value = '';
      });
    }

    /* ============================== *
     *  View switching & tabs         *
     * ============================== */
    function switchToList() { el('listView').classList.remove('hidden'); el('editorView').classList.add('hidden'); editingId = null; }
    function switchToEditor() {
      el('listView').classList.add('hidden'); el('editorView').classList.remove('hidden');
      requestAnimationFrame(() => {
        renderAllCharts(el('descEditor'));
        ensureTrailingParagraph();
      });
    }
    function toggleSectionBody(key, on) { const body = el('sec_' + key + '_body'); if (!body) return; body.classList.toggle('hidden', !on); }
    function activateTab(name) {
      const panels = { details: el('tab_details'), analysis: el('tab_analysis'), plan: el('tab_plan') };
      const buttons = Array.from(document.querySelectorAll('[data-tab]'));
      buttons.forEach(b => {
        const on = (b.dataset.tab === name);
        b.classList.toggle('tab-btn-active', on);
        b.classList.toggle('tab-btn-inactive', !on);
        b.setAttribute('aria-selected', on ? 'true' : 'false');
      });
      Object.entries(panels).forEach(([k, panel]) => { if (panel) panel.classList.toggle('active', k === name); });
      if (name === 'plan') { renderPlanGaps(); renderMitigationsList(); }
    }
    function wireEditorTabs() {
      const list = el('editorTabs');
      list?.addEventListener('keydown', (e) => {
        const tabs = Array.from(list.querySelectorAll('[role="tab"]'));
        const idx = tabs.indexOf(document.activeElement);
        if (idx < 0) return;
        if (['ArrowRight','ArrowLeft','Home','End'].includes(e.key)){
          e.preventDefault();
          let next = idx + (e.key==='ArrowRight'?1:e.key==='ArrowLeft'?-1:0);
          if (e.key==='Home') next = 0;
          if (e.key==='End') next = tabs.length-1;
          tabs[(next+tabs.length)%tabs.length].focus();
        }
      });
      document.querySelectorAll('[data-tab]').forEach(btn => { btn.addEventListener('click', () => activateTab(btn.dataset.tab)); });
      activateTab('details');
    }

    /* ============================== *
     *  Objectives tab (grouped + Gantt) *
     * ============================== */
    let OBJECTIVES_TIMELINE_OFFSET = 0;        // months offset (0 = last month..+12 ahead)
    const OBJECTIVES_MONTHS = 14;

    async function setMitigationObjective(itemId, mitigationId, objectiveId) {
      const it = await SessionStore.get(itemId);
      if (!it) return;
      it.planMitigations = (it.planMitigations || []).map(m =>
        m.id === mitigationId ? { ...m, objectiveId: objectiveId ?? null } : m
      );
      await SessionStore.put(it);
    }

    function startOfMonth(d){ const x = new Date(d); x.setDate(1); x.setHours(0,0,0,0); return x; }
    function addMonths(d,n){ const x = new Date(d); x.setMonth(x.getMonth()+n); return x; }
    function monthDiff(a,b){ return (b.getFullYear()-a.getFullYear())*12 + (b.getMonth()-a.getMonth()); }
    function daysInMonth(d){ return new Date(d.getFullYear(), d.getMonth()+1, 0).getDate(); }
    function parseDate(v){ if(!v) return null; const t = new Date(v); return Number.isNaN(t.getTime()) ? null : t; }
    function clamp(n,min,max){ return Math.max(min, Math.min(max, n)); }

    async function renderObjectivesBoard(){
      const root = el('hazardsAccordion');
      const [objectives, items] = await Promise.all([SessionStore.getAllObjectives(), SessionStore.getAll()]);
      const allMits = [];
      items.forEach(it => (it.planMitigations||[]).forEach(m => allMits.push({ itemId: it.id, eventTitle: it.title || '', mitigation: m })));

      const monthsBase = startOfMonth(new Date());
      const windowStart = addMonths(monthsBase, OBJECTIVES_TIMELINE_OFFSET - 1); // last month
      const months = Array.from({length: OBJECTIVES_MONTHS}, (_,i)=> addMonths(windowStart, i));

      const headerMonthCells = months.map(dt =>
        `<div class="gantt-cell header">${dt.toLocaleString(undefined,{month:'short'})}<br><span class="text-slate-400">${String(dt.getFullYear()).slice(2)}</span></div>`
      ).join('');

      function timelineBoundsText(){
        const a = months[0].toLocaleString(undefined,{month:'short',year:'numeric'});
        const b = months[months.length-1].toLocaleString(undefined,{month:'short',year:'numeric'});
        return `${a} → ${b}`;
      }

      function totalDaysInRange(){
        return months.reduce((acc,dt)=> acc + daysInMonth(dt),0);
      }

      function todayPercent(){
        const today = new Date();
        let daysFromStart = 0;
        for (let i=0;i<months.length;i++){
          const ms = months[i]; const me = addMonths(ms,1);
          if (today < ms) break;
          if (today >= me){ daysFromStart += daysInMonth(ms); continue; }
          daysFromStart += Math.max(0, today.getDate() - 1);
          break;
        }
        return (daysFromStart / Math.max(1,totalDaysInRange())) * 100;
      }

      function barStyleFor(m){
        const s = parseDate(m.start);
        const e = parseDate(m.end);
        const startIdx = s ? clamp(monthDiff(windowStart, startOfMonth(s)), 0, months.length-1)
                           : clamp(monthDiff(windowStart, startOfMonth(new Date())), 0, months.length-1);
        const endIdx   = e ? clamp(monthDiff(windowStart, startOfMonth(e)), 0, months.length-1)
                           : startIdx;

        const startMonth = addMonths(windowStart, startIdx);
        const endMonth   = addMonths(windowStart, endIdx);

        let daysBefore = 0;
        for (let i=0;i<startIdx;i++) daysBefore += daysInMonth(addMonths(windowStart,i));
        if (s && s > startMonth) daysBefore += (s.getDate() - 1);

        let daysSpan = 0;
        if (startIdx === endIdx){
          const base = s ? s.getDate() : 1;
          const endd = e ? e.getDate() : base;
          daysSpan = Math.max(1, endd - base + 1);
        } else {
          const dimStart = daysInMonth(startMonth);
          daysSpan += s ? (dimStart - s.getDate() + 1) : Math.ceil(dimStart/2);
          for (let i=startIdx+1;i<endIdx;i++) daysSpan += daysInMonth(addMonths(windowStart,i));
          daysSpan += e ? e.getDate() : Math.ceil(daysInMonth(endMonth)/2);
        }

        const total = Math.max(1,totalDaysInRange());
        const left = (daysBefore / total) * 100;
        const width = (daysSpan / total) * 100;
        return `left:${left.toFixed(3)}%;width:${Math.max(0.8,width).toFixed(3)}%;`;
      }

      function rowHtml(rec){
        const m = rec.mitigation;
        const title = escapeHtml(m.title || '(untitled mitigation)');
        const dates = (m.start || m.end) ? `${escapeHtml(m.start||'?')} → ${escapeHtml(m.end||'?')}` : 'No dates';
        const objectiveControls = m.objectiveId
          ? `<button class="row-btn" data-detach data-item="${rec.itemId}" data-mid="${m.id}">Detach</button>`
          : `<span class="text-xs text-slate-500">Attach to:</span>
             <select class="border rounded px-2 py-1 text-sm" data-attach-select data-item="${rec.itemId}" data-mid="${m.id}">
               <option value="">Choose…</option>
               ${objectives.map(a=> `<option value="${a.id}">${escapeHtml(a.title)}</option>`).join('')}
             </select>
             <button class="row-btn" data-attach data-item="${rec.itemId}" data-mid="${m.id}">Attach</button>`;

        return `
          <div class="gantt-title">
            ${title}
            <div class="text-xs text-slate-500">${escapeHtml(rec.eventTitle)} • ${dates}</div>
            <div class="mt-1 flex items-center gap-2 flex-wrap">${objectiveControls}</div>
          </div>
          <div class="gantt-track" style="grid-column: 2 / -1;">
            <div class="gantt-bar" style="${barStyleFor(m)}"></div>
          </div>
        `;
      }

      function boardFor(title, list, objectiveId = null){
        const colStyle = `grid-template-columns: 18rem repeat(${months.length}, minmax(60px, 1fr))`;
        const headerRow = `
          <div class="gantt-title font-semibold">${escapeHtml(title)}</div>
          ${headerMonthCells}
        `;
        const rows = list.length ? list.map(rowHtml).join('') : `
          <div class="gantt-title text-slate-500">No mitigations</div>
          ${months.map(()=>'<div class="gantt-cell"></div>').join('')}
        `;
        const controls = objectiveId ? `<div class="flex gap-2">
            <button class="row-btn" data-action="objective-edit" data-id="${objectiveId}">Edit</button>
            <button class="row-btn" data-action="objective-del" data-id="${objectiveId}">Delete</button>
          </div>` : '';

        return `
          <details class="rounded-xl border">
            <summary class="cursor-pointer select-none flex items-center justify-between p-4">
              <div class="flex items-center gap-3">
                <svg class="chev transition-transform" width="16" height="16" viewBox="0 0 20 20"><path fill="currentColor" d="M7 5l6 5-6 5V5z"/></svg>
                <div>
                  <div class="font-semibold">${escapeHtml(title)}</div>
                  <div class="text-sm text-slate-500">${list.length} mitigation${list.length===1?'':'s'}</div>
                </div>
              </div>
              ${controls}
            </summary>
            <div class="p-4 pt-0">
              <div class="relative gantt-wrap">
                <div class="gantt-grid" style="${colStyle}">
                  <div class="gantt-header">${headerRow}</div>
                  <div class="gantt-rows">${rows}</div>
                </div>
                <div class="today-line" style="left: calc(18rem + (100% - 18rem) * ${(todayPercent()/100).toFixed(6)})"></div>
              </div>
            </div>
          </details>
        `;
      }

      const byObjective = new Map();
      objectives.forEach(a => byObjective.set(a.id, []));
      const unassigned = [];
      allMits.forEach(rec => {
        const aid = rec.mitigation.objectiveId || null;
        if (aid && byObjective.has(aid)) byObjective.get(aid).push(rec); else unassigned.push(rec);
      });

      const sortMit = (list)=> list.slice().sort((r1,r2)=>{
        const a = parseDate(r1.mitigation.start)?.getTime() || 0;
        const b = parseDate(r2.mitigation.start)?.getTime() || 0;
        return a - b;
      });

      const blocks = [];
      blocks.push(`
        <div class="objectives-toolbar gantt-toolbar">
          <div class="flex items-center gap-2">
            <button class="row-btn" id="objectivesPrev" title="Back one month">←</button>
            <div class="text-sm text-slate-500">Timeline: ${timelineBoundsText()}</div>
            <button class="row-btn" id="objectivesNext" title="Forward one month">→</button>
          </div>
          <div class="flex items-center gap-2">
            <button class="px-3 py-1.5 rounded-lg bg-sky-600 hover:bg-sky-700 text-white text-sm" data-action="objective-add">+ Add Objective</button>
          </div>
        </div>
      `);

      blocks.push(boardFor('Not attached to Objective', sortMit(unassigned), null));
      const objectivesAZ = objectives.slice().sort((a,b)=>(a.title||'').localeCompare(b.title||'',undefined,{sensitivity:'base'}));
      objectivesAZ.forEach(a => blocks.push(boardFor(a.title, sortMit(byObjective.get(a.id)||[]), a.id)));

      root.innerHTML = blocks.join('');

      el('objectivesPrev')?.addEventListener('click', ()=> { OBJECTIVES_TIMELINE_OFFSET -= 1; renderObjectivesBoard(); });
      el('objectivesNext')?.addEventListener('click', ()=> { OBJECTIVES_TIMELINE_OFFSET += 1; renderObjectivesBoard(); });
    }

    async function renderObjectivesList(){
      await renderObjectivesBoard();
    }

    (function extendListDelegationForObjectives(){
      const listRoot = el('hazardsAccordion');
      if (!listRoot) return;
      listRoot.addEventListener('click', async (e) => {
        const attachBtn = e.target.closest('button[data-attach]');
        const detachBtn = e.target.closest('button[data-detach]');
        if (attachBtn){
          const itemId = Number(attachBtn.getAttribute('data-item'));
          const mid = Number(attachBtn.getAttribute('data-mid'));
          const sel = listRoot.querySelector(`select[data-attach-select][data-item="${itemId}"][data-mid="${mid}"]`);
          const val = sel?.value ? Number(sel.value) : NaN;
          if (!Number.isFinite(val)) { alert('Choose an objective to attach to.'); return; }
          await setMitigationObjective(itemId, mid, val);
          await recomputeDirtyAgainstSnapshot();
          renderObjectivesBoard();
          return;
        }
        if (detachBtn){
          const itemId = Number(detachBtn.getAttribute('data-item'));
          const mid = Number(detachBtn.getAttribute('data-mid'));
          await setMitigationObjective(itemId, mid, null);
          await recomputeDirtyAgainstSnapshot();
          renderObjectivesBoard();
          return;
        }
      });
    })();

    /* ============================== *
     *  RTE                           *
     * ============================== */
    function exec(cmd, val = null) { document.execCommand(cmd, false, val); el('descEditor').focus(); markDirtySoon(); }

    function sanitizePaste(e){
      e.preventDefault();
      const text = (e.clipboardData || window.clipboardData).getData('text/plain') || '';
      if (document.queryCommandSupported && document.queryCommandSupported('insertText')) {
        document.execCommand('insertText', false, text);
      } else {
        const sel = window.getSelection();
        if (!sel.rangeCount) return;
        const range = sel.getRangeAt(0);
        range.deleteContents();
        range.insertNode(document.createTextNode(text));
        range.collapse(false);
        sel.removeAllRanges(); sel.addRange(range);
      }
      markDirtySoon();
    }

    function wireRte() {
      const tb = document.querySelector('.rte-toolbar');
      if (!tb) return;
      tb.addEventListener('click', (e) => {
        const btn = e.target.closest('button'); if (!btn) return;
        if (btn.hasAttribute('data-cmd')) exec(btn.getAttribute('data-cmd'));
        else if (btn.hasAttribute('data-link')) {
          const url = prompt('Enter URL (https://…):');
          if (url) exec('createLink', safeUrl(url));
        } else if (btn.hasAttribute('data-clear')) {
          if (confirm('Clear description?')) { el('descEditor').innerHTML = ''; ensureTrailingParagraph(); markDirtySoon(); }
        }
      });
      el('descEditor').addEventListener('input', markDirtySoon);
      el('descEditor').addEventListener('paste', sanitizePaste);
    }

    /* ============================== *
     *  Dirty helpers                 *
     * ============================== */
    let markDirtyTimer = null;
    function markDirty() { fileDirty = true; refreshStatus(); }
    function markDirtySoon() { clearTimeout(markDirtyTimer); markDirtyTimer = setTimeout(markDirty, 150); }

    /* ============================== *
     *  File helpers (.litl)          *
     * ============================== */
    function createLitlBlob(items, hazards, objectives) {
      const litl = { litlVersion: 1, appId: 'crr-v1', title: 'Community Risk Register', data: { items, hazards, objectives } };
      return new Blob([JSON.stringify(litl, null, 2)], { type: 'application/x-litl' });
    }

    async function buildCurrentLitlBlob() {
      const [items, hazards, objectives] = await Promise.all([SessionStore.getAll(), SessionStore.getAllHazards(), SessionStore.getAllObjectives()]);
      return createLitlBlob(items, hazards, objectives);
    }

    async function applyLitlPayload(payload, sourceName) {
      SessionStore.setAll(payload);
      refreshHazardsAccordion();
      switchToList();
      if (sourceName) currentFileName = sourceName;
      await updateSavedSnapshot();
      updateFileMenuState();
    }

    function normalizeLitlJson(data) {
      if (!data) return { items: [], hazards: [], objectives: [] };
      const payload = data?.data
        ? { items: Array.isArray(data.data.items) ? data.data.items : [],
            hazards: Array.isArray(data.data.hazards) ? data.data.hazards : [],
            objectives: Array.isArray(data.data.objectives) ? data.data.objectives : [] }
        : { items: Array.isArray(data.items) ? data.items : [],
            hazards: Array.isArray(data.hazards) ? data.hazards : [],
            objectives: Array.isArray(data.objectives) ? data.objectives : [] };
      return payload;
    }

    async function openDriveFile(rawInput, opts = {}) {
      const fileId = parseDriveFileId(rawInput);
      if (!fileId) throw new Error('Unable to determine Google Drive file ID.');
      const { json, meta } = await downloadDriveLitl(fileId, opts);
      const payload = normalizeLitlJson(json);
      currentFileHandle = null;
      currentDriveFileId = fileId;
      const fallbackName = json?.title ? `${json.title}.litl` : `${fileId}.litl`;
      currentFileName = (meta?.name || fallbackName || 'google_drive.litl');
      await applyLitlPayload(payload, currentFileName);
      updateDriveUrlParam(fileId);
    }

    async function promptOpenDriveFile() {
      const input = prompt('Enter the Google Drive file link or ID to open:');
      if (!input) return;
      try {
        await openDriveFile(input, { interactive: true });
      } catch (err) {
        console.error('Open Drive file failed', err);
        alert('Failed to open Google Drive file: ' + (err?.message || err));
      }
    }

    async function saveLitlToDriveExisting(blob) {
      if (!hasDriveTarget()) throw new Error('No Google Drive file selected.');
      const result = await uploadDriveLitl(currentDriveFileId, blob, currentFileName, { interactive: true });
      if (result?.name) currentFileName = result.name;
      await updateSavedSnapshot();
      updateDriveUrlParam(currentDriveFileId);
      updateFileMenuState();
    }

    async function saveLitlToDriveAs(blob) {
      if (!driveConfigured()) throw new Error('Google Drive client ID is not configured.');
      let defaultName = currentFileName || 'community_risk_register.litl';
      if (!/\.litl$/i.test(defaultName)) defaultName += '.litl';
      const name = prompt('Save to Google Drive as:', defaultName);
      if (!name) return;
      const result = await createDriveLitl(blob, name.trim(), { interactive: true });
      currentDriveFileId = result?.id || null;
      currentFileHandle = null;
      currentFileName = result?.name || name.trim();
      await updateSavedSnapshot();
      updateDriveUrlParam(currentDriveFileId);
      updateFileMenuState();
    }

    async function copyDriveShareLink() {
      const share = driveShareableUrl();
      if (!share) { alert('No Google Drive file to share. Save to Google Drive first.'); return; }
      try {
        if (navigator.clipboard?.writeText) {
          await navigator.clipboard.writeText(share);
          alert('Share link copied to clipboard.');
          return;
        }
      } catch (err) {
        console.warn('Clipboard write failed', err);
      }
      prompt('Copy this shareable link:', share);
    }

    async function openLitlWithPicker() {
      const [handle] = await window.showOpenFilePicker({ multiple: false, types: [{ description: 'litl files', accept: { 'application/x-litl': ['.litl'] } }] });
      currentFileHandle = handle;
      currentFileName = handle.name || 'untitled.litl';
      currentDriveFileId = null;
      updateDriveUrlParam(null);
      const file = await handle.getFile();
      const text = await file.text();
      const data = JSON.parse(text);
      const payload = normalizeLitlJson(data);
      await applyLitlPayload(payload, currentFileName);
    }
    function openLitlWithInput(file) {
      const reader = new FileReader();
      reader.onload = async () => {
        try {
          const json = JSON.parse(reader.result);
          const payload = normalizeLitlJson(json);
          currentFileHandle = null;
          currentFileName = file.name || 'loaded.litl';
          currentDriveFileId = null;
          updateDriveUrlParam(null);
          await applyLitlPayload(payload, currentFileName);
        } catch (err) { console.error(err); alert('Invalid .litl file.'); }
      };
      reader.readAsText(file);
    }
    async function saveLitl() {
      try {
        saving = true; refreshStatus();
        const blob = await buildCurrentLitlBlob();
        if (hasDriveTarget()) {
          await saveLitlToDriveExisting(blob);
          saving = false; refreshStatus(); return;
        }
        if (hasLocalHandle()) {
          const writable = await currentFileHandle.createWritable();
          await writable.write(blob); await writable.close();
          await updateSavedSnapshot(); saving = false; refreshStatus(); return;
        }
        await saveLitlAs(); saving = false; refreshStatus();
      } catch (err) { console.error('Save .litl failed', err); saving = false; refreshStatus(); alert('Save failed: ' + (err?.message || err)); }
    }
    async function saveLitlAs() {
      try {
        saving = true; refreshStatus();
        const blob = await buildCurrentLitlBlob();
        if (window.showSaveFilePicker) {
          const handle = await window.showSaveFilePicker({
            suggestedName: currentFileName || 'community_risk_register.litl',
            types: [{ description: 'litl files', accept: { 'application/x-litl': ['.litl'] } }]
          });
          currentFileHandle = handle; currentFileName = handle.name || 'community_risk_register.litl';
          currentDriveFileId = null; updateDriveUrlParam(null);
          const writable = await handle.createWritable(); await writable.write(blob); await writable.close();
          await updateSavedSnapshot(); updateFileMenuState(); saving = false; refreshStatus(); return;
        }
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = currentFileName || 'community_risk_register.litl';
        document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
        if (!currentFileName) currentFileName = 'community_risk_register.litl';
        currentDriveFileId = null; updateDriveUrlParam(null);
        await updateSavedSnapshot(); updateFileMenuState(); saving = false; refreshStatus();
      } catch (err) { console.error('Save As .litl failed', err); saving = false; refreshStatus(); alert('Save As failed: ' + (err?.message || err)); }
    }
    function newRegister() {
      SessionStore.setAll({ items: [], hazards: [], objectives: [] });
      currentFileHandle = null;
      currentFileName = null;
      currentDriveFileId = null;
      updateDriveUrlParam(null);
      refreshHazardsAccordion();
      switchToList();
      lastSavedSnapshot = '';
      fileDirty = true;
      updateFileMenuState();
      refreshStatus();
    }

    /* ============================== *
     *  File menu wiring              *
     * ============================== */
    function wireFileMenu() {
      const root = document.getElementById('fileMenuRoot');
      const btn = document.getElementById('fileMenuBtn');
      const panel = document.getElementById('fileMenuPanel');
      const openInput = document.getElementById('openLitlInput');
      function closeMenu() { root.classList.remove('menu-open'); panel.classList.add('hidden'); }
      function openMenu() { root.classList.add('menu-open'); panel.classList.remove('hidden'); }
      btn.addEventListener('click', (e) => { e.stopPropagation(); if (root.classList.contains('menu-open')) closeMenu(); else openMenu(); });
      document.addEventListener('click', closeMenu);
      document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeMenu(); });
      document.getElementById('menuNew').addEventListener('click', () => { closeMenu(); newRegister(); });
      document.getElementById('menuOpen').addEventListener('click', async () => {
        closeMenu();
        try { if (window.showOpenFilePicker) await openLitlWithPicker(); else openInput.click(); }
        catch (err) { if (err?.name !== 'AbortError') { console.error(err); } }
      });
      openInput.addEventListener('change', (e) => { const f = e.target.files?.[0]; if (f) openLitlWithInput(f); e.target.value = ''; });
      document.getElementById('menuSave').addEventListener('click', async () => { closeMenu(); await saveLitl(); });
      document.getElementById('menuSaveAs').addEventListener('click', async () => { closeMenu(); await saveLitlAs(); });
      document.getElementById('menuOpenDrive')?.addEventListener('click', async () => {
        closeMenu();
        if (!driveConfigured()) { warnDriveNotConfigured(); return; }
        try { await promptOpenDriveFile(); }
        catch (err) { if (err?.name !== 'AbortError') console.error(err); }
      });
      document.getElementById('menuSaveDrive')?.addEventListener('click', async () => {
        closeMenu();
        if (!driveConfigured()) { warnDriveNotConfigured(); return; }
        try {
          saving = true; refreshStatus();
          const blob = await buildCurrentLitlBlob();
          if (hasDriveTarget()) await saveLitlToDriveExisting(blob);
          else await saveLitlToDriveAs(blob);
        } catch (err) {
          console.error('Save to Google Drive failed', err);
          alert('Failed to save to Google Drive: ' + (err?.message || err));
        } finally {
          saving = false; refreshStatus();
        }
      });
      document.getElementById('menuSaveDriveAs')?.addEventListener('click', async () => {
        closeMenu();
        if (!driveConfigured()) { warnDriveNotConfigured(); return; }
        try {
          saving = true; refreshStatus();
          const blob = await buildCurrentLitlBlob();
          await saveLitlToDriveAs(blob);
        } catch (err) {
          console.error('Save As to Google Drive failed', err);
          alert('Failed to save to Google Drive: ' + (err?.message || err));
        } finally {
          saving = false; refreshStatus();
        }
      });
      document.getElementById('menuShareDrive')?.addEventListener('click', async () => {
        closeMenu();
        if (!driveConfigured()) { warnDriveNotConfigured(); return; }
        if (!hasDriveTarget()) { alert(DRIVE_NEEDS_FILE_REASON); return; }
        try { await copyDriveShareLink(); }
        catch (err) { console.error('Share link failed', err); alert('Unable to prepare share link: ' + (err?.message || err)); }
      });
      document.addEventListener('keydown', (e) => {
        const ctrl = e.ctrlKey || e.metaKey;
        if (ctrl && e.key.toLowerCase() === 's') { e.preventDefault(); saveLitl(); }
        if (ctrl && e.key.toLowerCase() === 'o') { e.preventDefault(); if (window.showOpenFilePicker) openLitlWithPicker(); else openInput.click(); }
        if (ctrl && e.key.toLowerCase() === 'n') { e.preventDefault(); newRegister(); }
      });
      updateFileMenuState();
    }

    /* ============================== *
     *  List tabs + search wiring     *
     * ============================== */
    function updateListTabsUI() {
      document.querySelectorAll('[data-listtab]').forEach(b => {
        const on = (b.getAttribute('data-listtab') === PRIMARY_TAB);
        b.classList.toggle('tab-btn-active', on);
        b.classList.toggle('tab-btn-inactive', !on);
        b.setAttribute('aria-selected', on ? 'true' : 'false');
      });
    }

    const applySearch = makeDebounced(() => {
      const searchEl = el('listSearch');
      LIST_SEARCH = (searchEl?.value || '');
      if (LIST_SEARCH.trim()) {
        PRIMARY_TAB = 'risks';
        RISKS_MODE = 'all';
        const sel = el('risksModeSelect'); if (sel) sel.value = 'all';
        updateListTabsUI();
      }
      renderListView();
    }, 120);

    /* ============================== *
     *  Event delegation (list area)  *
     * ============================== */
    function wireListDelegation() {
      const listRoot = el('hazardsAccordion');
      if (!listRoot) return;

      document.getElementById('listTabs')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-listtab]'); if (!btn) return;
        const tab = btn.getAttribute('data-listtab');
        if (tab === 'risks' || tab === 'objectives') {
          PRIMARY_TAB = tab;
          updateListTabsUI();
          renderListView();
        }
      });

      listRoot.addEventListener('click', async (e) => {
        const btn = e.target.closest('button');
        if (btn?.dataset?.action) {
          e.stopPropagation();
          const action = btn.dataset.action;
          const id = Number(btn.dataset.id);
          if (action === 'del') {
            if (!confirm('Delete this hazardous event? This cannot be undone.')) return;
            await SessionStore.delete(id);
            await recomputeDirtyAgainstSnapshot();
            renderListView();
            return;
          }

          // Objectives tab: main controls
          if (btn?.dataset?.action === 'objective-add') { openObjectiveModal(null); return; }
          if (btn?.dataset?.action === 'objective-edit') { openObjectiveModal(Number(btn.dataset.id)); return; }
          if (btn?.dataset?.action === 'objective-del') {
            if (!confirm('Delete this objective?')) return;
            await SessionStore.deleteObjective(Number(btn.dataset.id));
            await recomputeDirtyAgainstSnapshot();
            renderObjectivesList();
            return;
          }

          if (action === 'dup') { duplicateEvent(id); return; }
          if (action === 'edit') { loadEvent(id); return; }
          if (action === 'haz-rename') {
            const hid = btn.dataset.id === 'null' ? null : Number(btn.dataset.id);
            if (hid === null) { alert('Cannot rename the Uncategorised group.'); return; }
            const hzs = await SessionStore.getAllHazards();
            const hz = hzs.find(z => z.id === hid);
            if (!hz) return;
            const t = prompt('Rename hazard:', hz.title);
            if (t && t.trim()) { await SessionStore.putHazard({ id: hid, title: t.trim() }); await recomputeDirtyAgainstSnapshot(); refreshHazardsAccordion(); }
            return;
          }
          if (action === 'haz-delete') {
            const hid = btn.dataset.id === 'null' ? null : Number(btn.dataset.id);
            if (hid === null) { alert('Uncategorised cannot be deleted.'); return; }
            if (!confirm('Delete this hazard? All associated events will be moved to Uncategorised.')) return;
            await SessionStore.deleteHazard(hid);
            await recomputeDirtyAgainstSnapshot();
            refreshHazardsAccordion();
            return;
          }
        }
        const sortBtn = e.target.closest('.th-sort');
        if (sortBtn) {
          e.preventDefault();
          const key = sortBtn.getAttribute('data-sort');
          if (EVENT_SORT.key === key) EVENT_SORT.dir = (EVENT_SORT.dir === 'asc' ? 'desc' : 'asc');
          else { EVENT_SORT.key = key; EVENT_SORT.dir = (key === 'score' ? 'desc' : 'asc'); }
          updateSortIndicators(listRoot);

          if (PRIMARY_TAB === 'risks') {
            if (RISKS_MODE === 'all') await resortAllEventsInPlace();
            else await resortAllTablesInPlace();
          }
          return;
        }

        const tr = e.target.closest('tbody tr[data-id]');
        if (tr && !btn) loadEvent(Number(tr.dataset.id));
      });
    }

    /* ============================== *
     *  Mitigation buttons wiring     *
     * ============================== */
    function wireMitigationDelegation() {
      const root = el('mitigationsList');
      if (!root) return;
      root.addEventListener('click', (e) => {
        const btn = e.target.closest('button[data-mit]');
        if (!btn) return;
        const idx = Number(btn.dataset.index || btn.closest('[data-mitigation-index]')?.getAttribute('data-mitigation-index'));
        if (Number.isNaN(idx)) return;
        if (btn.dataset.mit === 'edit') { openMitigationModal(idx); markDirtySoon(); }
        if (btn.dataset.mit === 'del') {
          if (!confirm('Delete this mitigation?')) return;
          planMitigations.splice(idx, 1);
          renderMitigationsList(); renderPlanGaps(); markDirtySoon();
        }
      });
    }

    /* ============================== *
     *  Add / Duplicate helpers       *
     * ============================== */
    function duplicateEvent(id) {
      SessionStore.get(id).then(r => {
        if (!r) return;
        const { id: _omit, createdAt: __omit, ...rest } = r;
        const copy = { ...rest, title: r.title + ' (copy)', createdAt: new Date().toISOString() };
        return SessionStore.add(copy);
      }).then(async () => {
        refreshHazardsAccordion();
        await recomputeDirtyAgainstSnapshot();
      }).catch(err => { console.error('Duplicate failed', err); alert('Duplicate failed: ' + (err?.message || err)); });
    }

    /* ===== Chart helpers (Bar, Pie, Line) ===== */
    function makeEmptyPara() { const p = document.createElement('p'); p.innerHTML = '<br>'; return p; }
    function placeCaretIn(elm, atEnd = false) {
      const sel = window.getSelection(); const r = document.createRange();
      r.selectNodeContents(elm); r.collapse(!atEnd); sel.removeAllRanges(); sel.addRange(r);
    }
    function ensureRoomAroundFigure(fig) {
      const ed = document.getElementById('descEditor'); if (!ed) return;
      const prev = fig.previousSibling; if (!(prev && prev.nodeType === 1 && prev.tagName === 'P')) { ed.insertBefore(makeEmptyPara(), fig); }
      const next = fig.nextSibling; if (!(next && next.nodeType === 1 && next.tagName === 'P')) { if (fig.nextSibling) ed.insertBefore(makeEmptyPara(), fig.nextSibling); else ed.appendChild(makeEmptyPara()); }
    }
    function ensureTrailingParagraph() {
      const ed = document.getElementById('descEditor'); if (!ed) return;
      const last = ed.lastElementChild;
      const isEmptyPara = (el) => el && el.tagName === 'P' && (el.innerHTML === '' || el.innerHTML === '<br>');
      if (!isEmptyPara(last)) ed.appendChild(makeEmptyPara());
    }
    function sizeCanvasToParentWidth(canvas, desiredHeightPx) {
      const dpr = window.devicePixelRatio || 1;
      const parent = canvas.parentElement || canvas;
      const cssW = Math.max(260, Math.floor(parent.getBoundingClientRect().width || 600));
      const cssH = Math.max(120, Number(desiredHeightPx) || 320);
      canvas.style.width = '100%'; canvas.style.height = cssH + 'px';
      canvas.width = Math.round(cssW * dpr); canvas.height = Math.round(cssH * dpr);
      const ctx = canvas.getContext('2d'); ctx.setTransform(dpr,0,0,dpr,0,0); ctx.clearRect(0,0,cssW,cssH);
      return { ctx, W: cssW, H: cssH };
    }
    function parseCSVishLines(txt) {
      const lines = (txt || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      const labels = [], values = [];
      for (const line of lines) {
        const m = line.match(/^\s*(?:"([^"]+)"|([^,]+))\s*,\s*(-?\d+(?:\.\d+)?)\s*$/);
        if (!m) continue;
        const label = (m[1] ?? m[2] ?? '').trim();
        const val = Number(m[3]);
        if (!Number.isFinite(val)) continue;
        labels.push(label); values.push(val);
      }
      return { labels, values };
    }
    function drawBarChart(canvas, spec) {
      const { labels, values, h } = spec;
      const { ctx, W, H } = sizeCanvasToParentWidth(canvas, h);
      const pad = { t: 16, r: 12, b: 44, l: 40 };
      const plotW = W - pad.l - pad.r;
      const plotH = H - pad.t - pad.b;
      const maxV = Math.max(1, Math.max(...values, 0));
      const n = values.length, gap = 10;
      const barW = Math.max(6, (plotW - gap*(n-1)) / Math.max(1,n));
      ctx.strokeStyle = '#CBD5E1'; ctx.beginPath(); ctx.moveTo(pad.l, pad.t); ctx.lineTo(pad.l, pad.t + plotH); ctx.lineTo(pad.l + plotW, pad.t + plotH); ctx.stroke();
      ctx.fillStyle = 'rgba(99,102,241,0.85)';
      values.forEach((v, i) => { const x = pad.l + i*(barW + gap); const h = (v / maxV) * (plotH - 1); const y = pad.t + plotH - h; ctx.fillRect(x, y, barW, h); });
      ctx.fillStyle = '#475569'; ctx.font = '12px system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial';
      ctx.textAlign = 'center'; ctx.textBaseline = 'top';
      labels.forEach((lab, i) => { const x = pad.l + i*(barW + gap) + barW/2; const text = lab.length > 20 ? (lab.slice(0,17)+'…') : lab; ctx.fillText(text, x, pad.t + plotH + 6); });
    }
    function drawPieChart(canvas, spec) {
      const { labels, values, h } = spec;
      const { ctx, W, H } = sizeCanvasToParentWidth(canvas, h);
      const cx = W/2, cy = H/2, r = Math.min(W,H)*0.35;
      const sum = values.reduce((a,b)=>a+b,0) || 1; let start = -Math.PI/2;
      values.forEach((v, i) => { const frac = v / sum, end = start + frac * Math.PI * 2; ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, r, start, end); ctx.closePath(); ctx.fillStyle = `hsl(${(i*53)%360} 70% 55% / 0.85)`; ctx.fill(); start = end; });
      const lx = Math.min(W - 140, cx + r + 20); let ly = Math.max(20, cy - 0.5 * values.length * 18);
      ctx.font = '12px system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial';
      labels.forEach((lab, i) => { ctx.fillStyle = `hsl(${(i*53)%360} 70% 45%)`; ctx.fillRect(lx, ly + 2, 10, 10); ctx.fillStyle = '#334155'; const text = lab.length > 24 ? lab.slice(0,21)+'…' : lab; ctx.fillText(text, lx + 16, ly); ly += 18; });
    }
    function drawLineChart(canvas, spec) {
      const { labels, values, h } = spec;
      const { ctx, W, H } = sizeCanvasToParentWidth(canvas, h);
      const pad = { t: 16, r: 16, b: 44, l: 44 };
      const plotW = W - pad.l - pad.r; const plotH = H - pad.t - pad.b;
      if (!values.length) return;
      const minV = Math.min(...values); const maxV = Math.max(...values); const span = (maxV - minV) || 1;
      const xAt = (i) => pad.l + (i/(Math.max(1, values.length-1))) * plotW;
      const yAt = (v) => pad.t + plotH - ((v - minV)/span) * plotH;
      ctx.strokeStyle = '#CBD5E1'; ctx.beginPath(); ctx.moveTo(pad.l, pad.t); ctx.lineTo(pad.l, pad.t + plotH); ctx.lineTo(pad.l + plotW, pad.t + plotH); ctx.stroke();
      ctx.strokeStyle = '#E2E8F0';
      for (let g=1; g<=4; g++){ const y = pad.t + (plotH * g/4); ctx.beginPath(); ctx.moveTo(pad.l, y); ctx.lineTo(pad.l + plotW, y); ctx.stroke(); }
      ctx.strokeStyle = 'rgba(99,102,241,0.95)'; ctx.lineWidth = 2; ctx.beginPath();
      values.forEach((v,i) => { const x = xAt(i), y = yAt(v); if (i===0) ctx.moveTo(x,y); else ctx.lineTo(x,y); }); ctx.stroke();
      ctx.fillStyle = 'rgba(99,102,241,0.95)';
      values.forEach((v,i) => { const x = xAt(i), y = yAt(v); ctx.beginPath(); ctx.arc(x,y,3,0,Math.PI*2); ctx.fill(); });
      ctx.fillStyle = '#475569'; ctx.font = '12px system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial';
      ctx.textAlign = 'center'; ctx.textBaseline = 'top';
      const step = Math.ceil(labels.length / 8);
      labels.forEach((lab,i) => { if (i % step !== 0 && i !== labels.length-1) return; const text = lab.length > 20 ? lab.slice(0,17)+'…' : lab; ctx.fillText(text, xAt(i), pad.t + plotH + 6); });
      ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(String(minV), pad.l - 6, yAt(minV)); ctx.fillText(String(maxV), pad.l - 6, yAt(maxV));
    }
    function renderChart(canvas) {
      if (!canvas) return;
      const specStr = canvas.dataset.spec; if (!specStr) return;
      let spec; try { spec = JSON.parse(specStr); } catch { return; }
      const t = spec.type || 'bar';
      if (t === 'pie') return drawPieChart(canvas, spec);
      if (t === 'line') return drawLineChart(canvas, spec);
      return drawBarChart(canvas, spec);
    }
    function renderAllCharts(root = document) { root.querySelectorAll('figure.rte-chart canvas[data-spec]').forEach(renderChart); }
    function openChartModal(existingFigure) {
      const modal = document.getElementById('chartModal'); const ok = document.getElementById('chartOkBtn'); const cancel = document.getElementById('chartCancelBtn');
      function close(){ closeModal(modal); ok.onclick = null; cancel.onclick = null; }
      if (existingFigure) {
        const canvas = existingFigure.querySelector('canvas[data-spec]'); const cap = existingFigure.querySelector('figcaption');
        try { const spec = JSON.parse(canvas.dataset.spec);
          document.getElementById('chartType').value = spec.type || 'bar';
          document.getElementById('chartTitle').value = cap?.textContent || '';
          document.getElementById('chartH').value = spec.h || 320;
          const lines = spec.labels.map((lab, i) => {
            const needsQuotes = /[",]/.test(lab);
            const safe = needsQuotes ? `"${lab.replace(/"/g, '""')}"` : lab;
            return `${safe}, ${spec.values[i] ?? 0}`;
          }).join('\n');
          document.getElementById('chartData').value = lines;
        } catch {}
      } else {
        document.getElementById('chartType').value = 'bar';
        document.getElementById('chartTitle').value = '';
        document.getElementById('chartData').value = '';
        document.getElementById('chartH').value = 320;
      }
      ok.onclick = () => {
        const type = document.getElementById('chartType').value;
        const title = document.getElementById('chartTitle').value.trim();
        const { labels, values } = parseCSVishLines(document.getElementById('chartData').value);
        const h = parseInt(document.getElementById('chartH').value, 10) || 320;
        if (!labels.length) { alert('Please enter at least one "Label, Value" line.'); return; }
        const spec = { type, labels, values, h };
        if (existingFigure) {
          const canvas = existingFigure.querySelector('canvas[data-spec]');
          const cap = existingFigure.querySelector('figcaption');
          canvas.dataset.spec = JSON.stringify(spec);
          if (cap) cap.textContent = title || '';
          renderChart(canvas);
          ensureRoomAroundFigure(existingFigure); ensureTrailingParagraph();
        } else {
          const fig = document.createElement('figure');
          fig.className = 'rte-chart'; fig.setAttribute('contenteditable', 'false'); fig.tabIndex = 0;
          const canvas = document.createElement('canvas'); canvas.dataset.spec = JSON.stringify(spec); fig.appendChild(canvas);
          const cap = document.createElement('figcaption'); cap.textContent = title || ''; fig.appendChild(cap);
          const sel = window.getSelection(); const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
          const editor = document.getElementById('descEditor');
          if (range && editor.contains(range.startContainer)) { range.collapse(true); range.insertNode(fig); }
          else { editor.appendChild(fig); }
          renderChart(canvas); ensureRoomAroundFigure(fig); ensureTrailingParagraph();
        }
        markDirtySoon(); close();
      };
      cancel.onclick = close;
      openModal(modal, '#chartType');
    }
    function enhanceRTEWithCharts() {
      const editor = document.getElementById('descEditor'); const tb = document.querySelector('.rte-toolbar'); const chartBtn = tb.querySelector('[data-chart]');
      if (chartBtn) chartBtn.addEventListener('click', () => openChartModal(null));
      editor.addEventListener('dblclick', (e) => {
        const fig = e.target.closest('figure.rte-chart');
        if (fig && editor.contains(fig)) { e.preventDefault(); openChartModal(fig); }
      });
      editor.addEventListener('keydown', (e) => {
        const fig = document.activeElement?.closest?.('figure.rte-chart');
        if (!fig) return;

        if (e.key === 'Enter') {
          e.preventDefault();
          const p = makeEmptyPara();
          if (fig.nextSibling) {
            fig.parentNode.insertBefore(p, fig.nextSibling);
          } else {
            fig.parentNode.appendChild(p);
          }
          placeCaretIn(p, true);
        }

        if (e.key === 'Backspace' || e.key === 'Delete') {
          // Allow removing a selected chart with Backspace/Delete
          e.preventDefault();
          fig.remove();
          ensureTrailingParagraph();
        }
      });
    }

    /* ===== Objectives: modal editor (Add/Edit) ===== */
    let editingObjectiveId = null;

    function openObjectiveModal(id = null) {
      editingObjectiveId = id;
      const modal = el('objectiveModal');
      const title = el('objectiveModalTitle');

      if (id == null) {
        title.textContent = 'Add Objective';
        el('objectiveTitleInput').value = '';
        el('objectiveOwnerInput').value = '';
        el('objectiveStatusInput').value = 'Planned';
        el('objectiveColorInput').value = '#2563eb';
        el('objectiveDescInput').value = '';
      } else {
        SessionStore.getAllObjectives().then(list => {
          const a = list.find(x => x.id === id);
          if (!a) return;
          title.textContent = 'Edit Objective';
          el('objectiveTitleInput').value = a.title || '';
          el('objectiveOwnerInput').value = a.owner || '';
          el('objectiveStatusInput').value = a.status || 'Planned';
          el('objectiveColorInput').value = a.color || '#2563eb';
          el('objectiveDescInput').value = a.description || '';
        });
      }

      el('objectiveOkBtn').onclick = async () => {
        const rec = {
          ...(editingObjectiveId != null ? { id: editingObjectiveId } : {}),
          title: el('objectiveTitleInput').value.trim() || 'Untitled Objective',
          owner: el('objectiveOwnerInput').value.trim(),
          status: el('objectiveStatusInput').value,
          color: el('objectiveColorInput').value,
          description: el('objectiveDescInput').value.trim()
        };

        if (editingObjectiveId == null) await SessionStore.addObjective(rec);
        else await SessionStore.putObjective(rec);

        closeModal(modal);
        await recomputeDirtyAgainstSnapshot();

        if (PRIMARY_TAB === 'objectives') {
          await renderObjectivesList();
        } else {
          await renderListView();
        }
      };

      el('objectiveCancelBtn').onclick = () => closeModal(modal);
      openModal(modal, '#objectiveTitleInput');
    }

    /* ============================== *
     *  Risks toolbar objectives         *
     * ============================== */
    function newEvent() {
      editingId = null;
      populateHazardSelect(null);
      el('riskTitle').value = '';
      el('statusSel').value = 'Tolerate';
      el('descEditor').innerHTML = '';
      setLinks([]);
      planMitigations = [];
      renderMitigationsList();
      renderPlanGaps();
      CFG.categories.forEach(c => {
        el('en_' + c.key).checked = false;
        toggleSectionBody(c.key, false);
        c.sub.forEach(sf => { el(sf.id).value = 3; if (GUIDANCE[sf.id]) updateGuidanceHighlight(sf.id); });
        el('eff_' + c.key).value = 1;
        setMitigations(c.key, []);
        setGaps(c.key, []);
      });
      el(CFG.likelihoodId).value = 3;
      SUB_REASONING = {};
      // after building the category UI defaults…
      CFG.categories.forEach(c => c.sub.forEach(sf => updateReasonBadge(sf.id)));
      recalc();
      el('narrative').textContent = 'Press Generate Narrative to produce a tailored summary.';
      activateTab('details');
      switchToEditor();
    }
    async function addHazardFlow() {
      const t = prompt('New hazard title:');
      if (!t || !t.trim()) return;
      await SessionStore.addHazard(t.trim());
      await recomputeDirtyAgainstSnapshot();
      refreshHazardsAccordion();
      populateHazardSelect(null);
    }

    function updateReasonBadge(subId) {
      const badge = el(`rsn_${subId}_badge`);
      if (!badge) return;
      const has = (SUB_REASONING[subId] || '').trim().length > 0;
      badge.classList.toggle('hidden', !has);
    }

    function updateAllReasonBadges() {
      CFG.categories.forEach(c => c.sub.forEach(sf => updateReasonBadge(sf.id)));
    }

    function openReasonModal(sf) {
      el('reasonModalTitle').textContent = `${sf.label} Reasoning`;
      el('reasonTextInput').value = SUB_REASONING[sf.id] || '';
      openModal(el('reasonModal'), '#reasonTextInput');

      el('reasonOkBtn').onclick = () => {
        SUB_REASONING[sf.id] = el('reasonTextInput').value;
        updateReasonBadge(sf.id);
        closeModal(el('reasonModal'));
        markDirtySoon();
      };
      el('reasonCancelBtn').onclick = () => closeModal(el('reasonModal'));
    }


    /* ============================== *
     *  Wire page                     *
     * ============================== */
    async function wire() {
      buildCategoryBlocks();
      CFG.categories.forEach(c => {
        c.sub.forEach(sf => {
          const btn = el(`rsn_${sf.id}`);
          if (btn) btn.addEventListener('click', () => openReasonModal(sf));
        });
      });
      updateAllReasonBadges();
      wireEditorTabs();
      wireRte();
      enhanceRTEWithCharts();
      wireFileMenu();
      wireListDelegation();
      wireMitigationDelegation();

      // Chips add buttons
      CFG.categories.forEach(c => {
        const mitAdd = el(`mit_${c.key}_add`);
        const mitInp = el(`mit_${c.key}_input`);
        if (mitAdd && mitInp) {
          mitAdd.addEventListener('click', () => { const v = mitInp.value.trim(); if (v) { addTag('mit', c.key, v, ''); mitInp.value = ''; }});
          mitInp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); mitAdd.click(); }});
        }
        const gapAdd = el(`gap_${c.key}_add`);
        const gapInp = el(`gap_${c.key}__input`);
        if (gapAdd && gapInp) {
          gapAdd.addEventListener('click', () => { const v = gapInp.value.trim(); if (v) { addTag('gap', c.key, v, ''); gapInp.value=''; }});
          gapInp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); gapAdd.click(); }});
        }
        const eff = el(`eff_${c.key}`);
        if (eff) eff.addEventListener('input', recalc);
        const en = el(`en_${c.key}`);
        if (en) en.addEventListener('change', () => { toggleSectionBody(c.key, en.checked); recalc(); });
        c.sub.forEach(sf => {
          const input = el(sf.id);
          if (input) {
            input.addEventListener('input', () => { recalc(); if (GUIDANCE[sf.id]) updateGuidanceHighlight(sf.id); });
            const guide = el('g_'+sf.id);
            if (guide) {
              input.addEventListener('focus', () => showGuide('g_'+sf.id));
              input.addEventListener('blur', () => hideGuide('g_'+sf.id));
            }
          }
        });
      });

      // Likelihood
      el(CFG.likelihoodId)?.addEventListener('input', recalc);

      // RTE buttons
      el('addLinkBtn')?.addEventListener('click', () => openLinkModal('Add Link'));
      el('generateBtn')?.addEventListener('click', genNarrative);
      el('saveBtn')?.addEventListener('click', saveCurrent);
      el('resetBtn')?.addEventListener('click', () => loadEvent(editingId));
      el('cancelBtn')?.addEventListener('click', switchToList);
      el('backToListBtn')?.addEventListener('click', switchToList);

      // Risks toolbar
      el('risksModeSelect')?.addEventListener('change', (e) => {
        RISKS_MODE = e.target.value === 'grouped' ? 'grouped' : 'all';
        renderListView();
      });
      el('addHazardBtn')?.addEventListener('click', addHazardFlow);
      el('addEventBtn')?.addEventListener('click', newEvent);

      // Search
      el('listSearch')?.addEventListener('input', applySearch);
      el('listSearchClear')?.addEventListener('click', () => { el('listSearch').value=''; applySearch(); });

      // Chart modal
      el('chartOkBtn')?.addEventListener('click', () => {}); // wired inside openChartModal
      el('chartCancelBtn')?.addEventListener('click', () => {}); // wired inside openChartModal

      // Mitigation modal
      el('addMitigationBtn')?.addEventListener('click', () => openMitigationModal(null));

      // Objective modal button handlers are set in openObjectiveModal dynamically

      // Initial render
      await importDriveFromQuery();
      populateHazardSelect(null);
      recalc();
      renderListView();
      refreshStatus();

      // Import via hash if provided
      importFromHash();
    }

    /* ===== run ===== */
    wire().catch((err) => console.error('Initialization failed', err));

  })();
