/* ============================================================
   ClassTools — Teacher Account / Storage  (items 2,5,6,10,11,19)
   Obsidian-style markdown storage:
     • One chosen folder (File System Access API), handle persisted in
       IndexedDB so it reconnects across sessions.
     • The app reads EVERY .md file in the folder on connect and maps
       every list block it finds. It writes continuously on each change.
     • localStorage mirrors everything so it still works with no folder
       (PowerPoint add-in / unsupported browsers) via Export/Import .md.
   List block format (multiple lists can share one file):
     <!-- ct:list id="L.." type="class" title="Grade 7" subjects="Sci, Math" -->
     ## Grade 7
     - Ana
     - Ben
     <!-- ct:end id="L.." -->
   ============================================================ */
(function () {
  "use strict";
  if (window.__ctAccount) return; window.__ctAccount = true;

  var KEY = 'ct-account-v2';
  var fsSupported = (typeof window.showDirectoryPicker === 'function');
  var dirHandle = null;           // live folder handle (this session)
  var folderName = null;
  var lists = [];                 // [{id,type,title,subjects,items}]
  var listeners = [];

  /* ---------- tiny IndexedDB for the folder handle ---------- */
  function idb(op, val) {
    return new Promise(function (res) {
      var r = indexedDB.open('ct-store', 1);
      r.onupgradeneeded = function () { r.result.createObjectStore('kv'); };
      r.onsuccess = function () {
        var db = r.result, tx = db.transaction('kv', op === 'get' ? 'readonly' : 'readwrite');
        var st = tx.objectStore('kv'); var q;
        if (op === 'get') q = st.get('dir'); else if (op === 'set') q = st.put(val, 'dir'); else q = st.delete('dir');
        q.onsuccess = function () { res(op === 'get' ? q.result : true); };
        q.onerror = function () { res(null); };
      };
      r.onerror = function () { res(null); };
    });
  }

  /* ---------- helpers ---------- */
  function uid() { return 'L' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]; }); }
  function fireChange() { listeners.forEach(function (f) { try { f(); } catch (e) {} }); }
  var TYPES = [['class', 'Class list (students)'], ['words', 'Word / spelling list'], ['names', 'Name list'], ['questions', 'Questions'], ['items', 'General items']];
  function typeLabel(t) { for (var i = 0; i < TYPES.length; i++) if (TYPES[i][0] === t) return TYPES[i][1]; return t; }

  /* ---------- markdown (de)serialise ---------- */
  function blockFor(L) {
    var attrs = 'id="' + L.id + '" type="' + (L.type || 'items') + '" title="' + (L.title || '').replace(/"/g, "'") + '"';
    if (L.type === 'class' && L.subjects && L.subjects.length) attrs += ' subjects="' + L.subjects.join(', ').replace(/"/g, "'") + '"';
    return '<!-- ct:list ' + attrs + ' -->\n## ' + (L.title || 'List') + '\n' +
      (L.items || []).map(function (s) { return '- ' + s; }).join('\n') + '\n<!-- ct:end id="' + L.id + '" -->\n';
  }
  function parseMd(text) {
    var out = [], re = /<!--\s*ct:list\s+([^>]*?)-->([\s\S]*?)<!--\s*ct:end[^>]*-->/g, m;
    while ((m = re.exec(text))) {
      var a = m[1], body = m[2];
      function attr(k) { var mm = new RegExp(k + '="([^"]*)"').exec(a); return mm ? mm[1] : ''; }
      var items = []; body.split('\n').forEach(function (ln) { var mm = /^\s*[-*]\s+(.+)$/.exec(ln); if (mm) items.push(mm[1].trim()); });
      var subj = attr('subjects');
      out.push({ id: attr('id') || uid(), type: attr('type') || 'items', title: attr('title') || 'Untitled', subjects: subj ? subj.split(',').map(function (x) { return x.trim(); }).filter(Boolean) : [], items: items });
    }
    return out;
  }
  function fileFor(type) { return type === 'class' ? 'classes.md' : (type === 'words' ? 'word-lists.md' : 'lists.md'); }

  /* ---------- localStorage mirror ---------- */
  function saveLocal() { try { localStorage.setItem(KEY, JSON.stringify({ folderName: folderName, lists: lists })); } catch (e) {} }
  function loadLocal() { try { var d = JSON.parse(localStorage.getItem(KEY)); if (d) { folderName = d.folderName || null; lists = d.lists || []; } } catch (e) {} }

  /* ---------- folder read / write ---------- */
  async function ensurePerm(handle, write) {
    if (!handle) return false;
    var opts = { mode: write ? 'readwrite' : 'read' };
    try {
      if ((await handle.queryPermission(opts)) === 'granted') return true;
      if ((await handle.requestPermission(opts)) === 'granted') return true;
    } catch (e) {}
    return false;
  }
  async function readFolder() {
    if (!dirHandle) return false;
    if (!(await ensurePerm(dirHandle, false))) return false;
    var found = [];
    try {
      for await (var entry of dirHandle.values()) {
        if (entry.kind === 'file' && /\.md$/i.test(entry.name)) {
          try { var f = await entry.getFile(); var txt = await f.text(); parseMd(txt).forEach(function (b) { found.push(b); }); } catch (e) {}
        }
      }
    } catch (e) { return false; }
    // de-dupe by id (folder is source of truth when connected)
    var seen = {}; lists = found.filter(function (L) { if (seen[L.id]) return false; seen[L.id] = 1; return true; });
    folderName = dirHandle.name; saveLocal(); fireChange(); return true;
  }
  async function writeFolder() {
    if (!dirHandle) return false;
    if (!(await ensurePerm(dirHandle, true))) return false;
    var groups = {}; lists.forEach(function (L) { var fn = fileFor(L.type); (groups[fn] = groups[fn] || []).push(L); });
    // make sure known files get cleared if now empty
    ['classes.md', 'word-lists.md', 'lists.md'].forEach(function (fn) { if (!groups[fn]) groups[fn] = []; });
    try {
      for (var fn in groups) {
        var header = '# ClassTools — ' + fn.replace('.md', '') + '\n_Auto-managed by ClassTools. Each list is wrapped in ct:list … ct:end markers._\n\n';
        var body = groups[fn].map(blockFor).join('\n');
        var fh = await dirHandle.getFileHandle(fn, { create: true });
        var w = await fh.createWritable(); await w.write(header + body); await w.close();
      }
      return true;
    } catch (e) { return false; }
  }
  function persist() { saveLocal(); writeFolder(); fireChange(); }

  async function chooseFolder() {
    if (!fsSupported) { exportBundle(); return; }
    try {
      dirHandle = await window.showDirectoryPicker({ mode: 'readwrite', id: 'classtools' });
      await idb('set', dirHandle);
      await readFolder();              // map whatever is already in the folder
      await writeFolder();             // and write current in-memory lists too
      render();
    } catch (e) { /* cancelled */ }
  }
  async function reconnect() {
    if (!fsSupported) return;
    var h = await idb('get'); if (!h) return;
    dirHandle = h; folderName = h.name;
    if (await ensurePerm(h, false)) { await readFolder(); }
    render();
  }

  /* ---------- export / import (markdown only) ---------- */
  function allMd() {
    return '# ClassTools export — ' + new Date().toLocaleString() + '\n\n' + lists.map(blockFor).join('\n');
  }
  async function exportBundle() {
    if (dirHandle && await writeFolder()) { toast('Saved to folder “' + folderName + '”'); return; }
    var blob = new Blob([allMd()], { type: 'text/markdown' });
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'classtools.md';
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(a.href);
  }
  function importBundle() {
    var inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.md,.markdown,.txt';
    inp.onchange = function () { var f = inp.files[0]; if (!f) return; var r = new FileReader();
      r.onload = function () { var add = parseMd(String(r.result || '')); var seen = {}; lists.forEach(function (L) { seen[L.id] = 1; }); add.forEach(function (L) { if (!seen[L.id]) lists.push(L); }); persist(); render(); };
      r.readAsText(f); };
    inp.click();
  }

  /* tiny toast */
  function toast(msg) {
    var t = document.createElement('div'); t.textContent = msg;
    t.style.cssText = 'position:fixed;bottom:18px;left:50%;transform:translateX(-50%);background:#241b4e;color:#fff;padding:10px 16px;border-radius:12px;z-index:3000;font-weight:700;font-size:13px;box-shadow:0 10px 30px rgba(0,0,0,.3);opacity:0;transition:opacity .2s';
    document.body.appendChild(t); requestAnimationFrame(function () { t.style.opacity = '1'; });
    setTimeout(function () { t.style.opacity = '0'; setTimeout(function () { t.remove(); }, 250); }, 1900);
  }

  /* ---------- drawer UI ---------- */
  var scrim, drawer, editingId = null, isNew = false;
  function buildChrome() {
    scrim = document.createElement('div'); scrim.className = 'ct-acc-scrim'; scrim.addEventListener('click', close);
    drawer = document.createElement('div'); drawer.className = 'ct-acc-drawer';
    document.body.appendChild(scrim); document.body.appendChild(drawer);
    var bar = document.querySelector('.topbar');
    if (bar) {
      var b = document.createElement('button'); b.className = 'icon-btn ct-acc-btn'; b.type = 'button';
      b.title = 'My classes, lists & settings'; b.innerHTML = '<i class="fas fa-user"></i>';
      b.addEventListener('click', open); bar.appendChild(b);
    }
  }
  function open() { editingId = null; isNew = false; render(); scrim.classList.add('show'); drawer.classList.add('show'); }
  function close() { scrim.classList.remove('show'); drawer.classList.remove('show'); }

  function render() {
    if (!drawer) return;
    var storage = fsSupported
      ? (folderName ? ('<i class="fas fa-circle-check" style="color:var(--ok)"></i> Connected to <b>' + esc(folderName) + '</b> — lists auto-save here as Markdown.')
        : 'Choose a folder (ideally one that syncs, e.g. Google Drive / iCloud). Lists save there as Markdown and load automatically next time.')
      : 'This browser can’t auto-save to a folder. Your lists are saved on this device; use Export / Import to move a Markdown file.';
    var h = '';
    h += '<button class="ct-acc-x" onclick="ctAccount._close()" aria-label="Close"><i class="fas fa-xmark"></i></button>';
    h += '<h2><i class="fas fa-user" style="color:var(--primary)"></i> My ClassTools</h2>';
    h += '<div class="ct-acc-note"><i class="fas fa-shield-halved"></i> Everything stays on <b>your device / your folder</b> — nothing is uploaded. Point this at a synced folder to use the same lists on any computer.</div>';
    h += '<h3>Storage</h3><div class="ct-acc-note">' + storage + '</div><div class="ct-acc-row">';
    if (fsSupported) h += '<button class="btn btn-primary" style="width:auto" onclick="ctAccount.chooseFolder()"><i class="fas fa-folder-open"></i> ' + (folderName ? 'Change folder' : 'Choose folder') + '</button>';
    if (fsSupported && folderName) h += '<button class="btn btn-soft" onclick="ctAccount.refresh()"><i class="fas fa-rotate"></i> Re-read folder</button>';
    h += '<button class="btn btn-soft" onclick="ctAccount.export()"><i class="fas fa-file-arrow-down"></i> Export .md</button>';
    h += '<button class="btn btn-soft" onclick="ctAccount.import()"><i class="fas fa-file-arrow-up"></i> Import .md</button></div>';

    h += '<h3>My lists</h3>';
    if (editingId !== null || isNew) {
      var L = isNew ? { id: '', type: 'class', title: '', subjects: [], items: [] } : (getList(editingId) || {});
      h += '<div class="ct-acc-classcard">' +
        '<label class="ct-acc-lbl">List type</label><select id="ctAccType" onchange="ctAccount._typeChange()">' +
        TYPES.map(function (t) { return '<option value="' + t[0] + '"' + (L.type === t[0] ? ' selected' : '') + '>' + t[1] + '</option>'; }).join('') + '</select>' +
        '<label class="ct-acc-lbl" style="display:block;margin-top:8px">Title</label><input type="text" id="ctAccTitle" placeholder="e.g. Grade 7 - Rizal" value="' + esc(L.title) + '">' +
        '<div id="ctAccSubjWrap" style="' + (L.type === 'class' ? '' : 'display:none') + '"><label class="ct-acc-lbl" style="display:block;margin-top:8px">Subjects (comma-separated)</label><input type="text" id="ctAccSubj" placeholder="Science, Math" value="' + esc((L.subjects || []).join(', ')) + '"></div>' +
        '<label class="ct-acc-lbl" style="display:block;margin-top:8px">Items (one per line)</label>' +
        '<textarea id="ctAccItems" placeholder="Ana&#10;Ben&#10;Carlo">' + esc((L.items || []).join('\n')) + '</textarea>' +
        '<div class="ct-acc-row"><button class="btn btn-primary" style="width:auto" onclick="ctAccount._save()"><i class="fas fa-check"></i> Save list</button>' +
        '<button class="btn btn-soft" onclick="ctAccount._cancel()">Cancel</button></div></div>';
    } else {
      if (!lists.length) h += '<div class="ct-acc-note">No lists yet. Create a class roster or any list to reuse it in every tool.</div>';
      lists.forEach(function (L) {
        h += '<div class="ct-acc-classcard"><div class="nm"><i class="fas fa-list" style="color:var(--primary)"></i> ' + esc(L.title) +
          ' <span class="ct-acc-pill">' + (L.items || []).length + ' items</span></div>' +
          '<div class="meta">' + esc(typeLabel(L.type)) + (L.subjects && L.subjects.length ? ' · ' + esc(L.subjects.join(', ')) : '') + '</div>' +
          '<div class="acts"><button class="btn btn-soft" onclick="ctAccount._edit(\'' + L.id + '\')"><i class="fas fa-pen"></i> Edit</button>' +
          '<button class="btn btn-soft" onclick="ctAccount._del(\'' + L.id + '\')"><i class="fas fa-trash"></i></button></div></div>';
      });
      h += '<button class="btn btn-soft" style="margin-top:6px" onclick="ctAccount._new()"><i class="fas fa-plus"></i> New list</button>';
    }
    drawer.innerHTML = h;
  }

  function getList(id) { for (var i = 0; i < lists.length; i++) if (lists[i].id === id) return lists[i]; return null; }

  window.ctAccount = {
    chooseFolder: chooseFolder,
    refresh: function () { readFolder().then(function (ok) { toast(ok ? 'Folder re-read.' : 'Could not read folder.'); render(); }); },
    export: exportBundle, import: importBundle,
    listLists: function (type) { return type ? lists.filter(function (L) { return L.type === type; }) : lists.slice(); },
    listTypes: function () { var s = {}; lists.forEach(function (L) { s[L.type] = 1; }); return Object.keys(s); },
    getList: getList,
    onChange: function (fn) { listeners.push(fn); },
    open: open, _close: close,
    _new: function () { isNew = true; editingId = null; render(); },
    _edit: function (id) { editingId = id; isNew = false; render(); },
    _cancel: function () { editingId = null; isNew = false; render(); },
    _typeChange: function () { var t = document.getElementById('ctAccType').value; var w = document.getElementById('ctAccSubjWrap'); if (w) w.style.display = (t === 'class') ? '' : 'none'; },
    _del: function (id) { lists = lists.filter(function (L) { return L.id !== id; }); persist(); render(); },
    _save: function () {
      var type = document.getElementById('ctAccType').value;
      var title = (document.getElementById('ctAccTitle').value || '').trim();
      if (!title) { document.getElementById('ctAccTitle').focus(); return; }
      var subj = (document.getElementById('ctAccSubj') ? document.getElementById('ctAccSubj').value : '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
      var items = (document.getElementById('ctAccItems').value || '').split('\n').map(function (x) { return x.trim(); }).filter(Boolean);
      if (isNew) lists.push({ id: uid(), type: type, title: title, subjects: subj, items: items });
      else { var L = getList(editingId); if (L) { L.type = type; L.title = title; L.subjects = subj; L.items = items; } }
      persist(); toast(dirHandle ? 'Saved to folder “' + folderName + '”' : 'Saved on this device'); editingId = null; isNew = false; render(); buildListPickers();
    },
    // write any file into the connected folder (used by tools), markdown
    saveFile: async function (name, content) {
      if (dirHandle && await ensurePerm(dirHandle, true)) {
        try { var fh = await dirHandle.getFileHandle(name, { create: true }); var w = await fh.createWritable(); await w.write(content); await w.close(); toast('Saved ' + name + ' to folder'); return true; } catch (e) {}
      }
      var blob = new Blob([content], { type: 'text/markdown' });
      var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(a.href);
      return false;
    },
    // save a named list straight from a tool
    saveList: function (type, title, items) {
      var L = { id: uid(), type: type || 'items', title: title || 'List', subjects: [], items: items || [] };
      lists.push(L); persist(); buildListPickers(); toast('List “' + title + '” saved'); return L.id;
    }
  };

  /* ---------- "Load a list" control for tools ----------
     <textarea data-ct-list-target="#sel" [data-ct-list-type="class"]> gets a
     [type ▾][list ▾] picker inserted above it; choosing fills the target.   */
  function buildListPickers() {
    var targets = document.querySelectorAll('[data-ct-list-target],[data-ct-class-target]');
    targets.forEach(function (host) {
      var lockType = host.getAttribute('data-ct-class-target') ? 'class' : (host.getAttribute('data-ct-list-type') || '');
      var prev = host.previousElementSibling;
      var ctrl = (prev && prev.classList && prev.classList.contains('ct-classpick')) ? prev : null;
      if (!ctrl) {
        ctrl = document.createElement('div'); ctrl.className = 'ct-classpick';
        ctrl.innerHTML = '<i class="fas fa-folder-tree" style="color:var(--muted)"></i>' +
          (lockType ? '' : '<select class="ct-lp-type"></select>') + '<select class="ct-lp-list"></select>';
        host.parentNode.insertBefore(ctrl, host);
        var listSel = ctrl.querySelector('.ct-lp-list');
        listSel.addEventListener('change', function () {
          var L = getList(this.value); if (!L) return;
          var tgt = document.querySelector(host.getAttribute('data-ct-list-target') || host.getAttribute('data-ct-class-target')) || host;
          if (tgt && 'value' in tgt) { tgt.value = (L.items || []).join('\n'); tgt.dispatchEvent(new Event('input', { bubbles: true })); }
        });
        var typeSel = ctrl.querySelector('.ct-lp-type');
        if (typeSel) typeSel.addEventListener('change', function () { fillListOptions(ctrl, this.value); });
      }
      var lt = ctrl.querySelector('.ct-lp-type');
      if (lt) { lt.innerHTML = TYPES.map(function (t) { return '<option value="' + t[0] + '">' + t[1] + '</option>'; }).join(''); }
      fillListOptions(ctrl, lockType || (lt ? lt.value : ''));
    });
  }
  function fillListOptions(ctrl, type) {
    var listSel = ctrl.querySelector('.ct-lp-list');
    var rel = lists.filter(function (L) { return !type || L.type === type; });
    listSel.innerHTML = '<option value="">' + (rel.length ? 'Load a saved list…' : 'No saved lists of this type') + '</option>' +
      rel.map(function (L) { return '<option value="' + L.id + '">' + esc(L.title) + ' (' + (L.items || []).length + ')</option>'; }).join('');
  }
  listeners.push(buildListPickers);

  function boot() { loadLocal(); buildChrome(); render(); buildListPickers(); reconnect(); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
