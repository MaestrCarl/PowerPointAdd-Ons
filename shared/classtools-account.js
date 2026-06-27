/* ============================================================
   ClassTools — Teacher Account / Storage  (items 2,3,4,5,6,10,11,19)
   Obsidian-like local storage with a planned naming + tagging scheme.

   FOLDER LAYOUT (inside the folder the teacher picks):
     <folder>/Lists/            reusable lists (class rosters, word lists…)
     <folder>/<ToolName>/       per-tool saved setups (auto-created on save)

   FILE NAMING (so the app & a human can identify a file at a glance):
     ct__<kind>__<scope>__<type>__<slug>__<id>__<YYYY-MM-DD>.md
       kind  = list | setup
       scope = any | tool ids joined by + (e.g. wordgames+picker)  -> TAGGING
       type  = list type (class/words/…) or setup kind (hangman/groups…)
       slug  = kebab-case title
       id    = stable id (older dated files with same id are replaced)
       date  = last-updated (lets the app pick the newest version)

   FILE CONTENT (each file holds one or more delimited blocks):
     <!-- ct:list id=".." type=".." title=".." tools=".." subjects=".." updated=".." -->
     ## Title
     - item
     <!-- ct:end id=".." -->

   The app scans EVERY .md under the folder (recursively), parses blocks,
   de-dupes by id keeping the newest `updated`. localStorage mirrors all so
   it still works with no folder (PowerPoint add-in / unsupported browsers).
   ============================================================ */
(function () {
  "use strict";
  if (window.__ctAccount) return; window.__ctAccount = true;

  var KEY = 'ct-account-v3';
  var fsSupported = (typeof window.showDirectoryPicker === 'function');
  var dirHandle = null, folderName = null;
  var lists = [];           // [{id,type,title,tools:[],subjects:[],items:[],updated}]
  var listeners = [];

  /* ---- which tool are we in (for tagging / filtered loading) ---- */
  var TOOL = (function () {
    var p = location.pathname.toLowerCase();
    if (p.indexOf('wordgames') >= 0) return 'wordgames';
    if (p.indexOf('picker') >= 0 || p.indexOf('randompicker') >= 0) return 'picker';
    if (p.indexOf('timer') >= 0) return 'timer';
    if (p.indexOf('scoreboard') >= 0) return 'scoreboard';
    if (p.indexOf('intro') >= 0) return 'intro';
    if (p.indexOf('splitter') >= 0) return 'splitter';
    if (p.indexOf('spelling') >= 0) return 'spelling';
    return 'hub';
  })();
  var TOOLS_LIST = [['picker', 'Pickers & Groups'], ['timer', 'Timer Tools'], ['intro', 'Intro & Reveal'], ['wordgames', 'Word Games'], ['spelling', 'Spelling Helper'], ['splitter', 'Word Splitter'], ['scoreboard', 'Scoreboard']];
  var TOOL_FOLDER = { wordgames: 'WordGames', picker: 'Pickers', timer: 'Timer', scoreboard: 'Scoreboard', intro: 'IntroTools', splitter: 'WordSplitter', spelling: 'SpellingHelper', hub: 'ClassTools' };
  var TYPES = [['class', 'Class list (students)'], ['words', 'Word / spelling list'], ['names', 'Name list'], ['questions', 'Questions'], ['items', 'General items']];
  function typeLabel(t) { for (var i = 0; i < TYPES.length; i++) if (TYPES[i][0] === t) return TYPES[i][1]; return t; }

  /* ---- helpers ---- */
  function uid() { return 'L' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]; }); }
  function slug(s) { return String(s || 'untitled').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 48) || 'untitled'; }
  function today() { var d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
  function fireChange() { listeners.forEach(function (f) { try { f(); } catch (e) {} }); }
  function scopeStr(L) { return (L.tools && L.tools.length) ? L.tools.join('+') : 'any'; }

  /* ---- markdown block (de)serialise ---- */
  function blockFor(L) {
    var a = 'id="' + L.id + '" type="' + (L.type || 'items') + '" title="' + (L.title || '').replace(/"/g, "'") +
      '" tools="' + scopeStr(L) + '" updated="' + (L.updated || today()) + '"';
    if (L.type === 'class' && L.subjects && L.subjects.length) a += ' subjects="' + L.subjects.join(', ').replace(/"/g, "'") + '"';
    return '<!-- ct:list ' + a + ' -->\n## ' + (L.title || 'List') + '\n' +
      (L.items || []).map(function (s) { return '- ' + s; }).join('\n') + '\n<!-- ct:end id="' + L.id + '" -->\n';
  }
  function fileName(L) { return 'ct__list__' + scopeStr(L) + '__' + (L.type || 'items') + '__' + slug(L.title) + '__' + L.id + '__' + (L.updated || today()) + '.md'; }
  function parseMd(text) {
    var out = [], re = /<!--\s*ct:list\s+([^>]*?)-->([\s\S]*?)<!--\s*ct:end[^>]*-->/g, m;
    while ((m = re.exec(text))) {
      var a = m[1], body = m[2];
      function at(k) { var x = new RegExp(k + '="([^"]*)"').exec(a); return x ? x[1] : ''; }
      var items = []; body.split('\n').forEach(function (ln) { var x = /^\s*[-*]\s+(.+)$/.exec(ln); if (x) items.push(x[1].trim()); });
      var tools = at('tools'); var subj = at('subjects');
      out.push({
        id: at('id') || uid(), type: at('type') || 'items', title: at('title') || 'Untitled',
        tools: (tools && tools !== 'any') ? tools.split('+').map(function (x) { return x.trim(); }).filter(Boolean) : [],
        subjects: subj ? subj.split(',').map(function (x) { return x.trim(); }).filter(Boolean) : [],
        items: items, updated: at('updated') || ''
      });
    }
    return out;
  }

  /* ---- localStorage mirror ---- */
  function saveLocal() { try { localStorage.setItem(KEY, JSON.stringify({ folderName: folderName, lists: lists })); } catch (e) {} }
  function loadLocal() { try { var d = JSON.parse(localStorage.getItem(KEY)); if (d) { folderName = d.folderName || null; lists = (d.lists || []).map(norm); } } catch (e) {} }
  function norm(L) { L.tools = L.tools || []; L.subjects = L.subjects || []; L.items = L.items || []; L.updated = L.updated || today(); return L; }

  /* ---- IndexedDB: remember the folder handle across sessions ---- */
  function idb(op, val) {
    return new Promise(function (res) {
      var r; try { r = indexedDB.open('ct-store', 1); } catch (e) { return res(null); }
      r.onupgradeneeded = function () { r.result.createObjectStore('kv'); };
      r.onsuccess = function () {
        try { var tx = r.result.transaction('kv', op === 'get' ? 'readonly' : 'readwrite'), st = tx.objectStore('kv');
          var q = op === 'get' ? st.get('dir') : st.put(val, 'dir');
          q.onsuccess = function () { res(op === 'get' ? q.result : true); }; q.onerror = function () { res(null); };
        } catch (e) { res(null); }
      };
      r.onerror = function () { res(null); };
    });
  }
  async function ensurePerm(h, write) {
    if (!h) return false; var o = { mode: write ? 'readwrite' : 'read' };
    try { if ((await h.queryPermission(o)) === 'granted') return true; if ((await h.requestPermission(o)) === 'granted') return true; } catch (e) {}
    return false;
  }

  /* ---- recursive scan: read every .md, keep newest per id ---- */
  async function walk(dir, depth, cb) {
    if (depth > 4) return;
    try { for await (var e of dir.values()) {
      if (e.kind === 'file' && /\.md$/i.test(e.name)) { try { var f = await e.getFile(); cb(await f.text()); } catch (ex) {} }
      else if (e.kind === 'directory') { await walk(e, depth + 1, cb); }
    } } catch (ex) {}
  }
  async function scanFolder() {
    if (!dirHandle || !(await ensurePerm(dirHandle, false))) return false;
    var found = {}; // id -> list (newest)
    await walk(dirHandle, 0, function (txt) {
      parseMd(txt).forEach(function (b) { var ex = found[b.id]; if (!ex || (b.updated || '') >= (ex.updated || '')) found[b.id] = b; });
    });
    var arr = Object.keys(found).map(function (k) { return found[k]; });
    if (arr.length || lists.length === 0) lists = arr.map(norm);  // folder is source of truth
    folderName = dirHandle.name; saveLocal(); fireChange(); return true;
  }

  /* ---- write a single list file (replacing older versions of same id) ---- */
  async function getListsDir() { return await dirHandle.getDirectoryHandle('Lists', { create: true }); }
  async function removeOld(dir, id) {
    try { var del = []; for await (var e of dir.values()) { if (e.kind === 'file' && e.name.indexOf('__' + id + '__') >= 0) del.push(e.name); }
      for (var i = 0; i < del.length; i++) { try { await dir.removeEntry(del[i]); } catch (ex) {} } } catch (ex) {}
  }
  async function writeList(L) {
    if (!dirHandle || !(await ensurePerm(dirHandle, true))) return false;
    try { var dir = await getListsDir(); await removeOld(dir, L.id);
      var fh = await dir.getFileHandle(fileName(L), { create: true }); var w = await fh.createWritable(); await w.write(blockFor(L)); await w.close(); return true;
    } catch (e) { return false; }
  }
  async function deleteListFile(id) {
    if (!dirHandle) return; try { await removeOld(await getListsDir(), id); } catch (e) {}
  }

  function persist(changed) { saveLocal(); if (changed) writeList(changed); fireChange(); }

  /* ---- choose / reconnect folder ---- */
  async function chooseFolder() {
    if (!fsSupported) { exportBundle(); return; }
    try { dirHandle = await window.showDirectoryPicker({ mode: 'readwrite', id: 'classtools' });
      await idb('set', dirHandle);
      await scanFolder();                       // map what's already there
      for (var i = 0; i < lists.length; i++) await writeList(lists[i]); // and persist current
      toast('Folder “' + dirHandle.name + '” connected'); render();
    } catch (e) { /* cancelled */ }
  }
  async function reconnect() {
    if (!fsSupported) return; var h = await idb('get'); if (!h) return;
    dirHandle = h; folderName = h.name;
    if (await ensurePerm(h, false)) await scanFolder();
    render();
  }

  /* ---- export / import (markdown) ---- */
  function allMd() { return '# ClassTools export — ' + new Date().toLocaleString() + '\n\n' + lists.map(blockFor).join('\n'); }
  async function exportBundle() {
    if (dirHandle) { var ok = true; for (var i = 0; i < lists.length; i++) ok = (await writeList(lists[i])) && ok; if (ok) { toast('Saved to “' + folderName + '/Lists”'); return; } }
    download('classtools-lists.md', allMd());
  }
  function importBundle() {
    var inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.md,.markdown,.txt';
    inp.onchange = function () { var f = inp.files[0]; if (!f) return; var r = new FileReader();
      r.onload = function () { var add = parseMd(String(r.result || '')); var seen = {}; lists.forEach(function (L) { seen[L.id] = 1; });
        add.forEach(function (L) { norm(L); if (!seen[L.id]) { lists.push(L); writeList(L); } }); saveLocal(); fireChange(); render(); };
      r.readAsText(f); };
    inp.click();
  }
  function download(name, content) { var b = new Blob([content], { type: 'text/markdown' }); var a = document.createElement('a'); a.href = URL.createObjectURL(b); a.download = name; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(a.href); }
  function toast(msg) {
    var t = document.createElement('div'); t.textContent = msg;
    t.style.cssText = 'position:fixed;bottom:18px;left:50%;transform:translateX(-50%);background:#241b4e;color:#fff;padding:10px 16px;border-radius:12px;z-index:3000;font-weight:700;font-size:13px;box-shadow:0 10px 30px rgba(0,0,0,.3);opacity:0;transition:opacity .2s';
    document.body.appendChild(t); requestAnimationFrame(function () { t.style.opacity = '1'; }); setTimeout(function () { t.style.opacity = '0'; setTimeout(function () { t.remove(); }, 250); }, 2000);
  }

  /* ---- drawer ---- */
  var scrim, drawer, editingId = null, isNew = false;
  function buildChrome() {
    scrim = document.createElement('div'); scrim.className = 'ct-acc-scrim'; scrim.addEventListener('click', closeDr);
    drawer = document.createElement('div'); drawer.className = 'ct-acc-drawer'; document.body.appendChild(scrim); document.body.appendChild(drawer);
    var bar = document.querySelector('.topbar');
    if (bar) { var b = document.createElement('button'); b.className = 'icon-btn ct-acc-btn'; b.type = 'button'; b.title = 'My classes, lists & settings'; b.innerHTML = '<i class="fas fa-user"></i>'; b.addEventListener('click', openDr); bar.appendChild(b); }
  }
  function openDr() { editingId = null; isNew = false; render(); scrim.classList.add('show'); drawer.classList.add('show'); }
  function closeDr() { scrim.classList.remove('show'); drawer.classList.remove('show'); }

  function render() {
    if (!drawer) return;
    var storage = fsSupported
      ? (folderName ? ('<i class="fas fa-circle-check" style="color:var(--ok)"></i> Connected to <b>' + esc(folderName) + '</b>. Lists live in <b>' + esc(folderName) + '/Lists</b> as Markdown and load automatically.')
        : 'Pick a folder (ideally one that syncs — Drive / iCloud / OneDrive). Lists save there as Markdown and reload automatically next time.')
      : 'This browser can’t auto-save to a folder. Lists are kept on this device; use Export / Import to move a Markdown file.';
    var h = '';
    h += '<button class="ct-acc-x" onclick="ctAccount._close()" aria-label="Close"><i class="fas fa-xmark"></i></button>';
    h += '<h2><i class="fas fa-user" style="color:var(--primary)"></i> My ClassTools</h2>';
    h += '<div class="ct-acc-note"><i class="fas fa-shield-halved"></i> Everything stays on <b>your device / your folder</b> — nothing is uploaded. Point this at a synced folder to use the same lists on any computer.</div>';
    h += '<h3>Storage</h3><div class="ct-acc-note">' + storage + '</div><div class="ct-acc-row">';
    if (fsSupported) h += '<button class="btn btn-primary" style="width:auto" onclick="ctAccount.chooseFolder()"><i class="fas fa-folder-open"></i> ' + (folderName ? 'Change folder' : 'Choose folder') + '</button>';
    if (fsSupported) h += '<button class="btn btn-soft" onclick="ctAccount.scan()"><i class="fas fa-magnifying-glass"></i> Scan folder</button>';
    h += '<button class="btn btn-soft" onclick="ctAccount.export()"><i class="fas fa-file-arrow-down"></i> Export .md</button>';
    h += '<button class="btn btn-soft" onclick="ctAccount.import()"><i class="fas fa-file-arrow-up"></i> Import .md</button></div>';

    h += '<h3>My lists</h3>';
    if (editingId !== null || isNew) {
      var L = isNew ? { id: '', type: 'class', title: '', tools: [], subjects: [], items: [] } : (getList(editingId) || {});
      var thisOnly = L.tools && L.tools.length === 1 && L.tools[0] === TOOL;
      h += '<div class="ct-acc-classcard">' +
        '<label class="ct-acc-lbl">List type</label><select id="ctAccType" onchange="ctAccount._typeChange()">' + TYPES.map(function (t) { return '<option value="' + t[0] + '"' + (L.type === t[0] ? ' selected' : '') + '>' + t[1] + '</option>'; }).join('') + '</select>' +
        '<label class="ct-acc-lbl" style="display:block;margin-top:8px">Title</label><input type="text" id="ctAccTitle" placeholder="e.g. Grade 7 - Rizal" value="' + esc(L.title) + '">' +
        '<div id="ctAccSubjWrap" style="' + (L.type === 'class' ? '' : 'display:none') + '"><label class="ct-acc-lbl" style="display:block;margin-top:8px">Subjects (comma-separated)</label><input type="text" id="ctAccSubj" placeholder="Science, Math" value="' + esc((L.subjects || []).join(', ')) + '"></div>' +
        (TOOL !== 'hub' ? '<label class="chk" style="margin-top:8px;display:flex;gap:7px;align-items:center;color:var(--muted);font-weight:700;font-size:13px"><input type="checkbox" id="ctAccThisTool"' + (thisOnly ? ' checked' : '') + '> Tag to this tool only (otherwise available everywhere)</label>' : '') +
        '<label class="ct-acc-lbl" style="display:block;margin-top:8px">Items (one per line)</label><textarea id="ctAccItems" placeholder="Ana&#10;Ben&#10;Carlo">' + esc((L.items || []).join('\n')) + '</textarea>' +
        '<div class="ct-acc-row"><button class="btn btn-primary" style="width:auto" onclick="ctAccount._save()"><i class="fas fa-check"></i> Save list</button><button class="btn btn-soft" onclick="ctAccount._cancel()">Cancel</button></div></div>';
    } else {
      if (!lists.length) h += '<div class="ct-acc-note">No lists yet. Create a class roster or any list to reuse it in every tool.</div>';
      lists.forEach(function (L) {
        var scope = (L.tools && L.tools.length) ? L.tools.map(function (t) { var n = TOOLS_LIST.filter(function (x) { return x[0] === t; })[0]; return n ? n[1] : t; }).join(', ') : 'All tools';
        h += '<div class="ct-acc-classcard"><div class="nm"><i class="fas fa-list" style="color:var(--primary)"></i> ' + esc(L.title) + ' <span class="ct-acc-pill">' + (L.items || []).length + ' items</span></div>' +
          '<div class="meta">' + esc(typeLabel(L.type)) + ' · ' + esc(scope) + (L.updated ? ' · ' + esc(L.updated) : '') + (L.subjects && L.subjects.length ? ' · ' + esc(L.subjects.join(', ')) : '') + '</div>' +
          '<div class="acts"><button class="btn btn-soft" onclick="ctAccount._edit(\'' + L.id + '\')"><i class="fas fa-pen"></i> Edit</button><button class="btn btn-soft" onclick="ctAccount._del(\'' + L.id + '\')"><i class="fas fa-trash"></i></button></div></div>';
      });
      h += '<button class="btn btn-soft" style="margin-top:6px" onclick="ctAccount._new()"><i class="fas fa-plus"></i> New list</button>';
    }
    drawer.innerHTML = h;
  }
  function getList(id) { for (var i = 0; i < lists.length; i++) if (lists[i].id === id) return lists[i]; return null; }

  window.ctAccount = {
    tool: TOOL,
    chooseFolder: chooseFolder,
    scan: function () { if (!dirHandle) { chooseFolder(); return; } scanFolder().then(function (ok) { toast(ok ? 'Scanned — ' + lists.length + ' list(s) found.' : 'Could not read folder. Try “Change folder”.'); render(); }); },
    export: exportBundle, import: importBundle,
    listLists: function (type, forTool) { return lists.filter(function (L) { return (!type || L.type === type) && (!forTool || !L.tools.length || L.tools.indexOf(forTool) >= 0); }); },
    getList: getList, onChange: function (fn) { listeners.push(fn); },
    open: openDr, _close: closeDr,
    _new: function () { isNew = true; editingId = null; render(); },
    _edit: function (id) { editingId = id; isNew = false; render(); },
    _cancel: function () { editingId = null; isNew = false; render(); },
    _typeChange: function () { var t = document.getElementById('ctAccType').value, w = document.getElementById('ctAccSubjWrap'); if (w) w.style.display = (t === 'class') ? '' : 'none'; },
    _del: function (id) { lists = lists.filter(function (L) { return L.id !== id; }); deleteListFile(id); saveLocal(); fireChange(); render(); },
    _save: function () {
      var type = document.getElementById('ctAccType').value, title = (document.getElementById('ctAccTitle').value || '').trim();
      if (!title) { document.getElementById('ctAccTitle').focus(); return; }
      var subj = (document.getElementById('ctAccSubj') ? document.getElementById('ctAccSubj').value : '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
      var items = (document.getElementById('ctAccItems').value || '').split('\n').map(function (x) { return x.trim(); }).filter(Boolean);
      var thisTool = document.getElementById('ctAccThisTool'); var tools = (thisTool && thisTool.checked && TOOL !== 'hub') ? [TOOL] : [];
      var L;
      if (isNew) { L = norm({ id: uid(), type: type, title: title, tools: tools, subjects: subj, items: items, updated: today() }); lists.push(L); }
      else { L = getList(editingId); if (L) { L.type = type; L.title = title; L.tools = tools; L.subjects = subj; L.items = items; L.updated = today(); } }
      persist(L); toast(dirHandle ? 'Saved to “' + folderName + '/Lists”' : 'Saved on this device'); editingId = null; isNew = false; render(); buildListPickers();
    },
    /* save a list straight from a tool (tagged to this tool) */
    saveList: function (type, title, items, allTools) { var L = norm({ id: uid(), type: type || 'items', title: title || 'List', tools: allTools ? [] : (TOOL !== 'hub' ? [TOOL] : []), subjects: [], items: items || [], updated: today() }); lists.push(L); persist(L); buildListPickers(); toast('List “' + title + '” saved'); return L.id; },
    /* save a tool setup file into <folder>/<ToolName>/ (auto-creates folder) */
    saveSetup: async function (kind, title, content) {
      var fn = 'ct__setup__' + TOOL + '__' + slug(kind) + '__' + slug(title) + '__' + uid() + '__' + today() + '.md';
      if (dirHandle && await ensurePerm(dirHandle, true)) {
        try { var d = await dirHandle.getDirectoryHandle(TOOL_FOLDER[TOOL] || 'ClassTools', { create: true });
          var fh = await d.getFileHandle(fn, { create: true }); var w = await fh.createWritable(); await w.write(content); await w.close(); toast('Saved to “' + folderName + '/' + (TOOL_FOLDER[TOOL] || 'ClassTools') + '”'); return true; } catch (e) {}
      }
      download(fn, content); return false;
    }
  };

  /* ---- "Load a list" control injected into tools ---- */
  function buildListPickers() {
    var targets = document.querySelectorAll('[data-ct-list-target],[data-ct-class-target]');
    targets.forEach(function (host) {
      var lockType = host.getAttribute('data-ct-class-target') ? 'class' : (host.getAttribute('data-ct-list-type') || '');
      var prev = host.previousElementSibling;
      var ctrl = (prev && prev.classList && prev.classList.contains('ct-classpick')) ? prev : null;
      if (!ctrl) {
        ctrl = document.createElement('div'); ctrl.className = 'ct-classpick';
        ctrl.innerHTML = '<i class="fas fa-folder-tree" style="color:var(--muted)"></i>' + (lockType ? '' : '<select class="ct-lp-type"></select>') + '<select class="ct-lp-list"></select>';
        host.parentNode.insertBefore(ctrl, host);
        ctrl.querySelector('.ct-lp-list').addEventListener('change', function () {
          var L = getList(this.value); if (!L) return;
          var tgt = document.querySelector(host.getAttribute('data-ct-list-target') || host.getAttribute('data-ct-class-target')) || host;
          if (tgt && 'value' in tgt) { tgt.value = (L.items || []).join('\n'); tgt.dispatchEvent(new Event('input', { bubbles: true })); }
        });
        var ts = ctrl.querySelector('.ct-lp-type'); if (ts) ts.addEventListener('change', function () { fillOptions(ctrl, this.value, lockType); });
      }
      var lt = ctrl.querySelector('.ct-lp-type');
      if (lt) lt.innerHTML = '<option value="">Any type</option>' + TYPES.map(function (t) { return '<option value="' + t[0] + '">' + t[1] + '</option>'; }).join('');
      fillOptions(ctrl, lockType || (lt ? lt.value : ''), lockType);
    });
  }
  function fillOptions(ctrl, type, lockType) {
    var sel = ctrl.querySelector('.ct-lp-list');
    var rel = window.ctAccount.listLists(type || lockType || '', TOOL === 'hub' ? '' : TOOL);
    sel.innerHTML = '<option value="">' + (rel.length ? 'Load a saved list…' : 'No saved lists yet') + '</option>' +
      rel.map(function (L) { return '<option value="' + L.id + '">' + esc(L.title) + ' (' + (L.items || []).length + ')</option>'; }).join('');
  }
  listeners.push(buildListPickers);

  function boot() { loadLocal(); buildChrome(); render(); buildListPickers(); reconnect(); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
