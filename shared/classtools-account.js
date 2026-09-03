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
  var TYPES = [['class', 'Class list (students)'], ['words', 'Word / spelling list'], ['questions', 'Questions'], ['items', 'General items'], ['media', 'Media (images/links)']];
  /* Lists that are always there, whether or not a teacher has saved anything.
     Kept OUT of the `lists` array on purpose: that array is what gets written to
     the folder, exported and deleted, and a built-in has no business in any of
     those. It is merged in only where lists are OFFERED — listLists and getList
     — so save, scan, export and delete carry on seeing only real files.

     Numbers exists because a roster is not always names. Seat numbers, jersey
     numbers, a class whose list you do not have to hand: typing 1 to 30 into a
     textarea is a minute of nobody's time well spent. */
  var BUILTIN = [{
    id: 'ct:builtin:numbers-30',
    title: 'Numbers',
    /* Offered under every type that means "a list of things with names on
       them". A roster of thirty, thirty items to sort, thirty questions to
       hand out — the numbers serve all of them, and a teacher who has filtered
       the picker to "General items" should not be told there is nothing
       ready-made. Not `media`: a number is not a picture or a link. */
    type: 'class',
    types: ['class', 'items', 'words', 'questions'],
    tools: [],                      // empty means every tool
    builtin: true,
    items: (function () { var a = []; for (var i = 1; i <= 30; i++) a.push(String(i)); return a; })()
  }];
  function builtinFor(type, forTool) {
    return BUILTIN.filter(function (L) {
      var types = L.types || [L.type];
      return (!type || types.indexOf(type) >= 0) &&
             (!forTool || !L.tools.length || L.tools.indexOf(forTool) >= 0);
    });
  }

  var profile = { name: '', subjects: [], photo: '' };
  var settings = { defaultDark: false };
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
  function saveLocal() { try { localStorage.setItem(KEY, JSON.stringify({ folderName: folderName, lists: lists, profile: profile, settings: settings })); } catch (e) {} }
  function loadLocal() { try { var d = JSON.parse(localStorage.getItem(KEY)); if (d) { folderName = d.folderName || null; lists = (d.lists || []).map(norm); profile = d.profile || profile; settings = d.settings || settings; } } catch (e) {} }
  async function writeAccount() {
    if (!dirHandle || !(await ensurePerm(dirHandle, true))) return false;
    try {
      var dir = await dirHandle.getDirectoryHandle('Account', { create: true });
      var md = '---\nname: ' + (profile.name || '') + '\nsubjects: ' + (profile.subjects || []).join(', ') + '\nphoto: ' + (profile.photo ? 'profile-photo' : '') + '\ndefaultDark: ' + (!!settings.defaultDark) + '\nupdated: ' + today() + '\n---\n\n# ' + (profile.name || 'My ClassTools account') + '\n\nSubjects: ' + (profile.subjects || []).join(', ') + '\n';
      var fh = await dir.getFileHandle('account.md', { create: true }); var w = await fh.createWritable(); await w.write(md); await w.close();
      if (profile.photo && profile.photo.indexOf('data:') === 0) {
        try { var ext = (profile.photo.match(/^data:image\/(\w+)/) || [, 'png'])[1]; var blob = await (await fetch(profile.photo)).blob();
          var pf = await dir.getFileHandle('profile-photo.' + ext, { create: true }); var pw = await pf.createWritable(); await pw.write(blob); await pw.close(); } catch (e) {}
      }
      return true;
    } catch (e) { return false; }
  }
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

  /* ---- class → subject folder tree + activity saving (item 7) ---- */
  function safeName(s) { return String(s || '').replace(/[\/\\:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim().slice(0, 60) || 'Untitled'; }
  async function ensureDir(parts) { var d = dirHandle; for (var i = 0; i < parts.length; i++) { d = await d.getDirectoryHandle(safeName(parts[i]), { create: true }); } return d; }
  function classes() { return lists.filter(function (L) { return L.type === 'class'; }); }
  function subjectsFor(name) { var c = classes().filter(function (L) { return L.title === name; })[0]; return c ? (c.subjects || []) : []; }
  async function ensureClassFolders() {
    if (!dirHandle || !(await ensurePerm(dirHandle, true))) return;
    var cs = classes();
    for (var i = 0; i < cs.length; i++) { try { await ensureDir([cs[i].title]); var subs = cs[i].subjects || []; for (var j = 0; j < subs.length; j++) await ensureDir([cs[i].title, subs[j]]); } catch (e) {} }
  }
  /* activity index (so tools can offer "load previous activity") */
  var actIndex = []; try { actIndex = JSON.parse(localStorage.getItem('ct-activities') || '[]'); } catch (e) { actIndex = []; }
  function saveActIndex() { try { localStorage.setItem('ct-activities', JSON.stringify(actIndex.slice(-400))); } catch (e) {} }
  function activityMd(meta, content) {
    return '---\ntool: ' + TOOL + '\nkind: ' + (meta.kind || '') + '\ntitle: ' + (meta.title || '') + '\nclass: ' + (meta.className || '') + '\nsubject: ' + (meta.subject || '') + '\ngrade: ' + (meta.grade || '') + '\nsection: ' + (meta.section || '') + '\nupdated: ' + today() + '\n---\n\n' + content + '\n';
  }
  async function saveActivity(meta, content) {
    meta = meta || {}; var fn = 'ct__setup__' + TOOL + '__' + slug(meta.kind || 'activity') + '__' + slug(meta.title || 'activity') + '__' + uid() + '__' + today() + '.md';
    var md = activityMd(meta, content);
    if (!dirHandle || !(await ensurePerm(dirHandle, true))) { download(fn, md); return false; }
    if (!meta.className) { openChooser(meta, content); return false; }   // ask where to save
    var parts = [meta.className]; if (meta.subject) parts.push(meta.subject);
    try {
      var d = await ensureDir(parts); var fh = await d.getFileHandle(fn, { create: true }); var w = await fh.createWritable(); await w.write(md); await w.close();
      actIndex.push({ tool: TOOL, kind: meta.kind || '', title: meta.title || 'Activity', className: meta.className, subject: meta.subject || '', path: parts.concat(fn), date: today() }); saveActIndex();
      toast('Saved to “' + [folderName].concat(parts).join('/') + '”'); return true;
    } catch (e) { download(fn, md); return false; }
  }
  function listActivities(tool) { return actIndex.filter(function (a) { return !tool || a.tool === tool; }).slice().reverse(); }
  async function readActivity(entry) {
    if (!dirHandle || !entry || !entry.path) return null;
    try { var d = dirHandle; for (var i = 0; i < entry.path.length - 1; i++) d = await d.getDirectoryHandle(safeName(entry.path[i])); var fh = await d.getFileHandle(entry.path[entry.path.length - 1]); var f = await fh.getFile(); var t = await f.text(); return t.replace(/^---[\s\S]*?---\n+/, ''); } catch (e) { return null; }
  }

  /* ---- "where do you want to save?" chooser modal ---- */
  var chooserEl = null;
  function openChooser(meta, content) {
    if (!chooserEl) { chooserEl = document.createElement('div'); chooserEl.className = 'ct-acc-scrim'; chooserEl.style.zIndex = 3200; document.body.appendChild(chooserEl); }
    var cs = classes();
    var opts = cs.map(function (c) { return '<option value="' + esc(c.title) + '">' + esc(c.title) + '</option>'; }).join('');
    chooserEl.innerHTML = '<div class="ct-acc-drawer" style="position:fixed;left:50%;top:50%;transform:translate(-50%,-50%);right:auto;bottom:auto;width:min(420px,94vw);border-radius:18px;border:1px solid var(--line)">' +
      '<h2><i class="fas fa-folder-tree" style="color:var(--primary)"></i> Save activity</h2>' +
      '<div class="ct-acc-note">Choose where to file “' + esc(meta.title || 'this activity') + '”. It will be saved as Markdown under the class &amp; subject folder.</div>' +
      '<label class="ct-acc-lbl">Class</label><select id="ctChClass">' + (cs.length ? opts : '') + '<option value="">— General (no class) —</option></select>' +
      '<label class="ct-acc-lbl" style="display:block;margin-top:8px">Subject</label><input type="text" id="ctChSubj" placeholder="e.g. Science" value="' + esc(meta.subject || '') + '">' +
      '<div class="ct-acc-row" style="margin-top:12px"><button class="btn btn-primary" id="ctChSave"><i class="fas fa-check"></i> Save here</button><button class="btn btn-soft" id="ctChCancel">Cancel</button></div></div>';
    chooserEl.style.display = 'block'; chooserEl.classList.add('show');
    chooserEl.querySelector('#ctChCancel').onclick = function () { chooserEl.classList.remove('show'); chooserEl.style.display = 'none'; };
    chooserEl.querySelector('#ctChSave').onclick = function () {
      meta.className = chooserEl.querySelector('#ctChClass').value || 'General'; meta.subject = chooserEl.querySelector('#ctChSubj').value.trim();
      chooserEl.classList.remove('show'); chooserEl.style.display = 'none'; saveActivity(meta, content);
    };
  }

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
    var h = '';
    h += '<button class="ct-acc-x" onclick="ctAccount._close()" aria-label="Close"><i class="fas fa-xmark"></i></button>';
    h += '<h2><i class="fas fa-user" style="color:var(--primary)"></i> My ClassTools</h2>';

    /* ---- profile ---- */
    h += '<div class="ct-acc-profile">' +
      (profile.photo ? '<img class="ct-acc-photo" src="' + esc(profile.photo) + '" onclick="ctAccount._photo()" title="Change photo">' : '<div class="ct-acc-photo" onclick="ctAccount._photo()" title="Add photo"><i class="fas fa-camera"></i></div>') +
      '<div style="flex:1 1 auto;min-width:0"><input type="text" id="ctAccName" placeholder="Your name" value="' + esc(profile.name) + '" oninput="ctAccount._profileEdited()" style="margin-bottom:6px">' +
      '<input type="text" id="ctAccProfSubj" placeholder="Your subjects (comma-separated)" value="' + esc((profile.subjects || []).join(', ')) + '" oninput="ctAccount._profileEdited()"></div></div>';
    h += '<label class="chk" style="display:flex;gap:7px;align-items:center;color:var(--muted);font-weight:700;font-size:13px;margin:2px 0 6px"><input type="checkbox" id="ctAccDark"' + (settings.defaultDark ? ' checked' : '') + ' onchange="ctAccount._profileEdited()"> Default to dark theme</label>';
    h += '<div class="ct-acc-row"><button class="btn btn-soft" id="ctAccSaveProf" onclick="ctAccount._saveProfile()"><i class="fas fa-floppy-disk"></i> Save profile</button></div>';

    h += '<div class="ct-acc-note"><i class="fas fa-shield-halved"></i> Everything stays on <b>your device / your folder</b> — nothing is uploaded. Use a synced folder (Drive / iCloud) to share across computers.</div>';

    /* ---- storage ---- */
    h += '<h3>Storage</h3>';
    h += '<div style="margin-bottom:8px"><span class="ct-acc-badge ' + (folderName ? 'on' : 'off') + '"><i class="fas fa-' + (folderName ? 'circle-check' : 'circle-dot') + '"></i> ' + (folderName ? 'Connected: ' + esc(folderName) : (fsSupported ? 'No folder yet' : 'On this device')) + '</span></div>';
    if (fsSupported && folderName) h += '<div class="ct-acc-note">Lists live in <b>' + esc(folderName) + '/Lists</b>; classes/setups in their own folders.</div>';
    h += '<div class="ct-acc-row">';
    if (fsSupported) h += '<button class="btn btn-primary" onclick="ctAccount.chooseFolder()"><i class="fas fa-folder-open"></i> ' + (folderName ? 'Change' : 'Choose folder') + '</button>';
    if (fsSupported) h += '<button class="btn btn-soft" onclick="ctAccount.scan()"><i class="fas fa-magnifying-glass"></i> Scan</button>';
    h += '<button class="btn btn-soft" onclick="ctAccount.export()"><i class="fas fa-file-arrow-down"></i> Export</button>';
    h += '<button class="btn btn-soft" onclick="ctAccount.import()"><i class="fas fa-file-arrow-up"></i> Import</button></div>';

    /* ---- lists ---- */
    h += '<h3>My lists</h3>';
    if (editingId !== null || isNew) {
      var L = isNew ? { id: '', type: 'class', title: '', tools: [], subjects: [], items: [] } : (getList(editingId) || {});
      var thisOnly = L.tools && L.tools.length === 1 && L.tools[0] === TOOL;
      h += '<div class="ct-acc-classcard">' +
        '<label class="ct-acc-lbl">List type</label><select id="ctAccType" onchange="ctAccount._typeChange()">' + TYPES.map(function (t) { return '<option value="' + t[0] + '"' + (L.type === t[0] ? ' selected' : '') + '>' + t[1] + '</option>'; }).join('') + '</select>' +
        '<label class="ct-acc-lbl" style="display:block;margin-top:8px">Title</label><input type="text" id="ctAccTitle" placeholder="e.g. Grade 7 - Rizal" value="' + esc(L.title) + '">' +
        '<div id="ctAccSubjWrap" style="' + (L.type === 'class' ? '' : 'display:none') + '"><label class="ct-acc-lbl" style="display:block;margin-top:8px">Subjects (comma-separated)</label><input type="text" id="ctAccSubj" placeholder="Science, Math" value="' + esc((L.subjects || []).join(', ')) + '"></div>' +
        (TOOL !== 'hub' ? '<label class="chk" style="margin-top:8px;display:flex;gap:7px;align-items:center;color:var(--muted);font-weight:700;font-size:13px"><input type="checkbox" id="ctAccThisTool"' + (thisOnly ? ' checked' : '') + '> Tag to this tool only (otherwise available everywhere)</label>' : '') +
        '<label class="ct-acc-lbl" style="display:block;margin-top:8px">Items (one per line' + (L.type === 'media' ? '; paste image/video links' : '') + ')</label><textarea id="ctAccItems" placeholder="Ana&#10;Ben&#10;Carlo">' + esc((L.items || []).join('\n')) + '</textarea>' +
        '<div class="ct-acc-row"><button class="btn btn-primary" onclick="ctAccount._save()"><i class="fas fa-check"></i> Save list</button><button class="btn btn-soft" onclick="ctAccount._cancel()">Cancel</button></div></div>';
    } else {
      if (!lists.length) h += '<div class="ct-acc-note">No lists yet. Create a class roster or any list to reuse it in every tool.</div>';
      // group lists by type into collapsible sections
      var groups = {}; lists.forEach(function (L) { (groups[L.type] = groups[L.type] || []).push(L); });
      TYPES.forEach(function (t) {
        var arr = groups[t[0]]; if (!arr || !arr.length) return;
        h += '<details class="ct-acc-group" open><summary>' + t[1] + ' <span class="ct-acc-pill">' + arr.length + '</span></summary>';
        arr.forEach(function (L) {
          var scope = (L.tools && L.tools.length) ? 'this tool' : 'all tools';
          h += '<div class="ct-acc-classcard"><div class="nm"><i class="fas fa-list" style="color:var(--primary)"></i> <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + esc(L.title) + '</span><span class="ct-acc-pill">' + (L.items || []).length + '</span>' +
            '<span class="acts"><button class="btn btn-soft" onclick="ctAccount._edit(\'' + L.id + '\')" title="Edit"><i class="fas fa-pen"></i></button><button class="btn btn-soft" onclick="ctAccount._del(\'' + L.id + '\')" title="Delete"><i class="fas fa-trash"></i></button></span></div>' +
            '<div class="meta">' + esc(scope) + (L.updated ? ' · ' + esc(L.updated) : '') + (L.subjects && L.subjects.length ? ' · ' + esc(L.subjects.join(', ')) : '') + '</div></div>';
        });
        h += '</details>';
      });
      h += '<button class="btn btn-soft" style="margin-top:6px" onclick="ctAccount._new()"><i class="fas fa-plus"></i> New list</button>';
    }
    drawer.innerHTML = h;
  }
  function getList(id) {
    for (var i = 0; i < lists.length; i++) if (lists[i].id === id) return lists[i];
    for (var j = 0; j < BUILTIN.length; j++) if (BUILTIN[j].id === id) return BUILTIN[j];
    return null;
  }

  window.ctAccount = {
    tool: TOOL,
    chooseFolder: chooseFolder,
    scan: function () { if (!dirHandle) { chooseFolder(); return; } scanFolder().then(function (ok) { toast(ok ? 'Scanned — ' + lists.length + ' list(s) found.' : 'Could not read folder. Try “Change folder”.'); render(); }); },
    export: exportBundle, import: importBundle,
    listLists: function (type, forTool) { return lists.filter(function (L) { return (!type || L.type === type) && (!forTool || !L.tools.length || L.tools.indexOf(forTool) >= 0); }); },
    builtinLists: builtinFor,
    getList: getList, onChange: function (fn) { listeners.push(fn); },
    open: openDr, _close: closeDr,
    _new: function () { isNew = true; editingId = null; render(); },
    _edit: function (id) { editingId = id; isNew = false; render(); },
    _cancel: function () { editingId = null; isNew = false; render(); },
    _typeChange: function () { var t = document.getElementById('ctAccType').value, w = document.getElementById('ctAccSubjWrap'); if (w) w.style.display = (t === 'class') ? '' : 'none'; },
    _photo: function () {
      var inp = document.createElement('input'); inp.type = 'file'; inp.accept = 'image/*';
      inp.onchange = function () { var f = inp.files[0]; if (!f) return; var r = new FileReader(); r.onload = function () { profile.photo = String(r.result); saveLocal(); writeAccount(); render(); }; r.readAsDataURL(f); };
      inp.click();
    },
    _profileEdited: function () { var b = document.getElementById('ctAccSaveProf'); if (b) { b.classList.add('btn-primary'); b.innerHTML = '<i class="fas fa-floppy-disk"></i> Save profile *'; } },
    _saveProfile: function () {
      profile.name = (document.getElementById('ctAccName').value || '').trim();
      profile.subjects = (document.getElementById('ctAccProfSubj').value || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
      settings.defaultDark = !!document.getElementById('ctAccDark').checked;
      saveLocal(); writeAccount();
      try { if (settings.defaultDark && !localStorage.getItem('ct-theme')) { localStorage.setItem('ct-theme', 'dark'); } } catch (e) {}
      toast(dirHandle ? 'Profile saved to “' + folderName + '/Account”' : 'Profile saved on this device'); render();
    },
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
      persist(L); if (L.type === 'class') ensureClassFolders(); toast(dirHandle ? 'Saved to “' + folderName + '/Lists”' : 'Saved on this device'); editingId = null; isNew = false; render(); buildListPickers();
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
    },
    /* class/subject folder tree + activity saving (item 7) */
    classes: classes, subjectsFor: subjectsFor, ensureClassFolders: ensureClassFolders,
    saveActivity: saveActivity, listActivities: listActivities, readActivity: readActivity
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
          if (typeof window.ctOnListLoad === 'function') { try { window.ctOnListLoad(tgt, L); } catch (e) {} }
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
    var t = type || lockType || '', forTool = TOOL === 'hub' ? '' : TOOL;
    var rel = window.ctAccount.listLists(t, forTool);
    var built = builtinFor(t, forTool);
    var opt = function (L) {
      return '<option value="' + L.id + '">' + esc(L.title) + ' (' + (L.items || []).length + ')</option>';
    };
    /* Grouped, not merged. A built-in cannot be renamed or deleted, so showing
       it in the same run as the teacher's own lists invites them to try. */
    sel.innerHTML = '<option value="">' + (rel.length || built.length ? 'Load a list…' : 'No saved lists yet') + '</option>' +
      (built.length ? '<optgroup label="Ready made">' + built.map(opt).join('') + '</optgroup>' : '') +
      (rel.length ? '<optgroup label="Your lists">' + rel.map(opt).join('') + '</optgroup>' : '');
  }
  listeners.push(buildListPickers);

  /* ---- collapsible Activity bar (item 3): class/subject feed the save path ---- */
  function fillSubjects(host) {
    var cl = host.querySelector('.ct-act-class'), sub = host.querySelector('.ct-act-subject'); if (!sub) return;
    var subs = subjectsFor(cl.value); var cur = sub.value;
    sub.innerHTML = '<option value="">Subject…</option>' + subs.map(function (s) { return '<option' + (s === cur ? ' selected' : '') + '>' + esc(s) + '</option>'; }).join('');
  }
  function refreshActivity(host) {
    var cl = host.querySelector('.ct-act-class');
    if (cl) { var cur = cl.value; cl.innerHTML = '<option value="">Class…</option>' + classes().map(function (c) { return '<option' + (c.title === cur ? ' selected' : '') + '>' + esc(c.title) + '</option>'; }).join(''); fillSubjects(host); }
    var prev = host.querySelector('.ct-act-prev');
    if (prev) { var acts = listActivities(TOOL); prev.innerHTML = '<option value="">Load a previous activity…</option>' + acts.map(function (a) { return '<option>' + esc(a.title) + (a.className ? ' — ' + esc(a.className) : '') + (a.date ? ' (' + a.date + ')' : '') + '</option>'; }).join(''); }
  }
  function buildActivityBars() {
    document.querySelectorAll('[data-ct-activity]').forEach(function (host) {
      if (!host.__ctBuilt) {
        host.__ctBuilt = true;
        host.innerHTML = '<details class="ct-act"><summary><i class="fas fa-flag"></i> <span class="ct-act-sum">Activity</span></summary><div class="ct-act-body">' +
          '<input class="ct-act-title" placeholder="Activity title">' +
          '<div class="ct-act-row"><select class="ct-act-class" title="Class"></select><select class="ct-act-subject" title="Subject"></select></div>' +
          '<div class="ct-act-row"><input class="ct-act-grade" placeholder="Grade"><input class="ct-act-section" placeholder="Section"></div>' +
          '<select class="ct-act-prev"></select></div></details>';
        host.querySelector('.ct-act-class').addEventListener('change', function () { fillSubjects(host); });
        host.querySelector('.ct-act-title').addEventListener('input', function () { host.querySelector('.ct-act-sum').textContent = this.value || 'Activity'; });
        host.querySelector('.ct-act-prev').addEventListener('change', async function () {
          var e = listActivities(TOOL)[this.selectedIndex - 1]; if (!e) return;
          host.querySelector('.ct-act-title').value = e.title || ''; host.querySelector('.ct-act-sum').textContent = e.title || 'Activity';
          var cl = host.querySelector('.ct-act-class'); if (e.className) { cl.value = e.className; fillSubjects(host); } var sub = host.querySelector('.ct-act-subject'); if (e.subject) sub.value = e.subject;
          var content = await readActivity(e); if (content != null && typeof window.ctOnActivityLoad === 'function') { try { window.ctOnActivityLoad(content, e); } catch (ex) {} }
        });
      }
      refreshActivity(host);
    });
  }
  listeners.push(buildActivityBars);
  window.ctActivity = { get: function () { var h = document.querySelector('[data-ct-activity]'); if (!h) return {}; var q = function (c) { var e = h.querySelector(c); return e ? e.value : ''; }; return { title: q('.ct-act-title'), className: q('.ct-act-class'), subject: q('.ct-act-subject'), grade: q('.ct-act-grade'), section: q('.ct-act-section') }; } };

  function boot() { loadLocal(); buildChrome(); render(); buildListPickers(); buildActivityBars(); reconnect(); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
