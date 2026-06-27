/* ============================================================
   ClassTools — Teacher Account / Class storage (item 19)
   - Injects an account icon into every tool's .topbar
   - Right-side drawer to manage classes (students + subjects),
     settings, and the save folder
   - Storage: localStorage is canonical; if the browser supports the
     File System Access API the teacher can also pick a folder and the
     data is mirrored there as Markdown files (sync via Drive/iCloud).
     Falls back to download/upload of a .md bundle when unavailable.
   - window.ctAccount API + auto "Load class" dropdowns via
     data-ct-class-target="#namesTextarea".
   No external deps. Privacy: nothing is uploaded anywhere.
   ============================================================ */
(function () {
  "use strict";
  if (window.__ctAccount) return; window.__ctAccount = true;

  var KEY = 'ct-account-v1';
  var fsSupported = (typeof window.showDirectoryPicker === 'function');
  var dirHandle = null; // session-only folder handle

  function load() {
    try { return JSON.parse(localStorage.getItem(KEY)) || {}; } catch (e) { return {}; }
  }
  function save(d) { try { localStorage.setItem(KEY, JSON.stringify(d)); } catch (e) {} fireChange(); }
  function data() { var d = load(); if (!d.classes) d.classes = []; if (!d.settings) d.settings = {}; return d; }
  var listeners = [];
  function fireChange() { listeners.forEach(function (f) { try { f(); } catch (e) {} }); }

  function uid() { return 'c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]; }); }

  /* ---------- markdown (de)serialisation ---------- */
  function classToMd(c) {
    return '---\nname: ' + (c.name || '') + '\nsubjects: ' + (c.subjects || []).join(', ') +
      '\n---\n\n# ' + (c.name || 'Class') + '\n\n## Students\n' +
      (c.students || []).map(function (s) { return '- ' + s; }).join('\n') + '\n';
  }
  function slug(s) { return String(s || 'class').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'class'; }

  async function mirrorToFolder() {
    if (!dirHandle) return false;
    try {
      var d = data();
      for (var i = 0; i < d.classes.length; i++) {
        var c = d.classes[i];
        var fh = await dirHandle.getFileHandle('class-' + slug(c.name) + '.md', { create: true });
        var w = await fh.createWritable(); await w.write(classToMd(c)); await w.close();
      }
      var sh = await dirHandle.getFileHandle('settings.md', { create: true });
      var sw = await sh.createWritable();
      await sw.write('---\nfont-head: ' + (d.settings.fontHead || '') + '\n---\n\n# ClassTools settings\n\nSaved ' + new Date().toLocaleString() + '\n');
      await sw.close();
      return true;
    } catch (e) { return false; }
  }

  async function chooseFolder() {
    if (!fsSupported) { exportBundle(); return; }
    try {
      dirHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
      var d = data(); d.folderName = dirHandle.name; save(d);
      await mirrorToFolder();
      render();
    } catch (e) { /* user cancelled */ }
  }

  function exportBundle() {
    var d = data();
    var md = '# ClassTools — exported ' + new Date().toLocaleString() + '\n\n' +
      d.classes.map(classToMd).join('\n\n---\n\n');
    var blob = new Blob([md], { type: 'text/markdown' });
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    a.download = 'classtools-classes.md'; document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(a.href);
  }
  function importBundle() {
    var inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.md,.txt,.json';
    inp.onchange = function () {
      var f = inp.files[0]; if (!f) return; var r = new FileReader();
      r.onload = function () {
        var text = String(r.result || ''); var d = data();
        // parse "## Students" blocks per class chunk
        text.split(/\n---\n/).forEach(function (chunk) {
          var nm = (chunk.match(/name:\s*(.+)/) || [])[1];
          var subj = (chunk.match(/subjects:\s*(.+)/) || [])[1] || '';
          var studs = []; var inS = false;
          chunk.split('\n').forEach(function (ln) {
            if (/##\s*Students/i.test(ln)) { inS = true; return; }
            if (inS && /^\s*-\s+/.test(ln)) studs.push(ln.replace(/^\s*-\s+/, '').trim());
          });
          if (nm && nm.trim()) d.classes.push({ id: uid(), name: nm.trim(), subjects: subj.split(',').map(function (x) { return x.trim(); }).filter(Boolean), students: studs });
        });
        save(d); render();
      };
      r.readAsText(f);
    };
    inp.click();
  }

  /* ---------- drawer UI ---------- */
  var scrim, drawer, editingId = null;
  function buildChrome() {
    scrim = document.createElement('div'); scrim.className = 'ct-acc-scrim';
    scrim.addEventListener('click', closeDrawer);
    drawer = document.createElement('div'); drawer.className = 'ct-acc-drawer';
    document.body.appendChild(scrim); document.body.appendChild(drawer);
    // inject account button into the top bar
    var bar = document.querySelector('.topbar');
    if (bar) {
      var b = document.createElement('button');
      b.className = 'icon-btn ct-acc-btn'; b.type = 'button';
      b.title = 'My classes & settings'; b.innerHTML = '<i class="fas fa-user"></i>';
      b.addEventListener('click', openDrawer);
      bar.appendChild(b);
    }
  }
  function openDrawer() { editingId = null; render(); scrim.classList.add('show'); drawer.classList.add('show'); }
  function closeDrawer() { scrim.classList.remove('show'); drawer.classList.remove('show'); }

  function render() {
    var d = data();
    var folder = fsSupported
      ? (d.folderName ? ('Saving to folder: <b>' + esc(d.folderName) + '</b>') : 'No folder chosen yet — data is on this device only.')
      : 'This browser can’t auto-save to a folder. Use Export / Import to move a Markdown file.';
    var h = '';
    h += '<button class="ct-acc-x" onclick="ctAccount._close()"><i class="fas fa-xmark"></i></button>';
    h += '<h2><i class="fas fa-user" style="color:var(--primary)"></i> My ClassTools</h2>';
    h += '<div class="ct-acc-note"><i class="fas fa-shield-halved"></i> Your classes and settings stay on <b>this device</b> — nothing is uploaded. To use them on another computer, save into a synced folder (Google Drive, iCloud, OneDrive…) or Export the file.</div>';
    h += '<h3>Storage</h3>';
    h += '<div class="ct-acc-note">' + folder + '</div>';
    h += '<div class="ct-acc-row">';
    if (fsSupported) h += '<button class="btn btn-primary" style="width:auto" onclick="ctAccount.chooseFolder()"><i class="fas fa-folder-open"></i> ' + (d.folderName ? 'Change folder' : 'Choose a folder') + '</button>';
    h += '<button class="btn btn-soft" onclick="ctAccount.export()"><i class="fas fa-file-arrow-down"></i> Export</button>';
    h += '<button class="btn btn-soft" onclick="ctAccount.import()"><i class="fas fa-file-arrow-up"></i> Import</button>';
    h += '</div>';

    h += '<h3>Classes</h3>';
    if (editingId !== null) {
      var c = d.classes.filter(function (x) { return x.id === editingId; })[0] || { id: '', name: '', subjects: [], students: [] };
      h += '<div class="ct-acc-classcard">' +
        '<label class="ct-acc-lbl">Class / section name</label>' +
        '<input type="text" id="ctAccName" placeholder="e.g. Grade 7 - Rizal" value="' + esc(c.name) + '">' +
        '<label class="ct-acc-lbl" style="display:block;margin-top:8px">Subjects (comma-separated)</label>' +
        '<input type="text" id="ctAccSubj" placeholder="Science, Math" value="' + esc((c.subjects || []).join(', ')) + '">' +
        '<label class="ct-acc-lbl" style="display:block;margin-top:8px">Students (one per line)</label>' +
        '<textarea id="ctAccStud" placeholder="Ana\nBen\nCarlo">' + esc((c.students || []).join('\n')) + '</textarea>' +
        '<div class="ct-acc-row"><button class="btn btn-primary" style="width:auto" onclick="ctAccount._saveClass()"><i class="fas fa-check"></i> Save class</button>' +
        '<button class="btn btn-soft" onclick="ctAccount._cancel()">Cancel</button></div></div>';
    } else {
      if (!d.classes.length) h += '<div class="ct-acc-note">No classes yet. Add a class to reuse its student list in any tool.</div>';
      d.classes.forEach(function (c) {
        h += '<div class="ct-acc-classcard"><div class="nm"><i class="fas fa-users" style="color:var(--primary)"></i> ' + esc(c.name) +
          ' <span class="ct-acc-pill">' + (c.students || []).length + ' students</span></div>' +
          (c.subjects && c.subjects.length ? '<div class="meta">' + esc(c.subjects.join(' · ')) + '</div>' : '') +
          '<div class="acts"><button class="btn btn-soft" onclick="ctAccount._edit(\'' + c.id + '\')"><i class="fas fa-pen"></i> Edit</button>' +
          '<button class="btn btn-soft" onclick="ctAccount._del(\'' + c.id + '\')"><i class="fas fa-trash"></i></button></div></div>';
      });
      h += '<button class="btn btn-soft" style="margin-top:6px" onclick="ctAccount._new()"><i class="fas fa-plus"></i> Add a class</button>';
    }
    drawer.innerHTML = h;
  }

  /* ---------- public API ---------- */
  window.ctAccount = {
    chooseFolder: chooseFolder,
    export: exportBundle,
    import: importBundle,
    listClasses: function () { return data().classes; },
    getClass: function (id) { return data().classes.filter(function (c) { return c.id === id; })[0] || null; },
    onChange: function (fn) { listeners.push(fn); },
    open: openDrawer,
    _close: closeDrawer,
    _new: function () { editingId = ''; render(); },
    _edit: function (id) { editingId = id; render(); },
    _cancel: function () { editingId = null; render(); },
    _del: function (id) { var d = data(); d.classes = d.classes.filter(function (c) { return c.id !== id; }); save(d); render(); },
    _saveClass: function () {
      var d = data();
      var name = (document.getElementById('ctAccName').value || '').trim();
      if (!name) { document.getElementById('ctAccName').focus(); return; }
      var subjects = (document.getElementById('ctAccSubj').value || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
      var students = (document.getElementById('ctAccStud').value || '').split('\n').map(function (x) { return x.trim(); }).filter(Boolean);
      if (editingId) { var c = d.classes.filter(function (x) { return x.id === editingId; })[0]; if (c) { c.name = name; c.subjects = subjects; c.students = students; } }
      else d.classes.push({ id: uid(), name: name, subjects: subjects, students: students });
      save(d); mirrorToFolder(); editingId = null; render();
      buildClassPickers();
    }
  };

  /* ---------- "Load class" dropdowns for tools ----------
     Any element with data-ct-class-target="#sel" gets a class dropdown
     inserted before it; choosing a class fills that target (textarea)
     with the student list (newline-joined).                          */
  function buildClassPickers() {
    var targets = document.querySelectorAll('[data-ct-class-target]');
    var classes = data().classes;
    targets.forEach(function (host) {
      var sel = host.previousElementSibling && host.previousElementSibling.classList && host.previousElementSibling.classList.contains('ct-classpick')
        ? host.previousElementSibling : null;
      if (!classes.length) { if (sel) sel.remove(); return; }
      if (!sel) {
        sel = document.createElement('div'); sel.className = 'ct-classpick';
        sel.innerHTML = '<i class="fas fa-graduation-cap" style="color:var(--muted)"></i><select><option value="">Load a class…</option></select>';
        host.parentNode.insertBefore(sel, host);
        sel.querySelector('select').addEventListener('change', function () {
          var c = window.ctAccount.getClass(this.value); if (!c) return;
          var tgtSel = host.getAttribute('data-ct-class-target');
          var tgt = document.querySelector(tgtSel) || host;
          if (tgt && 'value' in tgt) { tgt.value = (c.students || []).join('\n'); tgt.dispatchEvent(new Event('input', { bubbles: true })); }
        });
      }
      var dd = sel.querySelector('select');
      dd.innerHTML = '<option value="">Load a class…</option>' + classes.map(function (c) { return '<option value="' + c.id + '">' + esc(c.name) + ' (' + (c.students || []).length + ')</option>'; }).join('');
    });
  }
  listeners.push(buildClassPickers);

  function boot() { buildChrome(); buildClassPickers(); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
