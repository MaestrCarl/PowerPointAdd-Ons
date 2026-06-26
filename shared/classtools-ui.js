/* ============================================================
   ClassTools shared UI behaviour (loaded by every tool)
   - Lifts any visible [data-ct-play] surface into a blurred modal
   - Injects a floating close (×) that returns to settings
   - Keeps [data-ct-shell] previews from looking empty/collapsed
   - window.ctEaseStroke(): smooth, organic ring progress (item 6)
   No dependencies. Safe to load in any tool (even embedded).
   ============================================================ */
(function () {
  "use strict";
  if (window.__ctUI) return; window.__ctUI = true;

  /* ---- one-time DOM: scrim + close button ---- */
  var scrim = document.createElement('div'); scrim.className = 'ct-scrim';
  var closeBtn = document.createElement('button');
  closeBtn.className = 'ct-play-close'; closeBtn.type = 'button';
  closeBtn.setAttribute('aria-label', 'Close');
  closeBtn.innerHTML = '<i class="fas fa-xmark"></i>';
  function addChrome() {
    if (!document.body) return;
    if (!scrim.parentNode) document.body.appendChild(scrim);
    if (!closeBtn.parentNode) document.body.appendChild(closeBtn);
  }

  var current = null; // the surface currently shown as a modal

  function visible(el) {
    if (!el) return false;
    if (el === current) return true;            // already lifted (it's fixed)
    // These surfaces are shown/hidden via display/visibility, so key on that.
    var n = el;
    while (n && n.nodeType === 1) {
      var cs = getComputedStyle(n);
      if (cs.display === 'none' || cs.visibility === 'hidden') return false;
      n = n.parentElement;
    }
    return true;
  }

  function enter(el) {
    if (current === el) return;
    if (current) current.classList.remove('ct-play-surface');
    current = el;
    document.body.classList.add('ct-playing');
    el.classList.add('ct-play-surface');
  }
  function exit() {
    if (!current) { document.body.classList.remove('ct-playing'); return; }
    current.classList.remove('ct-play-surface');
    current = null;
    document.body.classList.remove('ct-playing');
  }

  closeBtn.addEventListener('click', function () {
    var el = current;
    exit();
    if (!el) return;
    // 1) explicit handler  2) data-ct-back selector  3) common back buttons
    if (typeof window.ctClosePlay === 'function') { try { window.ctClosePlay(); return; } catch (e) {} }
    var sel = el.getAttribute('data-ct-back');
    var target = sel ? document.querySelector(sel) : null;
    if (!target) target = document.querySelector('#backBtn, [data-ct-back-btn], .ct-back');
    if (target) target.click();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && current) { e.preventDefault(); closeBtn.click(); }
  });

  /* ---- preview shells: flag emptiness so the placeholder shows ---- */
  function refreshShells() {
    var shells = document.querySelectorAll('[data-ct-shell]');
    for (var i = 0; i < shells.length; i++) {
      var s = shells[i];
      var empty = s.childElementCount === 0 && !s.textContent.trim();
      s.setAttribute('data-ct-empty', empty ? '1' : '0');
    }
  }

  /* ---- scan: decide which surface (if any) is the active modal ---- */
  var raf = 0;
  function scan() {
    raf = 0;
    var plays = document.querySelectorAll('[data-ct-play]');
    var active = null;
    for (var i = 0; i < plays.length; i++) {
      // ignore the lifted one's fixed state when judging others
      if (plays[i] !== current && visible(plays[i])) { active = plays[i]; break; }
    }
    if (!active && current && visible(current)) active = current;
    if (active) enter(active); else exit();
    refreshShells();
  }
  function schedule() { if (!raf) raf = requestAnimationFrame(scan); }

  function boot() {
    addChrome();
    schedule();
    var obs = new MutationObserver(schedule);
    obs.observe(document.body, { attributes: true, childList: true, subtree: true,
      attributeFilter: ['style', 'class', 'hidden'] });
    window.addEventListener('hashchange', schedule);
    window.addEventListener('resize', schedule);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();

  /* ---- item 6: smooth, semi-organic ring progress ----
     ctEaseStroke(circleEl, toOffset, [ms], [circumference])
     Animates stroke-dashoffset with an eased curve via rAF so the
     colour sweep glides instead of jumping each tick.               */
  var ringTimers = new WeakMap();
  function easeInOut(t) { return t < 0.5 ? 2 * t * t : 1 - Math.pow(-2 * t + 2, 2) / 2; }
  window.ctEaseStroke = function (circle, toOffset, ms) {
    if (!circle) return;
    ms = ms || 950;
    var from = parseFloat(circle.style.strokeDashoffset);
    if (isNaN(from)) { circle.style.strokeDashoffset = toOffset; return; }
    var prev = ringTimers.get(circle); if (prev) cancelAnimationFrame(prev);
    var t0 = performance.now(), delta = toOffset - from;
    if (Math.abs(delta) < 0.5) { circle.style.strokeDashoffset = toOffset; return; }
    (function step(now) {
      var k = Math.min(1, (now - t0) / ms);
      circle.style.strokeDashoffset = (from + delta * easeInOut(k));
      if (k < 1) ringTimers.set(circle, requestAnimationFrame(step));
    })(t0);
  };

  window.ctRefreshUI = schedule; // tools can poke a re-scan after big DOM swaps
})();
