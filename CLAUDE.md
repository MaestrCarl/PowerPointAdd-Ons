# ClassTools — PowerPoint Add-in suite (handoff notes)

Static HTML/CSS/JS classroom tools, hosted on GitHub Pages, embedded in PowerPoint as a **single content add-in**. No build step, no framework, no backend. Each tool is one self-contained `.html` file. Owner: "Mr. Carlo" (`MaestrCarl`).

Live base: `https://maestrcarl.github.io/PowerPointAdd-Ons/`
Repo: `https://github.com/MaestrCarl/PowerPointAdd-Ons` (push from the user's Mac; this sandbox has no git credentials).

## File map
- `Toolbox/index.html` — **the hub**. The only add-in loaded by PowerPoint. Shows a card gallery, opens each tool in an iframe, owns the bookmark (saved into the .pptx via Office settings), and is the only file that loads `office.js`.
- `Toolbox/manifest.xml` — the single manifest the user sideloads into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`. Points at `Toolbox/index.html`.
- `RandomPicker/picker.html` — Picker / Groups / Pairs / Spinner.
- `MyTimerAddin/timer.html` — multi-timer (stopwatch, countdown, loop, pomodoro, split, silent, race, quiz).
- `WordGames/games.html` — game gallery: Hangman, Word Scramble, Sentence Unscramble, Category Sort, Word Search (Crossword = "coming soon").
- `Scoreboard/scoreboard.html` — teams + live scoring.
- `IntroTools/intro.html` — reveal/pick activities (mystery box, envelopes, doors, wheel, scratch, hidden reveal).

## Shared conventions (copy these patterns; they are duplicated per file on purpose — keeps each tool standalone)
- **Theme**: CSS vars on `:root` + `.dark-mode`. Fonts: `--font-head:'Bricolage Grotesque'`, `--font-body:'Hanken Grotesk'` (Google Fonts). Primary purple `#6d28d9`/`#9333ea`. Palette array `PAL`/`PALETTE` = `['#6d28d9','#ff3d7f','#14d6c4','#ffb020','#2f6bff','#22c55e','#ef4444','#f97316',...]`.
- **Dark mode sync** across tools via `localStorage['ct-theme']` + a `storage` event listener. `applyTheme(dark)` / `toggleDark()`.
- **Top bar** is one line: back (only when embedded), brand, large bookmark star + tip, spacer, then icon buttons (share, mute, dark, etc.).
- **Hub nav**: tools detect embedding with `EMBEDDED = window.parent && window.parent !== window`. Back button posts `{ns:'classtools',cmd:'close'}` to parent; hub listens and calls `closeTool()`.
- **Bookmark**: star posts `{ns:'classtools',cmd:'bookmark',url,label}` (url=`location.href`, or `null` to clear). Hub saves it to Office settings key `ct-bookmark` (falls back to localStorage) and auto-opens it on load. The star pulses (`star-hint`) when there is content but it isn't bookmarked yet.
- **Sound**: tiny WebAudio engine — `ea()`/`ensureAudio()`, `tn()`/`tone()`, `mar()`/`marimba()`, master gain 0.16, `toggleMute()`. Keep gentle/low volume.
- **Confetti**: `confetti(ms)` draws on a fixed `#confetti` canvas.
- **Link-encoded state** (games + intro tools): config lives in the URL hash, no storage needed.
  - `enc(o)=btoa(unescape(encodeURIComponent(JSON.stringify(o))))`, `dec` is the inverse. Unicode-safe.
  - Seeded RNG `mulberry(seed)` + `shuf(arr,seed)` for **reproducible layouts** (word-search grid, scrambles, item order). Always store enough in the payload (or the seed) to reproduce the exact layout.
  - QR of the current link via `qrcodejs` (cdnjs). Share modal shows link + QR.
- **Masking**: builder textareas blur by default (`.masked{filter:blur(5px)}`, NOT `-webkit-text-security` — that broke Enter). Eye toggle sits just above the textarea, `confirm()` before revealing, resets to masked on every form load.

## Adding a new game (Word Games) or activity (Intro Tools)
Both files use a **registry**:
1. Add a meta entry to the `GAMES`/`ACTS` array (`id,name,icon,c1,c2,desc,label,ph,...`).
2. Add a `play<Name>()` function and map it in `PLAYERS`/`ACTPLAYERS`.
3. Add a `buildPayload` branch for its `id`. The gallery, form, share/QR, bookmark, and player framework are generic.
Set `soon:true` to show a disabled "coming soon" tile.

## Adding a whole new tool
1. New folder + `tool.html` (copy an existing tool's head/topbar/core helpers).
2. Register it in `Toolbox/index.html` → `SECTIONS` array (group `name` + `tools:[{title,desc,icon,url,c1,c2}]`). `url` is relative, e.g. `../NewTool/tool.html`.

## Constraints / gotchas
- Tools run **embedded in the hub iframe**, where `office.js` never initializes — so **do not** put `office.js` in tool files (it blanks the page in a plain browser). Only the hub uses it. Tools persist via `localStorage`; durable in-file persistence rides on the hub bookmark.
- Real-time multi-device sync is **not possible** (static hosting, no server). QR handoff (open on phone with preloaded config) is the supported substitute.
- Layout: `body{height:100vh;display:flex;flex-direction:column}`, `main{flex:1;min-height:0}`, panes `height:100%`. Avoid absolutely-positioned full-bleed panes (they collapsed inside the iframe).
- Everything is responsive to the add-in box via `vmin`/`clamp`/`vh`.

## Deploy
Push to GitHub `main`; GitHub Pages serves it within ~1 min. The user runs `git push` from their Mac (clear stale `.git/*.lock` first if the sandbox left any).

## Status / backlog
Done: hub, picker (+spinner), timer, scoreboard, word games (hangman/scramble/sentence/category/word-search), intro tools, bookmark, QR, dark sync, link encoding.
Next: **Crossword** (generator + player, lock layout in link), and flesh out any "coming soon" tiles. Consider extracting the duplicated core into `shared/core.js` + `shared/theme.css` if maintenance cost grows (currently kept inline so each tool is standalone).
