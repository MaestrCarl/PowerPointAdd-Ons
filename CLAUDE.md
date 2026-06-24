# ClassTools — PowerPoint Add-in suite (handoff notes)

Static HTML/CSS/JS classroom tools, hosted on GitHub Pages, embedded in PowerPoint as a **single content add-in**. No build step, no framework, no backend. Each tool is one self-contained `.html` file. Owner: "Mr. Carlo" (`MaestrCarl`).

Live base: `https://maestrcarl.github.io/PowerPointAdd-Ons/`
Repo: `https://github.com/MaestrCarl/PowerPointAdd-Ons` (push from the user's Mac; this sandbox has no git credentials).

## File map
- `Toolbox/index.html` — **the hub**. The only add-in loaded by PowerPoint. Card gallery, opens each tool in an iframe, owns the bookmark (saved into the .pptx via Office settings), and is the only file that loads `office.js`. Also owns the **starred-tools history** panel.
- `Toolbox/manifest.xml` — the single manifest the user sideloads into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`. Points at `Toolbox/index.html`.
- `RandomPicker/picker.html` — "Pickers, Pairings & Groupings": Picker / Groups / Pairs / Spinner.
- `MyTimerAddin/timer.html` — "Timer Tools": stopwatch, countdown, loop, pomodoro, split, silent, race, quiz buzzer.
- `WordGames/games.html` — "Word Games & Builders": Hangman, Word Scramble, Sentence Unscramble, Category Sort, Word Search, **Crossword** (all live).
- `Scoreboard/scoreboard.html` — "Scoreboard & Scoring": teams + live scoring.
- `IntroTools/intro.html` — "Intro & Reveal Tools": mystery box, envelopes, doors, wheel, scratch, hidden reveal.

## Shared conventions (duplicated per file on purpose — keeps each tool standalone)
- **Theme & fonts**: CSS vars on `:root` + `.dark-mode`. `--font-head:'Baloo 2'` (headings), `--font-body:'Nunito'` (body), loaded from Google Fonts. **Luckiest Guy is reserved for the hub "ClassTools" logo only** (`--font-logo` in `Toolbox/index.html`, applied to `.topbar .brand` + `.hub-header h1`); do NOT use it elsewhere — it reads as wacky at body sizes. To re-skin, change the two `--font-*` vars + the Google Fonts `<link>` (kept per-file). Primary purple `#6d28d9`/`#9333ea`. Palette `PAL`/`PALETTE` = `['#6d28d9','#ff3d7f','#14d6c4','#ffb020','#2f6bff','#22c55e','#ef4444','#f97316']`.
- **Dark mode sync** across tools via `localStorage['ct-theme']` + `storage` listener. `applyTheme(dark)` / `toggleDark()`.
- **Top bar** one line: back (only when relevant), brand, bookmark star + tip, spacer, icon buttons (share, mute, dark…).
- **Stepwise Back** (`smartBack()` in games + intro): player/form → gallery list; only the gallery → hub home (`postMessage {ns:'classtools',cmd:'close'}`). `updateBack()` shows the back button when embedded OR not on the gallery.
- **Hub home**: the "ClassTools" logo is clickable (`goHome()`).
- **Bookmark**: star posts `{ns:'classtools',cmd:'bookmark',url,label}` (url=`location.href`, or `null`). Hub saves to Office settings key `ct-bookmark` (falls back to localStorage) and auto-opens on load. Star pulses (`star-hint`) when there's content but no bookmark.
- **Starred history (hub)**: every bookmark is also appended to `localStorage['ct-star-history']` (deduped, max 30). The clock-rotate icon opens a panel to re-open or remove entries; removing the active one also clears `ct-bookmark`.
- **Sound**: tiny WebAudio engine — `ea()`/`ensureAudio()`, `tn()`/`tone()`, `mar()`/`marimba()`, master gain 0.16, `toggleMute()`. Keep gentle.
- **Confetti**: `confetti(ms)` on a fixed `#confetti` canvas.
- **Link-encoded state** (games + intro): config in the URL hash. `enc/dec` = base64 of UTF-8 JSON. Seeded RNG `mulberry(seed)` + `shuf(arr,seed)` for **reproducible layouts**; store enough in the payload (or the seed) to reproduce. QR of the current link via `qrcodejs`.
- **Masking**: builder inputs are masked with **`-webkit-text-security:disc` / `text-security:disc`** (dots) — NOT blur — so a teacher can type and press Enter live without students seeing the letters. The `.masked` class is on `#items`; in games the `mrows` class also masks the row inputs (`#rowsEditor.mrows input`, `#catCats.mrows input`). On by default; the eye toggle reveals after an **in-page confirm modal** (`uiConfirm({title,body,yes,onYes})`, never the native `confirm()`).
- **In-page dialogs**: `uiConfirm()` + `#confirmModal` replace `window.confirm`. Generic modal helpers `openModal(id)`/`closeModal(id)` + `.modal.show`.

## Word Games builder architecture (`WordGames/games.html`)
- **Dual-panel form**: `.builder2` = `.b2-left` (settings) + `.b2-right` (`#previewPanel` → `#b2preview` live preview, always visible, + `#exportRow` for WS/CW). Every game shows a live preview (`previewUpdate()` dispatches: ws→`wsRender`, cw→`cwPreview`, hg→`hgPreview`, sc→`scPreview`, se→`sePreview`, cat→`catPreview`). Preview chips/slots use `.pvtile/.pvslot/.pvword/.pvcats…`.
- **Unified entry editor** with a **Rows ⇄ Paste** toggle (`setMode`, `MODE`). `#items` textarea is the single canonical source (one item per line). Rows mode is a friendly per-game editor that writes canonical lines via `rowSync()`/`rowsToText()`; Paste mode edits `#items` directly; `parseToRows()` converts back. `ENTRY` config maps each game's row fields. Category Sort has its own `#catCats` (chips: `catAdd/catDel`) + word rows with a category dropdown.
- **Canonical line formats**: hg `word | hint`; sc `word`; se `sentence`; ws `word | definition | flags` (flags ⊂ `d`,`r`); cw `word | clue`; cat `Category = a, b, c`.
- **Word Search**: `buildWordSearch(wordObjs,size,snake,seed)` where each `wordObj={w,dirs}`. Global direction toggles (`wsH/V/D/R`) + **per-word D/R ticks** combine via `wordDirs()`. `snake` lets words bend (self-avoiding path). Grid auto-sizes to longest word + 2 padding (`wsAutoSize`), overridable (`wsSize`/`wsSizeManual`). Player traces a path (`playWordSearch`) so straight AND snake words work; shows definitions.
- **Crossword**: `buildCrossword(entries)` interlocks words at shared letters (falls back to placing unmatched words below), numbers start cells, returns `{cells,entries,H,W}` locked into the link. `playCrossword` renders numbered Across/Down inputs with check + reveal.
- **Letter case**: `CAPS` ('upper'/'lower') for ws/cw/hg — preview + player apply `text-transform`. **Printable B&W PNG export** (`exportWsPNG`/`exportCwPNG`, puzzle + answer key) via a canvas → `toDataURL` download.
- **Timer overlay** (per game): `timer` payload `{on,mode:'up'|'down',secs,style:'num'|'ring'}`; engine `timerInit/timerTick/timerBar/timerPaint`; ring style drains/fills colour.
- **Drag tiles**: scramble + sentence default to draggable tiles (`makeSortable`, pointer-based), toggle `optDrag`.
- **Drafts**: `DRAFT` in memory + `localStorage['ct-wg-draft']`; the gallery is the landing view, draft is applied only when its game is opened.

## Intro Tools (`IntroTools/intro.html`)
- **Dual-panel form** too: `.builder2` left settings + right `#introPreview` (`introPreview()` renders a representative mini visual per activity — 3D box, mini envelopes/doors, conic-gradient `wheelMini`, scratch cards, blurred sample).
- **3D room stage** (`.room` + `roomDeco()`): floor/wall gradient, window/frame/bunting/plant/rug decorations. Box is a real `preserve-3d` cube with a lifting lid; sits above the floor line (`z-index`).
- **Modal Reveal effect**: results fly out into a 3D modal (`showReveal(it,{door,replay})` / `showRevealMulti`, `#revealModal`/`.reveal3d.fly`). Doors use `{door:true}` → dark-room skin with an animated ceiling **lamp + spotlight**, after the door swings open (`.door.opened` → `doorSwing`). A **Revealed history** strip (`REVEALED`, `pushReveal`/`renderHistory`/`reopenReveal`) and a **"take items out once revealed" toggle** (`removeRevealed`, `removeBar()`/`setRemove`) apply across box/envelopes/doors/wheel/scratch.
- **Scratch** is a multi-card picker → pick a card → scratch it inside `#scratchModal` with **brush size + sharpness** sliders (`scratchSize`, `scratchSharp`, `initScratch`).
- **Wheel** keeps a live `wheelLive` index list so revealed options can be removed; `wheelReset()` refills.

## Other tools
- **Timer & Picker** present their modes as a **card gallery** (not header tabs): `TMODES`/`PMODES` → `buildModeGallery()`, `openMode(id)` shows the panel + a back bar, `showModeGallery()` returns. `openTab()` is kept as a thin alias. Boot builds the gallery and lands on it.
- **Scoreboard** is single-mode (no tabs to convert). Spinner (in Picker) is the live "preview" for that mode; Timer's present view is its live display.

## Adding a new game / activity
Registry-driven. Games: add to `GAMES`, a `play<Name>()` in `PLAYERS`, an `ENTRY` entry (or custom builder), and a `buildPayload` branch + `previewUpdate` case. Intro: add to `ACTS`, a `play<Name>()` in `PLAYERS`, and an `introPreview()` case.

## Constraints / gotchas
- Tools run **embedded in the hub iframe**, where `office.js` never initializes — **do not** put `office.js` in tool files. Tools persist via `localStorage`; durable persistence rides on the hub bookmark.
- No real-time multi-device sync (static hosting). QR handoff is the substitute.
- Layout: `body{height:100vh;flex column}`, `main{flex:1;min-height:0}`. Responsive via `vmin`/`clamp`/`vh`. `.builder2` wraps to one column on narrow add-in boxes.
- PNG export uses `canvas.toDataURL` + an `<a download>`; works in a browser/embedded WebView.

## Deploy
Push to GitHub `main`; Pages serves within ~1 min. Run `git push` from the Mac (clear stale `.git/*.lock` first if needed).

## Status / backlog
Done: hub (+ home logo + starred history), fonts (Baloo 2 / Nunito, Luckiest Guy logo), tool renames, card-gallery Timer/Picker, masking-as-dots + in-page confirm, stepwise Back. Word Games: unified rows/paste builder, live preview for every game, **Word Search** (per-word ticks, snake, directions, auto-size) + **Crossword** + caps toggle + printable PNG export + in-game timers + drag tiles. Intro: 3D room/box, Modal Reveal + history + remove toggle, multi-card scratch with brush controls, doors zoom-in lit room, per-activity live preview.
Ideas next: a live preview for Picker Groups/Pairs; extract the duplicated core into `shared/` if maintenance grows; optional crossword auto-relayout when a word can't interlock.
