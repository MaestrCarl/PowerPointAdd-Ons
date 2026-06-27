# ClassTools — PowerPoint Add-in suite (handoff notes)

Static HTML/CSS/JS classroom tools, hosted on GitHub Pages, embedded in PowerPoint as a **single content add-in**. No build step, no framework, no backend. Each tool is one self-contained `.html` file. Owner: "Mr. Carlo" (`MaestrCarl`).

Live base: `https://maestrcarl.github.io/PowerPointAdd-Ons/`
Repo: `https://github.com/MaestrCarl/PowerPointAdd-Ons` (push from the user's Mac; this sandbox has no git credentials).

## File map
- `Toolbox/index.html` — **the hub**. The only add-in loaded by PowerPoint. Card gallery, opens each tool in an iframe, owns the bookmark (saved into the .pptx via Office settings), and is the only file that loads `office.js`. Also owns the **starred-tools history** panel.
- `Toolbox/manifest.xml` — the single manifest the user sideloads into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`. Points at `Toolbox/index.html`.
- `RandomPicker/picker.html` — "Pickers, Pairings & Groupings": Picker / Groups / Pairs / Spinner.
- `MyTimerAddin/timer.html` — "Timer Tools": stopwatch, countdown, loop, pomodoro, split, silent, race. **Race is now a circular dial** (conic-gradient progress + animated rainbow `.ring-glow` sheen + an orbiting runner); the same circular ring is shown on every countdown-style timer and in the present view (`.ringwrap/.ring-prog/.ring-glow/.ring-hole`, driven by `--p`). **Quiz Buzzer was removed** — to be relocated into WordGames (gaming section).
- `RandomPicker/picker.html` — rebuilt: each mode uses a left **sidebar (settings)** + right **viewer**. Picker has remove-on-select + no-repeat pool + picked history. Groups deal cards over ~3s and support **pre-assigned leaders** (crown toggles in a roster). Pairs have **keep-apart** marking (ban toggles), trio/solo leftover option, ~3s animation, no truncated chips. Spinner shows a **winner modal**; the name is only removed (if opted) **after the modal closes** so the wheel lingers on the correct result. Per-mode **history export/import** (JSON/txt). Theme always follows live `ct-theme` (never the saved file theme) — fixes the dark-card-in-light-hub bug.
- `WordGames/games.html` — "Word Games & Builders": Hangman, Word Scramble, Sentence Unscramble, Category Sort, Word Search, **Crossword**, **Quiz Buzzer** (a `special:true` registry tile → `openBuzzer()`, multi-team first-to-buzz pads using the `#player` container + `inBuzzer` flag in smartBack/updateBack/showGallery). Builder now has **Save/Load list (.txt)** (`exportSetup`/`importSetup`, `#setupFile`) and **Print / Save as PDF** for WS/CW (`printPuzzle()` reuses the PNG canvas builders via a `_capture` flag on `dlCanvas`, opens a print window with puzzle + answer key pages).
- `WordSplitter/splitter.html` — "Word Splitter": split words into parts by a delimiter (`but-ter-fly`) or auto N parts; two play modes (**Arrange parts**, **Fill the blanks**); optional TTS audio; per-word hint + **3 global hint tokens** (reveal next part); shareable `#play=` link; student gate + certificate with `#cert=` QR. Self-contained; no office.js.
- `SpellingHelper/spelling.html` — "Spelling Helper": listen-and-spell. Natural Web-Speech TTS (voice/rate/pitch pickers, prefers neural/en voices), three answer modes (type / click letters / letter blocks), per-word hint + example sentence, shareable `#play=` link (base64), student gate (name/number/grade/class) on shared links, and an auto completion certificate with a `#cert=` QR that rebuilds it. Self-contained; no office.js.
- `Scoreboard/scoreboard.html` — rebuilt with **sidebar (settings) + viewer** layout. Adds quick multi-add, configurable quick-point buttons (`incs`), step, allow-negative, rank-by-score, show/hide members & rank, **win target** with auto **Announce winner** modal + confetti, inline rename / recolour / set-exact-score (contenteditable), export/import (JSON), share link + QR. Theme follows live `ct-theme`.
- `IntroTools/intro.html` — "Intro & Reveal Tools": mystery box, envelopes, doors, wheel, scratch, hidden reveal.

## Shared UI layer (`shared/classtools-ui.css` + `classtools-ui.js`)
Loaded by every tool via `../shared/classtools-ui.(css|js)` (defer). Auto-upgrading, opt-in by attribute — no per-tool rewrites:
- **Full-screen blurred play modal** (item 1): any element with `data-ct-play` is lifted into a centered card over a blurred scrim whenever it becomes visible (detected by `display`/`visibility` via a MutationObserver). A floating × is injected; closing calls `window.ctClosePlay()` if defined, else clicks the `data-ct-back` selector, else `#backBtn`. Surfaces: WordGames/Intro `#player`, Timer `#presentView`, Picker `#presentView`, Scoreboard `#scorePresent` (live board moved in/out), Splitter/Spelling `#playView`.
- **Left-sidebar settings** (item 2): `.b2-left` (games/intro) gets a sidebar card look; `.ct-build/.ct-side/.ct-view` available for any tool. Picker/Scoreboard/Splitter/Spelling already use `.side`.
- **Non-collapsing preview shell** (item 3): `data-ct-shell` + `data-ct-hint` gives a dashed min-height canvas with a placeholder while empty (`#b2preview`, `#introPreview`, `#buildPreview`).
- **Smooth ring** (item 6): `window.ctEaseStroke(circle,offset,ms)` rAF-eases `stroke-dashoffset`. Timer drives its conic `--p` via a local rAF easing loop (`smoothRings()`); games timer ring uses `ctEaseStroke`.
- **Field placement** (item 5): Timer & Picker hide the `.classbar` (Activity/Grade/Section/Subject) on the mode gallery; it shows only inside an open mode.
- **Hub title** (item 4): the duplicate `.hub-header h1` was removed; only the top-bar brand reads "ClassTools".

## Shared storage / account layer (`shared/classtools-account.js`, styles in `classtools-ui.css`)
Loaded by every tool (and the hub) via `../shared/classtools-account.js` (defer). Injects a **user/account icon** into `.topbar` → a right-side drawer for managing reusable **lists** + the save folder. Goal: an Obsidian-like local store, no database.
- **Tool detection**: `TOOL` is derived from the URL path (`wordgames|picker|timer|scoreboard|intro|splitter|spelling|hub`). `TOOL_FOLDER` maps each to a folder name.
- **List model**: `{id,type,title,tools:[],subjects:[],items:[],updated}`. `type` ∈ class/words/names/questions/items. `tools` empty = available everywhere; otherwise TAGS the list to specific tools.
- **Folder layout** (inside the chosen folder): `Lists/` for reusable lists; `<ToolName>/` for per-tool saved setups (auto-created by `saveSetup`).
- **File naming** (parseable, dated for versioning): `ct__<kind>__<scope>__<type>__<slug>__<id>__<YYYY-MM-DD>.md` where kind=list|setup, scope=`any` or tool-ids joined by `+`. On save, older files with the same `__<id>__` are deleted (no orphan versions).
- **File content** = one or more delimited blocks: `<!-- ct:list id=".." type=".." title=".." tools=".." subjects=".." updated=".." -->` then `## Title`, `- item` lines, then `<!-- ct:end id=".." -->`. Parsing is marker-based (`parseMd`), so any `.md` anywhere in the folder is understood regardless of filename.
- **Storage**: localStorage (`ct-account-v3`) is the always-on mirror. If the **File System Access API** is available, the folder handle is persisted in **IndexedDB** (`ct-store`/`dir`); `reconnect()` re-reads on boot, `scanFolder()` recursively walks (≤4 deep) and de-dupes by id keeping newest `updated`. A **Scan folder** button forces a re-read (needed when the browser doesn't auto-restore permission). PowerPoint add-in / unsupported browsers fall back to Export/Import `.md`.
- **APIs for tools**: `window.ctAccount.listLists(type, forTool)`, `getList(id)`, `saveList(type,title,items,allTools)`, `saveSetup(kind,title,mdContent)` (writes into `<folder>/<ToolName>/`), `scan()`, `onChange(fn)`, `.tool`.
- **"Load a list" control**: any element with `data-ct-list-target="#sel"` (or legacy `data-ct-class-target` to lock type=class, or `data-ct-list-type="words"`) gets a `[type ▾][list ▾]` picker injected above it, filtered to lists tagged for the current tool (or "all tools"). Wired into Picker's 4 name textareas + WordGames `#grpRoster`. Choosing a list fills the target textarea and fires `input`.
- **Per-tool save example**: WordGames `exportSetup()` now writes a markdown setup via `ctAccount.saveSetup('hangman', title, md)` → `<folder>/WordGames/ct__setup__…md` (else downloads `.md`). Other tools' export buttons should be migrated the same way (markdown + `saveSetup`/`saveList`).
- **All exports/imports are Markdown** going forward (no more .txt/.json for new code).
- **Profile & settings**: the drawer also has a teacher profile (name, subjects, photo) + settings (`defaultDark`), saved to `<folder>/Account/account.md` (+ `profile-photo.*`) and mirrored to localStorage. Connected state shows a green `.ct-acc-badge`. Drawer buttons are normalised via `.ct-acc-drawer .btn{width:auto…}` (tools force `.btn-primary{width:100%}`). Lists render grouped by type in collapsible `<details class="ct-acc-group">` with edit/delete pinned right. List types: class/words/questions/items/media ("Name list" was removed).
- **List-load hook**: after a list is loaded into a target, the shared loader calls `window.ctOnListLoad(target,list)` if defined (WordGames uses it to switch to paste mode + refresh preview). WordGames has a top-of-sidebar loader via a hidden `<span data-ct-list-target="#items" data-ct-list-type="words">`.
- **Auto-fit text** (item 5): `window.ctFitText(el)` / `ctFitAll(root)` shrink `.ct-fit` elements to fit one line; the play-modal scan auto-fits on open. Picker winner/shuffle names use `.ct-fit data-ct-fit-max`.
- **Play-modal × bug** (recurring): the surface rule must be `> *:not(.ct-play-close){width:100%}` — without the `:not`, the close button stretches full width. The × is `position:absolute` and is appended INTO the surface by the shared JS.

### Storage — class/subject tree + Activity bar (DONE 2026-06-27 s8)
- **Class/subject folder tree** (item 7): saving a class list auto-creates `<folder>/<Class>/` + `<Class>/<Subject>/` (`ensureClassFolders`). `ctAccount.saveActivity(meta, content)` writes the activity md into `<Class>/<Subject>/`; if `meta.className` is missing it opens a **chooser modal** (pick class + subject, or "General") — `openChooser`. `safeName()` keeps folders human-readable. Activities are indexed in `localStorage['ct-activities']`; `listActivities(tool)` + `readActivity(entry)` power "load previous".
- **Activity bar** (item 3): drop `<div data-ct-activity></div>` into a sidebar → shared script renders a collapsible **Activity** panel (title, Class ▾ from account classes, Subject ▾ from that class, Grade, Section, + "Load previous activity ▾"). `window.ctActivity.get()` returns `{title,className,subject,grade,section}`. A tool can define `window.ctOnActivityLoad(content,entry)` to consume a reloaded activity. Added to WordGames/Intro/Splitter/Spelling/Scoreboard. **WordGames is the full template**: Save → `ctAccount.saveActivity({kind,title,className,subject,…}, md)`; previous activities reload via `ctOnActivityLoad` (parses `- ` lines back into `#items`).

### Storage — still TODO
- **Per-tool Save migration** (item 6/14): WordGames is wired to `saveActivity`. Picker/Scoreboard/Splitter/Spelling/Timer still need their Save/Export buttons migrated to `ctAccount.saveActivity` (markdown into class/subject) + `ctActivity` bar usage; Timer/Picker still use their own `.classbar` instead of `data-ct-activity`.
- **Media lists** (item 2): a `media` type exists but binary drag-drop into a per-list media folder is not built yet.
- **Reveal 3D** (items 10–12,15), **WebRTC buzzer w/ QR + PeerJS** (item 16), **shared-link truncation** (item 17): not yet addressed.

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
