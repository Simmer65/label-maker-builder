---
name: label-maker-builder
description: Generate a customized Windows label-printing desktop app for the user (Python/tkinter + pywin32, direct-to-printer GDI). Prints Avery-style file-folder label sheets on 4x6 media with a WYSIWYG preview, three format presets, printer calibration, and zero-dialog Quick Print. Use when the user asks to "build me a label printing app", "create a file folder label maker", "set up the TRI label maker", "Avery label printing app", or similar requests for an app that prints label sheets.
---

# label-maker-builder

This skill generates a complete, customized copy of a proven label-printing app
(origin: the TRI Label Maker) on the user's machine. The generated project is a
**starting point the user owns** — after generation they customize it freely,
and it ships with its own CLAUDE.md so future Claude Code sessions understand it.

Files bundled with this skill (in `reference/` next to this SKILL.md):

- `reference/tri_label_maker.py` — the full working source of the original app.
  It still carries the original TRI naming, icon/tutorial references, and Avery
  5209 geometry; the rules below say exactly what to rename, strip, and
  parameterize.
- `reference/app_spec.md` — the condensed functional spec of the app you are
  generating. Read it before generating; it defines every behavior the copy
  must keep.
- `reference/project_claude_md.md` — template for the generated project's
  CLAUDE.md, with `{{PLACEHOLDER}}` tokens you fill in.

Perform the three phases below **in order**. Do not generate anything before
Phase 1 and Phase 2 are complete.

## Phase 1 — Interview the user (before touching anything)

Ask these questions, offering the defaults. Accept "just use the defaults" as
an answer to all of them at once.

1. **Label sheet** — default is Avery 5209: 4" × 6" sheet, 7 labels of
   3.4375" × 2/3" (0.6667"), stacked vertically with no gutter, 0.65"
   top/bottom sheet margins, 0.3" side margins. If they use a different sheet,
   collect: sheet width × height, number of labels, label width × height,
   top and left margins, and the vertical gutter between labels (0 if none).
   Sanity-check that `margin_top + n*label_h + (n-1)*gutter` fits the sheet
   height before accepting.
2. **App name** — default "Label Maker". Used for the window title, the doc
   name in the print spooler, and the config folder
   (`%APPDATA%\<App Name>\config.json`).
3. **Theme** — dark (charcoal + amber, the reference default) or light.
4. **Default font** — default Lucida Sans Typewriter (a monospace font
   present on Windows). Each label can pick its own font from a dropdown in
   the style editor at runtime; this answer only sets the default. Also ask
   **max lines per label** — default 2.
5. **The three format presets** — names and styles for presets 1/2/3 (e.g.
   "Job folders" / "Vendor folders" / "Commercial folders"). Accepting the
   defaults is fine; presets are edited later in-app via the Formats dialog.

## Phase 2 — Verify the environment

1. Confirm the OS is Windows. If not, warn that direct printing (pywin32 raw
   GDI) will not work, and **stop** — unless the user explicitly wants a
   preview-only app with printing stubbed out.
2. `python --version` — need **3.10+**. If Python is missing or too old, offer
   `winget install Python.Python.3.12` (only run it with the user's approval).
3. Verify with `python -c "import win32print"`; only if the import fails, run
   `pip install pywin32` and verify again.
4. **Do not proceed to Phase 3 until the import succeeds.**

## Phase 3 — Generate the project

Ask where to put the project; default is a new folder named after the app
(e.g. `.\Label Maker\` or a snake_case variant in the current directory).
Then generate the app from `reference/tri_label_maker.py`, adapted per the
interview. The reference source is ~1,900 lines and works as-is for the
default answers — your job is a careful adaptation, not a rewrite.

### 3.1 Rename and rebrand

- Name the main file after the app in snake_case (e.g. `label_maker.py`).
- Set `APP_NAME` to the user's app name; it flows to the window title, spooler
  doc names, and message boxes automatically. `CONFIG_DIR` does **not** follow
  automatically — the reference hardcodes the literal folder name — so edit it
  (or make it derive from `APP_NAME`).
- Set `APP_VERSION = "1.0"` — the generated project starts its own history;
  don't carry over the original's version number.
- Remove `ICON_FILE`, the `resource_path()` helper, and the
  `root.iconbitmap(...)` call — the generated app has no bundled icon (the
  user can add one later; mention it in the hand-off).
- Update user-visible captions that name the sheet: the preview caption
  (`"Preview · 4 × 6 in"`), the rows caption (`"Labels · Avery 5209"`), and
  the About text. Because geometry is config-driven (§3.3), compute the size
  part from `SHEET_W`/`SHEET_H` at build time rather than hardcoding it, and
  put the sheet's display name (e.g. "Avery 5209") in a `SHEET_NAME` constant
  so it can't go stale if the user later edits `config.json`.

### 3.2 Strip the Tutorial feature (always)

The narrated video is specific to the original TRI app. Remove:

- the `TUTORIAL_FILE` constant,
- the `_mci()` helper, `tutorial_video_path()`, and the entire
  `TutorialDialog` class (the whole "Tutorial video player" section),
- `LabelApp.show_tutorial()` and the `hbtn("Tutorial", ...)` header button.

Header buttons in the generated app are: Formats / Print… / Quick Print
(Quick Print keeps the amber accent). Do not copy any tutorial media or
tutorial scripts into the skill output.

### 3.3 Make the sheet geometry config-driven

All Phase-1 geometry goes into **config defaults, not hardcoded constants**,
so the user can change sheets later by editing `config.json`:

- Add a `"geometry"` key to `default_config()` holding the interview values:
  `sheet_w, sheet_h, label_w, label_h, num_labels, gutter, margin_top,
  margin_left` (inches; `gutter` = vertical gap between labels, 0 for
  Avery 5209).
- Keep the existing module-level names (`SHEET_W`, `SHEET_H`, `LABEL_W`,
  `LABEL_H`, `NUM_LABELS`, `MARGIN_TOP`, `MARGIN_LEFT`, `MARGIN_RIGHT`,
  `ROW_PITCH`, `ALIGN_BOX_LEFT`, `ALIGN_TARGET_LR`, `ALIGN_TARGET_BOTTOM`) —
  plus one **new** module-level constant `GUTTER` (vertical gap between
  labels; the reference has none) — so the rest of the code is untouched, but
  assign them all in one `_apply_geometry(cfg)` function called at startup
  right after `load_config()` and **before any UI is built**. `load_config()`
  should heal missing/invalid geometry keys back to the interview defaults
  (mirror how it heals format slots). Recompute the derived values there:
  `MARGIN_RIGHT = SHEET_W - MARGIN_LEFT - LABEL_W`,
  `ROW_PITCH = round((LABEL_H + GUTTER) * PREVIEW_SCALE)`,
  `ALIGN_BOX_LEFT = (8.5 - SHEET_W) / 2` (subject to the wide-sheet rule in
  §3.4), `ALIGN_TARGET_LR = ALIGN_BOX_LEFT`,
  `ALIGN_TARGET_BOTTOM = ALIGN_BOX_TOP + SHEET_H`.
- Also call `_apply_geometry` once **at module scope** with the default
  geometry, so every module-level name exists even before `main()` runs —
  otherwise importing the module (e.g. to call `compute_calibration` from a
  test) raises `NameError`.
- **Label gutter support** (the original has none): the top of label *i* is
  `MARGIN_TOP + i * (LABEL_H + GUTTER)`. That expression appears in four
  places that must stay in sync — `print_sheet`, `print_alignment_sheet`,
  `update_preview`, and the row pitch. Add a single `label_top(i)` helper and
  use it in all of them.

### 3.4 Printing geometry derivations (knowledge you need for non-4x6 sheets)

- **DEVMODE custom paper size** is in **0.1 mm units**: `inches × 254`,
  rounded. The reference hardcodes 1016 × 1524 (= 4" × 6") in
  `_create_printer_dc` and matches the same values in `_find_4x6_paper`.
  Generalize both to `round(SHEET_W * 254)` / `round(SHEET_H * 254)` (keep the
  ±30 = 3 mm match tolerance, and rename `_find_4x6_paper` to something like
  `_find_sheet_paper`). Sheets are described in feed orientation (width across
  the feed); the app always forces portrait (`Orientation = 1`).
- **The alignment test always prints on plain letter paper** (that path is
  `_create_printer_dc(letter=True)` — keep it separate). The sheet outline is
  centered horizontally on the 8.5" page and its top sits at 0.25" (the
  paper's very top edge is unprintable). If the user's sheet is **wider than
  8"**, don't center — place the outline's left edge at 0.25" and let the
  right side run off; the Calibration dialog labels already display the
  computed targets, so they stay correct. If the sheet is **taller than
  10.5"**, warn that the bottom-line measurement will be off-page (X-only
  calibration still works).
- `compute_calibration()` works purely off the `ALIGN_TARGET_*` values —
  no changes needed once those are derived.

### 3.5 Preview scale

`PREVIEW_SCALE` is 90 px/inch; a 6"-tall sheet gives a ~540 px preview.
If `sheet_h × 90 > ~720` px, lower `PREVIEW_SCALE` (keep it an integer) so the
window fits on screen. Watch the other side of the coupling: entry rows are
frozen to `ROW_PITCH` px each, and rows need roughly `24 × max_lines + 10` px
to be usable. If `ROW_PITCH` lands below that, tell the user their label pitch
is too small for comfortable entry at this scale and pick the best compromise
(smaller row font/padding first, then accept preview taller than 720 px).

### 3.6 Theme

- **Dark** (default): keep the `CLR_*` palette exactly as in the reference.
- **Light**: replace the palette, keeping the same constant names. A tested
  starting point:

  ```python
  CLR_BG = "#f2f0ec"; CLR_HEADER = "#e6e2da"; CLR_ROW = "#ffffff"
  CLR_ROW_ACTIVE = "#fff3dd"; CLR_FIELD = "#ffffff"; CLR_BTN = "#dcd8d0"
  CLR_BTN_HOVER = "#cfcabf"; CLR_FG = "#26241f"; CLR_DIM = "#6f6a60"
  CLR_ACCENT = "#d98a00"; CLR_ACCENT_ACTIVE = "#b87400"
  CLR_ON_ACCENT = "#ffffff"; CLR_PREVIEW_BG = "#e0dcd3"
  ```

  Also: make `enable_dark_titlebar()` a no-op (or don't call it), and sweep
  the code for hardcoded `"#ffffff"` / near-black literals used as hover/text
  colors (`style_button`, `apply_theme`, header title label, menu colors, the
  preview drop-shadow) and re-fit them to the light palette.

### 3.7 Font and max lines

- Set `FONT_NAME` to the interview answer — it is the **default** family for
  the per-label font dropdown (and `DEFAULT_STYLE["font"]`); keep
  `FALLBACK_FONT = "Courier New"` and the existing "fall back if not
  installed" check in `LabelApp.__init__`. The font selector itself needs no
  adaptation.
- Set `MAX_LINES` to the interview answer, and make everything that assumes 2
  follow it: build `DEFAULT_STYLE["lines"]` programmatically with `MAX_LINES`
  entries, set the entry `tk.Text(height=MAX_LINES)`, and update the About
  text. The layout engine, style editor, and line-limit key handling already
  iterate over `MAX_LINES`.

### 3.8 Format presets

Put the interview's three preset names (and any style choices the user
described) into `DEFAULT_FORMATS`. A preset holds **style only, never text**.
Don't bake anything machine-specific into the code: no printer names, no
calibration values — those live only in the user's own `config.json` at
runtime (`printer` starts `None`, `calibration` starts `{x: 0, y: 0}`).

### 3.9 Architecture — keep intact (do not "improve" these)

- **Single layout engine**: `layout_label()` computes text runs in inches,
  backend-agnostic via a measurer (point size → width/line-height in inches).
  `TkMeasurer` backs the preview, `GdiMeasurer` backs the printer DC — that is
  what keeps preview and print WYSIWYG-consistent. The font family is part of
  each label's style and flows through both measurers (metrics cached per
  family). Layout behavior changes happen only there, never in preview or
  print code separately.
- **Raw-GDI printing** with a DEVMODE forcing portrait sheet media (driver
  paper code if one matches, else custom size), drawing at device DPI,
  subtracting the printer's physical unprintable offset
  (`PHYSICALOFFSETX/Y`) and adding the stored calibration offset. The driver
  dialog is never opened. The alignment test keeps its separate letter-media
  path.
- **Entry-rows-to-preview vertical coupling**: rows are locked to the
  preview's label pitch (`ROW_PITCH`, the top spacer matching the sheet's
  first-label offset, fixed row height via `pack_propagate(False)` after
  measuring natural width) so row N sits level with label N. Preserve it when
  adapting geometry.
- **Keyboard model**: Enter/Tab = next label, Shift+Enter = newline (up to
  max lines), **Ctrl+Tab = literal tab character** (required for the
  tab-L/R split mode), Ctrl+1/2/3 = assign format, Ctrl+P = Print…,
  Ctrl+Shift+P = Quick Print, Ctrl+N = new sheet.
- **Drag-reorder moves data, not widgets**: the ⋮ handle drag re-inserts the
  label's dump (text/style/preset) at the drop position; the fixed row frames
  never move, which is what preserves the row-to-preview coupling. The Style…
  dialog's Format row uses the same link/detach rule as the row's 1/2/3
  buttons (untouched OK links, a tweak after picking detaches).
- **Toast confirmations** (`show_toast`) for Quick Print (5 s) and the
  alignment test (10 s).
- Persisted state (formats, printer, calibration, autosaved sheet, named
  label sets, geometry) all lives in `%APPDATA%\<App Name>\config.json`.

### 3.10 Smoke test — generate, then prove it

Keep the `--smoke` flag and `_run_smoke()`. If `MAX_LINES` ≠ 2, adapt the
line-limit assertions to the new value; if geometry changed, the calibration
assertions still hold because they use the `ALIGN_TARGET_*` constants. After
generating, run:

```
python <app_file>.py --smoke
```

It must print `SMOKE OK` and exit 0 (it briefly opens a real window — needs a
desktop session). **Fix failures before declaring the project done.**

### 3.11 Write the project's CLAUDE.md

Fill `reference/project_claude_md.md` — replace every `{{PLACEHOLDER}}` with
the real values (app name, file name, geometry table, font, max lines, theme)
— and write it as `CLAUDE.md` in the generated project folder.

### 3.12 Hand-off

Tell the user, concretely:

- Run it with `python <app_file>.py` (needs `pip install pywin32` for
  printing; the GUI runs without it).
- First print: use **Print…**, pick their label printer, and leave "Remember
  as Quick Print printer" checked — after that **Quick Print**
  (Ctrl+Shift+P) is zero-dialog.
- If prints land off the labels: **Print menu → Print alignment test** (plain
  letter paper), measure the solid box with a ruler, then **Print menu →
  Calibration…** → enter the measurements → Calculate → Save, and reprint the
  test to verify.
- The project has its own CLAUDE.md — they can open Claude Code in the
  project folder and ask for any customization.

### 3.13 Optional: build a standalone EXE

Offer (don't push) a PyInstaller build:

```
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name <AppName> <app_file>.py
```

Add `--icon <file>.ico` only if the user supplies an icon. If they take the
offer, apply this build hygiene:

1. Before building: kill any running copy of the EXE, delete the old
   `dist\<AppName>.exe`, and **verify the delete succeeded** — if the EXE is
   locked, PyInstaller fails quietly and the stale EXE keeps looking like a
   fresh build.
2. After building: confirm the EXE's LastWriteTime matches the current clock,
   then run `dist\<AppName>.exe --smoke` and check the exit code is 0 (it's a
   windowed app, so there is no console output — the exit code is the
   verdict).
