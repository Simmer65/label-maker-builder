# Functional spec — generated label-printing app

The behaviors the generated app must keep. Geometry values below are the
Avery 5209 defaults; the generated app takes its actual geometry from the
Phase-1 interview and stores it in config (see SKILL.md §3.3).

## Label sheet geometry (default: Avery 5209)

- Sheet: 4.0" wide × 6.0" tall — the printer is fed sheet-size media and the
  job prints at actual size, no scaling.
- 7 labels of 3.4375" × 2/3", stacked vertically with no gutter, centered on
  the sheet: 0.65" top/bottom margins, 0.3" side margins.
- Blank rows print nothing, so a partially used sheet is reused by leaving the
  used positions empty.

## Entry and layout

- The GUI starts blank with one editable text entry per label position, each
  sitting level with its label in the live WYSIWYG preview.
- A label holds at most MAX_LINES lines (default 2); extra lines are
  blocked/truncated, including on paste.
- Font: per-label font family, chosen from a dropdown of installed fonts at
  the top of the Style… dialog. `FONT_NAME` is the default family (default
  Lucida Sans Typewriter, falls back to Courier New if not installed); the
  font is part of the style, so presets carry it and it flows through both
  measurers (metrics cached per family) to keep preview and print WYSIWYG.

## Per-label style (the Style… button)

- A **Format:** row at the top of the dialog shows the three preset buttons
  (number + name). Clicking one loads that format's settings into the
  controls; an untouched OK **links** the label to the preset, while tweaking
  any control after picking saves the style **detached** (same rule as the
  row's 1/2/3 buttons). A label already linked opens with its format
  pre-highlighted.
- A font-family dropdown (see Entry and layout above).
- Per line: font size — explicit points or auto-sized to fit — and a
  bold/normal toggle. Auto-sized lines share the height left over after
  fixed-size lines claim theirs. Typing in a line's size spinbox
  automatically switches that line from Auto to Points mode.
- A configurable vertical gutter (inches, default 0) between lines;
  auto-sizing accounts for it.
- Horizontal alignment: left, center, right (left/right use a margin inset),
  plus **tab left/right mode**: the label's single text field is split at a
  tab character — text before the tab is left-aligned at the left margin,
  text after it right-aligned at the right margin
  (e.g. `Smith, John<TAB>2026`).
- Vertical alignment: top, middle, bottom.

## Format presets 1 / 2 / 3

- Three user-defined style presets, **global to the app** — the same three
  definitions for every label, not per-row.
- A preset holds style settings only (font family, per-line size/auto and
  bold, line gutter, H alignment incl. tab mode, V alignment, margin) —
  **never text**.
- Presets are defined and edited in the Formats dialog (not by saving from a
  label) and persist in the config file.
- Labels stay **linked** to their assigned preset number: editing a format's
  definition immediately restyles every label assigned to it. Manually
  tweaking a linked label via its Style… button **detaches** it from the
  preset.
- A "default format for blank labels" setting styles new/blank labels, and an
  "apply format N to all labels" action (Formats menu) formats the whole
  sheet with no per-row clicks.

## Per-label row buttons

- **▲ / ▼** duplicate the label's text *and* style to the adjacent position.
- **✕** clears the label's text; its style and format link are kept.
- **⋮ drag handle** (far right of the row): drag vertically to move the label
  to a new position — the dragged label's **data** (text, style, preset link)
  is re-inserted at the drop position and the labels in between shift, while
  the row widgets stay put (preserving the row-to-preview coupling). Amber
  accent marks show the source and drop target while dragging; the status bar
  confirms the move.

## Printing

- **Print…** (Ctrl+P) opens printer selection with a "Remember as Quick Print
  printer" checkbox.
- **Quick Print** (Ctrl+Shift+P) sends directly to the stored printer with
  zero dialogs. The paper setup (sheet-size media, portrait) is applied
  programmatically via DEVMODE **on every job** — never rely on the driver's
  saved defaults, never open the driver dialog. Falls back to the system
  default printer if none is stored.
- Both Quick Print and the alignment test confirm with an auto-hiding toast
  (`show_toast`, 5 s / 10 s).
- Printing is raw GDI at device DPI: subtract the printer's physical
  unprintable offset, add the stored calibration offset.

## Alignment test + calibration

- The alignment test prints on plain **letter paper** (a separate
  letter-media DC path), all on one page:
  - a **solid** sheet outline with the current calibration offset applied —
    the box to measure — and a **dashed** outline (1/4" dashes, 1/8" gaps,
    drawn segment-by-segment because GDI stock pens can't do custom dash
    lengths) at raw printer position with no offset;
  - centered horizontally (default targets: 2.250" left/right margins) with
    the outline top at 0.25" (the paper's top edge is unprintable) and the
    bottom line at 6.250" from the paper top (targets derive from the actual
    sheet geometry);
  - the label boxes drawn inside the solid outline, numbered;
  - a legend below the last label ("Measure the SOLID box.");
  - centered instructions below the box pointing at the Calibration dialog.
    There is no printed measuring worksheet.
- The **Calibration dialog** computes the X/Y offsets from ruler
  measurements — paper left edge → box left side (target 2.250"), paper
  right edge → box right side (optional, averaged), paper top edge → box
  bottom line (target 6.250", optional) — via `compute_calibration()`.
  Direct X/Y entry remains available in the same dialog. Offsets persist in
  config and apply to every job.

## Sheet operations

- **New sheet** (Ctrl+N) clears all labels (confirms if any have text).
- **Start at position N** shifts the entered labels down so a half-used sheet
  can start at the first unused label (warns if labels would fall off).
- **Auto-save**: the current sheet is saved on exit and restored on launch.
- **Named label sets** can be saved, loaded, and deleted (e.g. "standard
  vendor folders").

## Keyboard model

- Enter / Tab — next label (wraps); Shift+Tab — previous.
- Shift+Enter — newline (blocked at the line limit).
- **Ctrl+Tab — literal tab character** (required for tab-L/R split mode).
- Ctrl+1/2/3 — assign that format preset to the focused label.
- Ctrl+P — Print…; Ctrl+Shift+P — Quick Print; Ctrl+N — new sheet.

## Persistence

Everything lives in `%APPDATA%\<App Name>\config.json`: formats, default
format, printer name, calibration offsets, autosaved sheet, named label
sets, and (in generated apps) the sheet geometry. Loading heals missing or
partial entries back to defaults.

## Smoke test

`python <app_file>.py --smoke` builds the real UI, exercises the layout
engine (tab split, per-line auto/fixed/bold, old-config style migration,
line limit, gutter spacing), the style-editor mode flip, `compute_calibration`
round-trips, and the toast — prints `SMOKE OK` and exits 0.

## Intentionally omitted — implement only if the user asks

- Pasting a multi-line list to fill labels.
- Batches beyond one sheet (the app handles exactly one sheet per print).
- Sequence/series generation (auto-incrementing numbered labels).
- The original app's narrated tutorial video and its player.
