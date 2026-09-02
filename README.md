# label-maker-builder

A [Claude Code](https://claude.com/claude-code) skill that generates a
customized Windows label-printing desktop app (Python/tkinter + pywin32,
direct-to-printer GDI).

The generated app prints Avery-style file-folder label sheets on 4x6 media
(or any sheet geometry you describe) with:

- a live WYSIWYG preview — what you see is exactly what prints
- per-label styling: font, per-line size (fixed or auto-fit), bold,
  alignment (including a tab left/right split), and three global format
  presets
- drag-and-drop label reordering
- zero-dialog **Quick Print** (paper size forced programmatically — the
  printer driver dialog is never opened)
- a printed alignment test + calibration dialog so output lands exactly on
  the labels
- auto-save, named label sets, and partial-sheet reuse

## Installing the skill

Copy this folder into your Claude Code skills directory:

- **Per project**: `<your project>\.claude\skills\label-maker-builder\`
- **All projects**: `%USERPROFILE%\.claude\skills\label-maker-builder\`

Then, in Claude Code, ask something like *"build me a label printing app"*
or invoke it directly with `/label-maker-builder`. The skill interviews you
(sheet geometry, app name, theme, default font, presets), verifies Python
and pywin32, then generates a complete project you own — with its own
CLAUDE.md so future Claude Code sessions understand it.

## Contents

- `SKILL.md` — the skill definition and generation instructions
- `reference/tri_label_maker.py` — the full working source of the origin
  app (TRI Label Maker) that generation adapts
- `reference/app_spec.md` — the functional spec the generated app must keep
- `reference/project_claude_md.md` — template for the generated project's
  CLAUDE.md

## Requirements (for the generated app)

- Windows (printing is raw GDI via pywin32)
- Python 3.10+ and `pip install pywin32`
