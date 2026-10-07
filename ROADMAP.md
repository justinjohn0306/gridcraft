# SheetCraft roadmap

SheetCraft aims at full Microsoft Excel parity — and to be better: faster, open (XLSX is the
native format), fully scriptable by agents (MCP + control channel + CLI), and available on the
web.

## Status (2026-10-07)

**Parity estimate**

| Measure | Value | Notes |
|---|---|---|
| Ribbon/menu command catalog | **60%** (169 / 281) | `cargo xtask parity` → [`docs/parity.md`](docs/parity.md). Home is 97%. |
| Worksheet functions | **~93%** (478 + 32 evaluator built-ins of ~545) | Missing: ODDF*/ODDL*, BAHTTEXT, IMAGE, TRANSLATE, COPILOT/PY, STOCKHISTORY, RTD |
| Feature depth (weighted) | **~38%** | Pivot tables, Power Query, print/PDF, page layout view, chart formatting depth, ink, macros/scripts, threaded collaboration and solver/analysis tools are the big missing areas |

**Estimated remaining effort to full parity: ~220 wall-clock hours of Claude Opus 5.5 agent
work** (≈ 9–10 days running continuously with 3–4 parallel agents), broken down below.

**Working today**
- Engine: A1/R1C1 formula parser with reference adjustment, dependency-graph recalculation with
  dynamic arrays (spill + `#SPILL!`), cycle detection, LET/LAMBDA/MAP/REDUCE/SCAN/BYROW/BYCOL/
  MAKEARRAY, structured table references, defined names, INDIRECT/OFFSET/INDEX references,
  478 worksheet functions (math, statistics and distributions, text and regex, lookup incl.
  XLOOKUP/FILTER/SORT/UNIQUE/VSTACK, dates, financial incl. bonds, engineering incl. complex
  numbers and Bessel, database, information).
- Number formats: Excel's full format-code language (sections, conditions, colours, dates and
  elapsed time, fractions, scientific, accounting fill), General narrowing to column width.
- Files: XLSX read/write (styles, themes, merges, CF, validation, tables, comments, hyperlinks,
  charts, pictures, sparklines, print settings, defined names, shared/array/dynamic formulas),
  CSV/TSV with encoding and delimiter detection, HTML export, JSON.
- Editing: cell entry with type detection and auto-formats, in-cell and formula-bar editing,
  point mode (click/drag/arrow references), reference colouring, F4 anchor cycling, function
  autocomplete and argument hints, copy/cut/paste and Paste Special (values, formats, transpose,
  operations, link, skip blanks), AutoFill series (numbers, dates, months, weekdays, custom lists,
  text+number), Fill Series, Flash Fill, find/replace with wildcards, Go To / Go To Special,
  undo/redo with history, format painter.
- Formatting: fonts, fills, all border styles, alignment (wrap, indent, rotation, merge), number
  formats, 47 cell styles, 60 table styles, conditional formatting (cell rules, text, dates,
  duplicates, top/bottom, averages, data bars, colour scales, icon sets, formulas), automatic
  row heights, autofit, hide/unhide, freeze panes.
- Data: sort (multi-level, by colour, custom lists), AutoFilter (values, custom, top 10,
  average, colour), tables with totals rows, remove duplicates, text to columns, data validation
  with error alerts, grouping/outline, subtotals.
- Charts: column, bar, line, area, pie, doughnut, scatter, bubble, radar, histogram, waterfall,
  funnel, treemap, sunburst, box & whisker, stock and combo charts; sparklines.
- UI (egui): Excel-style title bar with Quick Access Toolbar and AutoSave, the full ribbon with
  galleries (cell styles, table styles, colour palettes, conditional formatting menus),
  contextual Table/Chart Design tabs, formula bar with Name Box, virtualised grid (smooth
  scrolling, zoom, frozen panes, text overflow), sheet tabs (rename, reorder, colour, hide),
  status bar with Sum/Count/Average and zoom, dialogs (Format Cells, Formula Builder, Find and
  Replace, Sort, Name Manager, Data Validation, Paste Special…), command palette, dark mode.
- Agents: every action is an engine command; JSON control channel in the desktop app (pointer,
  keyboard, dialogs, screenshots); MCP server (headless or bridged); CLI (`info`, `convert`,
  `eval`, `cat`, `run`, `commands`, `functions`, `mcp`, `send`).
- Platforms: macOS, Windows, Linux, FreeBSD; web (WASM via trunk). Release workflows for
  signed/notarized macOS universal DMG, signed Windows x64/x86 MSI + zip, Linux AppImage/deb/rpm/
  tar.gz + Flatpak manifest, FreeBSD tarball and a web zip.

## Milestones

| # | Milestone | Status | Est. hours left |
|---|---|---|---|
| M0 | Foundation: core, number formats, formula language, functions, model, calc, XLSX/CSV, engine | ✅ done | — |
| M1 | Excel look: chrome, ribbon, formula bar, grid, tabs, status bar, control channel | ✅ done (polish continues) | 8 |
| M2 | Formatting completeness (Format Cells parity, borders drawing, themes fonts/effects, styles authoring) | 🟡 most | 10 |
| M3 | Editing power (drag-move/copy with Alt, insert copied cells, AutoComplete in column, AutoCorrect, spelling) | 🟡 most | 12 |
| M4 | Data (advanced filter, consolidate, what-if: goal seek, scenarios, data tables; Get & Transform lite) | 🟡 partial | 24 |
| M5 | Formulas tab (trace arrows on the grid, watch window, evaluate stepper UI, error-check UI, remaining functions) | 🟡 partial | 12 |
| M6 | Conditional formatting manager parity (rule editor dialog, all icon sets drawn, stop-if-true UI) | 🟡 most | 6 |
| M7 | Charts parity (element selection & Format pane, axes options, trendlines, error bars, chart sheets, styles/colours galleries, PivotCharts) | 🟡 partial | 28 |
| M8 | Insert & Review (icons library of our own, SmartArt-lite, equations, threaded comments pane, show changes, accessibility pane) | 🟡 partial | 18 |
| M9 | Page layout & print (Page Layout and Page Break Preview views, header/footer editor, print to PDF, print preview) | ⬜ | 26 |
| M10 | PivotTables (field list, layouts, grouping, calculated fields, slicers, timelines), Solver, Analysis ToolPak | ⬜ | 40 |
| M11 | Automation (record actions → replayable scripts, Rhai script editor, form controls) | ⬜ | 14 |
| M12 | Performance (1M-row files, parallel recalc, incremental geometry caches), accessibility, i18n | 🟡 partial | 16 |
| M13 | Release & polish (signed builds green on every platform, side-by-side pixel tuning, README screenshots) | 🟡 tooling ready | 6 |
| | **Total** | | **≈ 220** |

## How to measure

- `cargo xtask parity` — catalog coverage (feature names from Excel's ribbon and menus).
- `cargo test --workspace` — engine, formula, function, XLSX and UI behaviour.
- `cargo run --release -p sheetcraft-ui-egui --example snapshot -- --sample sales out.png` — look at the UI offscreen.
