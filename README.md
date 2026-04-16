# CAP 741 Logbook App

This project is a small browser-based app for editing and printing CAP 741 Section 3.1 logbook pages.

It is built as a single-page app with:

- `cap741-logbook-fixed.html`: the page shell and modal containers
- `cap741-logbook-fixed.css`: the printable CAP 741 layout and UI styling
- `cap741-logbook-fixed.js`: all application logic
- `cap741-data.xlsx`: the workbook used as the app's persistent storage

## What the app does

The app reads data from `cap741-data.xlsx`, turns it into editable CAP 741 pages, lets the user edit those pages in the browser, and then writes the updated state back into the same workbook.

The JavaScript file is large, but the behavior is straightforward:

1. Load workbook sheets into memory.
2. Normalize rows into a consistent row model.
3. Group rows by `Aircraft Type + Chapter`.
4. Paginate those groups into fixed CAP 741 page slots.
5. Render the pages as editable HTML.
6. Auto-save or manually save back to Excel.

## Workbook structure

The workbook is treated like a lightweight database with five sheets:

- `Logbook`: the main maintenance experience rows
- `Aircraft`: registration-to-aircraft-type mapping
- `Chapters`: ATA chapter reference data
- `Supervisors`: supervisor names, stamps, and licence numbers
- `Info`: owner metadata such as name and signature

## Core mental model

The most important thing to know is that the UI is not editing the Excel file directly.

Instead, the app keeps an in-memory `rows` array. Every edit updates that array first. The screen is then re-rendered from `rows`, and saves write the current in-memory state back to `cap741-data.xlsx`.

That means:

- `rows` is the working state
- `renderAll()` is the main screen refresh
- `writeXlsx()` is the persistence step

## Important code sections

Use these sections in `cap741-logbook-fixed.js` as your map:

- `Filter state`: filter chips, draft filters, and page filtering
- `Row model`: creation, normalization, lookup, and grouping of rows
- `Render helpers`: grouping, pagination, and page HTML generation
- `Row editing`: how inputs update in-memory rows
- `XLSX core`: load from workbook and build workbook output
- `Event handlers`: button clicks, inline edits, modal behavior
- `Startup`: boot sequence and workbook loading

## Data flow

Here is the full lifecycle in plain English:

1. On startup, the app tries to load `cap741-data.xlsx`.
2. Workbook sheets are parsed into browser-side state.
3. `renderAll()` rebuilds the visible pages from that state.
4. User edits a field.
5. `updateRowFromEditor()` syncs that field into the matching row object.
6. The app re-renders pages when layout-affecting fields change.
7. Auto-save writes the rebuilt workbook back to disk.

## Why the file feels complex

Most of the complexity comes from three responsibilities living in one file:

- Excel import/export
- printable page layout logic
- inline editing and modal interactions

There is not a deep architecture here. It is mostly one stateful UI script with a few clear subsystems.

## Best places to start reading

If you want to understand the project quickly, read in this order:

1. `loadWorkbookFromArrayBuffer()`
2. `groupRows()`
3. `paginate()`
4. `renderAll()`
5. `updateRowFromEditor()`
6. `writeXlsx()`

That path gives you the whole app: load, shape data, display it, edit it, save it.

## Simplification ideas for a future refactor

If we want to simplify the code further later, the cleanest next steps would be:

- split workbook I/O into a separate file
- split rendering helpers from event handlers
- replace string-built HTML with smaller rendering helpers or templates
- move static aircraft reference data into JSON
- introduce a small state object instead of many top-level `var` declarations

These are optional improvements, not urgent fixes. The current code is workable once you understand the `rows -> renderAll() -> writeXlsx()` flow.
