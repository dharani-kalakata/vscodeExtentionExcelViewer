# JSON FortuneSheet Editor

A VS Code custom text editor that displays JSON and CSV files in a FortuneSheet-powered spreadsheet with full type preservation and support for structured data shapes.

## Setup

1. Install dependencies:
   ```bash
   npm install
   ```
2. Build the webview bundle:
   ```bash
   npm run build:webview
   ```
3. Launch the extension for development (F5 in VS Code) and open a `.json` file. Choose **Open With → JSON FortuneSheet** when prompted.

## Features

- **Unified editor** for `.json`, `.csv`, and `.xlsx` files with FortuneSheet spreadsheet UI
- Backed by [FortuneSheet](https://github.com/ruilisi/fortune-sheet) (React)
- **One right-click "Open With → FortuneSheet Viewer"** works for all three file types
- Same layout and features for all file types: filters, formulas, sorting, editing
- Type preservation for JSON: numbers, booleans, null, and strings round-trip correctly
- Supports structured JSON shapes:
  - Top-level arrays of objects (table view)
  - Object-of-objects with "key" column (table view)
  - Key-value pairs (two-column view)
  - Optional wrapper objects with `data` property
- CSV files: full edit and save support
- Excel files: view and explore data (read-only; edits show warning)
- Keyboard shortcuts: `Ctrl+S` to save (JSON/CSV only)
- Dirty tracking: prompts to save when closing with unsaved edits

## Development notes

- Webview source lives in `webview-src/` and is bundled to `media/webview.js` and `media/webview.css` via esbuild (`scripts/build-webview.js`).
- The React app uses `Workbook` from `@fortune-sheet/react` with `onChange` and `onOp` to track edits and notify the extension.
- Type mapping is tracked by key paths (e.g., `[0].price`, `cbd_score_total.scale`) to cast edited strings back to JSON values on save.
- JSON shape validation is centralized in `json-sheet-rules.js` to reject deeply nested structures.

## Testing

No automated tests are defined yet. You can manually validate by:
1. Opening a JSON/CSV file via "Open With → JSON FortuneSheet"
2. Editing cells, toggling booleans, changing numbers
3. Pressing `Ctrl+S` to save or closing the editor to verify save prompts work correctly
