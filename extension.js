const vscode = require('vscode');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const jsonRules = require('./json-sheet-rules');

function activate(context) {
  const provider = new FortuneSheetEditorProvider(context);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider('jsonFortuneSheet.editor', provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: false,
    }),
  );
}

class FortuneSheetDocument {
  constructor(uri, fileType, content, initialText, initialMatrix) {
    this.uri = uri;
    this.fileType = fileType;
    this.content = content;
    this.currentText = typeof initialText === 'string' ? initialText : '';
    this.currentMatrix = Array.isArray(initialMatrix)
      ? initialMatrix
      : content && content.dataKind === 'xlsx' && Array.isArray(content.matrix)
        ? content.matrix
        : content && content.dataKind === 'xlsx' && Array.isArray(content.sheets) && Array.isArray(content.sheets[0]?.matrix)
          ? content.sheets[0].matrix
          : [[]];
  }

  markSaved() {
    // VS Code clears dirty state when saveCustomDocument completes.
    // This method exists so the provider can keep internal state if needed.
  }

  resetFrom(otherDocument) {
    this.fileType = otherDocument.fileType;
    this.content = otherDocument.content;
    this.currentText = otherDocument.currentText;
    this.currentMatrix = otherDocument.currentMatrix;
  }

  dispose() {}
}

class FortuneSheetEditorProvider {
  constructor(context) {
    this.extensionUri = context.extensionUri;
    this.globalStorageUri = context.globalStorageUri;
    this._onDidChangeCustomDocument = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;
    this._webviews = new Map();
  }

  async openCustomDocument(uri, openContext, token) {
    const fileType = detectKind(uri.fsPath);
    let content;
    let initialText;
    let initialMatrix;

    try {
      const fileContent = openContext?.untitledDocumentData
        ? openContext.untitledDocumentData
        : openContext?.backupId
          ? await vscode.workspace.fs.readFile(vscode.Uri.parse(openContext.backupId))
          : await vscode.workspace.fs.readFile(uri);
      const buffer = Buffer.from(fileContent);

      if (fileType === 'xlsx') {
        try {
          const workbook = XLSX.read(buffer, { type: 'buffer' });
          // Multi-sheet support: collect all sheets
          const sheets = workbook.SheetNames.map((sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            return { name: sheetName, matrix };
          });
          content = { dataKind: 'xlsx', sheets };
          initialMatrix = sheets[0]?.matrix || [[]];
        } catch (error) {
          console.error('Error reading Excel file:', error);
          content = { dataKind: 'xlsx', sheets: [{ name: 'Sheet1', matrix: [[]] }] };
          initialMatrix = [[]];
        }
      } else if (fileType === 'csv') {
        const text = buffer.toString('utf8');
        initialText = text;
        content = parseCsv(text);
      } else {
        const text = buffer.toString('utf8');
        initialText = text;
        try {
          content = JSON.parse(text);
        } catch (error) {
          console.warn('Invalid JSON, using empty object');
          content = {};
        }
      }
    } catch (error) {
      console.error('Error reading file:', error);
      if (fileType === 'xlsx') {
        content = { dataKind: 'xlsx', sheets: [{ name: 'Sheet1', matrix: [[]] }] };
        initialMatrix = [[]];
      } else if (fileType === 'csv') {
        content = { dataKind: 'csv', matrix: [[]], text: '' };
        initialText = '';
      } else {
        content = {};
        initialText = '{}';
      }
    }

    return new FortuneSheetDocument(uri, fileType, content, initialText, initialMatrix);
  }

  async resolveCustomEditor(document, webviewPanel, token) {
    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: [
        vscode.Uri.joinPath(this.extensionUri, 'media'),
        vscode.Uri.joinPath(this.extensionUri, 'node_modules'),
      ],
    };

    webviewPanel.webview.html = this.getHtmlForWebview(webviewPanel.webview);

    this._trackWebview(document, webviewPanel);

    const updateWebview = () => {
      if (!document || !document.content) {
        vscode.window.showErrorMessage('Failed to load document content');
        return;
      }
      const payload = toSheetPayloadFromContent(document.content);
      webviewPanel.webview.postMessage({ type: 'init', payload });
    };

    webviewPanel.webview.onDidReceiveMessage(async (message) => {
      switch (message.type) {
        case 'ready': {
          updateWebview();
          break;
        }
        case 'webviewError': {
          const msg = message?.message ? String(message.message) : 'Webview error';
          const stack = message?.stack ? String(message.stack) : '';
          console.error('Webview error:', msg, stack);
          vscode.window.showErrorMessage(`FortuneSheet Viewer error: ${msg}`);
          break;
        }
        case 'edit': {
          this._applyEdit(document, message);
          break;
        }
        case 'save': {
          // Ensure the latest content is recorded before save.
          this._applyEdit(document, message);
          // Trigger VS Code's save pipeline so dirty state is cleared.
          await vscode.commands.executeCommand('workbench.action.files.save');
          break;
        }
        default:
          break;
      }
    });
  }

  async saveCustomDocument(document, cancellation) {
    if (document.fileType === 'xlsx') {
      // Save all sheets if present
      if (document.content && document.content.sheets) {
        await this.saveXlsxFileMulti(document.uri, document.content.sheets);
      } else {
        await this.saveXlsxFile(document.uri, document.currentMatrix);
      }
    } else {
      await this.saveTextFile(document.uri, document.currentText ?? '', document.fileType);
    }
    document.markSaved();
  }

  async saveCustomDocumentAs(document, destination, cancellation) {
    if (document.fileType === 'xlsx') {
      if (document.content && document.content.sheets) {
        await this.saveXlsxFileMulti(destination, document.content.sheets);
      } else {
        await this.saveXlsxFile(destination, document.currentMatrix);
      }
    } else {
      await this.saveTextFile(destination, document.currentText ?? '', document.fileType);
    }
    // Don't mark as saved: this is a different URI.
  }

  // Save all sheets to Excel file
  async saveXlsxFileMulti(uri, sheets) {
    const workbook = XLSX.utils.book_new();
    for (const sheet of sheets) {
      const safeMatrix = Array.isArray(sheet?.matrix) ? sheet.matrix : [[]];
      const normalized = normalizeMatrixForXlsx(safeMatrix);
      const ws = XLSX.utils.aoa_to_sheet(normalized);
      XLSX.utils.book_append_sheet(workbook, ws, String(sheet?.name || 'Sheet1'));
    }
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    await vscode.workspace.fs.writeFile(uri, buffer);
  }

  async revertCustomDocument(document, cancellation) {
    const fresh = await this.openCustomDocument(document.uri, { backupId: undefined, untitledDocumentData: undefined }, cancellation);
    document.resetFrom(fresh);
    this._updateAllWebviews(document);
  }

  async backupCustomDocument(document, context, cancellation) {
    // Persist current state to the suggested destination.
    const destination = context.destination;
    try {
      if (document.fileType === 'xlsx') {
        if (document.content && Array.isArray(document.content.sheets)) {
          await this.saveXlsxFileMulti(destination, document.content.sheets);
        } else {
          const safeMatrix = document.currentMatrix;
          const normalized = normalizeMatrixForXlsx(safeMatrix);
          const worksheet = XLSX.utils.aoa_to_sheet(normalized);
          const workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
          const wbout = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
          await vscode.workspace.fs.writeFile(destination, wbout);
        }
      } else {
        const buffer = Buffer.from(document.currentText ?? '', 'utf8');
        await vscode.workspace.fs.writeFile(destination, buffer);
      }
    } catch (error) {
      console.error('Backup failed:', error);
    }

    const id = destination.toString();
    return {
      id,
      delete: () => {
        vscode.workspace.fs.delete(destination).then(
          () => {},
          () => {},
        );
      },
    };
  }

  _applyEdit(document, message) {
    if (!document) {
      return;
    }

    if (document.fileType === 'xlsx') {
      const xlsxSheets = Array.isArray(message.xlsxSheets) ? message.xlsxSheets : null;
      if (xlsxSheets) {
        document.content = { dataKind: 'xlsx', sheets: xlsxSheets };
        document.currentMatrix = Array.isArray(xlsxSheets[0]?.matrix) ? xlsxSheets[0].matrix : [[]];
        document.currentText = typeof message.text === 'string' ? message.text : '';
      } else {
        // Backward compatibility: older webview sends a single matrix
        const matrix = Array.isArray(message.matrix) ? message.matrix : [[]];
        document.currentMatrix = matrix;
        document.content = { dataKind: 'xlsx', sheets: [{ name: 'Sheet1', matrix }] };
        document.currentText = typeof message.text === 'string' ? message.text : JSON.stringify(matrix);
      }
    } else {
      const text = typeof message.text === 'string' ? message.text : '';
      document.currentText = text;
      if (document.fileType === 'csv') {
        document.content = parseCsv(text);
      } else {
        try {
          document.content = JSON.parse(text || '{}');
        } catch {
          // Keep previous content if JSON invalid; still allow saving raw text.
        }
      }
    }

    // Signal a content change so VS Code marks the editor dirty (no custom undo/redo stack).
    this._onDidChangeCustomDocument.fire({ document });
  }

  _trackWebview(document, webviewPanel) {
    const key = document.uri.toString();
    let set = this._webviews.get(key);
    if (!set) {
      set = new Set();
      this._webviews.set(key, set);
    }
    set.add(webviewPanel);
    webviewPanel.onDidDispose(() => {
      const current = this._webviews.get(key);
      if (current) {
        current.delete(webviewPanel);
        if (current.size === 0) {
          this._webviews.delete(key);
        }
      }
    });
  }

  _updateAllWebviews(document) {
    const key = document.uri.toString();
    const set = this._webviews.get(key);
    if (!set || set.size === 0) {
      return;
    }
    const payload = toSheetPayloadFromContent(document.content);
    for (const panel of set) {
      try {
        panel.webview.postMessage({ type: 'init', payload });
      } catch {
        // ignore
      }
    }
  }

  async saveTextFile(uri, content, fileType) {
    try {
      const buffer = Buffer.from(content, 'utf8');
      await vscode.workspace.fs.writeFile(uri, buffer);
    } catch (error) {
      vscode.window.showErrorMessage(`Failed to save file: ${error.message}`);
    }
  }

  async saveXlsxFile(uri, matrix) {
    try {
      const safeMatrix = Array.isArray(matrix) ? matrix : [[]];
      const normalized = normalizeMatrixForXlsx(safeMatrix);
      const worksheet = XLSX.utils.aoa_to_sheet(normalized);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
      const wbout = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
      await vscode.workspace.fs.writeFile(uri, wbout);
    } catch (error) {
      vscode.window.showErrorMessage(`Failed to save Excel file: ${error.message}`);
    }
  }

  getHtmlForWebview(webview) {
    const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(this.extensionUri, 'media', 'webview.js'));
    const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(this.extensionUri, 'media', 'webview.css'));
    const nonce = getNonce();
    const csp = [
      "default-src 'none'",
      `img-src ${webview.cspSource} data:`,
      `style-src ${webview.cspSource} 'unsafe-inline'`,
      `script-src 'nonce-${nonce}'`,
    ].join('; ');

    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta http-equiv="Content-Security-Policy" content="${csp}" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <link rel="stylesheet" href="${styleUri}" />
  <title>JSON FortuneSheet Editor</title>
</head>
<body>
  <div id="root"></div>
  <script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
  }
}

function toSheetPayloadFromContent(content) {
  if (content && content.dataKind === 'csv') {
    const { sheets, typeMap } = createSheetFromMatrix(content.matrix);
    return { sheets, typeMap, dataKind: 'csv', text: content.text };
  }

  if (content && content.dataKind === 'xlsx') {
    const rawSheets = Array.isArray(content.sheets)
      ? content.sheets
      : Array.isArray(content.matrix)
        ? [{ name: 'Sheet1', matrix: content.matrix }]
        : [{ name: 'Sheet1', matrix: [[]] }];

    const sheets = [];
    rawSheets.forEach((s, idx) => {
      const built = createSheetFromMatrix(s.matrix, s.name);
      const sheetObj = built.sheets[0];
      sheetObj.order = idx;
      sheets.push(sheetObj);
    });

    return { sheets, typeMap: {}, dataKind: 'xlsx', text: '' };
  }

  const extracted = jsonRules.validateAndExtract(content);
  const text = JSON.stringify(content ?? {}, null, 2);

  if (!extracted.ok) {
    return {
      sheets: [],
      typeMap: {},
      dataKind: 'unsupported',
      error: extracted.reason,
      text,
    };
  }

  if (extracted.kind === 'objectKV') {
    const { sheets, dataKind } = createSheetsFromJson(extracted.object);
    return { sheets, typeMap: {}, dataKind, text };
  }

  if (extracted.kind === 'objectOfObjects') {
    const prep = prepareObjectOfObjects(extracted.object);
    if (!prep.ok) {
      return {
        sheets: [],
        typeMap: {},
        dataKind: 'unsupported',
        error: prep.reason,
        text,
      };
    }
    const { sheets } = createSheetFromObjectOfObjects(prep.keys, prep.fieldHeaders, prep.rows);
    return {
      sheets,
      typeMap: {},
      dataKind: 'objectOfObjects',
      text,
    };
  }

  // Array-like shapes (top-level array or wrapper {data:[...]})
  const inputRows = extracted.kind === 'wrappedArray' ? extracted.data : extracted.data;
  const kindForWrapper = extracted.kind;

  const flattenedRows = [];
  const headers = [];
  const headerSet = new Set();

  for (let i = 0; i < inputRows.length; i += 1) {
    const row = inputRows[i];
    const flatRes = jsonRules.flattenRowObject(row);
    if (!flatRes.ok) {
      return {
        sheets: [],
        typeMap: {},
        dataKind: 'unsupported',
        error: `Unsupported JSON at row[${i}]: ${flatRes.errors.join(' ')}`,
        text,
      };
    }

    const flat = flatRes.flat;
    Object.keys(flat).forEach((key) => {
      if (!headerSet.has(key)) {
        headerSet.add(key);
        headers.push(key);
      }
    });
    flattenedRows.push(flat);
  }

  const { sheets } = createSheetFromFlatRows(headers, flattenedRows);

  if (kindForWrapper === 'wrappedArray') {
    return {
      sheets,
      typeMap: {},
      dataKind: 'wrappedArray',
      wrapper: { meta: extracted.meta, dataProp: extracted.dataProp },
      text,
    };
  }

  return {
    sheets,
    typeMap: {},
    dataKind: 'array',
    text,
  };
}

function prepareObjectOfObjects(obj) {
  const keys = Object.keys(obj || {});
  const rows = [];
  const fieldHeaders = [];
  const fieldSet = new Set();

  for (let i = 0; i < keys.length; i += 1) {
    const k = keys[i];
    const inner = obj[k];
    const flatRes = jsonRules.flattenInnerObject(inner, k);
    if (!flatRes.ok) {
      return { ok: false, reason: flatRes.errors.join(' ') };
    }
    const flat = flatRes.flat;
    Object.keys(flat).forEach((field) => {
      if (!fieldSet.has(field)) {
        fieldSet.add(field);
        fieldHeaders.push(field);
      }
    });
    rows.push({ key: k, fields: flat });
  }

  return { ok: true, keys, fieldHeaders, rows };
}

function createSheetFromObjectOfObjects(keys, fieldHeaders, rows) {
  const celldata = [];
  const headers = ['key', ...fieldHeaders];

  headers.forEach((h, c) => {
    celldata.push(createCell(0, c, h));
  });

  rows.forEach((row, idx) => {
    const r = idx + 1;
    celldata.push(createCell(r, 0, row.key));
    fieldHeaders.forEach((field, cIdx) => {
      const v = Object.prototype.hasOwnProperty.call(row.fields, field) ? row.fields[field] : '';
      celldata.push(createCell(r, cIdx + 1, v));
    });
  });

  const data = celldataToMatrixFromCells(celldata);
  return {
    sheets: [
      {
        name: 'Sheet1',
        order: 0,
        status: 1,
        celldata,
        data,
        row: Math.max(rows.length + 10, 20),
        column: Math.max(headers.length + 5, 10),
        defaultColWidth: 120,
        showGridLines: true,
        filter: {},
      },
    ],
  };
}

function toSheetPayloadFromObject(obj) {
  const typeMap = {};
  buildTypeMap(obj, '', typeMap);
  const { sheets, dataKind } = createSheetsFromJson(obj);
  return { sheets, typeMap, dataKind };
}

function buildTypeMap(value, pathPrefix, typeMap) {
  if (value === null) {
    typeMap[pathPrefix || '$root'] = 'null';
    return;
  }
  const valueType = typeof value;
  if (valueType === 'number' || valueType === 'boolean' || valueType === 'string') {
    typeMap[pathPrefix || '$root'] = valueType;
    return;
  }
  if (Array.isArray(value)) {
    const base = pathPrefix || '$root';
    typeMap[base] = 'array';
    value.forEach((item, index) => {
      const nextPrefix = pathPrefix ? `${pathPrefix}[${index}]` : `[${index}]`;
      buildTypeMap(item, nextPrefix, typeMap);
    });
    return;
  }
  if (value && valueType === 'object') {
    const base = pathPrefix || '$root';
    typeMap[base] = 'object';
    Object.entries(value).forEach(([key, val]) => {
      const nextPrefix = pathPrefix ? `${pathPrefix}.${key}` : key;
      buildTypeMap(val, nextPrefix, typeMap);
    });
  }
}

function createSheetsFromJson(json) {
  if (Array.isArray(json)) {
    return createSheetFromArray(json);
  }
  if (json && typeof json === 'object') {
    return createSheetFromObject(json);
  }
  return createSheetFromObject({});
}

function createSheetFromFlatRows(headers, rows) {
  const celldata = [];

  headers.forEach((key, c) => {
    celldata.push(createCell(0, c, key));
  });

  rows.forEach((rowObj, rowIndex) => {
    const r = rowIndex + 1;
    headers.forEach((key, c) => {
      const value = Object.prototype.hasOwnProperty.call(rowObj, key) ? rowObj[key] : '';
      celldata.push(createCell(r, c, value));
    });
  });

  const data = celldataToMatrixFromCells(celldata);

  return {
    sheets: [
      {
        name: 'Sheet1',
        order: 0,
        status: 1,
        celldata,
        data,
        row: Math.max(rows.length + 10, 20),
        column: Math.max(headers.length + 5, 10),
        showGridLines: true,
        filter: {},
      },
    ],
  };
}

function createSheetFromMatrix(matrix, sheetName = 'Sheet1') {
  const typeMap = {};
  const celldata = [];
  (matrix || []).forEach((row, r) => {
    (row || []).forEach((value, c) => {
      const hint = inferSimpleHint(value);
      typeMap[`${r},${c}`] = hint;

      // Keep primitive values as-is; default to empty string.
      let cellValue = value;
      if (cellValue === undefined || cellValue === null) {
        cellValue = '';
      }
      celldata.push(createCell(r, c, cellValue));
    });
  });

  const data = celldataToMatrixFromCells(celldata);

  return {
    sheets: [
      {
        name: sheetName,
        order: 0,
        status: 1,
        celldata,
        data,
        row: Math.max(matrix.length + 10, 20),
        column: Math.max((matrix[0] || []).length + 5, 10),
        showGridLines: true,
        filter: {},
      },
    ],
    typeMap,
  };
}

function createSheetFromArray(arr) {
  const headers = collectHeaders(arr);
  const celldata = [];

  // Create header row
  headers.forEach((key, idx) => {
    celldata.push(createCell(0, idx, key));
  });

  // Create data rows
  arr.forEach((row, rowIndex) => {
    const targetRow = rowIndex + 1;
    if (row && typeof row === 'object' && !Array.isArray(row)) {
      headers.forEach((key, colIndex) => {
        const value = row[key];
        celldata.push(createCell(targetRow, colIndex, value));
      });
    } else {
      // Handle primitive values in array
      celldata.push(createCell(targetRow, 0, rowIndex));
      celldata.push(createCell(targetRow, 1, row));
    }
  });

  const data = celldataToMatrixFromCells(celldata);

  return {
    sheets: [
      {
        name: 'Sheet1',
        order: 0,
        status: 1,
        celldata,
        data,
        row: Math.max(arr.length + 10, 20),
        column: Math.max(headers.length + 2, 10),
        showGridLines: true,
        filter: {},
      },
    ],
    dataKind: 'array',
  };
}

function createSheetFromObject(obj) {
  const entries = Object.entries(obj ?? {});
  const celldata = [];

  entries.forEach(([key, value], rowIndex) => {
    celldata.push(createCell(rowIndex, 0, key));
    celldata.push(createCell(rowIndex, 1, value));
  });

  if (entries.length === 0) {
    celldata.push(createCell(0, 0, 'key'));
    celldata.push(createCell(0, 1, 'value'));
  }

  const data = celldataToMatrixFromCells(celldata);

  return {
    sheets: [
      {
        name: 'Sheet1',
        order: 0,
        status: 1,
        celldata,
        data,
        row: Math.max(entries.length + 10, 20),
        column: 10,
        showGridLines: true,
      },
    ],
    dataKind: 'object',
  };
}

function createCell(r, c, value) {
  let display;
  let rawValue;
  
  if (value === null) {
    display = 'null';
    rawValue = null;
  } else if (value === undefined) {
    display = '';
    rawValue = '';
  } else if (typeof value === 'object') {
    // Convert objects and arrays to JSON string for display
    display = JSON.stringify(value);
    rawValue = display;
  } else {
    display = String(value);
    rawValue = value;
  }
  
  return {
    r,
    c,
    v: {
      v: rawValue,
      m: display,
    },
  };
}

// Convert a celldata array into a 2D matrix structure FortuneSheet can also consume
function celldataToMatrixFromCells(cells) {
  const matrix = [];
  cells.forEach((cell) => {
    if (!matrix[cell.r]) {
      matrix[cell.r] = [];
    }
    matrix[cell.r][cell.c] = cell.v;
  });
  return matrix;
}

function parseCsv(text) {
  const lines = text.split(/\r?\n/);
  const matrix = lines
    .filter((line, idx, arr) => !(idx === arr.length - 1 && line.trim() === ''))
    .map((line) => line.split(',').map((cell) => cell.trim()));
  return { dataKind: 'csv', matrix, text };
}

function detectKind(fsPath) {
  const lower = (fsPath || '').toLowerCase();
  if (lower.endsWith('.csv')) {
    return 'csv';
  }
  if (lower.endsWith('.xlsx')) {
    return 'xlsx';
  }
  return 'json';
}

function inferSimpleHint(value) {
  if (value === null) return 'null';
  if (typeof value === 'number') return 'number';
  if (typeof value === 'boolean') return 'boolean';
  return 'string';
}

function collectHeaders(arr) {
  const headers = new Set();
  arr.forEach((item) => {
    if (item && typeof item === 'object' && !Array.isArray(item)) {
      Object.keys(item).forEach((key) => headers.add(key));
    }
  });
  if (headers.size === 0) {
    headers.add('index');
    headers.add('value');
  }
  return Array.from(headers);
}

function normalizeMatrixForXlsx(matrix) {
  const raw = [];
  (matrix || []).forEach((row, rIdx) => {
    const outRow = [];
    (row || []).forEach((cell, cIdx) => {
      if (cell === null || cell === undefined) {
        outRow[cIdx] = '';
      } else if (typeof cell === 'object') {
        // FortuneSheet stores cell values as { v: { v: raw, m: display } } or similar shapes.
        if (Object.prototype.hasOwnProperty.call(cell, 'v')) {
          const maybeVal = cell.v;
          if (maybeVal && typeof maybeVal === 'object') {
            outRow[cIdx] = maybeVal.v ?? maybeVal.m ?? '';
          } else {
            outRow[cIdx] = maybeVal ?? '';
          }
        } else if (Object.prototype.hasOwnProperty.call(cell, 'm')) {
          outRow[cIdx] = cell.m ?? '';
        } else {
          outRow[cIdx] = '';
        }
      } else {
        outRow[cIdx] = cell;
      }
    });
    raw[rIdx] = outRow;
  });
  return raw;
}

function getNonce() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let nonce = '';
  for (let i = 0; i < 32; i += 1) {
    nonce += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return nonce;
}

function deactivate() {}

module.exports = {
  activate,
  deactivate,
};
