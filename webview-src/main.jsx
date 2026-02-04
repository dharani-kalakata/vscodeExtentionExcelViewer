import React, { useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Workbook } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import './style.css';

const vscode = acquireVsCodeApi();

const defaultSheets = [];

function App() {
  const [sheets, setSheets] = useState(defaultSheets);
  const [typeMap, setTypeMap] = useState({});
  const [dataKind, setDataKind] = useState('object');
  const [isInitialized, setIsInitialized] = useState(false);
  const [wrapper, setWrapper] = useState(null);
  const [error, setError] = useState('');
  const lastTextRef = useRef('');
  const ignoreChangesRef = useRef(true);
  const pendingFlushRef = useRef(null);
  const currentSheetsRef = useRef(defaultSheets);
  const currentTypeMapRef = useRef({});
  const workbookRef = useRef(null);
  const currentDataKindRef = useRef('object');
  const currentWrapperRef = useRef(null);
  const initUnlockTimerRef = useRef(null);

  useEffect(() => {
    const saved = vscode.getState();
    if (saved?.typeMap) {
      setTypeMap(saved.typeMap);
    }
    if (saved?.dataKind) {
      setDataKind(saved.dataKind);
      currentDataKindRef.current = saved.dataKind;
    }

    const handleKeyDown = (e) => {
      const key = String(e.key || '').toLowerCase();
      if ((e.ctrlKey || e.metaKey) && key === 's') {
        e.preventDefault();
        e.stopPropagation();
        flushPendingChanges();

        const api = workbookRef.current;
        const latestSheets = api && typeof api.getAllSheets === 'function' ? api.getAllSheets() : currentSheetsRef.current;
        const { textOut, matrix, xlsxSheets } = sheetToText(
          latestSheets,
          currentTypeMapRef.current,
          currentDataKindRef.current,
          currentWrapperRef.current,
        );

        if (currentDataKindRef.current === 'xlsx') {
          vscode.postMessage({ type: 'save', dataKind: 'xlsx', xlsxSheets });
        } else {
          vscode.postMessage({ type: 'save', text: textOut, matrix, dataKind: currentDataKindRef.current });
        }
      }
    };

    // Capture phase so FortuneSheet/internal handlers can't swallow Ctrl+S.
    window.addEventListener('keydown', handleKeyDown, true);

    const onError = (event) => {
      try {
        const message = event?.message || 'Unknown webview error';
        const stack = event?.error?.stack || '';
        vscode.postMessage({ type: 'webviewError', message, stack });
      } catch {
        // ignore
      }
    };

    const onUnhandledRejection = (event) => {
      try {
        const reason = event?.reason;
        const message = typeof reason === 'string' ? reason : reason?.message || 'Unhandled promise rejection';
        const stack = reason?.stack || '';
        vscode.postMessage({ type: 'webviewError', message, stack });
      } catch {
        // ignore
      }
    };

    window.addEventListener('error', onError);
    window.addEventListener('unhandledrejection', onUnhandledRejection);

    const handler = (event) => {
      const message = event.data;
      if (message.type === 'init') {
        const nextError = message.payload.error || '';
        const nextWrapper = message.payload.wrapper || null;
        const nextSheets = message.payload.sheets || [];
        const nextTypeMap = message.payload.typeMap || {};
        const nextDataKind = message.payload.dataKind || 'object';

        setError(nextError);
        setWrapper(nextWrapper);
        setSheets(nextSheets);
        setTypeMap(nextTypeMap);
        setDataKind(nextDataKind);
        currentSheetsRef.current = nextSheets;
        currentTypeMapRef.current = nextTypeMap;
        currentDataKindRef.current = nextDataKind;
        currentWrapperRef.current = nextWrapper;

        // Establish a baseline text that matches how we serialize the sheet.
        // This prevents "dirty" prompts caused by FortuneSheet emitting onChange during init.
        if (nextError) {
          lastTextRef.current = message.payload.text || '';
        } else {
          lastTextRef.current = sheetToText(nextSheets, nextTypeMap, nextDataKind, nextWrapper).textOut;
        }

        // FortuneSheet can emit onChange during initialization.
        ignoreChangesRef.current = true;
        if (initUnlockTimerRef.current !== null) {
          clearTimeout(initUnlockTimerRef.current);
        }
        // Give FortuneSheet a moment to finish internal recalcs/ops.
        initUnlockTimerRef.current = setTimeout(() => {
          ignoreChangesRef.current = false;
          initUnlockTimerRef.current = null;
        }, 350);
        setIsInitialized(true);
        vscode.setState({
          typeMap: message.payload.typeMap,
          dataKind: message.payload.dataKind,
        });
      }
      if (message.type === 'updateFromPython') {
        setSheets(message.payload.sheets);
        setTypeMap(message.payload.typeMap || {});
        setDataKind(message.payload.dataKind || currentDataKindRef.current);
        vscode.setState({
          typeMap: message.payload.typeMap,
          dataKind: message.payload.dataKind || currentDataKindRef.current,
        });
      }
    };

    window.addEventListener('message', handler);
    vscode.postMessage({ type: 'ready' });

    return () => {
      window.removeEventListener('message', handler);
      window.removeEventListener('keydown', handleKeyDown, true);
      window.removeEventListener('error', onError);
      window.removeEventListener('unhandledrejection', onUnhandledRejection);
      if (initUnlockTimerRef.current !== null) {
        clearTimeout(initUnlockTimerRef.current);
        initUnlockTimerRef.current = null;
      }
    };
  }, []);

  const flushPendingChanges = () => {
    if (pendingFlushRef.current !== null) {
      clearTimeout(pendingFlushRef.current);
      pendingFlushRef.current = null;
    }

    if (!isInitialized || error) {
      return;
    }

    if (ignoreChangesRef.current) {
      return;
    }

    const api = workbookRef.current;
    const latestSheets = api && typeof api.getAllSheets === 'function' ? api.getAllSheets() : currentSheetsRef.current;

    const { textOut, nextTypeMap, matrix, xlsxSheets } = sheetToText(
      latestSheets,
      currentTypeMapRef.current,
      currentDataKindRef.current,
      currentWrapperRef.current
    );

    if (textOut !== lastTextRef.current) {
      lastTextRef.current = textOut;
      currentTypeMapRef.current = nextTypeMap;
      vscode.setState({ typeMap: nextTypeMap, dataKind: currentDataKindRef.current });
      if (currentDataKindRef.current === 'xlsx') {
        vscode.postMessage({ type: 'edit', dataKind: 'xlsx', xlsxSheets });
      } else {
        vscode.postMessage({ type: 'edit', text: textOut, typeMap: nextTypeMap, matrix, dataKind: currentDataKindRef.current });
      }
    }
  };

  const handleChange = (nextSheets) => {
    // Don't send updates until data is initialized
    if (!isInitialized) {
      return;
    }

    // Ignore any init-triggered changes.
    if (ignoreChangesRef.current) {
      currentSheetsRef.current = nextSheets;
      return;
    }
    
    currentSheetsRef.current = nextSheets;
    const { textOut, nextTypeMap, matrix, xlsxSheets } = sheetToText(
      nextSheets,
      typeMap,
      currentDataKindRef.current,
      currentWrapperRef.current,
    );

    // Don't rewrite the document if nothing actually changed.
    if (textOut === lastTextRef.current) {
      currentTypeMapRef.current = nextTypeMap;
      vscode.setState({ typeMap: nextTypeMap, dataKind: currentDataKindRef.current });
      return;
    }

    lastTextRef.current = textOut;
    currentTypeMapRef.current = nextTypeMap;
    vscode.setState({ typeMap: nextTypeMap, dataKind: currentDataKindRef.current });
    if (currentDataKindRef.current === 'xlsx') {
      vscode.postMessage({ type: 'edit', dataKind: 'xlsx', xlsxSheets });
    } else {
      vscode.postMessage({ type: 'edit', text: textOut, typeMap: nextTypeMap, matrix, dataKind: currentDataKindRef.current });
    }
  };

  const handleOp = () => {
    if (!isInitialized || error) {
      return;
    }

    if (ignoreChangesRef.current) {
      return;
    }

    const api = workbookRef.current;
    if (api && typeof api.getAllSheets === 'function') {
      currentSheetsRef.current = api.getAllSheets();
    }

    if (pendingFlushRef.current !== null) {
      clearTimeout(pendingFlushRef.current);
    }

    pendingFlushRef.current = setTimeout(() => {
      flushPendingChanges();
    }, 300);
  };

  return (
    <div className="app">
      <div className="app__body">
        {isInitialized && error ? (
          <div className="app__error">
            <div className="app__errorTitle">Unsupported JSON for grid view</div>
            <div className="app__errorBody">{error}</div>
          </div>
        ) : isInitialized ? (
          <Workbook
            ref={workbookRef}
            data={sheets}
            showToolbar
            showSheetTabs={dataKind === 'xlsx'}
            showFormulaBar
            allowEdit
            onChange={handleChange}
            onOp={handleOp}
            defaultColWidth={120}
          />
        ) : (
          <div className="app__loading">Loading…</div>
        )}
      </div>
    </div>
  );
}

function sheetToText(sheets, currentTypeMap, dataKind, wrapper) {
  const sheet = sheets[0] || {};
  const matrix = celldataToMatrix(sheet);
  const nextTypeMap = { ...currentTypeMap };

  if (dataKind === 'array') {
    const headers = (matrix[0] || []).map((cell) => getCellText(cell)).filter((value) => value !== '');
    const firstElementType = currentTypeMap['[0]'];
    const rows = [];

    for (let r = 1; r < matrix.length; r += 1) {
      const row = matrix[r] || [];
      const rowObj = {};
      let hasValue = false;

      if (firstElementType && firstElementType !== 'object') {
        const raw = getCellText(row[1] ?? row[0]);
        const path = `[${r - 1}]`;
        const { casted, type } = castWithType(raw, nextTypeMap[path]);
        nextTypeMap[path] = nextTypeMap[path] || type;
        if (raw !== '') {
          rows.push(casted);
        }
        continue;
      }

      headers.forEach((header, c) => {
        const raw = getCellText(row[c]);
        const path = `[${r - 1}].${header}`;
        const { casted, type } = castWithType(raw, nextTypeMap[path]);
        nextTypeMap[path] = nextTypeMap[path] || type;
        if (raw !== '') {
          rowObj[header] = casted;
          hasValue = true;
        }
      });

      if (hasValue) {
        rows.push(rowObj);
      }
    }

    return { textOut: JSON.stringify(rows, null, 2), nextTypeMap, matrix };
  }

  if (dataKind === 'wrappedArray') {
    const headers = (matrix[0] || []).map((cell) => getCellText(cell)).filter((value) => value !== '');
    const rows = [];

    for (let r = 1; r < matrix.length; r += 1) {
      const row = matrix[r] || [];
      const rowObj = {};
      let hasValue = false;

      headers.forEach((header, c) => {
        const raw = getCellText(row[c]);
        const path = `[${r - 1}].${header}`;
        const { casted, type } = castWithType(raw, nextTypeMap[path]);
        nextTypeMap[path] = nextTypeMap[path] || type;
        // Preserve keys even when value is cleared.
        setNestedValue(rowObj, header, casted);
        if (raw !== '') {
          hasValue = true;
        }
      });

      if (hasValue) {
        rows.push(rowObj);
      }
    }

    const dataProp = wrapper?.dataProp || 'data';
    const meta = wrapper?.meta || {};
    const out = { ...meta, [dataProp]: rows };
    return { textOut: JSON.stringify(out, null, 2), nextTypeMap, matrix };
  }

  if (dataKind === 'objectOfObjects') {
    const headers = (matrix[0] || []).map((cell) => getCellText(cell));
    const keyColIndex = headers.findIndex((h) => h === 'key');
    const fieldHeaders = headers.filter((h, idx) => idx !== keyColIndex && h);

    const out = {};

    for (let r = 1; r < matrix.length; r += 1) {
      const row = matrix[r] || [];
      const key = getCellText(row[keyColIndex >= 0 ? keyColIndex : 0]);
      if (!key) {
        continue;
      }

      const inner = {};
      fieldHeaders.forEach((field, c) => {
        // if key column is 0, fields start at 1
        const cellIndex = keyColIndex === 0 ? c + 1 : headers.indexOf(field);
        const raw = getCellText(row[cellIndex]);
        const path = `${key}.${field}`;
        const { casted, type } = castWithType(raw, nextTypeMap[path]);
        nextTypeMap[path] = nextTypeMap[path] || type;
        inner[field] = casted;
      });

      out[key] = inner;
    }

    return { textOut: JSON.stringify(out, null, 2), nextTypeMap, matrix };
  }

  if (dataKind === 'csv') {
    const csvRows = [];
    for (let r = 0; r < matrix.length; r += 1) {
      const row = matrix[r] || [];
      const outRow = row.map((cell, c) => {
        const raw = getCellText(cell);
        const path = `${r},${c}`;
        const { casted, type } = castWithType(raw, nextTypeMap[path]);
        nextTypeMap[path] = nextTypeMap[path] || type;
        return escapeCsv(casted);
      });
      csvRows.push(outRow.join(','));
    }
    return { textOut: csvRows.join('\n'), nextTypeMap, matrix };
  }

  if (dataKind === 'xlsx') {
    const xlsxSheets = (sheets || []).map((s, idx) => ({
      name: String(s?.name || `Sheet${idx + 1}`),
      matrix: celldataToMatrix(s || {}),
    }));

    // Use a deterministic string for dirty detection (not written to disk).
    const serialized = JSON.stringify(
      xlsxSheets.map((s) => ({ name: s.name, matrix: s.matrix.map((row) => (row || []).map(getCellText)) }))
    );

    return { textOut: serialized, nextTypeMap, matrix, xlsxSheets };
  }

  const result = {};
  for (let r = 0; r < matrix.length; r += 1) {
    const key = getCellText(matrix[r]?.[0]);
    if (!key) {
      continue;
    }
    const rawValue = getCellText(matrix[r]?.[1]);
    const { casted, type } = castWithType(rawValue, nextTypeMap[key]);
    nextTypeMap[key] = nextTypeMap[key] || type;
    result[key] = casted;
  }

  return { textOut: JSON.stringify(result, null, 2), nextTypeMap, matrix };
}

function celldataToMatrix(sheet) {
  const matrix = [];
  const applyValue = (r, c, value) => {
    if (!matrix[r]) {
      matrix[r] = [];
    }
    matrix[r][c] = value;
  };

  if (Array.isArray(sheet.data) && sheet.data.length) {
    sheet.data.forEach((row, r) => {
      (row || []).forEach((cell, c) => applyValue(r, c, cell));
    });
  }

  if (Array.isArray(sheet.celldata)) {
    sheet.celldata.forEach((cell) => applyValue(cell.r, cell.c, cell.v ?? cell));
  }

  return matrix;
}

function getCellText(cell) {
  if (!cell) {
    return '';
  }
  if (typeof cell === 'string' || typeof cell === 'number' || typeof cell === 'boolean') {
    return String(cell);
  }
  if (cell && typeof cell === 'object') {
    if (typeof cell.m !== 'undefined') {
      return String(cell.m ?? '');
    }
    if (typeof cell.v !== 'undefined') {
      return String(cell.v ?? '');
    }
  }
  return '';
}

function castWithType(raw, hint) {
  const trimmed = (raw ?? '').trim();
  if (trimmed === '""' || trimmed === "''") {
    return { casted: '', type: 'string' };
  }
  // Clearing a cell should become an empty string, not null, and should never drop keys.
  if (trimmed === '') {
    return { casted: '', type: 'string' };
  }
  if (!hint || hint === 'string') {
    const inferred = inferType(trimmed);
    return inferred;
  }

  if (hint === 'number') {
    const n = Number(trimmed);
    if (!Number.isFinite(n)) {
      return { casted: trimmed, type: 'string' };
    }
    return { casted: n, type: 'number' };
  }
  if (hint === 'boolean') {
    if (/^true$/i.test(trimmed)) {
      return { casted: true, type: 'boolean' };
    }
    if (/^false$/i.test(trimmed)) {
      return { casted: false, type: 'boolean' };
    }
    return { casted: Boolean(trimmed), type: 'boolean' };
  }
  if (hint === 'null') {
    return { casted: null, type: 'null' };
  }
  return { casted: trimmed, type: 'string' };
}

function setNestedValue(target, dottedPath, value) {
  const parts = String(dottedPath).split('.').filter(Boolean);
  if (parts.length === 0) {
    return;
  }

  let obj = target;
  for (let i = 0; i < parts.length - 1; i += 1) {
    const key = parts[i];
    if (!obj[key] || typeof obj[key] !== 'object') {
      obj[key] = {};
    }
    obj = obj[key];
  }
  obj[parts[parts.length - 1]] = value;
}

function inferType(raw) {
  if (raw === '' || raw === undefined) {
    return { casted: '', type: 'string' };
  }
  if (/^true$/i.test(raw)) {
    return { casted: true, type: 'boolean' };
  }
  if (/^false$/i.test(raw)) {
    return { casted: false, type: 'boolean' };
  }
  if (/^null$/i.test(raw)) {
    return { casted: null, type: 'null' };
  }
  const num = Number(raw);
  if (Number.isFinite(num)) {
    return { casted: num, type: 'number' };
  }
  return { casted: raw, type: 'string' };
}

function escapeCsv(value) {
  if (value === null || value === undefined) return '';
  const str = String(value);
  if (/[",\n]/.test(str)) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

const root = createRoot(document.getElementById('root'));
root.render(<App />);
