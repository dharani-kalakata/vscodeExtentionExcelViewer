'use strict';

// Centralized constraints for what JSON shapes this extension supports.
// Goal: make JSON editable in a spreadsheet-like grid without embedding complex documents inside cells.

const RULES = {
  // Supported roots:
  // - Array of objects (table)
  // - Object of primitives (key/value)
  // - Object of objects with primitive fields (table)
  // - Object with a `data` array (optional wrapper)
  optionalArrayWrapperProperty: 'data',

  // For table-like shapes, cells must be primitives or arrays of primitives.
  // Nested objects are only allowed when the ROOT is "object-of-objects" (one level).
  allowArrayOfPrimitives: true,
};

function isPrimitive(value) {
  return (
    value === null ||
    value === undefined ||
    typeof value === 'string' ||
    typeof value === 'number' ||
    typeof value === 'boolean'
  );
}

function isPlainObject(value) {
  return value && typeof value === 'object' && !Array.isArray(value);
}

function isArrayOfPrimitives(value) {
  return Array.isArray(value) && value.every(isPrimitive);
}

function validateAndExtract(content) {
  // Optional wrapper: { ...meta, data: [ ... ] }
  if (isPlainObject(content) && Array.isArray(content[RULES.optionalArrayWrapperProperty])) {
    const dataProp = RULES.optionalArrayWrapperProperty;
    const data = content[dataProp];
    const meta = { ...content };
    delete meta[dataProp];
    return { ok: true, kind: 'wrappedArray', data, meta, dataProp };
  }

  // Top-level array: [ { ... }, ... ]
  if (Array.isArray(content)) {
    return { ok: true, kind: 'array', data: content };
  }

  // Top-level object: either key/value or object-of-objects
  if (isPlainObject(content)) {
    const values = Object.values(content);
    const allPrimitivesOrPrimitiveArrays = values.every(
      (v) => isPrimitive(v) || (RULES.allowArrayOfPrimitives && isArrayOfPrimitives(v)),
    );
    if (allPrimitivesOrPrimitiveArrays) {
      return { ok: true, kind: 'objectKV', object: content };
    }

    const allObjects = values.every((v) => isPlainObject(v));
    if (allObjects) {
      return { ok: true, kind: 'objectOfObjects', object: content };
    }

    return {
      ok: false,
      reason:
        'Unsupported JSON root object. Use either (1) key/value pairs with primitive values, or (2) key/object pairs where the inner objects only contain primitive fields.',
    };
  }

  return {
    ok: false,
    reason: 'Unsupported JSON. Expected an array or an object.',
  };
}

function flattenRowObject(row) {
  const errors = [];
  if (!isPlainObject(row)) {
    return { ok: false, errors: ['Each row must be an object.'] };
  }

  const out = {};

  Object.keys(row).forEach((key) => {
    const value = row[key];

    if (isPrimitive(value)) {
      out[key] = value;
      return;
    }

    if (Array.isArray(value)) {
      if (RULES.allowArrayOfPrimitives && isArrayOfPrimitives(value)) {
        out[key] = JSON.stringify(value);
        return;
      }
      errors.push(`Unsupported array value at "${key}". Only arrays of primitive values are supported.`);
      return;
    }

    // No nested objects inside table rows.
    if (isPlainObject(value)) {
      errors.push(
        `Nested object at "${key}" is not supported in table rows. Flatten it first (e.g. use "${key}.field" columns) or change your JSON shape.`,
      );
      return;
    }

    errors.push(`Unsupported value type at "${key}".`);
  });

  if (errors.length) {
    return { ok: false, errors };
  }

  return { ok: true, flat: out };
}

function flattenInnerObject(inner, parentKeyForError) {
  const errors = [];
  if (!isPlainObject(inner)) {
    return { ok: false, errors: [`Value at "${parentKeyForError}" must be an object.`] };
  }
  const out = {};
  Object.keys(inner).forEach((key) => {
    const value = inner[key];
    const fullKey = `${parentKeyForError}.${key}`;
    if (isPrimitive(value)) {
      out[key] = value;
      return;
    }
    if (Array.isArray(value)) {
      if (RULES.allowArrayOfPrimitives && isArrayOfPrimitives(value)) {
        out[key] = JSON.stringify(value);
        return;
      }
      errors.push(`Unsupported array value at "${fullKey}". Only arrays of primitive values are supported.`);
      return;
    }
    if (isPlainObject(value)) {
      errors.push(`Nested object too deep at "${fullKey}". Inner objects must be flat.`);
      return;
    }
    errors.push(`Unsupported value type at "${fullKey}".`);
  });

  if (errors.length) {
    return { ok: false, errors };
  }
  return { ok: true, flat: out };
}

module.exports = {
  RULES,
  validateAndExtract,
  flattenRowObject,
  flattenInnerObject,
  isPrimitive,
  isPlainObject,
};
