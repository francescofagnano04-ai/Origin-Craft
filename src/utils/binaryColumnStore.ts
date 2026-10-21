import { DataColumn, WorkbookSheet } from '../types/origin';

/**
 * RULE 1: Contiguous Binary Data Structure (Float64Array + Uint8Array)
 *
 * Mask / Status Byte Encoding in Uint8Array:
 *   0 = Valid numeric point
 *   1 = Masked / Excluded point (highlighted in red, excluded from plots & fits)
 *   2 = NaN / Missing value
 */
export const POINT_VALID = 0;
export const POINT_MASKED = 1;
export const POINT_NAN = 2;

export interface ColumnBinaryBuffers {
  colId: string;
  length: number;
  /** Contiguous 64-bit IEEE 754 floating point values */
  float64: Float64Array;
  /** Contiguous 8-bit status mask: 0 = valid, 1 = masked, 2 = NaN */
  maskUint8: Uint8Array;
  /** Version counter to invalidate caches in O(1) */
  version: number;
}

const binaryRegistry = new Map<string, ColumnBinaryBuffers>();

/**
 * Allocates a backing ArrayBuffer or SharedArrayBuffer (if crossOriginIsolated is enabled)
 */
export function allocateContiguousBuffer(byteLength: number): ArrayBufferLike {
  if (typeof SharedArrayBuffer !== 'undefined' && globalThis.crossOriginIsolated) {
    try {
      return new SharedArrayBuffer(byteLength);
    } catch {
      return new ArrayBuffer(byteLength);
    }
  }
  return new ArrayBuffer(byteLength);
}

/**
 * Synchronizes or retrieves the contiguous Float64Array & Uint8Array buffers for a DataColumn.
 * If the column already has typed arrays attached (`_float64` and `_maskUint8`), returns them in O(1).
 */
export function getOrSyncColumnBuffers(col: DataColumn): ColumnBinaryBuffers {
  const existing = binaryRegistry.get(col.id);
  const colTypedF64 = (col as DataColumnWithBinary)._float64;
  const colTypedMask = (col as DataColumnWithBinary)._maskUint8;

  if (colTypedF64 && colTypedMask && colTypedF64.length === colTypedMask.length) {
    if (!existing || existing.float64 !== colTypedF64 || existing.maskUint8 !== colTypedMask) {
      const entry: ColumnBinaryBuffers = {
        colId: col.id,
        length: colTypedF64.length,
        float64: colTypedF64,
        maskUint8: colTypedMask,
        version: (existing?.version || 0) + 1,
      };
      binaryRegistry.set(col.id, entry);
      return entry;
    }
    return existing;
  }

  // Convert legacy or newly created col.values + col.maskedIndices into contiguous Float64Array + Uint8Array
  const len = col.values ? col.values.length : 0;
  const f64Buf = allocateContiguousBuffer(len * Float64Array.BYTES_PER_ELEMENT);
  const u8Buf = allocateContiguousBuffer(len * Uint8Array.BYTES_PER_ELEMENT);
  const float64 = new Float64Array(f64Buf);
  const maskUint8 = new Uint8Array(u8Buf);

  const maskedSet = new Set(col.maskedIndices || []);

  for (let i = 0; i < len; i++) {
    const v = col.values[i];
    if (v === null || v === undefined || !Number.isFinite(v)) {
      float64[i] = Number.NaN;
      maskUint8[i] = POINT_NAN;
    } else {
      float64[i] = v;
      maskUint8[i] = maskedSet.has(i) ? POINT_MASKED : POINT_VALID;
    }
  }

  (col as DataColumnWithBinary)._float64 = float64;
  (col as DataColumnWithBinary)._maskUint8 = maskUint8;
  (col as DataColumnWithBinary)._rowCount = len;

  const entry: ColumnBinaryBuffers = {
    colId: col.id,
    length: len,
    float64,
    maskUint8,
    version: (existing?.version || 0) + 1,
  };
  binaryRegistry.set(col.id, entry);
  return entry;
}

export interface DataColumnWithBinary extends DataColumn {
  _float64?: Float64Array;
  _maskUint8?: Uint8Array;
  _rowCount?: number;
}

/**
 * Returns the row count of a column in O(1) using its binary buffer length
 */
export function getColumnRowCount(col: DataColumn): number {
  const b = (col as DataColumnWithBinary)._float64;
  if (b) return b.length;
  return Math.max(col.values?.length || 0, col.rawStrings?.length || 0);
}

/**
 * Returns the row count of an entire sheet in O(cols)
 */
export function getSheetRowCount(sheet: WorkbookSheet): number {
  let max = 0;
  for (let i = 0; i < sheet.columns.length; i++) {
    const len = getColumnRowCount(sheet.columns[i]);
    if (len > max) max = len;
  }
  return max;
}

/**
 * Reads a single numeric value at rowIndex in O(1) directly from the Float64Array
 */
export function readCellNumericO1(col: DataColumn, rowIndex: number): number | null {
  const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
  if (rowIndex < 0 || rowIndex >= length) return null;
  if (maskUint8[rowIndex] === POINT_NAN) return null;
  const val = float64[rowIndex];
  return Number.isFinite(val) ? val : null;
}

/**
 * Checks if a cell is masked in O(1) directly from the Uint8Array mask buffer
 */
export function isCellMaskedO1(col: DataColumn, rowIndex: number): boolean {
  const { maskUint8, length } = getOrSyncColumnBuffers(col);
  if (rowIndex < 0 || rowIndex >= length) return false;
  return maskUint8[rowIndex] === POINT_MASKED;
}

/**
 * Counts masked cells in a column using the Uint8Array buffer
 */
export function countColumnMaskedCells(col: DataColumn): number {
  const { maskUint8, length } = getOrSyncColumnBuffers(col);
  let count = 0;
  for (let i = 0; i < length; i++) {
    if (maskUint8[i] === POINT_MASKED) count++;
  }
  return count;
}

/**
 * Creates a lightweight Array-like Proxy over Float64Array & Uint8Array so any existing code
 * accessing `col.values[r]` or `col.values.length` works in O(1) without allocating a JS Array!
 */
function createVirtualValuesView(float64: Float64Array, maskUint8: Uint8Array): (number | null)[] {
  // For small datasets (<= 5000 rows), a standard array is tiny (< 40 KB) and JSON-serializable.
  // For large datasets (> 5000 rows), we use a Proxy backed directly by Float64Array so 0 JS objects are allocated.
  if (float64.length <= 5000) {
    const arr = new Array<number | null>(float64.length);
    for (let i = 0; i < float64.length; i++) {
      arr[i] = maskUint8[i] === POINT_NAN || !Number.isFinite(float64[i]) ? null : float64[i];
    }
    return arr;
  }

  const target: (number | null)[] = [];
  return new Proxy(target, {
    get(_t, prop) {
      if (prop === 'length') return float64.length;
      if (typeof prop === 'string') {
        const idx = Number(prop);
        if (Number.isInteger(idx) && idx >= 0 && idx < float64.length) {
          return maskUint8[idx] === POINT_NAN || !Number.isFinite(float64[idx])
            ? null
            : float64[idx];
        }
      }
      if (prop === Symbol.iterator) {
        return function* () {
          for (let i = 0; i < float64.length; i++) {
            yield maskUint8[i] === POINT_NAN || !Number.isFinite(float64[i]) ? null : float64[i];
          }
        };
      }
      if (prop === 'slice') {
        return (start = 0, end = float64.length) => {
          const s = Math.max(0, start < 0 ? float64.length + start : start);
          const e = Math.min(float64.length, end < 0 ? float64.length + end : end);
          const out = new Array<number | null>(Math.max(0, e - s));
          for (let i = s; i < e; i++) {
            out[i - s] =
              maskUint8[i] === POINT_NAN || !Number.isFinite(float64[i]) ? null : float64[i];
          }
          return out;
        };
      }
      if (prop === 'map') {
        return <U>(fn: (v: number | null, i: number) => U): U[] => {
          const out = new Array<U>(float64.length);
          for (let i = 0; i < float64.length; i++) {
            const v =
              maskUint8[i] === POINT_NAN || !Number.isFinite(float64[i]) ? null : float64[i];
            out[i] = fn(v, i);
          }
          return out;
        };
      }
      if (prop === 'filter') {
        return (fn: (v: number | null, i: number) => boolean) => {
          const out: (number | null)[] = [];
          for (let i = 0; i < float64.length; i++) {
            const v =
              maskUint8[i] === POINT_NAN || !Number.isFinite(float64[i]) ? null : float64[i];
            if (fn(v, i)) out.push(v);
          }
          return out;
        };
      }
      return Reflect.get(_t, prop);
    },
  });
}

/**
 * Builds a DataColumn backed directly by contiguous Float64Array and Uint8Array buffers
 */
export function attachBinaryBuffersToColumn(
  col: DataColumn,
  float64: Float64Array,
  maskUint8?: Uint8Array
): DataColumnWithBinary {
  const len = float64.length;
  let mask = maskUint8;
  if (!mask || mask.length !== len) {
    const u8Buf = allocateContiguousBuffer(len);
    mask = new Uint8Array(u8Buf);
    const prevMasked = new Set(col.maskedIndices || []);
    for (let i = 0; i < len; i++) {
      if (!Number.isFinite(float64[i])) {
        mask[i] = POINT_NAN;
      } else if (prevMasked.has(i)) {
        mask[i] = POINT_MASKED;
      } else {
        mask[i] = POINT_VALID;
      }
    }
  }

  const maskedIndices: number[] = [];
  for (let i = 0; i < len; i++) {
    if (mask[i] === POINT_MASKED) {
      maskedIndices.push(i);
    }
  }

  const virtualValues = createVirtualValuesView(float64, mask);

  const updated: DataColumnWithBinary = {
    ...col,
    values: virtualValues,
    maskedIndices,
    _float64: float64,
    _maskUint8: mask,
    _rowCount: len,
  };

  binaryRegistry.set(col.id, {
    colId: col.id,
    length: len,
    float64,
    maskUint8: mask,
    version: (binaryRegistry.get(col.id)?.version || 0) + 1,
  });

  return updated;
}

/**
 * Mutates a single cell inside a column's binary buffers (creates a fast typed-array copy for React immutability)
 */
export function updateColumnCellBinary(
  col: DataColumn,
  rowIndex: number,
  numVal: number | null,
  strVal: string | null = null
): DataColumnWithBinary {
  const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
  const newLen = Math.max(length, rowIndex + 1);

  const nextF64 = new Float64Array(allocateContiguousBuffer(newLen * 8));
  const nextMask = new Uint8Array(allocateContiguousBuffer(newLen));

  nextF64.set(float64, 0);
  nextMask.set(maskUint8, 0);

  for (let i = length; i < newLen; i++) {
    nextF64[i] = Number.NaN;
    nextMask[i] = POINT_NAN;
  }

  const wasMasked = rowIndex < length && maskUint8[rowIndex] === POINT_MASKED;
  if (numVal === null || !Number.isFinite(numVal)) {
    nextF64[rowIndex] = Number.NaN;
    nextMask[rowIndex] = POINT_NAN;
  } else {
    nextF64[rowIndex] = numVal;
    nextMask[rowIndex] = wasMasked ? POINT_MASKED : POINT_VALID;
  }

  let nextRawStrings = col.rawStrings ? [...col.rawStrings] : undefined;
  if (strVal !== null || nextRawStrings) {
    if (!nextRawStrings) {
      nextRawStrings = new Array(newLen).fill(null);
    }
    while (nextRawStrings.length < newLen) {
      nextRawStrings.push(null);
    }
    nextRawStrings[rowIndex] = strVal;
  }

  return attachBinaryBuffersToColumn(
    { ...col, rawStrings: nextRawStrings },
    nextF64,
    nextMask
  );
}

/**
 * Toggles the mask bit (0 <-> 1) for a specific rowIndex in O(1) binary copy
 */
export function toggleColumnCellMaskBinary(
  col: DataColumn,
  rowIndex: number
): DataColumnWithBinary {
  const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
  if (rowIndex < 0 || rowIndex >= length) return col as DataColumnWithBinary;

  const nextMask = new Uint8Array(allocateContiguousBuffer(length));
  nextMask.set(maskUint8);

  if (nextMask[rowIndex] === POINT_MASKED) {
    nextMask[rowIndex] = Number.isFinite(float64[rowIndex]) ? POINT_VALID : POINT_NAN;
  } else {
    nextMask[rowIndex] = POINT_MASKED;
  }

  return attachBinaryBuffersToColumn(col, float64, nextMask);
}

/**
 * Extracts contiguous Float64Arrays for WebGL Plotly (`scattergl`) rendering:
 * - Unmasked valid points (xValid, yValid, xErrValid, yErrValid, rowIndicesValid)
 * - Masked points (xMasked, yMasked, rowIndicesMasked)
 * Zero intermediate JS objects!
 */
export function extractWebGLSeriesTypedArrays(
  xCol: DataColumn,
  yCol: DataColumn,
  xErrCol?: DataColumn,
  yErrCol?: DataColumn
): {
  xValid: Float64Array;
  yValid: Float64Array;
  xErrValid?: Float64Array;
  yErrValid?: Float64Array;
  validRowIndices: Int32Array;
  xMasked: Float64Array;
  yMasked: Float64Array;
  maskedRowIndices: Int32Array;
  minX: number;
  maxX: number;
  minY: number;
  maxY: number;
} {
  const xBuf = getOrSyncColumnBuffers(xCol);
  const yBuf = getOrSyncColumnBuffers(yCol);
  const xErrBuf = xErrCol ? getOrSyncColumnBuffers(xErrCol) : undefined;
  const yErrBuf = yErrCol ? getOrSyncColumnBuffers(yErrCol) : undefined;

  const len = Math.min(xBuf.length, yBuf.length);
  const xF64 = xBuf.float64;
  const yF64 = yBuf.float64;
  const xMask = xBuf.maskUint8;
  const yMask = yBuf.maskUint8;

  // Pass 1: count valid and masked points (pure typed-array loop, < 1ms for 1M points)
  let validCount = 0;
  let maskedCount = 0;

  for (let i = 0; i < len; i++) {
    const xv = xF64[i];
    const yv = yF64[i];
    if (!Number.isFinite(xv) || !Number.isFinite(yv)) continue;
    if (xMask[i] === POINT_MASKED || yMask[i] === POINT_MASKED) {
      maskedCount++;
    } else if (xMask[i] === POINT_VALID && yMask[i] === POINT_VALID) {
      validCount++;
    }
  }

  const xValid = new Float64Array(validCount);
  const yValid = new Float64Array(validCount);
  const xErrValid = xErrBuf ? new Float64Array(validCount) : undefined;
  const yErrValid = yErrBuf ? new Float64Array(validCount) : undefined;
  const validRowIndices = new Int32Array(validCount);

  const xMasked = new Float64Array(maskedCount);
  const yMasked = new Float64Array(maskedCount);
  const maskedRowIndices = new Int32Array(maskedCount);

  let vIdx = 0;
  let mIdx = 0;
  let minX = Infinity;
  let maxX = -Infinity;
  let minY = Infinity;
  let maxY = -Infinity;

  for (let i = 0; i < len; i++) {
    const xv = xF64[i];
    const yv = yF64[i];
    if (!Number.isFinite(xv) || !Number.isFinite(yv)) continue;

    if (xMask[i] === POINT_MASKED || yMask[i] === POINT_MASKED) {
      xMasked[mIdx] = xv;
      yMasked[mIdx] = yv;
      maskedRowIndices[mIdx] = i;
      mIdx++;
    } else if (xMask[i] === POINT_VALID && yMask[i] === POINT_VALID) {
      xValid[vIdx] = xv;
      yValid[vIdx] = yv;
      validRowIndices[vIdx] = i;

      let dy = 0;
      if (yErrValid && yErrBuf && i < yErrBuf.length) {
        const ey = yErrBuf.float64[i];
        dy = Number.isFinite(ey) ? Math.abs(ey) : 0;
        yErrValid[vIdx] = dy;
      }
      if (xErrValid && xErrBuf && i < xErrBuf.length) {
        const ex = xErrBuf.float64[i];
        xErrValid[vIdx] = Number.isFinite(ex) ? Math.abs(ex) : 0;
      }

      if (xv < minX) minX = xv;
      if (xv > maxX) maxX = xv;
      if (yv - dy < minY) minY = yv - dy;
      if (yv + dy > maxY) maxY = yv + dy;

      vIdx++;
    }
  }

  return {
    xValid,
    yValid,
    xErrValid,
    yErrValid,
    validRowIndices,
    xMasked,
    yMasked,
    maskedRowIndices,
    minX,
    maxX,
    minY,
    maxY,
  };
}

/**
 * Generates a high-density synthetic scientific dataset (e.g. 100,000 or 500,000 or 1,000,000 points)
 * directly in contiguous Float64Arrays in < 25ms to benchmark 60 FPS WebGL + Virtual Grid
 */
export function generateHighPerformanceDatasetSheet(
  pointCount: number,
  sheetName = `Segnale_${(pointCount / 1000).toFixed(0)}k`
): WorkbookSheet {
  const n = Math.max(100, Math.min(1_000_000, pointCount));

  const xF64 = new Float64Array(allocateContiguousBuffer(n * 8));
  const xMask = new Uint8Array(allocateContiguousBuffer(n));

  const y1F64 = new Float64Array(allocateContiguousBuffer(n * 8));
  const y1Mask = new Uint8Array(allocateContiguousBuffer(n));

  const y2F64 = new Float64Array(allocateContiguousBuffer(n * 8));
  const y2Mask = new Uint8Array(allocateContiguousBuffer(n));

  const dt = 100 / n;
  for (let i = 0; i < n; i++) {
    const t = i * dt;
    xF64[i] = t;
    xMask[i] = POINT_VALID;

    // Damped chirped oscillation + resonant Gaussian peak + controlled noise
    const noise1 = (Math.sin(i * 12.9898) * 43758.5453) % 1;
    const gaussianPeak = 4.5 * Math.exp(-Math.pow((t - 52) / 8, 2));
    y1F64[i] =
      2.0 +
      0.08 * t +
      gaussianPeak +
      1.2 * Math.sin(0.45 * t) * Math.exp(-0.015 * t) +
      noise1 * 0.22;
    y1Mask[i] = POINT_VALID;

    const noise2 = (Math.cos(i * 78.233) * 19341.123) % 1;
    y2F64[i] =
      1.5 +
      3.2 * Math.cos(0.3 * t) * Math.exp(-0.01 * t) +
      2.8 * Math.exp(-Math.pow((t - 30) / 6, 2)) +
      noise2 * 0.18;
    y2Mask[i] = POINT_VALID;
  }

  const colA = attachBinaryBuffersToColumn(
    {
      id: `col_hp_x_${Date.now()}`,
      shortName: 'A',
      longName: 'Tempo Acquisizione',
      unit: 'ms',
      comment: `${n.toLocaleString('it-IT')} campioni Float64`,
      role: 'X',
      dataType: 'numeric',
      values: [],
    },
    xF64,
    xMask
  );

  const colB = attachBinaryBuffersToColumn(
    {
      id: `col_hp_y1_${Date.now()}`,
      shortName: 'B',
      longName: 'Canale 1 (Risonanza)',
      unit: 'mV',
      comment: 'WebGL scattergl @ 60 FPS',
      role: 'Y',
      dataType: 'numeric',
      linkedXColId: colA.id,
      values: [],
    },
    y1F64,
    y1Mask
  );

  const colC = attachBinaryBuffersToColumn(
    {
      id: `col_hp_y2_${Date.now()}`,
      shortName: 'C',
      longName: 'Canale 2 (Inviluppo)',
      unit: 'mV',
      comment: 'Buffer binario contiguo',
      role: 'Y',
      dataType: 'numeric',
      linkedXColId: colA.id,
      values: [],
    },
    y2F64,
    y2Mask
  );

  return {
    id: `sheet_hp_${Date.now()}`,
    name: sheetName,
    columns: [colA, colB, colC],
  };
}

/**
 * Ensures all columns in a WorkbookSheet have contiguous Float64Array and Uint8Array binary buffers attached
 */
export function ensureSheetBinaryBuffers(sheet: WorkbookSheet): WorkbookSheet {
  return {
    ...sheet,
    columns: sheet.columns.map((col) => {
      const { float64, maskUint8 } = getOrSyncColumnBuffers(col);
      return attachBinaryBuffersToColumn(col, float64, maskUint8);
    }),
  };
}

