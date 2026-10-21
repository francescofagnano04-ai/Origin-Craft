import { DataColumn, FitModelType, SmoothingFilterMethod, WorkbookSheet } from '../types/origin';
import {
  attachBinaryBuffersToColumn,
  getOrSyncColumnBuffers,
} from './binaryColumnStore';
import { rebindSheetMultiXAssociations } from './mathEngine';
import type { IOWorkerParseRequest, IOWorkerParseResponse } from '../workers/io.worker';
import type { MathWorkerRequest, MathWorkerResponse } from '../workers/math.worker';

let ioWorker: Worker | null = null;
let mathWorker: Worker | null = null;

const pendingIORequests = new Map<
  string,
  { resolve: (res: IOWorkerParseResponse) => void; reject: (err: Error) => void }
>();

const pendingMathRequests = new Map<
  string,
  { resolve: (res: MathWorkerResponse) => void; reject: (err: Error) => void }
>();

function getIOWorker(): Worker | null {
  if (typeof Worker === 'undefined') return null;
  if (!ioWorker) {
    ioWorker = new Worker(new URL('../workers/io.worker.ts', import.meta.url), {
      type: 'module',
    });
    ioWorker.onmessage = (e: MessageEvent<IOWorkerParseResponse>) => {
      const data = e.data;
      const pending = pendingIORequests.get(data.id);
      if (pending) {
        pendingIORequests.delete(data.id);
        pending.resolve(data);
      }
    };
  }
  return ioWorker;
}

function getMathWorker(): Worker | null {
  if (typeof Worker === 'undefined') return null;
  if (!mathWorker) {
    mathWorker = new Worker(new URL('../workers/math.worker.ts', import.meta.url), {
      type: 'module',
    });
    mathWorker.onmessage = (e: MessageEvent<MathWorkerResponse>) => {
      const data = e.data;
      const pending = pendingMathRequests.get(data.id);
      if (pending) {
        pendingMathRequests.delete(data.id);
        pending.resolve(data);
      }
    };
  }
  return mathWorker;
}

/**
 * Offloads CSV/TSV parsing to io.worker.ts and returns a WorkbookSheet backed by Float64Array & Uint8Array
 */
export async function parseCsvInWorker(
  rawText: string,
  sheetName: string,
  options: IOWorkerParseRequest['options']
): Promise<WorkbookSheet> {
  const worker = getIOWorker();
  const reqId = `io_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;

  if (!worker) {
    throw new Error('Web Worker non supportato');
  }

  const res = await new Promise<IOWorkerParseResponse>((resolve, reject) => {
    pendingIORequests.set(reqId, { resolve, reject });
    worker.postMessage({
      id: reqId,
      type: 'PARSE_CSV_BINARY',
      rawText,
      sheetName,
      options,
    } satisfies IOWorkerParseRequest);
  });

  if (!res.ok) {
    throw new Error(res.error || 'Errore parsing CSV nel Web Worker');
  }

  const cols: DataColumn[] = res.columns.map((c, idx) => {
    const f64 = new Float64Array(c.float64Buffer);
    const u8 = new Uint8Array(c.maskUint8Buffer);
    return attachBinaryBuffersToColumn(
      {
        id: `col_wk_${Date.now()}_${idx}`,
        shortName: c.shortName,
        longName: c.longName,
        unit: c.unit,
        comment: c.comment,
        role: c.role,
        dataType: c.dataType,
        values: [],
      },
      f64,
      u8
    );
  });

  return rebindSheetMultiXAssociations({
    id: `sheet_wk_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    name: res.sheetName,
    columns: cols,
  });
}

/**
 * Offloads Set Column Values (F(x)=) formula evaluation to math.worker.ts using binary Float64Array buffers
 */
export async function evaluateFormulaInWorker(
  formula: string,
  targetCol: DataColumn,
  allColumns: DataColumn[]
): Promise<{ float64: Float64Array; maskUint8: Uint8Array }> {
  const worker = getMathWorker();
  if (!worker) {
    throw new Error('Math Worker non disponibile');
  }

  const reqId = `math_fx_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
  const transferables: Transferable[] = [];

  const columnsPayload = allColumns.map((col) => {
    const { float64 } = getOrSyncColumnBuffers(col);
    // Clone ArrayBuffer slice for transfer so main thread retains its buffer without detaching
    const copyBuf = float64.buffer.slice(0) as ArrayBuffer;
    transferables.push(copyBuf);
    return {
      id: col.id,
      shortName: col.shortName,
      float64Buffer: copyBuf,
    };
  });

  const rowCount = Math.max(20, ...allColumns.map((c) => getOrSyncColumnBuffers(c).length));

  const res = await new Promise<MathWorkerResponse>((resolve, reject) => {
    pendingMathRequests.set(reqId, { resolve, reject });
    worker.postMessage(
      {
        id: reqId,
        type: 'EVALUATE_FORMULA_BINARY',
        formula,
        targetColShortName: targetCol.shortName,
        rowCount,
        columns: columnsPayload,
      } satisfies MathWorkerRequest,
      transferables
    );
  });

  if (!res.ok || res.type !== 'EVALUATE_FORMULA_BINARY') {
    throw new Error(!res.ok ? res.error : 'Risposta inattesa dal Math Worker');
  }

  return {
    float64: new Float64Array(res.resultBuffer),
    maskUint8: new Uint8Array(res.maskBuffer),
  };
}

/**
 * Offloads Curve Fitting (Linear, Polynomial, Spline, Gaussian, Exponential) to math.worker.ts
 */
export async function computeCurveFitInWorker(params: {
  xCol: DataColumn;
  yCol: DataColumn;
  model: FitModelType;
  polyDegree: number;
  rangeC: number;
  rangeD: number;
  numPointsN: number;
}) {
  const worker = getMathWorker();
  if (!worker) {
    throw new Error('Math Worker non disponibile');
  }

  const xBuf = getOrSyncColumnBuffers(params.xCol);
  const yBuf = getOrSyncColumnBuffers(params.yCol);
  const len = Math.min(xBuf.length, yBuf.length);

  const combinedMask = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    if (xBuf.maskUint8[i] === 1 || yBuf.maskUint8[i] === 1) {
      combinedMask[i] = 1;
    } else if (xBuf.maskUint8[i] === 2 || yBuf.maskUint8[i] === 2) {
      combinedMask[i] = 2;
    } else {
      combinedMask[i] = 0;
    }
  }

  const xCopy = xBuf.float64.buffer.slice(0) as ArrayBuffer;
  const yCopy = yBuf.float64.buffer.slice(0) as ArrayBuffer;
  const maskCopy = combinedMask.buffer;

  const reqId = `math_fit_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;

  const res = await new Promise<MathWorkerResponse>((resolve, reject) => {
    pendingMathRequests.set(reqId, { resolve, reject });
    worker.postMessage(
      {
        id: reqId,
        type: 'CURVE_FIT_BINARY',
        xBuffer: xCopy,
        yBuffer: yCopy,
        maskBuffer: maskCopy,
        model: params.model,
        polyDegree: params.polyDegree,
        rangeC: params.rangeC,
        rangeD: params.rangeD,
        numPointsN: params.numPointsN,
      } satisfies MathWorkerRequest,
      [xCopy, yCopy, maskCopy]
    );
  });

  if (!res.ok || res.type !== 'CURVE_FIT_BINARY') {
    throw new Error(!res.ok ? res.error : 'Risposta inattesa dal Math Worker');
  }

  const fitX = new Float64Array(res.fitXBuffer);
  const fitY = new Float64Array(res.fitYBuffer);
  const generatedPoints: { x: number; y: number }[] = new Array(fitX.length);
  for (let i = 0; i < fitX.length; i++) {
    generatedPoints[i] = { x: fitX[i], y: fitY[i] };
  }

  return {
    equationText: res.equationText,
    coefficients: res.coefficients,
    rSquared: res.rSquared,
    rmse: res.rmse,
    sourceDomainMin: res.sourceDomainMin,
    sourceDomainMax: res.sourceDomainMax,
    sourceValidPointsCount: res.sourceValidPointsCount,
    generatedPoints,
    fitX,
    fitY,
    engineUsed: res.engineUsed,
  };
}

/**
 * Offloads Signal Filtering & Smoothing to math.worker.ts
 */
export async function applySignalFilterInWorker(params: {
  col: DataColumn;
  method: SmoothingFilterMethod;
  windowSize: number;
  polyOrder: number;
  cutoffRatio: number;
  outlierSigma: number;
}): Promise<{ float64: Float64Array; outliersRemovedCount: number }> {
  const worker = getMathWorker();
  if (!worker) {
    throw new Error('Math Worker non disponibile');
  }

  const yBuf = getOrSyncColumnBuffers(params.col);
  const yCopy = yBuf.float64.buffer.slice(0) as ArrayBuffer;
  const reqId = `math_flt_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;

  const res = await new Promise<MathWorkerResponse>((resolve, reject) => {
    pendingMathRequests.set(reqId, { resolve, reject });
    worker.postMessage(
      {
        id: reqId,
        type: 'SIGNAL_FILTER_BINARY',
        yBuffer: yCopy,
        method: params.method,
        windowSize: params.windowSize,
        polyOrder: params.polyOrder,
        cutoffRatio: params.cutoffRatio,
        outlierSigma: params.outlierSigma,
      } satisfies MathWorkerRequest,
      [yCopy]
    );
  });

  if (!res.ok || res.type !== 'SIGNAL_FILTER_BINARY') {
    throw new Error(!res.ok ? res.error : 'Risposta inattesa dal Math Worker');
  }

  return {
    float64: new Float64Array(res.filteredBuffer),
    outliersRemovedCount: res.outliersRemovedCount,
  };
}
