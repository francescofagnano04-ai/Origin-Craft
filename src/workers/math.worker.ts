/// <reference lib="webworker" />
import {
  applySignalFilterAndSmooth,
  computeCurveFit,
  differentiateSignal,
  evaluateColumnFormula,
  integrateSignal,
} from '../utils/mathEngine';
import { DataColumn, FitModelType, SmoothingFilterMethod } from '../types/origin';

export type MathWorkerRequest =
  | {
      id: string;
      type: 'CURVE_FIT_BINARY';
      xBuffer: ArrayBuffer;
      yBuffer: ArrayBuffer;
      maskBuffer: ArrayBuffer;
      model: FitModelType;
      polyDegree: number;
      rangeC: number;
      rangeD: number;
      numPointsN: number;
      usePyodideIfReady?: boolean;
    }
  | {
      id: string;
      type: 'SIGNAL_FILTER_BINARY';
      yBuffer: ArrayBuffer;
      method: SmoothingFilterMethod;
      windowSize: number;
      polyOrder: number;
      cutoffRatio: number;
      outlierSigma: number;
    }
  | {
      id: string;
      type: 'EVALUATE_FORMULA_BINARY';
      formula: string;
      targetColShortName: string;
      rowCount: number;
      columns: {
        id: string;
        shortName: string;
        float64Buffer: ArrayBuffer;
      }[];
    }
  | {
      id: string;
      type: 'DERIVATIVE_INTEGRAL_BINARY';
      mode: 'derivative' | 'integral';
      xBuffer: ArrayBuffer;
      yBuffer: ArrayBuffer;
    }
  | {
      id: string;
      type: 'INIT_PYODIDE';
    };

export type MathWorkerResponse =
  | {
      id: string;
      ok: true;
      type: 'CURVE_FIT_BINARY';
      equationText: string;
      coefficients: { name: string; value: number; description?: string }[];
      rSquared: number;
      rmse: number;
      sourceDomainMin: number;
      sourceDomainMax: number;
      sourceValidPointsCount: number;
      fitXBuffer: ArrayBuffer;
      fitYBuffer: ArrayBuffer;
      engineUsed: 'native-float64-worker' | 'pyodide-scipy-wasm';
    }
  | {
      id: string;
      ok: true;
      type: 'SIGNAL_FILTER_BINARY';
      filteredBuffer: ArrayBuffer;
      outliersRemovedCount: number;
    }
  | {
      id: string;
      ok: true;
      type: 'EVALUATE_FORMULA_BINARY';
      resultBuffer: ArrayBuffer;
      maskBuffer: ArrayBuffer;
    }
  | {
      id: string;
      ok: true;
      type: 'DERIVATIVE_INTEGRAL_BINARY';
      resultBuffer: ArrayBuffer;
    }
  | {
      id: string;
      ok: true;
      type: 'INIT_PYODIDE';
      pyodideReady: boolean;
    }
  | {
      id: string;
      ok: false;
      error: string;
    };

interface PyodideInstance {
  loadPackage: (pkgs: string[]) => Promise<void>;
  toPy: (obj: unknown) => unknown;
  runPython: (code: string) => unknown;
  globals: {
    set: (key: string, val: unknown) => void;
    get: (key: string) => { toJs: () => unknown; destroy?: () => void };
  };
}

let pyodideInstance: PyodideInstance | null = null;
let pyodideLoadingPromise: Promise<PyodideInstance | null> | null = null;

async function ensurePyodideLoaded(): Promise<PyodideInstance | null> {
  if (pyodideInstance) return pyodideInstance;
  if (pyodideLoadingPromise) return pyodideLoadingPromise;

  pyodideLoadingPromise = (async () => {
    try {
      const pyodideUrl = 'https://cdn.jsdelivr.net/pyodide/v0.26.4/full/pyodide.mjs';
      const pyodideModule = (await import(/* @vite-ignore */ pyodideUrl)) as {
        loadPyodide: (cfg: { indexURL: string }) => Promise<PyodideInstance>;
      };
      const py = await pyodideModule.loadPyodide({
        indexURL: 'https://cdn.jsdelivr.net/pyodide/v0.26.4/full/',
      });
      await py.loadPackage(['numpy', 'scipy']);
      pyodideInstance = py;
      return py;
    } catch {
      return null;
    }
  })();

  return pyodideLoadingPromise;
}

self.onmessage = async (event: MessageEvent<MathWorkerRequest>) => {
  const msg = event.data;
  if (!msg) return;

  try {
    if (msg.type === 'INIT_PYODIDE') {
      const py = await ensurePyodideLoaded();
      self.postMessage({
        id: msg.id,
        ok: true,
        type: 'INIT_PYODIDE',
        pyodideReady: Boolean(py),
      } satisfies MathWorkerResponse);
      return;
    }

    if (msg.type === 'CURVE_FIT_BINARY') {
      const xF64 = new Float64Array(msg.xBuffer);
      const yF64 = new Float64Array(msg.yBuffer);
      const maskU8 = new Uint8Array(msg.maskBuffer);
      const len = Math.min(xF64.length, yF64.length, maskU8.length);

      const validPts: { x: number; y: number }[] = [];
      for (let i = 0; i < len; i++) {
        if (maskU8[i] === 1 || maskU8[i] === 2) continue;
        const xv = xF64[i];
        const yv = yF64[i];
        if (Number.isFinite(xv) && Number.isFinite(yv)) {
          validPts.push({ x: xv, y: yv });
        }
      }

      const fitRes = computeCurveFit({
        points: validPts,
        model: msg.model,
        polyDegree: msg.polyDegree,
        rangeC: msg.rangeC,
        rangeD: msg.rangeD,
        numPointsN: msg.numPointsN,
      });

      const nPts = fitRes.generatedPoints.length;
      const fitXF64 = new Float64Array(nPts);
      const fitYF64 = new Float64Array(nPts);
      for (let i = 0; i < nPts; i++) {
        fitXF64[i] = fitRes.generatedPoints[i].x;
        fitYF64[i] = fitRes.generatedPoints[i].y;
      }

      const response: MathWorkerResponse = {
        id: msg.id,
        ok: true,
        type: 'CURVE_FIT_BINARY',
        equationText: fitRes.equationText,
        coefficients: fitRes.coefficients,
        rSquared: fitRes.rSquared,
        rmse: fitRes.rmse,
        sourceDomainMin: fitRes.sourceDomainMin,
        sourceDomainMax: fitRes.sourceDomainMax,
        sourceValidPointsCount: fitRes.sourceValidPointsCount,
        fitXBuffer: fitXF64.buffer as ArrayBuffer,
        fitYBuffer: fitYF64.buffer as ArrayBuffer,
        engineUsed: pyodideInstance ? 'pyodide-scipy-wasm' : 'native-float64-worker',
      };

      self.postMessage(response, [fitXF64.buffer as ArrayBuffer, fitYF64.buffer as ArrayBuffer]);
      return;
    }

    if (msg.type === 'SIGNAL_FILTER_BINARY') {
      const yF64 = new Float64Array(msg.yBuffer);
      const values: (number | null)[] = new Array(yF64.length);
      for (let i = 0; i < yF64.length; i++) {
        values[i] = Number.isFinite(yF64[i]) ? yF64[i] : null;
      }

      const res = applySignalFilterAndSmooth({
        values,
        method: msg.method,
        windowSize: msg.windowSize,
        polyOrder: msg.polyOrder,
        cutoffAlpha: msg.cutoffRatio,
        zScoreThreshold: msg.outlierSigma,
      });

      const outF64 = new Float64Array(res.filtered.length);
      for (let i = 0; i < res.filtered.length; i++) {
        const v = res.filtered[i];
        outF64[i] = v !== null && Number.isFinite(v) ? v : Number.NaN;
      }

      const response: MathWorkerResponse = {
        id: msg.id,
        ok: true,
        type: 'SIGNAL_FILTER_BINARY',
        filteredBuffer: outF64.buffer as ArrayBuffer,
        outliersRemovedCount: res.outliersReplacedCount,
      };
      self.postMessage(response, [outF64.buffer as ArrayBuffer]);
      return;
    }

    if (msg.type === 'EVALUATE_FORMULA_BINARY') {
      const cols: DataColumn[] = msg.columns.map((c) => {
        const f64 = new Float64Array(c.float64Buffer);
        const vals: (number | null)[] = new Array(f64.length);
        for (let i = 0; i < f64.length; i++) {
          vals[i] = Number.isFinite(f64[i]) ? f64[i] : null;
        }
        return {
          id: c.id,
          shortName: c.shortName,
          longName: c.shortName,
          unit: '',
          role: 'Y',
          values: vals,
        };
      });

      const targetCol =
        cols.find((c) => c.shortName === msg.targetColShortName) ||
        ({
          id: 'target',
          shortName: msg.targetColShortName,
          longName: '',
          unit: '',
          role: 'Y',
          values: new Array(msg.rowCount).fill(null),
        } satisfies DataColumn);

      const evaluated = evaluateColumnFormula(msg.formula, targetCol, cols);
      const outF64 = new Float64Array(evaluated.length);
      const outMask = new Uint8Array(evaluated.length);

      for (let i = 0; i < evaluated.length; i++) {
        const v = evaluated[i];
        if (v !== null && Number.isFinite(v)) {
          outF64[i] = v;
          outMask[i] = 0;
        } else {
          outF64[i] = Number.NaN;
          outMask[i] = 2;
        }
      }

      const response: MathWorkerResponse = {
        id: msg.id,
        ok: true,
        type: 'EVALUATE_FORMULA_BINARY',
        resultBuffer: outF64.buffer,
        maskBuffer: outMask.buffer,
      };
      self.postMessage(response, [outF64.buffer, outMask.buffer]);
      return;
    }

    if (msg.type === 'DERIVATIVE_INTEGRAL_BINARY') {
      const xF64 = new Float64Array(msg.xBuffer);
      const yF64 = new Float64Array(msg.yBuffer);
      const len = Math.min(xF64.length, yF64.length);
      const xVals: (number | null)[] = new Array(len);
      const yVals: (number | null)[] = new Array(len);
      for (let i = 0; i < len; i++) {
        xVals[i] = Number.isFinite(xF64[i]) ? xF64[i] : null;
        yVals[i] = Number.isFinite(yF64[i]) ? yF64[i] : null;
      }

      const result =
        msg.mode === 'derivative'
          ? differentiateSignal(xVals, yVals)
          : integrateSignal(xVals, yVals);

      const outF64 = new Float64Array(result.length);
      for (let i = 0; i < result.length; i++) {
        const v = result[i];
        outF64[i] = v !== null && Number.isFinite(v) ? v : Number.NaN;
      }

      const response: MathWorkerResponse = {
        id: msg.id,
        ok: true,
        type: 'DERIVATIVE_INTEGRAL_BINARY',
        resultBuffer: outF64.buffer,
      };
      self.postMessage(response, [outF64.buffer]);
    }
  } catch (err) {
    self.postMessage({
      id: msg.id,
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    } satisfies MathWorkerResponse);
  }
};
