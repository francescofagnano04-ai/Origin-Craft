import {
  ColumnDataType,
  ColumnStatistics,
  DataColumn,
  FitModelType,
  FittedCurve,
  GraphWindow,
  NumericFormatMode,
  NumericPrecisionMode,
  RowFilterConfig,
  SmoothingFilterMethod,
  WorkbookSheet,
} from '../types/origin';
import {
  POINT_MASKED,
  POINT_NAN,
  POINT_VALID,
  attachBinaryBuffersToColumn,
  getColumnRowCount,
  getOrSyncColumnBuffers,
  getSheetRowCount,
} from './binaryColumnStore';

/**
 * Generates column letters like OriginLab: A, B, C, ... Z, AA, AB
 */
export function getColumnShortName(index: number): string {
  let name = '';
  let n = index;
  while (n >= 0) {
    name = String.fromCharCode((n % 26) + 65) + name;
    n = Math.floor(n / 26) - 1;
  }
  return name;
}

/**
 * Resolves the effective X column for a given Y (or Y_ERR / X_ERR / LABEL) column in a sheet.
 * Follows OriginLab Multi-X convention: associates automatically with the nearest X column to its left
 * (e.g. X1, Y1, X2, Y2), unless an explicit valid linkedXColId was set and there is only 1 X or it matches.
 */
export function resolveXColumnForY(sheet: WorkbookSheet, yCol: DataColumn): DataColumn | undefined {
  const yIndex = sheet.columns.findIndex((c) => c.id === yCol.id);

  // 1. Search backwards from yCol to find the nearest X column to its left (OriginLab Multi-X rule)
  if (yIndex > 0) {
    for (let i = yIndex - 1; i >= 0; i--) {
      if (sheet.columns[i].role === 'X') {
        return sheet.columns[i];
      }
    }
  }

  // 2. Fallback to explicit linkedXColId if present
  if (yCol.linkedXColId) {
    const explicit = sheet.columns.find((c) => c.id === yCol.linkedXColId && c.role === 'X');
    if (explicit) return explicit;
  }

  // 3. Fallback to first X column in sheet or first column
  return sheet.columns.find((c) => c.role === 'X') || sheet.columns[0];
}

/**
 * Resolves companion columns (xErr, yErr, Label) for a given Y column in a sheet
 */
export function resolveErrorAndLabelColumnsForY(
  sheet: WorkbookSheet,
  yCol: DataColumn
): {
  xErrCol?: DataColumn;
  yErrCol?: DataColumn;
  labelCol?: DataColumn;
} {
  const yIdx = sheet.columns.findIndex((c) => c.id === yCol.id);
  const xCol = resolveXColumnForY(sheet, yCol);
  const xIdx = xCol ? sheet.columns.findIndex((c) => c.id === xCol.id) : -1;

  let yErrCol = yCol.linkedErrColId
    ? sheet.columns.find((c) => c.id === yCol.linkedErrColId && c.role === 'Y_ERR')
    : undefined;
  if (!yErrCol && yIdx >= 0) {
    // Look to the right of yCol before the next Y or X column
    for (let i = yIdx + 1; i < sheet.columns.length; i++) {
      const col = sheet.columns[i];
      if (col.role === 'Y_ERR') {
        yErrCol = col;
        break;
      }
      if (col.role === 'Y' || col.role === 'X') break;
    }
  }

  let xErrCol = yCol.linkedXErrColId
    ? sheet.columns.find((c) => c.id === yCol.linkedXErrColId && c.role === 'X_ERR')
    : undefined;
  if (!xErrCol) {
    // Look to the right of xCol or yCol for an X_ERR column in the same X group
    const startSearch = xIdx >= 0 ? xIdx + 1 : 0;
    for (let i = startSearch; i < sheet.columns.length; i++) {
      const col = sheet.columns[i];
      if (col.role === 'X' && i !== xIdx) break;
      if (col.role === 'X_ERR') {
        xErrCol = col;
        break;
      }
    }
  }

  let labelCol: DataColumn | undefined;
  const startLbl = xIdx >= 0 ? xIdx : 0;
  for (let i = startLbl; i < sheet.columns.length; i++) {
    const col = sheet.columns[i];
    if (i > startLbl && col.role === 'X') break;
    if (col.role === 'LABEL') {
      labelCol = col;
      break;
    }
  }
  if (!labelCol) {
    labelCol = sheet.columns.find((c) => c.role === 'LABEL');
  }

  return { xErrCol, yErrCol, labelCol };
}

/**
 * Computes OriginLab-style column role badge with Multi-X group numbering:
 * e.g. X1, Y1, X2, Y2, xEr±, yEr±, L, Ign
 */
export function getOriginColumnRoleBadge(sheet: WorkbookSheet, colIndex: number): string {
  const col = sheet.columns[colIndex];
  if (!col) return '';
  const totalXCols = sheet.columns.filter((c) => c.role === 'X').length;
  const multiX = totalXCols > 1;

  // Count how many X columns appear up to colIndex
  let xGroupNumber = 0;
  for (let i = 0; i <= colIndex; i++) {
    if (sheet.columns[i].role === 'X') {
      xGroupNumber++;
    }
  }
  if (xGroupNumber === 0 && totalXCols > 0) xGroupNumber = 1;

  switch (col.role) {
    case 'X':
      return multiX ? `X${xGroupNumber}` : 'X';
    case 'Y':
      return multiX ? `Y${xGroupNumber}` : 'Y';
    case 'X_ERR':
      return multiX ? `xEr±${xGroupNumber}` : 'xEr±';
    case 'Y_ERR':
      return multiX ? `yEr±${xGroupNumber}` : 'yEr±';
    case 'LABEL':
      return 'L';
    case 'IGNORE':
      return 'Ign';
    default:
      return col.role;
  }
}

/**
 * Rebinds all columns in a sheet after role changes or Drag & Drop column reordering:
 * - Updates shortName (A, B, C, ...)
 * - Associates each Y / Y_ERR / X_ERR / LABEL column with the nearest X column to its left
 */
export function rebindSheetMultiXAssociations(sheet: WorkbookSheet): WorkbookSheet {
  const reindexed = sheet.columns.map((col, idx) => ({
    ...col,
    shortName: getColumnShortName(idx),
  }));

  const firstX = reindexed.find((c) => c.role === 'X');
  let currentLeftXId: string | undefined = firstX?.id;
  let lastYColIndex = -1;

  for (let i = 0; i < reindexed.length; i++) {
    const col = reindexed[i];
    if (col.role === 'X') {
      currentLeftXId = col.id;
      col.linkedXColId = undefined;
    } else {
      col.linkedXColId = currentLeftXId;
      if (col.role === 'Y') {
        lastYColIndex = i;
      } else if (col.role === 'Y_ERR' && lastYColIndex >= 0) {
        reindexed[lastYColIndex].linkedErrColId = col.id;
      } else if (col.role === 'X_ERR' && lastYColIndex >= 0) {
        reindexed[lastYColIndex].linkedXErrColId = col.id;
      }
    }
  }

  return {
    ...sheet,
    columns: reindexed,
  };
}

/**
 * Formats a number cleanly for scientific display and equations
 */
export function formatSciNumber(val: number, digits = 4): string {
  if (!Number.isFinite(val)) return 'NaN';
  if (Math.abs(val) === 0) return '0';
  if (Math.abs(val) >= 1e5 || (Math.abs(val) > 0 && Math.abs(val) < 1e-3)) {
    return val.toExponential(digits - 1);
  }
  return Number(val.toPrecision(digits)).toString();
}

/**
 * Solves a linear system A * x = b using Gaussian elimination with partial pivoting
 */
function solveLinearSystem(A: number[][], b: number[]): number[] {
  const n = b.length;
  const M = A.map((row, i) => [...row, b[i]]);

  for (let col = 0; col < n; col++) {
    let maxRow = col;
    let maxVal = Math.abs(M[col][col]);
    for (let row = col + 1; row < n; row++) {
      const absVal = Math.abs(M[row][col]);
      if (absVal > maxVal) {
        maxVal = absVal;
        maxRow = row;
      }
    }

    if (maxRow !== col) {
      const tmp = M[col];
      M[col] = M[maxRow];
      M[maxRow] = tmp;
    }

    const pivot = M[col][col];
    if (Math.abs(pivot) < 1e-14) {
      continue;
    }

    for (let j = col; j <= n; j++) {
      M[col][j] /= pivot;
    }

    for (let row = 0; row < n; row++) {
      if (row !== col) {
        const factor = M[row][col];
        for (let j = col; j <= n; j++) {
          M[row][j] -= factor * M[col][j];
        }
      }
    }
  }

  return M.map((row) => row[n]);
}

/**
 * Natural Cubic Spline solver & evaluator
 */
export function createNaturalCubicSpline(rawPoints: { x: number; y: number }[]) {
  const sorted = [...rawPoints].sort((a, b) => a.x - b.x);
  const knots: { x: number; y: number }[] = [];
  for (const pt of sorted) {
    if (knots.length > 0 && Math.abs(knots[knots.length - 1].x - pt.x) < 1e-12) {
      knots[knots.length - 1].y = (knots[knots.length - 1].y + pt.y) / 2;
    } else {
      knots.push({ x: pt.x, y: pt.y });
    }
  }

  const n = knots.length;
  if (n === 0) {
    return { evaluate: () => 0, knotsCount: 0 };
  }
  if (n === 1) {
    return { evaluate: () => knots[0].y, knotsCount: 1 };
  }
  if (n === 2) {
    const slope = (knots[1].y - knots[0].y) / (knots[1].x - knots[0].x || 1e-9);
    return {
      evaluate: (x: number) => knots[0].y + slope * (x - knots[0].x),
      knotsCount: 2,
    };
  }

  const h: number[] = new Array(n - 1);
  for (let i = 0; i < n - 1; i++) {
    h[i] = knots[i + 1].x - knots[i].x;
    if (h[i] <= 0) h[i] = 1e-9;
  }

  const mSize = n - 2;
  const A: number[][] = Array.from({ length: mSize }, () => new Array(mSize).fill(0));
  const rhs: number[] = new Array(mSize).fill(0);

  for (let i = 0; i < mSize; i++) {
    const hPrev = h[i];
    const hNext = h[i + 1];
    A[i][i] = 2 * (hPrev + hNext);
    if (i > 0) A[i][i - 1] = hPrev;
    if (i < mSize - 1) A[i][i + 1] = hNext;

    const dyNext = (knots[i + 2].y - knots[i + 1].y) / hNext;
    const dyPrev = (knots[i + 1].y - knots[i].y) / hPrev;
    rhs[i] = 6 * (dyNext - dyPrev);
  }

  const interiorM = solveLinearSystem(A, rhs);
  const M = [0, ...interiorM, 0];

  const evaluate = (x: number): number => {
    if (x <= knots[0].x) {
      const h0 = h[0];
      const slopeLeft =
        (knots[1].y - knots[0].y) / h0 - (h0 / 6) * (2 * M[0] + M[1]);
      return knots[0].y + slopeLeft * (x - knots[0].x);
    }
    if (x >= knots[n - 1].x) {
      const hLast = h[n - 2];
      const slopeRight =
        (knots[n - 1].y - knots[n - 2].y) / hLast +
        (hLast / 6) * (M[n - 2] + 2 * M[n - 1]);
      return knots[n - 1].y + slopeRight * (x - knots[n - 1].x);
    }

    let low = 0;
    let high = n - 2;
    while (low <= high) {
      const mid = (low + high) >> 1;
      if (x >= knots[mid].x && x <= knots[mid + 1].x) {
        low = mid;
        break;
      } else if (x < knots[mid].x) {
        high = mid - 1;
      } else {
        low = mid + 1;
      }
    }
    const i = Math.max(0, Math.min(n - 2, low));
    const hi = h[i];
    const dxLeft = knots[i + 1].x - x;
    const dxRight = x - knots[i].x;

    const term1 = (M[i] * Math.pow(dxLeft, 3)) / (6 * hi);
    const term2 = (M[i + 1] * Math.pow(dxRight, 3)) / (6 * hi);
    const term3 = ((knots[i].y - (M[i] * hi * hi) / 6) * dxLeft) / hi;
    const term4 = ((knots[i + 1].y - (M[i + 1] * hi * hi) / 6) * dxRight) / hi;

    return term1 + term2 + term3 + term4;
  };

  return { evaluate, knotsCount: n };
}

/**
 * Core Curve Fitting & Custom Resampling Engine
 */
export function computeCurveFit(params: {
  points: { x: number; y: number }[];
  model: FitModelType;
  polyDegree: number;
  rangeC: number;
  rangeD: number;
  numPointsN: number;
}): {
  equationText: string;
  coefficients: { name: string; value: number; description?: string }[];
  rSquared: number;
  rmse: number;
  sourceDomainMin: number;
  sourceDomainMax: number;
  sourceValidPointsCount: number;
  generatedPoints: { x: number; y: number }[];
} {
  const { points, model, polyDegree, rangeC, rangeD, numPointsN } = params;
  const valid = points.filter((p) => Number.isFinite(p.x) && Number.isFinite(p.y));
  const m = valid.length;

  if (m < 2) {
    return {
      equationText: 'Dati insufficienti (minimo 2 punti validi)',
      coefficients: [],
      rSquared: 0,
      rmse: 0,
      sourceDomainMin: 0,
      sourceDomainMax: 1,
      sourceValidPointsCount: m,
      generatedPoints: [],
    };
  }

  let minX = Infinity;
  let maxX = -Infinity;
  let sumY = 0;
  for (const p of valid) {
    if (p.x < minX) minX = p.x;
    if (p.x > maxX) maxX = p.x;
    sumY += p.y;
  }
  const meanY = sumY / m;

  let evaluateFn: (x: number) => number = () => 0;
  let equationText = '';
  const coefficients: { name: string; value: number; description?: string }[] = [];

  if (model === 'linear') {
    let sxx = 0;
    let sxy = 0;
    let sumX = 0;
    for (const p of valid) {
      sumX += p.x;
    }
    const meanX = sumX / m;
    for (const p of valid) {
      const dx = p.x - meanX;
      const dy = p.y - meanY;
      sxx += dx * dx;
      sxy += dx * dy;
    }
    const b = Math.abs(sxx) > 1e-15 ? sxy / sxx : 0;
    const a = meanY - b * meanX;

    evaluateFn = (x: number) => a + b * x;
    const signB = b >= 0 ? '+' : '-';
    equationText = `y = ${formatSciNumber(a, 5)} ${signB} ${formatSciNumber(Math.abs(b), 5)} · x`;
    coefficients.push(
      { name: 'a (Intercetta)', value: a, description: 'Valore di Y per X = 0' },
      { name: 'b (Pendenza)', value: b, description: 'Coeff. angolare ΔY/ΔX' }
    );
  } else if (model === 'polynomial') {
    const degree = Math.max(1, Math.min(8, Math.floor(polyDegree), m - 1));
    const size = degree + 1;

    const XtX: number[][] = Array.from({ length: size }, () => new Array(size).fill(0));
    const Xty: number[] = new Array(size).fill(0);

    for (const p of valid) {
      const xPowers: number[] = new Array(2 * degree + 1);
      xPowers[0] = 1;
      for (let k = 1; k <= 2 * degree; k++) {
        xPowers[k] = xPowers[k - 1] * p.x;
      }
      for (let r = 0; r < size; r++) {
        for (let c = 0; c < size; c++) {
          XtX[r][c] += xPowers[r + c];
        }
        Xty[r] += xPowers[r] * p.y;
      }
    }

    const cCoeffs = solveLinearSystem(XtX, Xty);
    evaluateFn = (x: number) => {
      let acc = 0;
      let xp = 1;
      for (let k = 0; k < cCoeffs.length; k++) {
        acc += cCoeffs[k] * xp;
        xp *= x;
      }
      return acc;
    };

    const parts: string[] = [];
    cCoeffs.forEach((cVal, idx) => {
      coefficients.push({
        name: `c${idx}`,
        value: cVal,
        description: idx === 0 ? 'Termine noto' : `Coeff. di x^${idx}`,
      });
      const absStr = formatSciNumber(Math.abs(cVal), 4);
      const sign = cVal >= 0 ? (idx === 0 ? '' : ' + ') : idx === 0 ? '-' : ' - ';
      if (idx === 0) parts.push(`${sign}${absStr}`);
      else if (idx === 1) parts.push(`${sign}${absStr}·x`);
      else parts.push(`${sign}${absStr}·x^${idx}`);
    });
    equationText = `y = ${parts.join('')}`;
  } else if (model === 'gaussian') {
    let maxPt = valid[0];
    let minPt = valid[0];
    for (const p of valid) {
      if (p.y > maxPt.y) maxPt = p;
      if (p.y < minPt.y) minPt = p;
    }
    let y0 = minPt.y;
    let A = Math.max(1e-6, maxPt.y - y0);
    let xc = maxPt.x;
    const halfAmp = y0 + A / 2;
    let w = (maxX - minX) / 6 || 1;
    const aboveHalf = valid.filter((p) => p.y >= halfAmp);
    if (aboveHalf.length >= 2) {
      const fwhmEst = aboveHalf[aboveHalf.length - 1].x - aboveHalf[0].x;
      if (fwhmEst > 1e-6) {
        w = fwhmEst / 2.354820045;
      }
    }

    // Damped Gauss-Newton (Levenberg-Marquardt) 4-parameter refinement: [y0, A, xc, w]
    let lambda = 1e-3;
    const calcSse = (pY0: number, pA: number, pXc: number, pW: number) => {
      let s = 0;
      for (const p of valid) {
        const diff = p.x - pXc;
        const e = Math.exp(-Math.min(80, (diff * diff) / (2 * pW * pW)));
        const r = p.y - (pY0 + pA * e);
        s += r * r;
      }
      return s;
    };
    let currentSse = calcSse(y0, A, xc, w);

    for (let iter = 0; iter < 60; iter++) {
      const JTJ = [
        [0, 0, 0, 0],
        [0, 0, 0, 0],
        [0, 0, 0, 0],
        [0, 0, 0, 0],
      ];
      const JTr = [0, 0, 0, 0];

      for (const p of valid) {
        const diff = p.x - xc;
        const w2 = w * w;
        const arg = (diff * diff) / (2 * w2);
        const e = Math.exp(-Math.min(80, arg));
        const yHat = y0 + A * e;
        const r = p.y - yHat;

        const j0 = 1;
        const j1 = e;
        const j2 = A * e * (diff / w2);
        const j3 = A * e * ((diff * diff) / (w2 * w));
        const J = [j0, j1, j2, j3];

        for (let row = 0; row < 4; row++) {
          JTr[row] += J[row] * r;
          for (let col = 0; col < 4; col++) {
            JTJ[row][col] += J[row] * J[col];
          }
        }
      }

      for (let d = 0; d < 4; d++) {
        JTJ[d][d] += lambda * (JTJ[d][d] + 1e-6);
      }

      const delta = solveLinearSystem(JTJ, JTr);
      const nextY0 = y0 + delta[0];
      const nextA = Math.max(1e-8, A + delta[1]);
      const nextXc = xc + delta[2];
      const nextW = Math.max(1e-5, Math.abs(w + delta[3]));
      const nextSse = calcSse(nextY0, nextA, nextXc, nextW);

      if (nextSse < currentSse) {
        y0 = nextY0;
        A = nextA;
        xc = nextXc;
        w = nextW;
        currentSse = nextSse;
        lambda = Math.max(1e-7, lambda * 0.3);
        if (
          Math.abs(delta[0]) + Math.abs(delta[1]) + Math.abs(delta[2]) + Math.abs(delta[3]) <
          1e-9
        ) {
          break;
        }
      } else {
        lambda = Math.min(1e5, lambda * 5);
      }
    }

    const fwhm = 2.354820045 * Math.abs(w);
    const area = A * Math.abs(w) * Math.sqrt(2 * Math.PI);

    evaluateFn = (x: number) => {
      const diff = x - xc;
      return y0 + A * Math.exp(-(diff * diff) / (2 * w * w));
    };

    equationText = `y = ${formatSciNumber(y0, 4)} + ${formatSciNumber(A, 4)}·exp(-(x - ${formatSciNumber(xc, 4)})² / (2·${formatSciNumber(w, 4)}²))`;
    coefficients.push(
      { name: 'y0 (Linea base)', value: y0, description: 'Livello di fondo' },
      { name: 'xc (Centro Picco)', value: xc, description: 'Posizione centro' },
      { name: 'A (Ampiezza Picco)', value: A, description: 'Altezza del picco rispetto alla base' },
      { name: 'w (Larghezza σ)', value: w, description: 'Deviazione standard Gaussiana' },
      { name: 'FWHM', value: fwhm, description: 'Larghezza a metà altezza (2.355·w)' },
      { name: 'Area Picco', value: area, description: "Integrale dell'area del picco" }
    );
  } else if (model === 'exponential') {
    let y0 = Math.min(...valid.map((p) => p.y)) - 1e-3;
    let A = Math.max(1e-4, valid[0].y - y0);
    let k = -0.5;

    const shifted = valid.filter((p) => p.y - y0 > 1e-4);
    if (shifted.length >= 3) {
      let sumX = 0,
        sumL = 0,
        sumXX = 0,
        sumXL = 0;
      for (const p of shifted) {
        const lx = Math.log(p.y - y0);
        sumX += p.x;
        sumL += lx;
        sumXX += p.x * p.x;
        sumXL += p.x * lx;
      }
      const count = shifted.length;
      const denom = count * sumXX - sumX * sumX;
      if (Math.abs(denom) > 1e-9) {
        k = (count * sumXL - sumX * sumL) / denom;
        const lnA = (sumL - k * sumX) / count;
        A = Math.exp(lnA);
      }
    }

    // Damped Gauss-Newton (Levenberg-Marquardt) 3-parameter refinement: [y0, A, k]
    let lambda = 1e-3;
    const calcExpSse = (pY0: number, pA: number, pK: number) => {
      let s = 0;
      for (const p of valid) {
        const e = Math.exp(Math.max(-60, Math.min(60, pK * p.x)));
        const r = p.y - (pY0 + pA * e);
        s += r * r;
      }
      return s;
    };
    let currentSse = calcExpSse(y0, A, k);

    for (let iter = 0; iter < 60; iter++) {
      const JTJ = [
        [0, 0, 0],
        [0, 0, 0],
        [0, 0, 0],
      ];
      const JTr = [0, 0, 0];

      for (const p of valid) {
        const e = Math.exp(Math.max(-60, Math.min(60, k * p.x)));
        const yHat = y0 + A * e;
        const r = p.y - yHat;
        const J = [1, e, A * p.x * e];

        for (let row = 0; row < 3; row++) {
          JTr[row] += J[row] * r;
          for (let col = 0; col < 3; col++) {
            JTJ[row][col] += J[row] * J[col];
          }
        }
      }

      for (let d = 0; d < 3; d++) {
        JTJ[d][d] += lambda * (JTJ[d][d] + 1e-6);
      }

      const delta = solveLinearSystem(JTJ, JTr);
      const nextY0 = y0 + delta[0];
      const nextA = A + delta[1];
      const nextK = Math.max(-25, Math.min(25, k + delta[2]));
      const nextSse = calcExpSse(nextY0, nextA, nextK);

      if (nextSse < currentSse) {
        y0 = nextY0;
        A = nextA;
        k = nextK;
        currentSse = nextSse;
        lambda = Math.max(1e-7, lambda * 0.3);
        if (Math.abs(delta[0]) + Math.abs(delta[1]) + Math.abs(delta[2]) < 1e-9) {
          break;
        }
      } else {
        lambda = Math.min(1e5, lambda * 5);
      }
    }

    evaluateFn = (x: number) => y0 + A * Math.exp(Math.max(-60, Math.min(60, k * x)));
    const tau = Math.abs(k) > 1e-9 ? 1 / Math.abs(k) : Infinity;

    equationText = `y = ${formatSciNumber(y0, 4)} + ${formatSciNumber(A, 4)}·exp(${formatSciNumber(k, 4)}·x)`;
    coefficients.push(
      { name: 'y0 (Offset)', value: y0, description: 'Asintoto orizzontale' },
      { name: 'A (Ampiezza)', value: A, description: 'Fattore moltiplicativo' },
      { name: 'k (Costante tasso)', value: k, description: 'Tasso esponenziale' },
      { name: 'τ (Costante tempo)', value: tau, description: 'Tempo caratteristico 1/|k|' }
    );
  } else {
    const spline = createNaturalCubicSpline(valid);
    evaluateFn = spline.evaluate;
    equationText = `Spline Cubica Naturale (${spline.knotsCount} nodi C²)`;
    coefficients.push(
      { name: 'Nodi (Knots)', value: spline.knotsCount, description: 'Punti sperimentali interpolati' },
      { name: 'X min nodo', value: minX, description: 'Inizio dominio nodi' },
      { name: 'X max nodo', value: maxX, description: 'Fine dominio nodi' }
    );
  }

  let ssRes = 0;
  let ssTot = 0;
  for (const p of valid) {
    const yPred = evaluateFn(p.x);
    const err = p.y - yPred;
    ssRes += err * err;
    const dev = p.y - meanY;
    ssTot += dev * dev;
  }

  const rSquared = ssTot > 1e-15 ? Math.max(0, 1 - ssRes / ssTot) : 1;
  const rmse = Math.sqrt(ssRes / m);

  const safeN = Math.max(2, Math.min(5000, Math.round(numPointsN)));
  const generatedPoints: { x: number; y: number }[] = [];
  const step = (rangeD - rangeC) / (safeN - 1);

  for (let i = 0; i < safeN; i++) {
    const xVal = i === safeN - 1 ? rangeD : rangeC + i * step;
    const yVal = evaluateFn(xVal);
    generatedPoints.push({
      x: Number(xVal.toPrecision(12)),
      y: Number(yVal.toPrecision(12)),
    });
  }

  return {
    equationText,
    coefficients,
    rSquared,
    rmse,
    sourceDomainMin: minX,
    sourceDomainMax: maxX,
    sourceValidPointsCount: m,
    generatedPoints,
  };
}

/**
 * Strips optional left-hand assignment like `col(C) = ...` or `C = ...`
 * and returns `{ explicitTargetRef?: string; rhsFormula: string }`.
 */
export function parseFormulaTargetColumnAndExpr(rawFormula: string): {
  explicitTargetRef?: string;
  rhsFormula: string;
} {
  const trimmed = rawFormula.trim();
  const assignMatch = trimmed.match(
    /^(?:col\s*\(\s*["']?([a-zA-Z0-9_\s]+)["']?\s*\)|([a-zA-Z]{1,3}))\s*=(?!=)\s*(.+)$/i
  );
  if (assignMatch) {
    const targetRef = (assignMatch[1] || assignMatch[2] || '').trim();
    const rhs = assignMatch[3].trim();
    return { explicitTargetRef: targetRef, rhsFormula: rhs };
  }
  return { rhsFormula: trimmed };
}

/**
 * OriginLab Vectorial Formula Calculation on Column (Set Column Values, F(x)=)
 * Supports:
 * - Full vectorial expressions: `col(C) = col(A) * 2 + sin(col(B))` or `col(A) * 2 + sin(col(B))`
 * - Power operator `^` (e.g. `col(A)^2`)
 * - Row index variables `i` (1..N) and `row` (1..N)
 * - Column aggregate functions: `mean(col(A))`, `sd(col(A))`, `min(col(A))`, `max(col(A))`, `sum(col(A))`
 * - Vector transforms: `norm(col(A))`, `zscore(col(A))`, `diff(col(A))`, `cumsum(col(A))`
 * - Elementary error propagation functions:
 *   - `err_add(col(A), col(dA), col(B), col(dB))`
 *   - `err_sub(col(A), col(dA), col(B), col(dB))`
 *   - `err_mul(col(A), col(dA), col(B), col(dB))`
 *   - `err_div(col(A), col(dA), col(B), col(dB))`
 *   - `err_pow(col(A), col(dA), n)`
 */
export function evaluateColumnFormula(
  formula: string,
  targetCol: DataColumn,
  allCols: DataColumn[]
): (number | null)[] {
  const { rhsFormula } = parseFormulaTargetColumnAndExpr(formula);
  const cleanFormula = rhsFormula.trim();
  if (!cleanFormula) return targetCol.values;

  const rowCount = Math.max(1, ...allCols.map((c) => Math.max(c.values.length, getColumnRowCount(c))));
  const result: (number | null)[] = new Array(rowCount).fill(null);

  // Build fast lookup map of columns by shortName and longName
  const colIndexMap = new Map<string, number>();
  const colValuesArray: (number | null)[][] = allCols.map((c, idx) => {
    colIndexMap.set(c.shortName.toUpperCase(), idx);
    if (c.longName && c.longName.trim()) {
      colIndexMap.set(c.longName.trim().toUpperCase(), idx);
    }
    return c.values;
  });

  // Lazy precomputed column statistics & vector transforms for aggregate calls
  const statsCache = new Map<number, ColumnStatistics>();
  const getColStats = (cIdx: number): ColumnStatistics => {
    if (!statsCache.has(cIdx)) {
      statsCache.set(cIdx, computeColumnStatistics(allCols[cIdx]));
    }
    return statsCache.get(cIdx)!;
  };

  const diffCache = new Map<number, (number | null)[]>();
  const cumsumCache = new Map<number, (number | null)[]>();

  const getColVal = (cIdx: number, r: number): number => {
    if (cIdx < 0 || cIdx >= colValuesArray.length) return Number.NaN;
    const arr = colValuesArray[cIdx];
    const v = r < arr.length ? arr[r] : null;
    return v !== null && v !== undefined && Number.isFinite(v) ? v : Number.NaN;
  };

  const resolveColIdx = (refStr: string): number => {
    const cleaned = refStr
      .replace(/^col\s*\(\s*["']?/i, '')
      .replace(/["']?\s*\)$/i, '')
      .replace(/^["']|["']$/g, '')
      .trim()
      .toUpperCase();
    return colIndexMap.get(cleaned) ?? -1;
  };

  // Transform aggregate/vector functions before row-level evaluation
  let jsExpr = cleanFormula;

  // Replace aggregate scalar functions: mean(col(A)), sd(col(A)), std(col(A)), min(col(A)), max(col(A)), sum(col(A)), median(col(A))
  jsExpr = jsExpr.replace(
    /\b(mean|sd|std|min|max|sum|median)\s*\(\s*col\s*\(\s*["']?([a-zA-Z0-9_\s]+)["']?\s*\)\s*\)/gi,
    (_, fnName, colRef) => {
      const cIdx = resolveColIdx(colRef);
      if (cIdx < 0) return 'NaN';
      const st = getColStats(cIdx);
      const fn = fnName.toLowerCase();
      if (fn === 'mean') return String(st.mean);
      if (fn === 'sd' || fn === 'std') return String(st.sd);
      if (fn === 'min') return String(st.min);
      if (fn === 'max') return String(st.max);
      if (fn === 'sum') return String(st.sum);
      if (fn === 'median') return String(st.median);
      return 'NaN';
    }
  );

  // Replace vector transform functions: norm(col(A)), zscore(col(A)), diff(col(A)), cumsum(col(A))
  jsExpr = jsExpr.replace(
    /\b(norm|zscore|diff|cumsum)\s*\(\s*col\s*\(\s*["']?([a-zA-Z0-9_\s]+)["']?\s*\)\s*\)/gi,
    (_, fnName, colRef) => {
      const cIdx = resolveColIdx(colRef);
      if (cIdx < 0) return 'NaN';
      const fn = fnName.toLowerCase();
      return `__vecFn("${fn}", ${cIdx}, __r)`;
    }
  );

  // Replace col(X)[rOffset] or col(X) with __col(cIdx, __r)
  jsExpr = jsExpr.replace(/col\s*\(\s*["']?([a-zA-Z0-9_\s]+)["']?\s*\)/gi, (_, colRef) => {
    const cIdx = resolveColIdx(colRef);
    return `__col(${cIdx}, __r)`;
  });

  // Replace ^ with ** for exponentiation
  jsExpr = jsExpr.replace(/\^/g, '**');

  // Replace scientific functions & constants
  jsExpr = jsExpr
    .replace(/\bsinh\b/gi, 'Math.sinh')
    .replace(/\bcosh\b/gi, 'Math.cosh')
    .replace(/\btanh\b/gi, 'Math.tanh')
    .replace(/\basin\b/gi, 'Math.asin')
    .replace(/\bacos\b/gi, 'Math.acos')
    .replace(/\batan2\b/gi, 'Math.atan2')
    .replace(/\batan\b/gi, 'Math.atan')
    .replace(/\bsin\b/gi, 'Math.sin')
    .replace(/\bcos\b/gi, 'Math.cos')
    .replace(/\btan\b/gi, 'Math.tan')
    .replace(/\bexp\b/gi, 'Math.exp')
    .replace(/\blog10\b/gi, 'Math.log10')
    .replace(/\blog2\b/gi, 'Math.log2')
    .replace(/\bln\b/gi, 'Math.log')
    .replace(/\blog\b/gi, 'Math.log')
    .replace(/\bsqrt\b/gi, 'Math.sqrt')
    .replace(/\bcbrt\b/gi, 'Math.cbrt')
    .replace(/\babs\b/gi, 'Math.abs')
    .replace(/\bpow\b/gi, 'Math.pow')
    .replace(/\bfloor\b/gi, 'Math.floor')
    .replace(/\bceil\b/gi, 'Math.ceil')
    .replace(/\bround\b/gi, 'Math.round')
    .replace(/\bsign\b/gi, 'Math.sign')
    .replace(/\bhypot\b/gi, 'Math.hypot')
    .replace(/\bpi\b/gi, 'Math.PI');

  const vecFn = (fn: string, cIdx: number, r: number): number => {
    const v = getColVal(cIdx, r);
    if (fn === 'norm') {
      const st = getColStats(cIdx);
      const span = st.max - st.min;
      return span > 0 ? (v - st.min) / span : 0;
    }
    if (fn === 'zscore') {
      const st = getColStats(cIdx);
      return st.sd > 0 ? (v - st.mean) / st.sd : 0;
    }
    if (fn === 'diff') {
      if (!diffCache.has(cIdx)) {
        const arr = colValuesArray[cIdx];
        const d: (number | null)[] = new Array(rowCount).fill(null);
        for (let k = 0; k < rowCount; k++) {
          if (k === 0) {
            const v0 = getColVal(cIdx, 0);
            const v1 = getColVal(cIdx, 1);
            d[0] = Number.isFinite(v0) && Number.isFinite(v1) ? v1 - v0 : null;
          } else {
            const vp = getColVal(cIdx, k - 1);
            const vc = getColVal(cIdx, k);
            d[k] = Number.isFinite(vp) && Number.isFinite(vc) ? vc - vp : null;
          }
        }
        diffCache.set(cIdx, d);
        void arr;
      }
      const dVal = diffCache.get(cIdx)![r];
      return dVal !== null ? dVal : Number.NaN;
    }
    if (fn === 'cumsum') {
      if (!cumsumCache.has(cIdx)) {
        const cs: (number | null)[] = new Array(rowCount).fill(null);
        let acc = 0;
        for (let k = 0; k < rowCount; k++) {
          const vc = getColVal(cIdx, k);
          if (Number.isFinite(vc)) {
            acc += vc;
            cs[k] = acc;
          } else {
            cs[k] = null;
          }
        }
        cumsumCache.set(cIdx, cs);
      }
      const cVal = cumsumCache.get(cIdx)![r];
      return cVal !== null ? cVal : Number.NaN;
    }
    return v;
  };

  // Elementary error propagation helpers usable directly inside F(x) formulas!
  const err_add = (_a: number, da: number, _b: number, db: number) =>
    Math.hypot(Number.isFinite(da) ? da : 0, Number.isFinite(db) ? db : 0);
  const err_sub = (_a: number, da: number, _b: number, db: number) =>
    Math.hypot(Number.isFinite(da) ? da : 0, Number.isFinite(db) ? db : 0);
  const err_mul = (a: number, da: number, b: number, db: number) => {
    if (!Number.isFinite(a) || !Number.isFinite(b)) return Number.NaN;
    const z = a * b;
    const relA = Math.abs(a) > 1e-15 && Number.isFinite(da) ? da / a : 0;
    const relB = Math.abs(b) > 1e-15 && Number.isFinite(db) ? db / b : 0;
    return Math.abs(z) * Math.hypot(relA, relB);
  };
  const err_div = (a: number, da: number, b: number, db: number) => {
    if (!Number.isFinite(a) || !Number.isFinite(b) || Math.abs(b) < 1e-15) return Number.NaN;
    const z = a / b;
    const relA = Math.abs(a) > 1e-15 && Number.isFinite(da) ? da / a : 0;
    const relB = Number.isFinite(db) ? db / b : 0;
    return Math.abs(z) * Math.hypot(relA, relB);
  };
  const err_pow = (a: number, da: number, n: number) => {
    if (!Number.isFinite(a) || !Number.isFinite(da) || !Number.isFinite(n)) return Number.NaN;
    return Math.abs(n * Math.pow(a, n - 1) * da);
  };
  const deg2rad = (deg: number) => (deg * Math.PI) / 180;
  const rad2deg = (rad: number) => (rad * 180) / Math.PI;

  try {
    // Compile the vectorized expression ONCE outside the loop for maximum performance on 10^6 rows
    // eslint-disable-next-line no-new-func
    const compiledFn = new Function(
      '__r',
      'i',
      'row',
      '__col',
      '__vecFn',
      'err_add',
      'err_sub',
      'err_mul',
      'err_div',
      'err_pow',
      'deg2rad',
      'rad2deg',
      `return (${jsExpr});`
    ) as (
      r: number,
      i: number,
      row: number,
      colFn: (cIdx: number, r: number) => number,
      vFn: (fn: string, cIdx: number, r: number) => number,
      eAdd: typeof err_add,
      eSub: typeof err_sub,
      eMul: typeof err_mul,
      eDiv: typeof err_div,
      ePow: typeof err_pow,
      d2r: typeof deg2rad,
      r2d: typeof rad2deg
    ) => unknown;

    for (let r = 0; r < rowCount; r++) {
      try {
        const calc = compiledFn(
          r,
          r + 1,
          r + 1,
          getColVal,
          vecFn,
          err_add,
          err_sub,
          err_mul,
          err_div,
          err_pow,
          deg2rad,
          rad2deg
        );
        if (typeof calc === 'number' && Number.isFinite(calc)) {
          result[r] = Number(calc.toPrecision(12));
        } else {
          result[r] = null;
        }
      } catch {
        result[r] = null;
      }
    }
  } catch {
    return targetCol.values;
  }

  return result;
}

/**
 * Computes Descriptive Statistics on Columns
 */
export function computeColumnStatistics(col: DataColumn): ColumnStatistics {
  const masked = new Set(col.maskedIndices || []);
  const valid = col.values.filter(
    (v, idx): v is number => !masked.has(idx) && v !== null && Number.isFinite(v)
  );
  const n = valid.length;
  const missing = col.values.length - n;

  if (n === 0) {
    return {
      colShortName: col.shortName,
      colLongName: col.longName || col.shortName,
      n: 0,
      missing,
      mean: 0,
      sd: 0,
      sem: 0,
      sum: 0,
      min: 0,
      median: 0,
      max: 0,
      variance: 0,
    };
  }

  const sum = valid.reduce((a, b) => a + b, 0);
  const mean = sum / n;
  let ss = 0;
  for (const v of valid) {
    ss += (v - mean) * (v - mean);
  }
  const variance = n > 1 ? ss / (n - 1) : 0;
  const sd = Math.sqrt(variance);
  const sem = sd / Math.sqrt(n);

  const sorted = [...valid].sort((a, b) => a - b);
  const min = sorted[0];
  const max = sorted[sorted.length - 1];
  const median =
    n % 2 === 1 ? sorted[Math.floor(n / 2)] : (sorted[n / 2 - 1] + sorted[n / 2]) / 2;

  return {
    colShortName: col.shortName,
    colLongName: col.longName || col.shortName,
    n,
    missing,
    mean,
    sd,
    sem,
    sum,
    min,
    median,
    max,
    variance,
  };
}

/**
 * Comprehensive Signal Processing & Filtering Suite (OriginLab Smooth & Filter)
 * Supports:
 * - 'savitzky_golay': Local polynomial least-squares smoothing (preserves peaks!)
 * - 'moving_average': Uniform boxcar smoothing
 * - 'gaussian': Gaussian kernel weighted smoothing
 * - 'median': Non-linear median filter (removes impulse spikes / bad points automatically!)
 * - 'lowpass': Zero-phase forward-backward RC Butterworth-like low-pass filter
 * - 'highpass': High-pass residual component (removes baseline drift)
 * - 'outlier_zscore': Automatic outlier rejection & interpolation where local residual > threshold * sigma
 */
export function applySignalFilterAndSmooth(params: {
  values: (number | null)[];
  method: SmoothingFilterMethod;
  windowSize?: number;
  polyOrder?: number;
  cutoffAlpha?: number;
  zScoreThreshold?: number;
}): {
  filtered: (number | null)[];
  outliersReplacedCount: number;
} {
  const {
    values,
    method,
    windowSize = 5,
    polyOrder = 2,
    cutoffAlpha = 0.25,
    zScoreThreshold = 2.2,
  } = params;

  const n = values.length;
  const half = Math.max(1, Math.floor(windowSize / 2));
  let outliersReplacedCount = 0;

  if (method === 'median') {
    const out: (number | null)[] = [];
    for (let i = 0; i < n; i++) {
      if (values[i] === null) {
        out.push(null);
        continue;
      }
      const neighborhood: number[] = [];
      for (let j = -half; j <= half; j++) {
        const idx = i + j;
        if (idx >= 0 && idx < n && values[idx] !== null && Number.isFinite(values[idx])) {
          neighborhood.push(values[idx] as number);
        }
      }
      neighborhood.sort((a, b) => a - b);
      const med =
        neighborhood.length > 0
          ? neighborhood[Math.floor(neighborhood.length / 2)]
          : (values[i] as number);
      out.push(Number(med.toFixed(6)));
    }
    return { filtered: out, outliersReplacedCount };
  }

  if (method === 'outlier_zscore') {
    // Compute local median & MAD (Median Absolute Deviation) to detect spikes robustly
    const out: (number | null)[] = [...values];
    const medWindow = Math.max(3, half);

    for (let i = 0; i < n; i++) {
      const val = values[i];
      if (val === null || !Number.isFinite(val)) continue;

      const neigh: number[] = [];
      for (let j = -medWindow; j <= medWindow; j++) {
        const idx = i + j;
        if (idx >= 0 && idx < n && values[idx] !== null && Number.isFinite(values[idx])) {
          neigh.push(values[idx] as number);
        }
      }
      if (neigh.length < 3) continue;
      const sorted = [...neigh].sort((a, b) => a - b);
      const localMed = sorted[Math.floor(sorted.length / 2)];
      const absDevs = sorted.map((v) => Math.abs(v - localMed)).sort((a, b) => a - b);
      const mad = absDevs[Math.floor(absDevs.length / 2)] * 1.4826 || 1e-5;

      if (Math.abs(val - localMed) > zScoreThreshold * mad) {
        out[i] = Number(localMed.toFixed(6));
        outliersReplacedCount++;
      }
    }
    return { filtered: out, outliersReplacedCount };
  }

  if (method === 'savitzky_golay') {
    // Local polynomial regression of degree `polyOrder` in [-half, +half]
    const deg = Math.max(1, Math.min(4, polyOrder, windowSize - 1));
    const out: (number | null)[] = [];

    for (let i = 0; i < n; i++) {
      if (values[i] === null) {
        out.push(null);
        continue;
      }
      const localPts: { x: number; y: number }[] = [];
      for (let j = -half; j <= half; j++) {
        const idx = i + j;
        if (idx >= 0 && idx < n && values[idx] !== null && Number.isFinite(values[idx])) {
          localPts.push({ x: j, y: values[idx] as number });
        }
      }
      if (localPts.length <= deg) {
        out.push(values[i]);
        continue;
      }
      const size = deg + 1;
      const XtX: number[][] = Array.from({ length: size }, () => new Array(size).fill(0));
      const Xty: number[] = new Array(size).fill(0);
      for (const p of localPts) {
        for (let r = 0; r < size; r++) {
          for (let c = 0; c < size; c++) {
            XtX[r][c] += Math.pow(p.x, r + c);
          }
          Xty[r] += Math.pow(p.x, r) * p.y;
        }
      }
      const coeffs = solveLinearSystem(XtX, Xty);
      // Evaluated at local x = 0 is simply coeffs[0]
      out.push(Number(coeffs[0].toFixed(6)));
    }
    return { filtered: out, outliersReplacedCount };
  }

  if (method === 'lowpass' || method === 'highpass') {
    // Zero-phase forward-backward exponential smoothing filter (filtfilt)
    const alpha = Math.max(0.02, Math.min(0.95, cutoffAlpha));
    const validIndices: number[] = [];
    const validVals: number[] = [];
    for (let i = 0; i < n; i++) {
      if (values[i] !== null && Number.isFinite(values[i])) {
        validIndices.push(i);
        validVals.push(values[i] as number);
      }
    }
    if (validVals.length < 2) return { filtered: [...values], outliersReplacedCount: 0 };

    const fwd = new Array(validVals.length);
    fwd[0] = validVals[0];
    for (let i = 1; i < validVals.length; i++) {
      fwd[i] = alpha * validVals[i] + (1 - alpha) * fwd[i - 1];
    }

    const bwd = new Array(validVals.length);
    bwd[validVals.length - 1] = fwd[validVals.length - 1];
    for (let i = validVals.length - 2; i >= 0; i--) {
      bwd[i] = alpha * fwd[i] + (1 - alpha) * bwd[i + 1];
    }

    const out: (number | null)[] = new Array(n).fill(null);
    for (let k = 0; k < validIndices.length; k++) {
      const idx = validIndices[k];
      const lp = bwd[k];
      out[idx] =
        method === 'lowpass'
          ? Number(lp.toFixed(6))
          : Number((validVals[k] - lp).toFixed(6));
    }
    return { filtered: out, outliersReplacedCount: 0 };
  }

  // Default: 'moving_average' or 'gaussian'
  const out: (number | null)[] = [];
  for (let i = 0; i < n; i++) {
    if (values[i] === null) {
      out.push(null);
      continue;
    }
    let sum = 0;
    let weightSum = 0;
    for (let j = -half; j <= half; j++) {
      const idx = i + j;
      if (idx >= 0 && idx < n && values[idx] !== null && Number.isFinite(values[idx])) {
        let w = 1;
        if (method === 'gaussian') {
          const sigma = half / 2 || 1;
          w = Math.exp(-(j * j) / (2 * sigma * sigma));
        }
        sum += (values[idx] as number) * w;
        weightSum += w;
      }
    }
    out.push(weightSum > 0 ? Number((sum / weightSum).toFixed(6)) : values[i]);
  }
  return { filtered: out, outliersReplacedCount };
}

export function smoothSignal(
  values: (number | null)[],
  windowSize = 5,
  method: 'moving_average' | 'gaussian' = 'moving_average'
): (number | null)[] {
  return applySignalFilterAndSmooth({ values, method, windowSize }).filtered;
}

/**
 * Numerical Differentiation: dy/dx
 */
export function differentiateSignal(
  xVals: (number | null)[],
  yVals: (number | null)[]
): (number | null)[] {
  const len = Math.min(xVals.length, yVals.length);
  const deriv: (number | null)[] = new Array(len).fill(null);

  for (let i = 0; i < len; i++) {
    if (i === 0) {
      if (xVals[0] !== null && xVals[1] !== null && yVals[0] !== null && yVals[1] !== null) {
        const dx = (xVals[1] as number) - (xVals[0] as number);
        deriv[0] = dx !== 0 ? ((yVals[1] as number) - (yVals[0] as number)) / dx : 0;
      }
    } else if (i === len - 1) {
      if (
        xVals[len - 2] !== null &&
        xVals[len - 1] !== null &&
        yVals[len - 2] !== null &&
        yVals[len - 1] !== null
      ) {
        const dx = (xVals[len - 1] as number) - (xVals[len - 2] as number);
        deriv[len - 1] =
          dx !== 0 ? ((yVals[len - 1] as number) - (yVals[len - 2] as number)) / dx : 0;
      }
    } else {
      if (
        xVals[i - 1] !== null &&
        xVals[i + 1] !== null &&
        yVals[i - 1] !== null &&
        yVals[i + 1] !== null
      ) {
        const dx = (xVals[i + 1] as number) - (xVals[i - 1] as number);
        deriv[i] = dx !== 0 ? ((yVals[i + 1] as number) - (yVals[i - 1] as number)) / dx : 0;
      }
    }
  }
  return deriv;
}

/**
 * Numerical Integration: Cumulative Trapezoidal Rule
 */
export function integrateSignal(
  xVals: (number | null)[],
  yVals: (number | null)[]
): (number | null)[] {
  const len = Math.min(xVals.length, yVals.length);
  const integ: (number | null)[] = new Array(len).fill(null);
  let total = 0;

  for (let i = 0; i < len; i++) {
    if (i === 0) {
      integ[0] = 0;
    } else {
      const x0 = xVals[i - 1],
        x1 = xVals[i];
      const y0 = yVals[i - 1],
        y1 = yVals[i];
      if (x0 !== null && x1 !== null && y0 !== null && y1 !== null) {
        const dx = x1 - x0;
        const avgY = (y0 + y1) / 2;
        total += dx * avgY;
        integ[i] = Number(total.toFixed(6));
      } else {
        integ[i] = total;
      }
    }
  }
  return integ;
}

/**
 * Intelligent TXT / CSV Parser for OriginLab-style laboratory files
 */
export function parseTabularText(
  rawText: string,
  sheetName = 'Dati_Importati'
): WorkbookSheet {
  const lines = rawText
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l.length > 0 && !l.startsWith('#') && !l.startsWith('//') && !l.startsWith('%'));

  if (lines.length === 0) {
    return createEmptySheet(sheetName);
  }

  const sampleLines = lines.slice(0, Math.min(10, lines.length));
  const tabCount = sampleLines.filter((l) => l.includes('\t')).length;
  const semiCount = sampleLines.filter((l) => l.includes(';')).length;
  const commaCount = sampleLines.filter((l) => l.includes(',')).length;

  let splitLine: (line: string) => string[];
  let usesCommaDecimal = false;

  if (tabCount >= sampleLines.length / 2) {
    splitLine = (l) => l.split('\t').map((s) => s.trim());
    usesCommaDecimal = true;
  } else if (semiCount >= sampleLines.length / 2) {
    splitLine = (l) => l.split(';').map((s) => s.trim());
    usesCommaDecimal = true;
  } else if (commaCount >= sampleLines.length / 2) {
    const hasSpaceAndCommaNumber = sampleLines.some((l) => /\d,\d+\s+\d/.test(l));
    if (hasSpaceAndCommaNumber) {
      splitLine = (l) => l.split(/\s+/).map((s) => s.trim());
      usesCommaDecimal = true;
    } else {
      splitLine = (l) => l.split(',').map((s) => s.trim());
      usesCommaDecimal = false;
    }
  } else {
    splitLine = (l) => l.split(/\s+/).map((s) => s.trim());
    usesCommaDecimal = true;
  }

  const parseNumberToken = (token: string): number | null => {
    if (!token || token === '--' || token === 'NaN' || token === 'null') return null;
    const cleaned = usesCommaDecimal ? token.replace(',', '.') : token;
    const num = Number(cleaned);
    return Number.isFinite(num) ? num : null;
  };

  const tokenizedRows = lines.map(splitLine);
  const maxCols = Math.max(...tokenizedRows.map((r) => r.length));

  const firstRow = tokenizedRows[0];
  const firstRowNumericCount = firstRow.filter((t) => parseNumberToken(t) !== null).length;
  const hasHeaderRow = firstRowNumericCount < firstRow.length / 2;

  let hasUnitRow = false;
  if (hasHeaderRow && tokenizedRows.length > 2) {
    const secondRow = tokenizedRows[1];
    const secondRowNumericCount = secondRow.filter((t) => parseNumberToken(t) !== null).length;
    if (secondRowNumericCount < secondRow.length / 2) {
      hasUnitRow = true;
    }
  }

  const dataStartIndex = hasHeaderRow ? (hasUnitRow ? 2 : 1) : 0;
  const dataRows = tokenizedRows.slice(dataStartIndex);

  const columns: DataColumn[] = [];
  for (let colIdx = 0; colIdx < maxCols; colIdx++) {
    const shortName = getColumnShortName(colIdx);
    let rawHeader = hasHeaderRow && firstRow[colIdx] ? firstRow[colIdx] : `Colonna ${shortName}`;
    let unit =
      hasUnitRow && tokenizedRows[1][colIdx]
        ? tokenizedRows[1][colIdx].replace(/^[\[\(]|[\]\)]$/g, '')
        : '';

    const unitMatch = rawHeader.match(/^(.+?)\s*[\(\[](.+?)[\)\]]$/);
    if (unitMatch && !unit) {
      rawHeader = unitMatch[1].trim();
      unit = unitMatch[2].trim();
    }

    const values: (number | null)[] = dataRows.map((row) =>
      colIdx < row.length ? parseNumberToken(row[colIdx]) : null
    );

    let role: DataColumn['role'] = colIdx === 0 ? 'X' : 'Y';
    if (/err|error|sigma|dev/i.test(rawHeader) || /err|error|sigma/i.test(unit)) {
      role = 'Y_ERR';
    }

    columns.push({
      id: `col_${Date.now()}_${colIdx}_${Math.random().toString(36).slice(2, 6)}`,
      shortName,
      longName: rawHeader,
      unit,
      role,
      values,
    });
  }

  return {
    id: `sheet_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    name: sheetName,
    columns,
  };
}

export function serializeSheetToText(sheet: WorkbookSheet, format: 'txt' | 'csv'): string {
  const sep = format === 'txt' ? '\t' : ',';
  const headers = sheet.columns.map((c) => {
    const roleTag = `[${c.role}]`;
    const unitTag = c.unit ? ` (${c.unit})` : '';
    const text = `${c.longName || c.shortName}${unitTag} ${roleTag}`;
    return format === 'csv' && text.includes(',') ? `"${text}"` : text;
  });

  const maxRows = Math.max(0, ...sheet.columns.map((c) => c.values.length));
  const lines: string[] = [headers.join(sep)];

  for (let r = 0; r < maxRows; r++) {
    const rowTokens = sheet.columns.map((c) => {
      const val = r < c.values.length ? c.values[r] : null;
      return val !== null && Number.isFinite(val) ? String(val) : '';
    });
    if (rowTokens.some((t) => t !== '')) {
      lines.push(rowTokens.join(sep));
    }
  }

  return lines.join('\n');
}

export function createEmptySheet(name = 'Foglio1', numRows = 30): WorkbookSheet {
  const colXId = `col_x_${Date.now()}`;
  const colY1Id = `col_y1_${Date.now()}`;
  const colY2Id = `col_y2_${Date.now()}`;
  return {
    id: `sheet_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    name,
    columns: [
      {
        id: colXId,
        shortName: 'A',
        longName: 'Asse X',
        unit: 'u.a.',
        role: 'X',
        values: new Array(numRows).fill(null),
      },
      {
        id: colY1Id,
        shortName: 'B',
        longName: 'Segnale Y1',
        unit: 'mV',
        role: 'Y',
        linkedXColId: colXId,
        values: new Array(numRows).fill(null),
      },
      {
        id: colY2Id,
        shortName: 'C',
        longName: 'Segnale Y2',
        unit: 'mV',
        role: 'Y',
        linkedXColId: colXId,
        values: new Array(numRows).fill(null),
      },
    ],
  };
}

/**
 * Creates a clean initial workspace with non-overlapping Legend & Fit Report boxes
 */
export function createInitialWorkspace(): {
  sheets: WorkbookSheet[];
  graphs: GraphWindow[];
} {
  const numExperimental = 100;
  const aStart = 0;
  const bEnd = 10;

  const xValues: (number | null)[] = [];
  const y1Values: (number | null)[] = [];
  const y2Values: (number | null)[] = [];
  const yErrValues: (number | null)[] = [];
  const yExpValues: (number | null)[] = [];

  for (let i = 0; i < numExperimental; i++) {
    const x = Number((aStart + (i * (bEnd - aStart)) / (numExperimental - 1)).toFixed(4));
    const noise1 = Math.sin(i * 1.73) * 0.28 + Math.cos(i * 3.11) * 0.15;
    const noise2 = Math.cos(i * 2.07) * 0.08 + Math.sin(i * 1.37) * 0.06;

    let y1 = Number((1.8 + 2.45 * x + noise1).toFixed(4));

    // 4 deliberate experimental outliers ("Bad Data Points")
    if (i === 18) y1 = 19.85;
    if (i === 44) y1 = 2.15;
    if (i === 71) y1 = 28.4;
    if (i === 86) y1 = 9.2;

    const diff = x - 5.2;
    const y2 = Number(
      (0.8 + 14.5 * Math.exp(-(diff * diff) / (2 * 1.2 * 1.2)) + noise2).toFixed(4)
    );

    const yErr = Number((0.4 + 0.12 * Math.sqrt(x) + Math.abs(noise1) * 0.3).toFixed(3));
    const yExp = Number((2.0 + 18.0 * Math.exp(-0.45 * x) + noise1 * 0.2).toFixed(4));

    xValues.push(x);
    y1Values.push(y1);
    y2Values.push(y2);
    yErrValues.push(yErr);
    yExpValues.push(yExp);
  }

  const colXId = 'col_exp_x';
  const colY1Id = 'col_exp_y1';
  const colY2Id = 'col_exp_y2';
  const colYErrId = 'col_exp_yerr';
  const colYExpId = 'col_exp_yexp';

  const expSheet: WorkbookSheet = {
    id: 'sheet_exp_1',
    name: 'Book1_Dati',
    columns: [
      {
        id: colXId,
        shortName: 'A',
        longName: 'Tensione Applicata',
        unit: 'V',
        comment: 'Potenziale sorgente',
        role: 'X',
        values: xValues,
      },
      {
        id: colY1Id,
        shortName: 'B',
        longName: 'Corrente Canale 1',
        unit: 'mA',
        comment: 'Con 4 outlier da filtrare',
        role: 'Y',
        linkedXColId: colXId,
        linkedErrColId: colYErrId,
        values: y1Values,
      },
      {
        id: colY2Id,
        shortName: 'C',
        longName: 'Picco Risonanza Ottica',
        unit: 'u.a.',
        comment: 'Profilo Gaussiano a 5.2 V',
        role: 'Y',
        linkedXColId: colXId,
        values: y2Values,
      },
      {
        id: colYErrId,
        shortName: 'D',
        longName: 'Errore Canale 1 (±σ)',
        unit: 'mA',
        comment: 'Incertezza strumentale',
        role: 'Y_ERR',
        linkedXColId: colXId,
        values: yErrValues,
      },
      {
        id: colYExpId,
        shortName: 'E',
        longName: 'Decadimento Transiente',
        unit: 'mV',
        comment: 'Curva esponenziale',
        role: 'Y',
        linkedXColId: colXId,
        values: yExpValues,
      },
    ],
  };

  const validPts1: { x: number; y: number }[] = [];
  for (let i = 0; i < numExperimental; i++) {
    if (xValues[i] !== null && y1Values[i] !== null) {
      validPts1.push({ x: xValues[i]!, y: y1Values[i]! });
    }
  }

  const initialFitCalc = computeCurveFit({
    points: validPts1,
    model: 'linear',
    polyDegree: 1,
    rangeC: 0,
    rangeD: 12,
    numPointsN: 50,
  });

  const fitId = 'fit_initial_1';
  const fitSheetId = 'sheet_fit_1';

  const fitSheet: WorkbookSheet = {
    id: fitSheetId,
    name: 'Book2_Fit50pt',
    isFitOutput: true,
    sourceFitId: fitId,
    columns: [
      {
        id: 'col_fit_x',
        shortName: 'A',
        longName: 'X Fittata [0 → 12]',
        unit: 'V',
        role: 'X',
        values: initialFitCalc.generatedPoints.map((p) => p.x),
      },
      {
        id: 'col_fit_y',
        shortName: 'B',
        longName: 'Y Fit Lineare (n=50)',
        unit: 'mA',
        role: 'Y',
        linkedXColId: 'col_fit_x',
        values: initialFitCalc.generatedPoints.map((p) => p.y),
      },
    ],
  };

  const initialFit: FittedCurve = {
    id: fitId,
    name: 'Fit Lineare (Corrente Ch1)',
    sourceSeriesId: 'series_1',
    sheetId: expSheet.id,
    xColId: colXId,
    yColId: colY1Id,
    model: 'linear',
    polyDegree: 1,
    rangeC: 0,
    rangeD: 12,
    numPointsN: 50,
    autoUpdateOnDataChange: true,
    outputSheetId: fitSheetId,
    equationText: initialFitCalc.equationText,
    coefficients: initialFitCalc.coefficients,
    rSquared: initialFitCalc.rSquared,
    rmse: initialFitCalc.rmse,
    sourceDomainMin: initialFitCalc.sourceDomainMin,
    sourceDomainMax: initialFitCalc.sourceDomainMax,
    sourceValidPointsCount: initialFitCalc.sourceValidPointsCount,
    generatedPoints: initialFitCalc.generatedPoints,
    color: '#DC2626',
    lineWidth: 2.2,
    lineStyle: 'solid',
    visible: true,
  };

  const mainGraph: GraphWindow = {
    id: 'graph_1',
    name: 'Graph1_Calibrazione',
    title: 'Caratteristica Tensione-Corrente & Risonanza',
    subtitle: 'Trascina Legenda o Riquadro Fit con il mouse per riposizionarli liberamente',
    xAxis: {
      label: 'Tensione Applicata (V)',
      autoLabel: true,
      scaleType: 'linear',
      min: 'auto',
      max: 'auto',
      showMajorGrid: true,
      showMinorGrid: true,
      tickCount: 7,
      tickDirection: 'both',
      labelFormat: 'auto',
    },
    yAxis: {
      label: 'Corrente Ch1 / Picco Risonanza (mA, u.a.)',
      autoLabel: true,
      scaleType: 'linear',
      min: 'auto',
      max: 'auto',
      showMajorGrid: true,
      showMinorGrid: true,
      tickCount: 6,
      tickDirection: 'both',
      labelFormat: 'auto',
    },
    showLegend: true,
    legendPosition: 'top-left',
    legendOffset: { x: 98, y: 76 },
    showFitReportBox: true,
    fitBoxOffset: { x: 545, y: 392 },
    annotations: [],
    series: [
      {
        id: 'series_1',
        sheetId: expSheet.id,
        xColId: colXId,
        yColId: colY1Id,
        yErrColId: colYErrId,
        label: 'Corrente Canale 1 (mA)',
        autoLabel: true,
        plotType: 'line_symbol',
        color: '#2563EB',
        lineWidth: 1.5,
        lineStyle: 'solid',
        connection: 'linear',
        markerShape: 'circle',
        markerSize: 4,
        symbolFill: '#2563EB',
        symbolBorder: '#ffffff',
        symbolBorderWidth: 1,
        fillArea: false,
        fillColor: '#2563EB',
        fillOpacity: 0.15,
        dropLines: false,
        showErrorBars: false,
        errorBarCap: 4,
        errorBarColor: '#93c5fd',
        visible: true,
      },
      {
        id: 'series_2',
        sheetId: expSheet.id,
        xColId: colXId,
        yColId: colY2Id,
        label: 'Picco Risonanza Ottica (u.a.)',
        autoLabel: true,
        plotType: 'area',
        color: '#0D9488',
        lineWidth: 2,
        lineStyle: 'solid',
        connection: 'spline',
        markerShape: 'none',
        markerSize: 4,
        symbolFill: '#0D9488',
        symbolBorder: '#ffffff',
        symbolBorderWidth: 1,
        fillArea: true,
        fillColor: '#0D9488',
        fillOpacity: 0.16,
        dropLines: false,
        showErrorBars: false,
        errorBarCap: 4,
        errorBarColor: '#0D9488',
        visible: true,
      },
    ],
    fits: [initialFit],
  };

  return {
    sheets: [expSheet, fitSheet],
    graphs: [mainGraph],
  };
}

/**
 * Parses a date/time string (ISO 2026-10-20, IT 20/10/2026, or time 14:30:05) into numeric seconds
 */
export function parseDateTimeTokenToNumeric(raw: string): number | null {
  const s = raw.trim();
  if (!s) return null;

  // Pure time HH:MM or HH:MM:SS(.ms)
  const timeMatch = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?$/);
  if (timeMatch) {
    const hh = Number(timeMatch[1]);
    const mm = Number(timeMatch[2]);
    const ss = timeMatch[3] ? Number(timeMatch[3]) : 0;
    return hh * 3600 + mm * 60 + ss;
  }

  // Italian/European date DD/MM/YYYY [HH:MM[:SS]]
  const euroMatch = s.match(
    /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (euroMatch) {
    const day = Number(euroMatch[1]);
    const month = Number(euroMatch[2]) - 1;
    const year = Number(euroMatch[3]);
    const hh = euroMatch[4] ? Number(euroMatch[4]) : 0;
    const mm = euroMatch[5] ? Number(euroMatch[5]) : 0;
    const ss = euroMatch[6] ? Number(euroMatch[6]) : 0;
    const ms = Date.UTC(year, month, day, hh, mm, ss);
    return Number.isFinite(ms) ? ms / 1000 : null;
  }

  // ISO date/time
  if (/\d{4}-\d{2}-\d{2}/.test(s)) {
    const parsed = Date.parse(s);
    if (Number.isFinite(parsed)) return parsed / 1000;
  }

  return null;
}

/**
 * Automatically detects the ColumnDataType ('numeric' | 'integer' | 'text' | 'datetime')
 */
export function detectColumnDataType(rawTokens: (string | number | null | undefined)[]): ColumnDataType {
  const nonEmpty = rawTokens
    .map((t) => (t === null || t === undefined ? '' : String(t).trim()))
    .filter((s) => s !== '' && s !== '--' && s.toLowerCase() !== 'nan');

  if (nonEmpty.length === 0) return 'numeric';

  let numCount = 0;
  let intCount = 0;
  let dateCount = 0;

  for (const tok of nonEmpty) {
    const cleanedNum = Number(tok.replace(',', '.'));
    if (Number.isFinite(cleanedNum) && !tok.includes('/') && !tok.includes(':')) {
      numCount++;
      if (Number.isInteger(cleanedNum) && !tok.includes('.') && !tok.includes(',')) {
        intCount++;
      }
    } else if (parseDateTimeTokenToNumeric(tok) !== null) {
      dateCount++;
    }
  }

  if (dateCount >= nonEmpty.length / 2) return 'datetime';
  if (numCount >= nonEmpty.length / 2) {
    return intCount === numCount ? 'integer' : 'numeric';
  }
  return 'text';
}

/**
 * Parses user input in a cell according to the column's dataType and role.
 * Missing or invalid values return numVal = null (NaN-safe).
 */
export function parseCellInputForColumn(
  rawInput: string,
  dataType: ColumnDataType = 'numeric',
  role: DataColumn['role'] = 'Y'
): { numVal: number | null; strVal: string | null } {
  const trimmed = rawInput.trim();
  if (trimmed === '' || trimmed === '--' || trimmed.toLowerCase() === 'nan') {
    return { numVal: null, strVal: null };
  }

  if (dataType === 'text' || role === 'LABEL') {
    const maybeNum = Number(trimmed.replace(',', '.'));
    return {
      numVal: Number.isFinite(maybeNum) ? maybeNum : null,
      strVal: trimmed,
    };
  }

  if (dataType === 'datetime') {
    const dtSec = parseDateTimeTokenToNumeric(trimmed);
    if (dtSec !== null) {
      return { numVal: dtSec, strVal: trimmed };
    }
    const directNum = Number(trimmed.replace(',', '.'));
    return {
      numVal: Number.isFinite(directNum) ? directNum : null,
      strVal: trimmed,
    };
  }

  const normalized = trimmed
    .replace(',', '.')
    .replace(/\s*[×xX\*]\s*10\^([+\-]?\d+)/, 'e$1');
  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) {
    // Graceful NaN handling: invalid strings in numeric columns become null (missing) without crashing
    return { numVal: null, strVal: null };
  }

  if (dataType === 'integer') {
    const rounded = Math.round(parsed);
    return { numVal: rounded, strVal: String(rounded) };
  }

  return { numVal: parsed, strVal: null };
}

/**
 * Formats a cell for display in WorkbookView
 */
export function formatCellForDisplay(col: DataColumn, rowIndex: number): string {
  const dt = col.dataType || (col.role === 'LABEL' ? 'text' : 'numeric');
  const rawStr = col.rawStrings && rowIndex < col.rawStrings.length ? col.rawStrings[rowIndex] : null;
  const numVal = rowIndex < col.values.length ? col.values[rowIndex] : null;

  if ((dt === 'text' || dt === 'datetime' || col.role === 'LABEL') && rawStr !== null && rawStr !== undefined) {
    return rawStr;
  }

  if (numVal === null || numVal === undefined || !Number.isFinite(numVal)) {
    return '';
  }

  if (dt === 'integer') {
    return String(Math.round(numVal));
  }

  return String(numVal);
}

/**
 * Sorts a SINGLE column in-place (keeping values, rawStrings, and maskedIndices aligned; nulls at the end)
 */
export function sortSingleColumn(
  sheet: WorkbookSheet,
  colId: string,
  direction: 'asc' | 'desc'
): WorkbookSheet {
  const updatedColumns = sheet.columns.map((col) => {
    if (col.id !== colId) return col;

    const len = Math.max(col.values.length, col.rawStrings?.length || 0);
    const maskedSet = new Set(col.maskedIndices || []);

    const items = Array.from({ length: len }, (_, idx) => ({
      num: idx < col.values.length ? col.values[idx] : null,
      str: col.rawStrings && idx < col.rawStrings.length ? col.rawStrings[idx] : null,
      masked: maskedSet.has(idx),
    }));

    items.sort((a, b) => {
      const aEmpty =
        (a.num === null || !Number.isFinite(a.num)) && (!a.str || a.str.trim() === '');
      const bEmpty =
        (b.num === null || !Number.isFinite(b.num)) && (!b.str || b.str.trim() === '');
      if (aEmpty && bEmpty) return 0;
      if (aEmpty) return 1; // Always push empty/NaN cells to the bottom
      if (bEmpty) return -1;

      if (col.dataType === 'text' || col.role === 'LABEL') {
        const sa = a.str ?? String(a.num ?? '');
        const sb = b.str ?? String(b.num ?? '');
        return direction === 'asc' ? sa.localeCompare(sb) : sb.localeCompare(sa);
      }

      const na = a.num ?? 0;
      const nb = b.num ?? 0;
      return direction === 'asc' ? na - nb : nb - na;
    });

    const newMasked: number[] = [];
    items.forEach((it, idx) => {
      if (it.masked) newMasked.push(idx);
    });

    return {
      ...col,
      values: items.map((it) => it.num),
      rawStrings: col.rawStrings ? items.map((it) => it.str) : undefined,
      maskedIndices: newMasked,
    };
  });

  return { ...sheet, columns: updatedColumns };
}

/**
 * Sorts the ENTIRE WorkbookSheet row-by-row according to a Key Column (keeping all columns & masks aligned)
 */
export function sortEntireSheetByColumn(
  sheet: WorkbookSheet,
  keyColId: string,
  direction: 'asc' | 'desc'
): WorkbookSheet {
  const keyCol = sheet.columns.find((c) => c.id === keyColId);
  if (!keyCol) return sheet;

  const maxRows = Math.max(
    0,
    ...sheet.columns.map((c) => Math.max(c.values.length, c.rawStrings?.length || 0))
  );
  if (maxRows <= 1) return sheet;

  const rowOrder = Array.from({ length: maxRows }, (_, r) => r);

  rowOrder.sort((rA, rB) => {
    const numA = rA < keyCol.values.length ? keyCol.values[rA] : null;
    const numB = rB < keyCol.values.length ? keyCol.values[rB] : null;
    const strA =
      keyCol.rawStrings && rA < keyCol.rawStrings.length ? keyCol.rawStrings[rA] : null;
    const strB =
      keyCol.rawStrings && rB < keyCol.rawStrings.length ? keyCol.rawStrings[rB] : null;

    const aEmpty =
      (numA === null || !Number.isFinite(numA)) && (!strA || strA.trim() === '');
    const bEmpty =
      (numB === null || !Number.isFinite(numB)) && (!strB || strB.trim() === '');

    if (aEmpty && bEmpty) return rA - rB;
    if (aEmpty) return 1; // Keep empty/NaN rows at the bottom
    if (bEmpty) return -1;

    if (keyCol.dataType === 'text' || keyCol.role === 'LABEL') {
      const sa = strA ?? String(numA ?? '');
      const sb = strB ?? String(numB ?? '');
      return direction === 'asc' ? sa.localeCompare(sb) : sb.localeCompare(sa);
    }

    const na = numA ?? 0;
    const nb = numB ?? 0;
    return direction === 'asc' ? na - nb : nb - na;
  });

  const sortedColumns = sheet.columns.map((col) => {
    const oldMasked = new Set(col.maskedIndices || []);
    const nextValues: (number | null)[] = [];
    const nextRawStrings: (string | null)[] | undefined = col.rawStrings ? [] : undefined;
    const nextMasked: number[] = [];

    rowOrder.forEach((origRowIdx, newRowIdx) => {
      nextValues.push(origRowIdx < col.values.length ? col.values[origRowIdx] : null);
      if (nextRawStrings && col.rawStrings) {
        nextRawStrings.push(
          origRowIdx < col.rawStrings.length ? col.rawStrings[origRowIdx] : null
        );
      }
      if (oldMasked.has(origRowIdx)) {
        nextMasked.push(newRowIdx);
      }
    });

    return {
      ...col,
      values: nextValues,
      rawStrings: nextRawStrings,
      maskedIndices: nextMasked,
    };
  });

  return { ...sheet, columns: sortedColumns };
}

export type AutoFillMode =
  | 'row_numbers'
  | 'linear_step'
  | 'linspace'
  | 'random_uniform'
  | 'random_normal';

/**
 * Generates sequences for Auto-Fill on a column (Row numbers, Linear step, Linspace, Uniform or Normal Random)
 */
export function applyAutoFillToColumn(
  col: DataColumn,
  params: {
    mode: AutoFillMode;
    rowCount: number;
    startValue?: number;
    stepValue?: number;
    stopValue?: number;
    minVal?: number;
    maxVal?: number;
    mean?: number;
    stdDev?: number;
  }
): DataColumn {
  const n = Math.max(1, Math.min(10000, Math.round(params.rowCount || col.values.length || 30)));
  const start = params.startValue ?? 0;
  const step = params.stepValue ?? 1;
  const stop = params.stopValue ?? 10;
  const minVal = params.minVal ?? 0;
  const maxVal = params.maxVal ?? 1;
  const mean = params.mean ?? 0;
  const stdDev = params.stdDev ?? 1;

  const newValues: number[] = [];

  for (let i = 0; i < n; i++) {
    if (params.mode === 'row_numbers') {
      newValues.push(i + 1);
    } else if (params.mode === 'linear_step') {
      newValues.push(Number((start + i * step).toPrecision(10)));
    } else if (params.mode === 'linspace') {
      const val = n <= 1 ? start : start + (i * (stop - start)) / (n - 1);
      newValues.push(Number(val.toPrecision(10)));
    } else if (params.mode === 'random_uniform') {
      const val = minVal + Math.random() * (maxVal - minVal);
      newValues.push(Number(val.toFixed(5)));
    } else if (params.mode === 'random_normal') {
      // Box-Muller transform for Gaussian random numbers
      const u1 = Math.max(1e-12, Math.random());
      const u2 = Math.random();
      const z0 = Math.sqrt(-2.0 * Math.log(u1)) * Math.cos(2.0 * Math.PI * u2);
      newValues.push(Number((mean + z0 * stdDev).toFixed(5)));
    }
  }

  return {
    ...col,
    dataType: params.mode === 'row_numbers' ? 'integer' : 'numeric',
    values: newValues,
    rawStrings: undefined,
  };
}

export type MaskOperationMode =
  | 'mask_range'
  | 'unmask_range'
  | 'mask_greater'
  | 'mask_less'
  | 'mask_outside'
  | 'invert_mask'
  | 'unmask_all';

/**
 * Applies range or conditional masking to a specific column or all Y columns in a sheet
 */
export function applyMaskingToSheet(
  sheet: WorkbookSheet,
  params: {
    targetColId: string | 'ALL_Y';
    mode: MaskOperationMode;
    startRow1?: number; // 1-indexed inclusive
    endRow1?: number; // 1-indexed inclusive
    threshold?: number;
    minThreshold?: number;
    maxThreshold?: number;
  }
): WorkbookSheet {
  const rStart = Math.max(0, (params.startRow1 ?? 1) - 1);
  const rEnd = Math.max(rStart, (params.endRow1 ?? 10) - 1);
  const thresh = params.threshold ?? 0;
  const minT = params.minThreshold ?? 0;
  const maxT = params.maxThreshold ?? 10;

  const updatedColumns = sheet.columns.map((col) => {
    const isTarget =
      params.targetColId === 'ALL_Y'
        ? col.role === 'Y' || col.role === 'Y_ERR'
        : col.id === params.targetColId;
    if (!isTarget) return col;

    if (params.mode === 'unmask_all') {
      return { ...col, maskedIndices: [] };
    }

    const masked = new Set(col.maskedIndices || []);
    const rowLen = col.values.length;

    if (params.mode === 'mask_range') {
      for (let r = rStart; r <= Math.min(rowLen - 1, rEnd); r++) {
        masked.add(r);
      }
    } else if (params.mode === 'unmask_range') {
      for (let r = rStart; r <= Math.min(rowLen - 1, rEnd); r++) {
        masked.delete(r);
      }
    } else if (params.mode === 'mask_greater') {
      for (let r = 0; r < rowLen; r++) {
        const v = col.values[r];
        if (v !== null && Number.isFinite(v) && v > thresh) masked.add(r);
      }
    } else if (params.mode === 'mask_less') {
      for (let r = 0; r < rowLen; r++) {
        const v = col.values[r];
        if (v !== null && Number.isFinite(v) && v < thresh) masked.add(r);
      }
    } else if (params.mode === 'mask_outside') {
      for (let r = 0; r < rowLen; r++) {
        const v = col.values[r];
        if (v !== null && Number.isFinite(v) && (v < minT || v > maxT)) masked.add(r);
      }
    } else if (params.mode === 'invert_mask') {
      for (let r = 0; r < rowLen; r++) {
        if (masked.has(r)) masked.delete(r);
        else if (col.values[r] !== null) masked.add(r);
      }
    }

    return {
      ...col,
      maskedIndices: Array.from(masked).sort((a, b) => a - b),
    };
  });

  return { ...sheet, columns: updatedColumns };
}

/**
 * Reorders columns in a sheet via Drag & Drop and rebinds Multi-X associations
 */
export function moveColumnInSheet(
  sheet: WorkbookSheet,
  fromColIndex: number,
  toColIndex: number
): WorkbookSheet {
  if (
    fromColIndex === toColIndex ||
    fromColIndex < 0 ||
    toColIndex < 0 ||
    fromColIndex >= sheet.columns.length ||
    toColIndex >= sheet.columns.length
  ) {
    return sheet;
  }

  const nextCols = [...sheet.columns];
  const [moved] = nextCols.splice(fromColIndex, 1);
  nextCols.splice(toColIndex, 0, moved);

  return rebindSheetMultiXAssociations({
    ...sheet,
    columns: nextCols,
  });
}

// ============================================================================
// BATCH 3: NUMERIC FORMATTING, SELECTION STATS, FIND & REPLACE, ROW FILTER,
//          SHEET SPLIT, AUTO-RECALCULATION & ELEMENTARY ERROR PROPAGATION
// ============================================================================

/**
 * Formats a numeric value according to the column's numericFormat, numericPrecision, and numericPrecisionMode:
 * - 'auto': standard clean representation
 * - 'decimal': fixed decimals or significant figures
 * - 'scientific': e.g. `1.23 × 10^-4` or `1.23E-4`
 * - 'engineering': exponent is always a multiple of 3 (e.g. `123.4E-6`, `1.23E+3`)
 */
export function formatNumericValueWithColumnSpec(
  val: number,
  col: Pick<DataColumn, 'dataType' | 'numericFormat' | 'numericPrecision' | 'numericPrecisionMode'>,
  useUnicodeSci = false
): string {
  if (!Number.isFinite(val)) return '';
  if (col.dataType === 'integer') return String(Math.round(val));

  const fmt: NumericFormatMode = col.numericFormat || 'auto';
  const prec = col.numericPrecision !== undefined ? Math.max(0, Math.min(12, col.numericPrecision)) : 4;
  const mode: NumericPrecisionMode = col.numericPrecisionMode || 'decimals';

  if (fmt === 'auto') {
    if (col.numericPrecision !== undefined) {
      return mode === 'sigfigs'
        ? Number(val.toPrecision(Math.max(1, prec))).toString()
        : Number(val.toFixed(prec)).toString();
    }
    return String(val);
  }

  if (fmt === 'decimal') {
    if (mode === 'sigfigs') {
      return Number(val.toPrecision(Math.max(1, prec))).toString();
    }
    return val.toFixed(prec);
  }

  if (fmt === 'scientific') {
    const expDigits = mode === 'sigfigs' ? Math.max(0, prec - 1) : prec;
    const rawExp = val.toExponential(expDigits); // e.g. "1.2300e-4"
    if (!useUnicodeSci) {
      return rawExp.replace('e', 'E');
    }
    const [mantissa, expPart] = rawExp.split('e');
    const expNum = Number(expPart);
    return `${mantissa}×10^${expNum}`;
  }

  if (fmt === 'engineering') {
    if (val === 0) {
      return mode === 'sigfigs' ? '0E+0' : `${(0).toFixed(prec)}E+0`;
    }
    const absV = Math.abs(val);
    const exp10 = Math.floor(Math.log10(absV));
    const engExp = Math.floor(exp10 / 3) * 3;
    const scaled = val / Math.pow(10, engExp);
    const mantissaStr =
      mode === 'sigfigs'
        ? Number(scaled.toPrecision(Math.max(1, prec))).toString()
        : scaled.toFixed(prec);
    const signStr = engExp >= 0 ? `+${engExp}` : `${engExp}`;
    return useUnicodeSci ? `${mantissaStr}×10^${engExp}` : `${mantissaStr}E${signStr}`;
  }

  return String(val);
}

export interface QuickSelectionStats {
  n: number;
  maskedCount: number;
  missingCount: number;
  totalCells: number;
  mean: number;
  sd: number;
  sem: number;
  sum: number;
  min: number;
  max: number;
  selectionLabel: string;
}

/**
 * Computes real-time Quick Statistics on a 2D cell selection (or active columns) directly from Float64Array & Uint8Array buffers
 */
export function computeQuickSelectionStats(
  sheet: WorkbookSheet,
  selection: {
    startColIdx: number;
    endColIdx: number;
    startRowIdx: number;
    endRowIdx: number;
  } | null
): QuickSelectionStats {
  const maxSheetRows = getSheetRowCount(sheet);
  if (sheet.columns.length === 0 || maxSheetRows === 0) {
    return {
      n: 0,
      maskedCount: 0,
      missingCount: 0,
      totalCells: 0,
      mean: 0,
      sd: 0,
      sem: 0,
      sum: 0,
      min: 0,
      max: 0,
      selectionLabel: 'Nessun dato',
    };
  }

  const cMin = selection
    ? Math.max(0, Math.min(selection.startColIdx, selection.endColIdx))
    : 0;
  const cMax = selection
    ? Math.min(sheet.columns.length - 1, Math.max(selection.startColIdx, selection.endColIdx))
    : sheet.columns.length - 1;
  const rMin = selection
    ? Math.max(0, Math.min(selection.startRowIdx, selection.endRowIdx))
    : 0;
  const rMax = selection
    ? Math.min(Math.max(maxSheetRows - 1, 0), Math.max(selection.startRowIdx, selection.endRowIdx))
    : Math.max(0, maxSheetRows - 1);

  let n = 0;
  let maskedCount = 0;
  let missingCount = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;

  // Welford's online algorithm for numerical stability in a single O(k) pass over Float64Array
  let meanW = 0;
  let m2 = 0;

  for (let c = cMin; c <= cMax; c++) {
    const col = sheet.columns[c];
    if (!col) continue;
    const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);

    for (let r = rMin; r <= rMax; r++) {
      if (r >= length) {
        missingCount++;
        continue;
      }
      const status = maskUint8[r];
      if (status === POINT_MASKED) {
        maskedCount++;
        continue;
      }
      const v = float64[r];
      if (status === POINT_NAN || !Number.isFinite(v)) {
        missingCount++;
        continue;
      }

      n++;
      sum += v;
      if (v < min) min = v;
      if (v > max) max = v;

      const delta = v - meanW;
      meanW += delta / n;
      const delta2 = v - meanW;
      m2 += delta * delta2;
    }
  }

  const totalCells = (cMax - cMin + 1) * (rMax - rMin + 1);
  const sd = n > 1 ? Math.sqrt(m2 / (n - 1)) : 0;
  const sem = n > 0 ? sd / Math.sqrt(n) : 0;

  const colStartName = sheet.columns[cMin]?.shortName || 'A';
  const colEndName = sheet.columns[cMax]?.shortName || colStartName;
  const selectionLabel = selection
    ? cMin === cMax
      ? `Col(${colStartName})[${rMin + 1}:${rMax + 1}]`
      : `Col(${colStartName}:${colEndName})[${rMin + 1}:${rMax + 1}]`
    : `Intero Foglio (${sheet.columns.length} col × ${maxSheetRows} righe)`;

  return {
    n,
    maskedCount,
    missingCount,
    totalCells,
    mean: n > 0 ? meanW : 0,
    sd,
    sem,
    sum: n > 0 ? sum : 0,
    min: n > 0 ? min : 0,
    max: n > 0 ? max : 0,
    selectionLabel,
  };
}

export type FindMatchMode =
  | 'exact'
  | 'instrument_errors'
  | 'contains'
  | 'lt'
  | 'gt'
  | 'outside_range';

export type ReplaceActionMode = 'replace_nan' | 'replace_value' | 'mask_matched';

/**
 * Find and Replace across a sheet or specific column (ideal for cleaning -9999 / 9999 instrument error codes)
 */
export function findAndReplaceInSheet(
  sheet: WorkbookSheet,
  params: {
    targetColId: string | 'ALL';
    matchMode: FindMatchMode;
    findQuery: string;
    findMin?: number;
    findMax?: number;
    replaceAction: ReplaceActionMode;
    replaceWith: string;
  }
): {
  updatedSheet: WorkbookSheet;
  matchCount: number;
  affectedColsCount: number;
} {
  let matchCount = 0;
  let affectedColsCount = 0;

  const queryTrimmed = params.findQuery.trim();
  const queryNum = Number(queryTrimmed.replace(',', '.'));
  const isQueryNumeric = queryTrimmed !== '' && Number.isFinite(queryNum);

  const instrumentErrorNumbers = new Set([-9999, -999, 9999, 99999, -32768, 65535, -1e30, 1e30]);
  const instrumentErrorStrings = new Set([
    '-9999',
    '9999',
    '-999',
    'nan',
    'inf',
    '-inf',
    'null',
    'err',
    'error',
    '#value!',
    '#div/0!',
    '#n/a',
    'overflow',
  ]);

  const updatedColumns = sheet.columns.map((col) => {
    if (params.targetColId !== 'ALL' && col.id !== params.targetColId) {
      return col;
    }

    const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
    const nextF64 = new Float64Array(length);
    const nextU8 = new Uint8Array(length);
    nextF64.set(float64);
    nextU8.set(maskUint8);

    let nextRawStrings = col.rawStrings ? [...col.rawStrings] : undefined;
    let colMatched = 0;

    for (let r = 0; r < length; r++) {
      const numVal = nextU8[r] === POINT_NAN ? Number.NaN : nextF64[r];
      const strVal = nextRawStrings && r < nextRawStrings.length ? nextRawStrings[r] : null;
      const cellStr =
        strVal !== null && strVal !== undefined
          ? strVal
          : Number.isFinite(numVal)
          ? String(numVal)
          : '';

      let isMatch = false;

      if (params.matchMode === 'instrument_errors') {
        if (Number.isFinite(numVal) && instrumentErrorNumbers.has(numVal)) {
          isMatch = true;
        } else if (cellStr && instrumentErrorStrings.has(cellStr.trim().toLowerCase())) {
          isMatch = true;
        }
      } else if (params.matchMode === 'exact') {
        if (isQueryNumeric && Number.isFinite(numVal)) {
          isMatch = Math.abs(numVal - queryNum) <= 1e-9;
        } else if (queryTrimmed !== '') {
          isMatch = cellStr.trim().toLowerCase() === queryTrimmed.toLowerCase();
        }
      } else if (params.matchMode === 'contains') {
        if (queryTrimmed !== '') {
          isMatch = cellStr.toLowerCase().includes(queryTrimmed.toLowerCase());
        }
      } else if (params.matchMode === 'lt') {
        if (isQueryNumeric && Number.isFinite(numVal)) {
          isMatch = numVal < queryNum;
        }
      } else if (params.matchMode === 'gt') {
        if (isQueryNumeric && Number.isFinite(numVal)) {
          isMatch = numVal > queryNum;
        }
      } else if (params.matchMode === 'outside_range') {
        const minB = params.findMin ?? -1000;
        const maxB = params.findMax ?? 1000;
        if (Number.isFinite(numVal)) {
          isMatch = numVal < minB || numVal > maxB;
        }
      }

      if (isMatch) {
        colMatched++;
        matchCount++;

        if (params.replaceAction === 'replace_nan') {
          nextF64[r] = Number.NaN;
          nextU8[r] = POINT_NAN;
          if (nextRawStrings && r < nextRawStrings.length) {
            nextRawStrings[r] = null;
          }
        } else if (params.replaceAction === 'mask_matched') {
          nextU8[r] = POINT_MASKED;
        } else {
          // replace_value
          const dt = col.dataType || (col.role === 'LABEL' ? 'text' : 'numeric');
          const parsed = parseCellInputForColumn(params.replaceWith, dt, col.role);
          if (parsed.numVal !== null && Number.isFinite(parsed.numVal)) {
            nextF64[r] = parsed.numVal;
            nextU8[r] = nextU8[r] === POINT_MASKED ? POINT_MASKED : POINT_VALID;
          } else {
            nextF64[r] = Number.NaN;
            nextU8[r] = POINT_NAN;
          }
          if (parsed.strVal !== null || nextRawStrings) {
            if (!nextRawStrings) nextRawStrings = new Array(length).fill(null);
            nextRawStrings[r] = parsed.strVal;
          }
        }
      }
    }

    if (colMatched > 0) {
      affectedColsCount++;
      return attachBinaryBuffersToColumn(
        { ...col, rawStrings: nextRawStrings },
        nextF64,
        nextU8
      );
    }
    return col;
  });

  const recalculated = recalculateSheetFormulas({
    ...sheet,
    columns: updatedColumns,
  });

  return {
    updatedSheet: recalculated,
    matchCount,
    affectedColsCount,
  };
}

/**
 * Evaluates a sheet's RowFilterConfig and returns an Int32Array of matching row indices
 */
export function evaluateSheetRowFilter(
  sheet: WorkbookSheet,
  maxRows: number
): {
  matchingRows: Int32Array | null;
  visibleCount: number;
} {
  const filter = sheet.rowFilter;
  if (!filter || !filter.enabled) {
    return { matchingRows: null, visibleCount: maxRows };
  }

  const matches: number[] = [];

  if (filter.operator === 'custom_expr' && filter.customExpression?.trim()) {
    // Evaluate custom boolean expression like `col(B) > 0 && col(A) <= 10`
    const colIndexMap = new Map<string, number>();
    const buffers = sheet.columns.map((c, idx) => {
      colIndexMap.set(c.shortName.toUpperCase(), idx);
      if (c.longName?.trim()) colIndexMap.set(c.longName.trim().toUpperCase(), idx);
      return getOrSyncColumnBuffers(c);
    });

    let boolExpr = filter.customExpression.trim();
    boolExpr = boolExpr.replace(/col\s*\(\s*["']?([a-zA-Z0-9_\s]+)["']?\s*\)/gi, (_, ref) => {
      const cIdx = colIndexMap.get(ref.trim().toUpperCase()) ?? -1;
      return `__val(${cIdx}, __r)`;
    });
    boolExpr = boolExpr.replace(/\^/g, '**');

    const getVal = (cIdx: number, r: number): number => {
      if (cIdx < 0 || cIdx >= buffers.length) return Number.NaN;
      const b = buffers[cIdx];
      if (r >= b.length || b.maskUint8[r] === POINT_NAN) return Number.NaN;
      return b.float64[r];
    };

    try {
      // eslint-disable-next-line no-new-func
      const predFn = new Function('__r', 'i', '__val', `return Boolean(${boolExpr});`) as (
        r: number,
        i: number,
        valFn: typeof getVal
      ) => boolean;

      for (let r = 0; r < maxRows; r++) {
        try {
          if (predFn(r, r + 1, getVal)) {
            matches.push(r);
          }
        } catch {
          // ignore row error
        }
      }
    } catch {
      return { matchingRows: null, visibleCount: maxRows };
    }

    return {
      matchingRows: new Int32Array(matches),
      visibleCount: matches.length,
    };
  }

  const targetCol = sheet.columns.find((c) => c.id === filter.colId) || sheet.columns[0];
  if (!targetCol) {
    return { matchingRows: null, visibleCount: maxRows };
  }

  const { float64, maskUint8, length } = getOrSyncColumnBuffers(targetCol);
  const v1Num = Number(filter.value1.replace(',', '.'));
  const v2Num = Number((filter.value2 || '0').replace(',', '.'));
  const op = filter.operator;

  for (let r = 0; r < maxRows; r++) {
    const isWithin = r < length;
    const status = isWithin ? maskUint8[r] : POINT_NAN;
    const numVal = isWithin && status !== POINT_NAN ? float64[r] : Number.NaN;
    const strVal =
      targetCol.rawStrings && r < targetCol.rawStrings.length ? targetCol.rawStrings[r] : null;
    const cellText = strVal ?? (Number.isFinite(numVal) ? String(numVal) : '');

    let keep = false;
    if (op === 'not_nan') {
      keep = Number.isFinite(numVal) || (strVal !== null && strVal.trim() !== '');
    } else if (op === 'unmasked') {
      keep = status === POINT_VALID;
    } else if (op === 'contains') {
      keep = cellText.toLowerCase().includes(filter.value1.trim().toLowerCase());
    } else if (op === 'gt') {
      keep = Number.isFinite(numVal) && Number.isFinite(v1Num) && numVal > v1Num;
    } else if (op === 'gte') {
      keep = Number.isFinite(numVal) && Number.isFinite(v1Num) && numVal >= v1Num;
    } else if (op === 'lt') {
      keep = Number.isFinite(numVal) && Number.isFinite(v1Num) && numVal < v1Num;
    } else if (op === 'lte') {
      keep = Number.isFinite(numVal) && Number.isFinite(v1Num) && numVal <= v1Num;
    } else if (op === 'eq') {
      if (Number.isFinite(v1Num) && filter.value1.trim() !== '' && Number.isFinite(numVal)) {
        keep = Math.abs(numVal - v1Num) <= 1e-9;
      } else {
        keep = cellText.trim().toLowerCase() === filter.value1.trim().toLowerCase();
      }
    } else if (op === 'neq') {
      if (Number.isFinite(v1Num) && filter.value1.trim() !== '' && Number.isFinite(numVal)) {
        keep = Math.abs(numVal - v1Num) > 1e-9;
      } else {
        keep = cellText.trim().toLowerCase() !== filter.value1.trim().toLowerCase();
      }
    } else if (op === 'between') {
      const low = Math.min(v1Num, v2Num);
      const high = Math.max(v1Num, v2Num);
      keep = Number.isFinite(numVal) && numVal >= low && numVal <= high;
    }

    if (keep) {
      matches.push(r);
    }
  }

  return {
    matchingRows: new Int32Array(matches),
    visibleCount: matches.length,
  };
}

/**
 * Extracts only the rows matching a filter into a brand new WorkbookSheet
 */
export function extractFilteredRowsToNewSheet(
  sheet: WorkbookSheet,
  matchingRows: Int32Array,
  newSheetName?: string
): WorkbookSheet {
  const m = matchingRows.length;
  const newCols = sheet.columns.map((col) => {
    const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
    const nextF64 = new Float64Array(m);
    const nextU8 = new Uint8Array(m);
    const nextStrs = col.rawStrings ? new Array<string | null>(m).fill(null) : undefined;

    for (let i = 0; i < m; i++) {
      const srcRow = matchingRows[i];
      if (srcRow < length) {
        nextF64[i] = float64[srcRow];
        nextU8[i] = maskUint8[srcRow];
      } else {
        nextF64[i] = Number.NaN;
        nextU8[i] = POINT_NAN;
      }
      if (nextStrs && col.rawStrings && srcRow < col.rawStrings.length) {
        nextStrs[i] = col.rawStrings[srcRow];
      }
    }

    return attachBinaryBuffersToColumn(
      {
        ...col,
        id: `col_filt_${Date.now()}_${col.shortName}`,
        rawStrings: nextStrs,
      },
      nextF64,
      nextU8
    );
  });

  return rebindSheetMultiXAssociations({
    id: `sheet_filt_${Date.now()}`,
    name: newSheetName || `${sheet.name}_Filtrato`,
    columns: newCols,
  });
}

export type SplitWorksheetMode = 'by_category_sheets' | 'by_category_columns' | 'by_row_chunks';

/**
 * Splits a WorkbookSheet either into multiple separate WorkbookSheets or into Multi-X columns in a new sheet
 * based on a categorical/grouping column or row interval
 */
export function splitWorksheetByColumn(
  sheet: WorkbookSheet,
  params: {
    mode: SplitWorksheetMode;
    groupColId: string;
    chunkSizeRows?: number;
  }
): {
  createdSheets: WorkbookSheet[];
  summaryText: string;
} {
  const maxRows = getSheetRowCount(sheet);

  if (params.mode === 'by_row_chunks') {
    const chunk = Math.max(2, Math.round(params.chunkSizeRows || 10));
    const createdSheets: WorkbookSheet[] = [];
    for (let start = 0, part = 1; start < maxRows; start += chunk, part++) {
      const end = Math.min(maxRows, start + chunk);
      const indices = new Int32Array(end - start);
      for (let r = start; r < end; r++) indices[r - start] = r;
      createdSheets.push(
        extractFilteredRowsToNewSheet(sheet, indices, `${sheet.name}_Blocco${part}`)
      );
    }
    return {
      createdSheets,
      summaryText: `Foglio diviso in ${createdSheets.length} fogli da max ${chunk} righe ciascuno`,
    };
  }

  const groupCol = sheet.columns.find((c) => c.id === params.groupColId) || sheet.columns[0];
  if (!groupCol) {
    return { createdSheets: [], summaryText: 'Colonna di raggruppamento non trovata' };
  }

  const { float64, maskUint8, length } = getOrSyncColumnBuffers(groupCol);
  const groupsMap = new Map<string, number[]>();

  for (let r = 0; r < maxRows; r++) {
    let key = '';
    if (
      groupCol.rawStrings &&
      r < groupCol.rawStrings.length &&
      groupCol.rawStrings[r] !== null &&
      groupCol.rawStrings[r]!.trim() !== ''
    ) {
      key = groupCol.rawStrings[r]!.trim();
    } else if (r < length && maskUint8[r] !== POINT_NAN && Number.isFinite(float64[r])) {
      key = String(float64[r]);
    }
    if (!key) continue;
    if (!groupsMap.has(key)) {
      groupsMap.set(key, []);
    }
    groupsMap.get(key)!.push(r);
  }

  const entries = Array.from(groupsMap.entries()).slice(0, 25); // cap to 25 groups for safety
  if (entries.length === 0) {
    return { createdSheets: [], summaryText: 'Nessuna categoria valida trovata nella colonna scelta' };
  }

  if (params.mode === 'by_category_sheets') {
    const createdSheets = entries.map(([catKey, rowList], idx) => {
      const safeCat = catKey.replace(/[^a-zA-Z0-9_\-.]/g, '_').slice(0, 16) || `Grp${idx + 1}`;
      return extractFilteredRowsToNewSheet(
        sheet,
        new Int32Array(rowList),
        `${sheet.name}_${safeCat}`
      );
    });
    return {
      createdSheets,
      summaryText: `Dataset diviso in ${createdSheets.length} fogli separati per "${groupCol.longName || groupCol.shortName}"`,
    };
  }

  // 'by_category_columns': Unstack into Multi-X columns (X_cat1, Y_cat1, X_cat2, Y_cat2, ...)
  const nonGroupCols = sheet.columns.filter((c) => c.id !== groupCol.id);
  const colsToUnstack = nonGroupCols.length > 0 ? nonGroupCols : sheet.columns;
  const unstackedCols: DataColumn[] = [];

  entries.forEach(([catKey, rowList], gIdx) => {
    const m = rowList.length;
    colsToUnstack.forEach((srcCol) => {
      const { float64: sF64, maskUint8: sU8, length: sLen } = getOrSyncColumnBuffers(srcCol);
      const nextF64 = new Float64Array(m);
      const nextU8 = new Uint8Array(m);
      const nextStrs = srcCol.rawStrings ? new Array<string | null>(m).fill(null) : undefined;

      for (let i = 0; i < m; i++) {
        const r = rowList[i];
        if (r < sLen) {
          nextF64[i] = sF64[r];
          nextU8[i] = sU8[r];
        } else {
          nextF64[i] = Number.NaN;
          nextU8[i] = POINT_NAN;
        }
        if (nextStrs && srcCol.rawStrings && r < srcCol.rawStrings.length) {
          nextStrs[i] = srcCol.rawStrings[r];
        }
      }

      const idx = unstackedCols.length;
      unstackedCols.push(
        attachBinaryBuffersToColumn(
          {
            ...srcCol,
            id: `col_unstack_${Date.now()}_${gIdx}_${idx}`,
            shortName: getColumnShortName(idx),
            longName: `${srcCol.longName || srcCol.shortName} [${catKey}]`,
            comment: `Gruppo ${groupCol.shortName} = ${catKey}`,
            formula: undefined,
            rawStrings: nextStrs,
          },
          nextF64,
          nextU8
        )
      );
    });
  });

  const unstackedSheet = rebindSheetMultiXAssociations({
    id: `sheet_unstack_${Date.now()}`,
    name: `${sheet.name}_SplitCol`,
    columns: unstackedCols,
  });

  return {
    createdSheets: [unstackedSheet],
    summaryText: `Dataset diviso in ${unstackedCols.length} colonne Multi-X (${entries.length} gruppi) nel foglio "${unstackedSheet.name}"`,
  };
}

/**
 * Automatically recalculates all columns in a sheet that have a non-empty `formula` (`F(x)=`)
 * Supports multi-step dependency chains (e.g., Col C depends on Col B which depends on Col A, or ΔZ depends on Z).
 */
export function recalculateSheetFormulas(sheet: WorkbookSheet): WorkbookSheet {
  if (sheet.autoRecalculate === false) return sheet;

  const hasFormulas = sheet.columns.some((c) => Boolean(c.formula && c.formula.trim()));
  if (!hasFormulas) return sheet;

  let currentCols = [...sheet.columns];
  // Run up to 2 topological evaluation passes so chained formulas (A -> B -> C or A,B -> Z -> dZ) stay in sync
  const passes = Math.min(3, currentCols.filter((c) => Boolean(c.formula?.trim())).length);

  for (let pass = 0; pass < passes; pass++) {
    currentCols = currentCols.map((col) => {
      if (!col.formula || !col.formula.trim()) return col;
      const evaluated = evaluateColumnFormula(col.formula, col, currentCols);
      const prevBuf = getOrSyncColumnBuffers(col);
      const len = evaluated.length;
      const nextF64 = new Float64Array(len);
      const nextU8 = new Uint8Array(len);

      for (let i = 0; i < len; i++) {
        const v = evaluated[i];
        const wasMasked = i < prevBuf.length && prevBuf.maskUint8[i] === POINT_MASKED;
        if (v !== null && v !== undefined && Number.isFinite(v)) {
          nextF64[i] = v;
          nextU8[i] = wasMasked ? POINT_MASKED : POINT_VALID;
        } else {
          nextF64[i] = Number.NaN;
          nextU8[i] = POINT_NAN;
        }
      }

      return attachBinaryBuffersToColumn(col, nextF64, nextU8);
    });
  }

  return {
    ...sheet,
    columns: currentCols,
  };
}

export type ErrorPropagationOp = 'add' | 'sub' | 'mul' | 'div' | 'pow' | 'linear_comb';

/**
 * Computes Elementary Error Propagation (Z and ΔZ) for basic operations:
 * - add: Z = A + B, ΔZ = sqrt(ΔA^2 + ΔB^2)
 * - sub: Z = A - B, ΔZ = sqrt(ΔA^2 + ΔB^2)
 * - mul: Z = A * B, ΔZ = |Z| * sqrt((ΔA/A)^2 + (ΔB/B)^2)
 * - div: Z = A / B, ΔZ = |Z| * sqrt((ΔA/A)^2 + (ΔB/B)^2)
 * - pow: Z = A^n,   ΔZ = |n * A^(n-1)| * ΔA
 * - linear_comb: Z = kA*A + kB*B, ΔZ = sqrt((kA*ΔA)^2 + (kB*ΔB)^2)
 *
 * Creates/updates both the result column Z (Role Y) and its linked error column ΔZ (Role Y_ERR)
 * with live F(x) formulas so Automatic Recalculation keeps Z and ΔZ updated whenever A, B, ΔA, or ΔB change!
 */
export function applyElementaryErrorPropagation(
  sheet: WorkbookSheet,
  params: {
    operation: ErrorPropagationOp;
    colAId: string;
    errAMode: 'column' | 'constant';
    colErrAId?: string;
    constErrA?: number;
    colBId?: string;
    errBMode?: 'column' | 'constant';
    colErrBId?: string;
    constErrB?: number;
    exponentN?: number;
    coeffA?: number;
    coeffB?: number;
    createZColumn?: boolean;
  }
): {
  updatedSheet: WorkbookSheet;
  zColShortName: string;
  errZColShortName: string;
  formulaZ: string;
  formulaErrZ: string;
} {
  const colA = sheet.columns.find((c) => c.id === params.colAId) || sheet.columns[0];
  const colB =
    sheet.columns.find((c) => c.id === params.colBId) ||
    sheet.columns[Math.min(1, sheet.columns.length - 1)];
  const colErrA =
    params.errAMode === 'column'
      ? sheet.columns.find((c) => c.id === params.colErrAId)
      : undefined;
  const colErrB =
    params.errBMode === 'column'
      ? sheet.columns.find((c) => c.id === params.colErrBId)
      : undefined;

  const refA = `col(${colA.shortName})`;
  const refB = colB ? `col(${colB.shortName})` : '1';
  const refErrA = colErrA ? `col(${colErrA.shortName})` : String(Math.abs(params.constErrA ?? 0.1));
  const refErrB = colErrB ? `col(${colErrB.shortName})` : String(Math.abs(params.constErrB ?? 0.1));
  const nExp = params.exponentN ?? 2;
  const kA = params.coeffA ?? 1;
  const kB = params.coeffB ?? 1;

  let formulaZ = '';
  let formulaErrZ = '';
  let opSymbol = '';
  let zUnit = colA.unit || '';

  switch (params.operation) {
    case 'add':
      formulaZ = `${refA} + ${refB}`;
      formulaErrZ = `err_add(${refA}, ${refErrA}, ${refB}, ${refErrB})`;
      opSymbol = `${colA.shortName} + ${colB.shortName}`;
      break;
    case 'sub':
      formulaZ = `${refA} - ${refB}`;
      formulaErrZ = `err_sub(${refA}, ${refErrA}, ${refB}, ${refErrB})`;
      opSymbol = `${colA.shortName} - ${colB.shortName}`;
      break;
    case 'mul':
      formulaZ = `${refA} * ${refB}`;
      formulaErrZ = `err_mul(${refA}, ${refErrA}, ${refB}, ${refErrB})`;
      opSymbol = `${colA.shortName} · ${colB.shortName}`;
      zUnit =
        colA.unit && colB?.unit ? `${colA.unit}·${colB.unit}` : colA.unit || colB?.unit || '';
      break;
    case 'div':
      formulaZ = `${refA} / ${refB}`;
      formulaErrZ = `err_div(${refA}, ${refErrA}, ${refB}, ${refErrB})`;
      opSymbol = `${colA.shortName} / ${colB.shortName}`;
      zUnit =
        colA.unit && colB?.unit ? `${colA.unit}/${colB.unit}` : colA.unit || colB?.unit || '';
      break;
    case 'pow':
      formulaZ = `(${refA})^${nExp}`;
      formulaErrZ = `err_pow(${refA}, ${refErrA}, ${nExp})`;
      opSymbol = `${colA.shortName}^${nExp}`;
      zUnit = colA.unit ? `${colA.unit}^${nExp}` : '';
      break;
    case 'linear_comb':
      formulaZ = `${kA} * ${refA} + (${kB}) * ${refB}`;
      formulaErrZ = `sqrt((${Math.abs(kA)} * ${refErrA})^2 + (${Math.abs(kB)} * ${refErrB})^2)`;
      opSymbol = `${kA}·${colA.shortName} + ${kB}·${colB.shortName}`;
      break;
  }

  const maxRows = Math.max(1, getSheetRowCount(sheet));
  const nextCols = [...sheet.columns];
  const defaultX = resolveXColumnForY(sheet, colA);

  let zColId = colA.id;
  let zColShortName = colA.shortName;

  if (params.createZColumn !== false) {
    zColShortName = getColumnShortName(nextCols.length);
    zColId = `col_z_${Date.now()}`;
    const emptyF64 = new Float64Array(maxRows).fill(Number.NaN);
    const emptyU8 = new Uint8Array(maxRows).fill(POINT_NAN);
    nextCols.push(
      attachBinaryBuffersToColumn(
        {
          id: zColId,
          shortName: zColShortName,
          longName: `Z = ${opSymbol}`,
          unit: zUnit,
          comment: `Calcolato con F(x): ${formulaZ}`,
          formula: formulaZ,
          role: 'Y',
          dataType: 'numeric',
          linkedXColId: defaultX?.id,
          values: [],
        },
        emptyF64,
        emptyU8
      )
    );
  }

  const errZColShortName = getColumnShortName(nextCols.length);
  const errZColId = `col_dz_${Date.now()}`;
  const emptyErrF64 = new Float64Array(maxRows).fill(Number.NaN);
  const emptyErrU8 = new Uint8Array(maxRows).fill(POINT_NAN);

  nextCols.push(
    attachBinaryBuffersToColumn(
      {
        id: errZColId,
        shortName: errZColShortName,
        longName: `Δ(${opSymbol})`,
        unit: zUnit,
        comment: `Propagazione errore: ${formulaErrZ}`,
        formula: formulaErrZ,
        role: 'Y_ERR',
        dataType: 'numeric',
        linkedXColId: defaultX?.id,
        values: [],
      },
      emptyErrF64,
      emptyErrU8
    )
  );

  // Link the newly created ΔZ column directly to Z
  const linkedCols = nextCols.map((c) =>
    c.id === zColId ? { ...c, linkedErrColId: errZColId } : c
  );

  const rebound = rebindSheetMultiXAssociations({
    ...sheet,
    columns: linkedCols,
  });

  const recalculated = recalculateSheetFormulas(rebound);

  return {
    updatedSheet: recalculated,
    zColShortName,
    errZColShortName,
    formulaZ,
    formulaErrZ,
  };
}

