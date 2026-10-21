import React, { useState, useMemo, useEffect } from 'react';
import { X, Sparkles, Check, Activity } from 'lucide-react';
import { SmoothingFilterMethod, WorkbookSheet } from '../types/origin';
import {
  applySignalFilterAndSmooth,
  formatSciNumber,
  resolveXColumnForY,
} from '../utils/mathEngine';

interface FilterSmoothModalProps {
  isOpen: boolean;
  onClose: () => void;
  sheets: WorkbookSheet[];
  activeSheetId: string;
  onApplyFilterResult: (params: {
    sheetId: string;
    sourceColId: string;
    filteredValues: (number | null)[];
    methodLabel: string;
    outputMode: 'new-col-and-plot' | 'new-col-only' | 'overwrite';
    color: string;
  }) => void;
}

export const FilterSmoothModal: React.FC<FilterSmoothModalProps> = ({
  isOpen,
  onClose,
  sheets,
  activeSheetId,
  onApplyFilterResult,
}) => {
  const [selectedSheetId, setSelectedSheetId] = useState<string>(activeSheetId);
  const currentSheet = sheets.find((s) => s.id === selectedSheetId) || sheets[0];

  const yCols = useMemo(
    () => currentSheet?.columns.filter((c) => c.role === 'Y') || [],
    [currentSheet]
  );

  const [selectedColId, setSelectedColId] = useState<string>(() => yCols[0]?.id || '');

  useEffect(() => {
    setSelectedSheetId(activeSheetId);
  }, [activeSheetId]);

  useEffect(() => {
    if (yCols.length > 0 && !yCols.some((c) => c.id === selectedColId)) {
      setSelectedColId(yCols[0].id);
    }
  }, [yCols, selectedColId]);

  const [method, setMethod] = useState<SmoothingFilterMethod>('savitzky_golay');
  const [windowSize, setWindowSize] = useState<number>(7);
  const [polyOrder, setPolyOrder] = useState<number>(2);
  const [cutoffAlpha, setCutoffAlpha] = useState<number>(0.22);
  const [zScoreThreshold, setZScoreThreshold] = useState<number>(2.0);
  const [outputMode, setOutputMode] = useState<'new-col-and-plot' | 'new-col-only' | 'overwrite'>(
    'new-col-and-plot'
  );
  const [curveColor, setCurveColor] = useState<string>('#16A34A');

  const sourceCol = currentSheet?.columns.find((c) => c.id === selectedColId) || yCols[0];
  const xCol =
    currentSheet && sourceCol ? resolveXColumnForY(currentSheet, sourceCol) : undefined;

  const previewResult = useMemo(() => {
    if (!sourceCol) return { filtered: [], outliersReplacedCount: 0 };
    return applySignalFilterAndSmooth({
      values: sourceCol.values,
      method,
      windowSize,
      polyOrder,
      cutoffAlpha,
      zScoreThreshold,
    });
  }, [sourceCol, method, windowSize, polyOrder, cutoffAlpha, zScoreThreshold]);

  // Build preview coordinates for live SVG comparison chart
  const previewChartData = useMemo(() => {
    if (!sourceCol || !xCol) return { rawPts: [], filtPts: [], xMin: 0, xMax: 10, yMin: 0, yMax: 10 };
    const len = Math.min(xCol.values.length, sourceCol.values.length);
    const rawPts: { x: number; y: number }[] = [];
    const filtPts: { x: number; y: number }[] = [];

    let minX = Infinity,
      maxX = -Infinity,
      minY = Infinity,
      maxY = -Infinity;

    for (let i = 0; i < len; i++) {
      const xv = xCol.values[i];
      const yRaw = sourceCol.values[i];
      const yFilt = previewResult.filtered[i];
      if (xv !== null && Number.isFinite(xv)) {
        if (yRaw !== null && Number.isFinite(yRaw)) {
          rawPts.push({ x: xv, y: yRaw });
          if (xv < minX) minX = xv;
          if (xv > maxX) maxX = xv;
          if (yRaw < minY) minY = yRaw;
          if (yRaw > maxY) maxY = yRaw;
        }
        if (yFilt !== null && yFilt !== undefined && Number.isFinite(yFilt)) {
          filtPts.push({ x: xv, y: yFilt });
          if (yFilt < minY) minY = yFilt;
          if (yFilt > maxY) maxY = yFilt;
        }
      }
    }

    if (!Number.isFinite(minX)) {
      minX = 0;
      maxX = 10;
      minY = 0;
      maxY = 10;
    }
    const dx = (maxX - minX) * 0.04 || 1;
    const dy = (maxY - minY) * 0.08 || 1;

    return {
      rawPts,
      filtPts,
      xMin: minX - dx,
      xMax: maxX + dx,
      yMin: minY - dy,
      yMax: maxY + dy,
    };
  }, [sourceCol, xCol, previewResult]);

  if (!isOpen) return null;

  const getMethodTitle = (): string => {
    switch (method) {
      case 'savitzky_golay':
        return `Savitzky-Golay (W=${windowSize}, g=${polyOrder})`;
      case 'moving_average':
        return `Media Mobile (W=${windowSize})`;
      case 'gaussian':
        return `Smoothing Gaussiano (W=${windowSize})`;
      case 'median':
        return `Filtro Mediano Anti-Spike (W=${windowSize})`;
      case 'lowpass':
        return `Passa-Basso Zero-Phase (α=${cutoffAlpha})`;
      case 'highpass':
        return `Passa-Alto Residui (α=${cutoffAlpha})`;
      case 'outlier_zscore':
        return `Filtro Outlier MAD (${zScoreThreshold}σ)`;
    }
  };

  const handleConfirm = () => {
    if (!currentSheet || !sourceCol) return;
    onApplyFilterResult({
      sheetId: currentSheet.id,
      sourceColId: sourceCol.id,
      filteredValues: previewResult.filtered,
      methodLabel: getMethodTitle(),
      outputMode,
      color: curveColor,
    });
    onClose();
  };

  const W = 520;
  const H = 220;
  const P = 28;
  const toX = (x: number) =>
    P +
    ((x - previewChartData.xMin) / (previewChartData.xMax - previewChartData.xMin || 1)) *
      (W - 2 * P);
  const toY = (y: number) =>
    H -
    P -
    ((y - previewChartData.yMin) / (previewChartData.yMax - previewChartData.yMin || 1)) *
      (H - 2 * P);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
      <div className="w-full max-w-3xl bg-white border border-slate-300 rounded shadow-xl flex flex-col max-h-[90vh] overflow-hidden text-xs">
        {/* Header */}
        <div className="flex items-center justify-between px-4 py-2.5 bg-[#eef2f6] border-b border-slate-300">
          <div className="flex items-center gap-2">
            <Sparkles className="w-4 h-4 text-emerald-700" />
            <span className="font-bold text-slate-900 text-sm">
              Origin Signal Processing — Filtro Dati & Smoothing
            </span>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1 text-slate-500 hover:text-slate-800 rounded"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-4 space-y-4">
          {/* Row 1: Source Sheet & Column */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 p-3 bg-slate-50 border border-slate-200 rounded">
            <div>
              <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                1. Foglio Dati (Workbook)
              </label>
              <select
                value={currentSheet?.id || ''}
                onChange={(e) => setSelectedSheetId(e.target.value)}
                className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded"
              >
                {sheets.map((s) => (
                  <option key={s.id} value={s.id}>
                    {s.name}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                2. Colonna Segnale Y da Filtrare / Levigare
              </label>
              <select
                value={sourceCol?.id || ''}
                onChange={(e) => setSelectedColId(e.target.value)}
                className="w-full px-2.5 py-1.5 text-xs font-mono bg-white border border-slate-300 rounded"
              >
                {yCols.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.shortName}(Y): {c.longName} {c.unit ? `[${c.unit}]` : ''}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Row 2: Filter / Smoothing Algorithm Selector */}
          <div className="space-y-2">
            <label className="block text-xs font-semibold text-slate-800">
              3. Algoritmo di Smoothing o Filtro Segnale
            </label>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-1.5">
              {[
                {
                  id: 'savitzky_golay',
                  title: 'Savitzky-Golay',
                  desc: 'Polinomiale locale (preserva picchi)',
                },
                {
                  id: 'moving_average',
                  title: 'Media Mobile',
                  desc: 'Media uniforme su finestra W',
                },
                {
                  id: 'gaussian',
                  title: 'Filtro Gaussiano',
                  desc: 'Convoluzione pesata morbida',
                },
                {
                  id: 'median',
                  title: 'Filtro Mediano',
                  desc: 'Elimina impulsi e spike isolati',
                },
                {
                  id: 'outlier_zscore',
                  title: 'Pulizia Outlier (MAD)',
                  desc: 'Sostituisce punti anomali > k·σ',
                },
                {
                  id: 'lowpass',
                  title: 'Filtro Passa-Basso',
                  desc: 'Rimuove rumore ad alta frequenza',
                },
                {
                  id: 'highpass',
                  title: 'Filtro Passa-Alto',
                  desc: 'Estrae oscillazioni / rimuove deriva',
                },
              ].map((item) => (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => setMethod(item.id as SmoothingFilterMethod)}
                  className={`p-2 rounded border text-left transition-colors ${
                    method === item.id
                      ? 'bg-emerald-50 border-emerald-600 text-emerald-950 font-semibold shadow-2xs'
                      : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
                  }`}
                >
                  <div className="text-xs">{item.title}</div>
                  <div className="text-[10px] font-normal text-slate-500 leading-tight mt-0.5">
                    {item.desc}
                  </div>
                </button>
              ))}
            </div>
          </div>

          {/* Row 3: Dynamic Parameters for Selected Filter */}
          <div className="p-3 bg-slate-50 border border-slate-200 rounded grid grid-cols-1 md:grid-cols-3 gap-4 items-center">
            {(method === 'savitzky_golay' ||
              method === 'moving_average' ||
              method === 'gaussian' ||
              method === 'median' ||
              method === 'outlier_zscore') && (
              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="font-semibold text-slate-700">
                    Ampiezza Finestra (W punti)
                  </label>
                  <span className="font-mono font-bold text-emerald-800">{windowSize} pt</span>
                </div>
                <input
                  type="range"
                  min={3}
                  max={31}
                  step={2}
                  value={windowSize}
                  onChange={(e) => setWindowSize(Number(e.target.value))}
                  className="w-full accent-emerald-600"
                />
              </div>
            )}

            {method === 'savitzky_golay' && (
              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="font-semibold text-slate-700">Ordine Polinomiale</label>
                  <span className="font-mono font-bold text-emerald-800">Grado {polyOrder}</span>
                </div>
                <select
                  value={polyOrder}
                  onChange={(e) => setPolyOrder(Number(e.target.value))}
                  className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                >
                  <option value={1}>1° Grado (Lineare locale)</option>
                  <option value={2}>2° Grado (Quadratico - Standard Origin)</option>
                  <option value={3}>3° Grado (Cubico)</option>
                  <option value={4}>4° Grado (Quartico)</option>
                </select>
              </div>
            )}

            {(method === 'lowpass' || method === 'highpass') && (
              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="font-semibold text-slate-700">
                    Frequenza di Taglio (α)
                  </label>
                  <span className="font-mono font-bold text-emerald-800">{cutoffAlpha}</span>
                </div>
                <input
                  type="range"
                  min={0.04}
                  max={0.85}
                  step={0.02}
                  value={cutoffAlpha}
                  onChange={(e) => setCutoffAlpha(Number(e.target.value))}
                  className="w-full accent-emerald-600"
                />
              </div>
            )}

            {method === 'outlier_zscore' && (
              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="font-semibold text-slate-700">Soglia Rifiuto Outlier</label>
                  <span className="font-mono font-bold text-red-700">
                    {zScoreThreshold.toFixed(1)} σ ({previewResult.outliersReplacedCount} rimossi)
                  </span>
                </div>
                <input
                  type="range"
                  min={1.0}
                  max={4.5}
                  step={0.1}
                  value={zScoreThreshold}
                  onChange={(e) => setZScoreThreshold(Number(e.target.value))}
                  className="w-full accent-red-600"
                />
              </div>
            )}

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                Colore Curva Filtrata
              </label>
              <div className="flex items-center gap-2">
                <input
                  type="color"
                  value={curveColor}
                  onChange={(e) => setCurveColor(e.target.value)}
                  className="w-8 h-7 rounded border border-slate-300 cursor-pointer"
                />
                <span className="font-mono text-slate-600">{curveColor}</span>
              </div>
            </div>
          </div>

          {/* Row 4: Live Interactive SVG Preview */}
          <div className="border border-slate-300 rounded bg-white p-2.5 space-y-1.5">
            <div className="flex items-center justify-between text-[11px]">
              <span className="font-semibold text-slate-800 flex items-center gap-1.5">
                <Activity className="w-3.5 h-3.5 text-emerald-600" />
                <span>Anteprima in Tempo Reale: Segnale Grezzo vs Curva Filtrata</span>
              </span>
              <div className="flex items-center gap-3 font-mono text-[10.5px]">
                <span className="flex items-center gap-1 text-slate-500">
                  <span className="w-2 h-2 rounded-full bg-slate-400 inline-block" /> Originale (
                  {previewChartData.rawPts.length} pt)
                </span>
                <span className="flex items-center gap-1 font-semibold" style={{ color: curveColor }}>
                  <span
                    className="w-3 h-0.5 inline-block"
                    style={{ backgroundColor: curveColor }}
                  />{' '}
                  {getMethodTitle()}
                </span>
              </div>
            </div>

            <svg viewBox={`0 0 ${W} ${H}`} className="w-full h-48 bg-slate-50/60 border border-slate-200 rounded">
              <rect x={P} y={P} width={W - 2 * P} height={H - 2 * P} fill="#ffffff" stroke="#cbd5e1" />
              {/* Original noisy points & thin line */}
              {previewChartData.rawPts.length > 1 && (
                <path
                  d={previewChartData.rawPts
                    .map((p, i) => `${i === 0 ? 'M' : 'L'} ${toX(p.x).toFixed(1)} ${toY(p.y).toFixed(1)}`)
                    .join(' ')}
                  fill="none"
                  stroke="#94a3b8"
                  strokeWidth={1}
                  strokeDasharray="2,2"
                />
              )}
              {previewChartData.rawPts.map((p, i) => (
                <circle
                  key={i}
                  cx={toX(p.x)}
                  cy={toY(p.y)}
                  r={2.3}
                  fill="#64748b"
                  opacity={0.65}
                />
              ))}

              {/* Filtered / Smoothed curve */}
              {previewChartData.filtPts.length > 1 && (
                <path
                  d={previewChartData.filtPts
                    .map((p, i) => `${i === 0 ? 'M' : 'L'} ${toX(p.x).toFixed(1)} ${toY(p.y).toFixed(1)}`)
                    .join(' ')}
                  fill="none"
                  stroke={curveColor}
                  strokeWidth={2.4}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              )}

              <text x={P + 6} y={P + 14} className="fill-slate-500 font-mono" style={{ fontSize: '9.5px' }}>
                Y max: {formatSciNumber(previewChartData.yMax, 4)}
              </text>
              <text x={P + 6} y={H - P - 6} className="fill-slate-500 font-mono" style={{ fontSize: '9.5px' }}>
                Y min: {formatSciNumber(previewChartData.yMin, 4)}
              </text>
            </svg>
          </div>

          {/* Row 5: Output Destination */}
          <div className="flex flex-wrap items-center justify-between gap-3 p-2.5 bg-slate-50 border border-slate-200 rounded">
            <span className="font-semibold text-slate-800">Destinazione Dati Filtrati:</span>
            <div className="flex items-center gap-3">
              <label className="flex items-center gap-1.5 cursor-pointer">
                <input
                  type="radio"
                  checked={outputMode === 'new-col-and-plot'}
                  onChange={() => setOutputMode('new-col-and-plot')}
                  className="text-emerald-600"
                />
                <span>Nuova Colonna + Traccia nel Grafico</span>
              </label>
              <label className="flex items-center gap-1.5 cursor-pointer">
                <input
                  type="radio"
                  checked={outputMode === 'new-col-only'}
                  onChange={() => setOutputMode('new-col-only')}
                  className="text-emerald-600"
                />
                <span>Solo Nuova Colonna in Tabella</span>
              </label>
              <label className="flex items-center gap-1.5 cursor-pointer">
                <input
                  type="radio"
                  checked={outputMode === 'overwrite'}
                  onChange={() => setOutputMode('overwrite')}
                  className="text-emerald-600"
                />
                <span>Sovrascrivi Colonna Originale</span>
              </label>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="flex items-center justify-end gap-2 px-4 py-2.5 bg-[#eef2f6] border-t border-slate-300">
          <button
            type="button"
            onClick={onClose}
            className="px-3.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50"
          >
            Annulla
          </button>
          <button
            type="button"
            onClick={handleConfirm}
            className="flex items-center gap-1.5 px-4 py-1.5 text-xs font-semibold text-white bg-emerald-600 rounded hover:bg-emerald-700 shadow-2xs"
          >
            <Check className="w-3.5 h-3.5" />
            <span>Applica {getMethodTitle()}</span>
          </button>
        </div>
      </div>
    </div>
  );
};
