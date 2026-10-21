import React, { useState, useEffect } from 'react';
import {
  Sliders,
  TrendingUp,
  Axis3d,
  Target,
  Plus,
  Trash2,
  Eye,
  EyeOff,
  Download,
  FileSpreadsheet,
  RotateCcw,
  Sparkles,
  Layers,
  CircleDot,
  Activity,
  BarChart2,
} from 'lucide-react';
import {
  BadPointRemovalMode,
  ConnectionType,
  FitModelType,
  FittedCurve,
  GraphWindow,
  LineStyle,
  MarkerShape,
  PlotSeries,
  PlotType,
  RemovedPointRecord,
  ToolMode,
  WorkbookSheet,
} from '../types/origin';
import { formatSciNumber } from '../utils/mathEngine';

interface InspectorPanelProps {
  activeTab: 'series' | 'fitting' | 'axes' | 'badpoints';
  onChangeTab: (tab: 'series' | 'fitting' | 'axes' | 'badpoints') => void;
  graph: GraphWindow;
  sheets: WorkbookSheet[];
  selectedSeriesId: string | null;
  onSelectSeries: (id: string) => void;
  onUpdateGraph: (updated: GraphWindow) => void;
  onAddSeriesToGraph: (params: {
    sheetId: string;
    xColId: string;
    yColId: string;
    plotType: PlotType;
    connection: ConnectionType;
    color: string;
  }) => void;
  onRunCurveFit: (params: {
    sourceSeriesId: string;
    model: FitModelType;
    polyDegree: number;
    rangeC: number;
    rangeD: number;
    numPointsN: number;
    color: string;
    lineWidth: number;
    lineStyle: LineStyle;
    createOutputSheet: boolean;
  }) => void;
  onDeleteFit: (fitId: string) => void;
  onOpenFitSheet: (sheetId: string) => void;
  onExportFitTxt: (fit: FittedCurve) => void;
  toolMode: ToolMode;
  onChangeToolMode: (mode: ToolMode) => void;
  badPointRemovalMode: BadPointRemovalMode;
  onChangeBadPointRemovalMode: (mode: BadPointRemovalMode) => void;
  removedPointsLog: RemovedPointRecord[];
  onRestoreRemovedPoint: (recordId: string) => void;
  onApplySmoothing: (sheetId: string, colId: string, windowSize: number) => void;
  onApplyDerivative: (sheetId: string, xColId: string, yColId: string) => void;
  onApplyIntegration: (sheetId: string, xColId: string, yColId: string) => void;
  onOpenFilterSmoothModal?: () => void;
}

const SCIENTIFIC_PALETTE = [
  '#2563EB',
  '#DC2626',
  '#0D9488',
  '#D97706',
  '#7C3AED',
  '#16A34A',
  '#0F172A',
  '#DB2777',
];

export const InspectorPanel: React.FC<InspectorPanelProps> = ({
  activeTab,
  onChangeTab,
  graph,
  sheets,
  selectedSeriesId,
  onSelectSeries,
  onUpdateGraph,
  onAddSeriesToGraph,
  onRunCurveFit,
  onDeleteFit,
  onOpenFitSheet,
  onExportFitTxt,
  toolMode,
  onChangeToolMode,
  badPointRemovalMode,
  onChangeBadPointRemovalMode,
  removedPointsLog,
  onRestoreRemovedPoint,
  onApplySmoothing,
  onApplyDerivative,
  onApplyIntegration,
  onOpenFilterSmoothModal,
}) => {
  // Add series state
  const [addSheetId, setAddSheetId] = useState<string>(() => sheets[0]?.id || '');
  const selectedAddSheet = sheets.find((s) => s.id === addSheetId) || sheets[0];
  const [addXColId, setAddXColId] = useState<string>('');
  const [addYColId, setAddYColId] = useState<string>('');
  const [addPlotType, setAddPlotType] = useState<PlotType>('line_symbol');
  const [addConnection, setAddConnection] = useState<ConnectionType>('linear');

  useEffect(() => {
    if (selectedAddSheet) {
      const xCol =
        selectedAddSheet.columns.find((c) => c.role === 'X') || selectedAddSheet.columns[0];
      const yCol =
        selectedAddSheet.columns.find((c) => c.role === 'Y') ||
        selectedAddSheet.columns[1] ||
        selectedAddSheet.columns[0];
      if (xCol) setAddXColId(xCol.id);
      if (yCol) setAddYColId(yCol.id);
    }
  }, [selectedAddSheet?.id]);

  // Curve fitting state
  const [fitSeriesId, setFitSeriesId] = useState<string>(() => graph.series[0]?.id || '');
  const [fitModel, setFitModel] = useState<FitModelType>('linear');
  const [fitPolyDegree, setFitPolyDegree] = useState<number>(2);
  const [fitRangeC, setFitRangeC] = useState<string>('0');
  const [fitRangeD, setFitRangeD] = useState<string>('12');
  const [fitNumPointsN, setFitNumPointsN] = useState<string>('100');
  const [fitColor, setFitColor] = useState<string>('#DC2626');
  const [fitLineWidth, setFitLineWidth] = useState<number>(2.5);
  const [fitLineStyle, setFitLineStyle] = useState<LineStyle>('solid');
  const [fitCreateSheet, setFitCreateSheet] = useState<boolean>(true);
  const [smoothWindow, setSmoothWindow] = useState<number>(5);

  useEffect(() => {
    if (!graph.series.some((s) => s.id === fitSeriesId) && graph.series[0]) {
      setFitSeriesId(graph.series[0].id);
    }
  }, [graph.series, fitSeriesId]);

  const sourceSeriesStats = React.useMemo(() => {
    const s = graph.series.find((item) => item.id === fitSeriesId) || graph.series[0];
    if (!s) return null;
    const sh = sheets.find((item) => item.id === s.sheetId);
    if (!sh) return null;
    const xCol = sh.columns.find((c) => c.id === s.xColId);
    const yCol = sh.columns.find((c) => c.id === s.yColId);
    if (!xCol || !yCol) return null;

    let minA = Infinity;
    let maxB = -Infinity;
    let count = 0;
    const len = Math.min(xCol.values.length, yCol.values.length);
    for (let i = 0; i < len; i++) {
      const xv = xCol.values[i];
      const yv = yCol.values[i];
      if (xv !== null && yv !== null && Number.isFinite(xv) && Number.isFinite(yv)) {
        if (xv < minA) minA = xv;
        if (xv > maxB) maxB = xv;
        count++;
      }
    }
    if (count === 0) return { minA: 0, maxB: 10, count: 0, series: s };
    return { minA, maxB, count, series: s };
  }, [fitSeriesId, graph.series, sheets]);

  const handleUpdateSeries = (seriesId: string, patch: Partial<PlotSeries>) => {
    const nextSeries = graph.series.map((s) => (s.id === seriesId ? { ...s, ...patch } : s));
    onUpdateGraph({ ...graph, series: nextSeries });
  };

  const handleRemoveSeries = (seriesId: string) => {
    onUpdateGraph({
      ...graph,
      series: graph.series.filter((s) => s.id !== seriesId),
      fits: graph.fits.filter((f) => f.sourceSeriesId !== seriesId),
    });
  };

  const handleExecuteFit = (e: React.FormEvent) => {
    e.preventDefault();
    const targetSeriesId = fitSeriesId || graph.series[0]?.id;
    if (!targetSeriesId) return;

    const cVal = Number(fitRangeC.replace(',', '.'));
    const dVal = Number(fitRangeD.replace(',', '.'));
    const nVal = Math.max(2, Math.min(5000, Math.round(Number(fitNumPointsN))));

    if (!Number.isFinite(cVal) || !Number.isFinite(dVal) || cVal === dVal) return;

    onRunCurveFit({
      sourceSeriesId: targetSeriesId,
      model: fitModel,
      polyDegree: fitPolyDegree,
      rangeC: Math.min(cVal, dVal),
      rangeD: Math.max(cVal, dVal),
      numPointsN: nVal,
      color: fitColor,
      lineWidth: fitLineWidth,
      lineStyle: fitLineStyle,
      createOutputSheet: fitCreateSheet,
    });
  };

  const activeSeriesObj =
    graph.series.find((s) => s.id === selectedSeriesId) || graph.series[0];

  return (
    <aside className="w-92 shrink-0 flex flex-col h-full bg-[#f8fafc] border-l border-slate-300 select-none overflow-hidden text-xs">
      {/* 4-Tab Navigation */}
      <div className="grid grid-cols-4 gap-1 p-1 bg-[#eef2f6] border-b border-slate-300">
        <button
          type="button"
          onClick={() => onChangeTab('series')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 ${
            activeTab === 'series'
              ? 'bg-white text-slate-900 font-semibold shadow-2xs'
              : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <Sliders className="w-3.5 h-3.5 mb-0.5 text-blue-700" />
          <span>Serie & Stile</span>
        </button>

        <button
          type="button"
          onClick={() => onChangeTab('fitting')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 ${
            activeTab === 'fitting'
              ? 'bg-white text-slate-900 font-semibold shadow-2xs'
              : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <TrendingUp className="w-3.5 h-3.5 mb-0.5 text-red-600" />
          <span>Fit & Analisi</span>
        </button>

        <button
          type="button"
          onClick={() => onChangeTab('axes')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 ${
            activeTab === 'axes'
              ? 'bg-white text-slate-900 font-semibold shadow-2xs'
              : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <Axis3d className="w-3.5 h-3.5 mb-0.5 text-teal-700" />
          <span>Assi & Griglia</span>
        </button>

        <button
          type="button"
          onClick={() => onChangeTab('badpoints')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 ${
            activeTab === 'badpoints'
              ? 'bg-white text-slate-900 font-semibold shadow-2xs'
              : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <Target className="w-3.5 h-3.5 mb-0.5 text-amber-600" />
          <span>Bad Points</span>
        </button>
      </div>

      {/* Tab Content */}
      <div className="flex-1 overflow-y-auto p-3 space-y-4">
        {/* TAB 1: SERIE & STILE (PLOT DETAILS) */}
        {activeTab === 'series' && (
          <>
            {/* Add series box */}
            <div className="p-2.5 bg-white border border-slate-300 rounded space-y-2">
              <div className="font-semibold text-slate-800 flex items-center justify-between">
                <span>Aggiungi Serie a {graph.name}</span>
              </div>

              <div className="space-y-1.5">
                <div>
                  <label className="block text-[10.5px] text-slate-600">Foglio</label>
                  <select
                    value={selectedAddSheet?.id || ''}
                    onChange={(e) => setAddSheetId(e.target.value)}
                    className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded text-slate-800"
                  >
                    {sheets.map((s) => (
                      <option key={s.id} value={s.id}>
                        {s.name} ({s.columns.length} col)
                      </option>
                    ))}
                  </select>
                </div>

                {selectedAddSheet && (
                  <div className="grid grid-cols-2 gap-1.5">
                    <div>
                      <label className="block text-[10.5px] text-slate-600">X</label>
                      <select
                        value={addXColId}
                        onChange={(e) => setAddXColId(e.target.value)}
                        className="w-full px-1.5 py-1 text-xs font-mono bg-white border border-slate-300 rounded"
                      >
                        {selectedAddSheet.columns.map((c) => (
                          <option key={c.id} value={c.id}>
                            {c.shortName}({c.role}): {c.longName}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10.5px] text-slate-600">Y</label>
                      <select
                        value={addYColId}
                        onChange={(e) => setAddYColId(e.target.value)}
                        className="w-full px-1.5 py-1 text-xs font-mono bg-white border border-slate-300 rounded"
                      >
                        {selectedAddSheet.columns.map((c) => (
                          <option key={c.id} value={c.id}>
                            {c.shortName}({c.role}): {c.longName}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>
                )}

                <div className="grid grid-cols-2 gap-1.5">
                  <div>
                    <label className="block text-[10.5px] text-slate-600">Tipo Grafico</label>
                    <select
                      value={addPlotType}
                      onChange={(e) => setAddPlotType(e.target.value as PlotType)}
                      className="w-full px-1.5 py-1 text-xs bg-white border border-slate-300 rounded"
                    >
                      <option value="line_symbol">Linea + Simboli</option>
                      <option value="line">Solo Linea</option>
                      <option value="scatter">Solo Punti (Scatter)</option>
                      <option value="spline">Spline Smussata</option>
                      <option value="step">Gradino (Step)</option>
                      <option value="area">Area Sottesa</option>
                      <option value="bar">Colonne / Barre</option>
                      <option value="stem">Impulsi (Stem)</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-[10.5px] text-slate-600">Connessione</label>
                    <select
                      value={addConnection}
                      onChange={(e) => setAddConnection(e.target.value as ConnectionType)}
                      className="w-full px-1.5 py-1 text-xs bg-white border border-slate-300 rounded"
                    >
                      <option value="linear">Lineare</option>
                      <option value="spline">Spline Cubica</option>
                      <option value="step">Step (Gradino)</option>
                      <option value="none">Nessuna</option>
                    </select>
                  </div>
                </div>

                <button
                  type="button"
                  onClick={() => {
                    if (!selectedAddSheet || !addXColId || !addYColId) return;
                    const nextColor =
                      SCIENTIFIC_PALETTE[graph.series.length % SCIENTIFIC_PALETTE.length];
                    onAddSeriesToGraph({
                      sheetId: selectedAddSheet.id,
                      xColId: addXColId,
                      yColId: addYColId,
                      plotType: addPlotType,
                      connection: addConnection,
                      color: nextColor,
                    });
                  }}
                  className="w-full mt-1 flex items-center justify-center gap-1.5 py-1.5 px-3 text-xs font-semibold text-white bg-blue-700 rounded hover:bg-blue-800 transition-colors shadow-2xs"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>Aggiungi al Grafico</span>
                </button>
              </div>
            </div>

            {/* Active series list with detailed customization */}
            <div className="space-y-2.5">
              <div className="font-semibold text-slate-800 flex items-center justify-between">
                <span>Curve Tracciate ({graph.series.length})</span>
                <span className="text-[10.5px] text-slate-500 font-normal">
                  Clicca per selezionare
                </span>
              </div>

              {graph.series.map((s) => {
                const isSelected = s.id === (selectedSeriesId || activeSeriesObj?.id);
                const sh = sheets.find((item) => item.id === s.sheetId);
                const yc = sh?.columns.find((c) => c.id === s.yColId);
                const xc = sh?.columns.find((c) => c.id === s.xColId);

                return (
                  <div
                    key={s.id}
                    onClick={() => onSelectSeries(s.id)}
                    className={`p-2.5 border rounded transition-colors cursor-pointer bg-white ${
                      isSelected
                        ? 'border-blue-500 ring-1 ring-blue-500 bg-blue-50/20'
                        : 'border-slate-300 hover:border-slate-400'
                    }`}
                  >
                    <div className="flex items-center justify-between gap-1 mb-2">
                      <div className="flex items-center gap-1.5 min-w-0">
                        <input
                          type="color"
                          value={s.color}
                          onChange={(e) => handleUpdateSeries(s.id, { color: e.target.value })}
                          className="w-5 h-5 rounded border border-slate-300 cursor-pointer shrink-0"
                          title="Cambia colore"
                        />
                        <span className="font-semibold text-slate-900 truncate">
                          {yc ? `${yc.shortName}(Y): ${yc.longName}` : s.label}
                        </span>
                      </div>

                      <div className="flex items-center gap-1 shrink-0">
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            handleUpdateSeries(s.id, { visible: !s.visible });
                          }}
                          className="p-1 text-slate-500 hover:text-slate-800"
                          title={s.visible ? 'Nascondi curva' : 'Mostra curva'}
                        >
                          {s.visible ? <Eye className="w-3.5 h-3.5" /> : <EyeOff className="w-3.5 h-3.5 text-slate-400" />}
                        </button>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            handleRemoveSeries(s.id);
                          }}
                          className="p-1 text-slate-400 hover:text-red-700"
                          title="Rimuovi curva dal grafico"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>

                    {/* Plot Type & Representation Selector */}
                    <div className="grid grid-cols-2 gap-1.5 mb-2">
                      <div>
                        <label className="block text-[10px] text-slate-500 mb-0.5">Rappresentazione</label>
                        <select
                          value={s.plotType}
                          onChange={(e) => {
                            const pt = e.target.value as PlotType;
                            handleUpdateSeries(s.id, {
                              plotType: pt,
                              connection:
                                pt === 'spline'
                                  ? 'spline'
                                  : pt === 'step'
                                  ? 'step'
                                  : pt === 'scatter' || pt === 'bar' || pt === 'stem'
                                  ? 'none'
                                  : s.connection === 'none'
                                  ? 'linear'
                                  : s.connection,
                            });
                          }}
                          className="w-full px-1.5 py-1 text-[11px] bg-slate-50 border border-slate-300 rounded font-medium"
                        >
                          <option value="line_symbol">Linea + Simboli (Entrambi)</option>
                          <option value="line">Solo Linea</option>
                          <option value="scatter">Solo Simboli (Scatter)</option>
                          <option value="spline">Spline Smussata</option>
                          <option value="step">Gradino (Step)</option>
                          <option value="area">Area Sottesa</option>
                          <option value="bar">Colonne / Barre</option>
                          <option value="stem">Impulsi (Stem)</option>
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] text-slate-500 mb-0.5">Connessione</label>
                        <select
                          value={s.connection}
                          onChange={(e) =>
                            handleUpdateSeries(s.id, {
                              connection: e.target.value as ConnectionType,
                            })
                          }
                          className="w-full px-1.5 py-1 text-[11px] bg-slate-50 border border-slate-300 rounded"
                        >
                          <option value="linear">Lineare (Spezzata)</option>
                          <option value="spline">Spline Cubica</option>
                          <option value="step">Step (Gradino)</option>
                          <option value="none">Nessuna</option>
                        </select>
                      </div>
                    </div>

                    {/* Line customization (Width & Style) */}
                    {(s.plotType === 'line' || s.plotType === 'line_symbol' || s.plotType === 'area') && (
                      <div className="grid grid-cols-2 gap-1.5 mb-2 p-1.5 bg-slate-50 rounded border border-slate-200">
                        <div>
                          <label className="block text-[10px] text-slate-600 mb-0.5">
                            Spessore Linea ({s.lineWidth} pt)
                          </label>
                          <input
                            type="range"
                            min={0.5}
                            max={8}
                            step={0.5}
                            value={s.lineWidth}
                            onChange={(e) =>
                              handleUpdateSeries(s.id, { lineWidth: Number(e.target.value) })
                            }
                            className="w-full accent-blue-700"
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] text-slate-600 mb-0.5">Stile Linea</label>
                          <select
                            value={s.lineStyle}
                            onChange={(e) =>
                              handleUpdateSeries(s.id, { lineStyle: e.target.value as LineStyle })
                            }
                            className="w-full px-1.5 py-0.5 text-[11px] bg-white border border-slate-300 rounded"
                          >
                            <option value="solid">Continua</option>
                            <option value="dashed">Tratteggiata</option>
                            <option value="dotted">Punteggiata</option>
                            <option value="dash_dot">Tratto-Punto</option>
                          </select>
                        </div>
                      </div>
                    )}

                    {/* Symbol / Marker customization */}
                    {(s.plotType === 'scatter' || s.plotType === 'line_symbol') && (
                      <div className="grid grid-cols-3 gap-1.5 mb-2 p-1.5 bg-slate-50 rounded border border-slate-200">
                        <div>
                          <label className="block text-[10px] text-slate-600 mb-0.5">Simbolo</label>
                          <select
                            value={s.markerShape}
                            onChange={(e) =>
                              handleUpdateSeries(s.id, {
                                markerShape: e.target.value as MarkerShape,
                              })
                            }
                            className="w-full px-1 py-0.5 text-[11px] bg-white border border-slate-300 rounded"
                          >
                            <option value="circle">Cerchio ●</option>
                            <option value="square">Quadrato ■</option>
                            <option value="triangle_up">Triangolo ▲</option>
                            <option value="triangle_down">Triangolo ▼</option>
                            <option value="diamond">Rombo ◆</option>
                            <option value="cross">Croce ✕</option>
                            <option value="star">Stella ★</option>
                            <option value="hexagon">Esagono ⬡</option>
                            <option value="none">Nessuno</option>
                          </select>
                        </div>

                        <div>
                          <label className="block text-[10px] text-slate-600 mb-0.5">
                            Dim. ({s.markerSize} pt)
                          </label>
                          <input
                            type="range"
                            min={1}
                            max={16}
                            step={0.5}
                            value={s.markerSize}
                            onChange={(e) =>
                              handleUpdateSeries(s.id, { markerSize: Number(e.target.value) })
                            }
                            className="w-full accent-blue-700"
                          />
                        </div>

                        <div>
                          <label className="block text-[10px] text-slate-600 mb-0.5">Riempimento</label>
                          <input
                            type="color"
                            value={s.symbolFill || s.color}
                            onChange={(e) =>
                              handleUpdateSeries(s.id, { symbolFill: e.target.value })
                            }
                            className="w-full h-6 rounded border border-slate-300 cursor-pointer"
                          />
                        </div>
                      </div>
                    )}

                    {/* Error bars & Drop lines toggles */}
                    <div className="flex items-center justify-between gap-2 pt-1 border-t border-slate-100">
                      <label className="flex items-center gap-1.5 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={s.showErrorBars}
                          onChange={(e) =>
                            handleUpdateSeries(s.id, { showErrorBars: e.target.checked })
                          }
                          className="rounded border-slate-300 text-blue-700"
                        />
                        <span className="text-[11px] text-slate-700">Barre di Errore (±σ)</span>
                      </label>

                      <label className="flex items-center gap-1.5 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={s.dropLines}
                          onChange={(e) =>
                            handleUpdateSeries(s.id, { dropLines: e.target.checked })
                          }
                          className="rounded border-slate-300 text-blue-700"
                        />
                        <span className="text-[11px] text-slate-700">Linee Cadenti (X)</span>
                      </label>
                    </div>
                  </div>
                );
              })}
            </div>
          </>
        )}

        {/* TAB 2: CURVE FITTING & ANALISI SCIENTIFICA */}
        {activeTab === 'fitting' && (
          <>
            <form
              onSubmit={handleExecuteFit}
              className="p-2.5 bg-white border border-slate-300 rounded space-y-2.5"
            >
              <div className="font-semibold text-slate-900 flex items-center justify-between">
                <span>Curve Fitting su Intervallo [C, D]</span>
                <span className="text-[10px] font-mono text-red-600 font-bold">Origin Fit</span>
              </div>

              {/* 1. Target Curve */}
              <div>
                <label className="block text-[10.5px] font-medium text-slate-700 mb-0.5">
                  Curva Sperimentale [A, B]
                </label>
                <select
                  value={fitSeriesId}
                  onChange={(e) => setFitSeriesId(e.target.value)}
                  className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                >
                  {graph.series.map((s) => {
                    const sh = sheets.find((item) => item.id === s.sheetId);
                    const yc = sh?.columns.find((c) => c.id === s.yColId);
                    return (
                      <option key={s.id} value={s.id}>
                        {yc ? `${yc.shortName}(Y): ${yc.longName}` : s.label} ({sh?.name})
                      </option>
                    );
                  })}
                </select>
                {sourceSeriesStats && (
                  <div className="mt-1 flex items-center justify-between text-[10.5px] font-mono text-slate-500">
                    <span>Validi: {sourceSeriesStats.count}</span>
                    <span>
                      [A, B]: [{formatSciNumber(sourceSeriesStats.minA, 3)},{' '}
                      {formatSciNumber(sourceSeriesStats.maxB, 3)}]
                    </span>
                  </div>
                )}
              </div>

              {/* 2. Model Selection: Linear, Polynomial, Gaussian Peak, Exponential, Spline */}
              <div>
                <label className="block text-[10.5px] font-medium text-slate-700 mb-1">
                  Modello Matematico
                </label>
                <div className="grid grid-cols-2 gap-1 text-[11px]">
                  <button
                    type="button"
                    onClick={() => setFitModel('linear')}
                    className={`py-1 px-1.5 rounded border transition-colors ${
                      fitModel === 'linear'
                        ? 'bg-blue-700 text-white font-semibold border-blue-700'
                        : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                    }`}
                  >
                    Lineare (a + bx)
                  </button>
                  <button
                    type="button"
                    onClick={() => setFitModel('polynomial')}
                    className={`py-1 px-1.5 rounded border transition-colors ${
                      fitModel === 'polynomial'
                        ? 'bg-blue-700 text-white font-semibold border-blue-700'
                        : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                    }`}
                  >
                    Polinomiale (g={fitPolyDegree})
                  </button>
                  <button
                    type="button"
                    onClick={() => setFitModel('gaussian')}
                    className={`py-1 px-1.5 rounded border transition-colors ${
                      fitModel === 'gaussian'
                        ? 'bg-blue-700 text-white font-semibold border-blue-700'
                        : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                    }`}
                  >
                    Picco Gaussiano
                  </button>
                  <button
                    type="button"
                    onClick={() => setFitModel('exponential')}
                    className={`py-1 px-1.5 rounded border transition-colors ${
                      fitModel === 'exponential'
                        ? 'bg-blue-700 text-white font-semibold border-blue-700'
                        : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                    }`}
                  >
                    Esponenziale (A e^kx)
                  </button>
                  <button
                    type="button"
                    onClick={() => setFitModel('spline')}
                    className={`col-span-2 py-1 px-1.5 rounded border transition-colors ${
                      fitModel === 'spline'
                        ? 'bg-blue-700 text-white font-semibold border-blue-700'
                        : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                    }`}
                  >
                    Spline Cubica Naturale C² (Ricampionamento esatto)
                  </button>
                </div>
              </div>

              {fitModel === 'polynomial' && (
                <div className="flex items-center justify-between gap-2 p-1.5 bg-slate-50 border border-slate-200 rounded">
                  <span className="text-[11px] font-medium text-slate-700">Grado Polinomio (1-8):</span>
                  <input
                    type="number"
                    min={1}
                    max={8}
                    value={fitPolyDegree}
                    onChange={(e) => setFitPolyDegree(Number(e.target.value))}
                    className="w-16 px-1.5 py-0.5 text-right font-mono text-xs border border-slate-300 rounded bg-white"
                  />
                </div>
              )}

              {/* 3. Custom Evaluation Range [C, D] with n points */}
              <div className="p-2 bg-slate-50 border border-slate-200 rounded space-y-1.5">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] font-semibold text-slate-800">
                    Intervallo [C, D] & Punti (n)
                  </span>
                  {sourceSeriesStats && (
                    <button
                      type="button"
                      onClick={() => {
                        setFitRangeC(String(sourceSeriesStats.minA));
                        setFitRangeD(String(sourceSeriesStats.maxB));
                      }}
                      className="text-[10px] text-blue-700 hover:underline"
                    >
                      Copia [A, B]
                    </button>
                  )}
                </div>

                <div className="grid grid-cols-3 gap-1.5">
                  <div>
                    <label className="block text-[10px] font-mono text-slate-500">Inizio C</label>
                    <input
                      type="text"
                      value={fitRangeC}
                      onChange={(e) => setFitRangeC(e.target.value)}
                      className="w-full px-1.5 py-0.5 text-right font-mono text-xs border border-slate-300 rounded bg-white"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-mono text-slate-500">Fine D</label>
                    <input
                      type="text"
                      value={fitRangeD}
                      onChange={(e) => setFitRangeD(e.target.value)}
                      className="w-full px-1.5 py-0.5 text-right font-mono text-xs border border-slate-300 rounded bg-white"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-mono text-slate-500">N. Punti (n)</label>
                    <input
                      type="number"
                      min={2}
                      max={5000}
                      value={fitNumPointsN}
                      onChange={(e) => setFitNumPointsN(e.target.value)}
                      className="w-full px-1.5 py-0.5 text-right font-mono text-xs border border-slate-300 rounded bg-white"
                    />
                  </div>
                </div>
              </div>

              {/* Fit styling */}
              <div className="grid grid-cols-3 gap-1.5 items-center">
                <div>
                  <label className="block text-[10px] text-slate-500">Colore</label>
                  <input
                    type="color"
                    value={fitColor}
                    onChange={(e) => setFitColor(e.target.value)}
                    className="w-full h-6 rounded border border-slate-300 cursor-pointer"
                  />
                </div>
                <div>
                  <label className="block text-[10px] text-slate-500">Spessore ({fitLineWidth}pt)</label>
                  <input
                    type="range"
                    min={1}
                    max={6}
                    step={0.5}
                    value={fitLineWidth}
                    onChange={(e) => setFitLineWidth(Number(e.target.value))}
                    className="w-full accent-red-600"
                  />
                </div>
                <div>
                  <label className="block text-[10px] text-slate-500">Stile</label>
                  <select
                    value={fitLineStyle}
                    onChange={(e) => setFitLineStyle(e.target.value as LineStyle)}
                    className="w-full px-1 py-0.5 text-xs bg-white border border-slate-300 rounded"
                  >
                    <option value="solid">Continua</option>
                    <option value="dashed">Tratteggiata</option>
                    <option value="dotted">Punteggiata</option>
                  </select>
                </div>
              </div>

              <label className="flex items-center gap-1.5 cursor-pointer">
                <input
                  type="checkbox"
                  checked={fitCreateSheet}
                  onChange={(e) => setFitCreateSheet(e.target.checked)}
                  className="rounded border-slate-300 text-blue-700"
                />
                <span className="text-[11px] text-slate-700">
                  Genera nuovo foglio tabella con gli <strong>{fitNumPointsN}</strong> punti in [C, D]
                </span>
              </label>

              <button
                type="submit"
                disabled={!graph.series.length}
                className="w-full py-2 px-3 text-xs font-bold text-white bg-red-600 rounded hover:bg-red-700 disabled:opacity-40 transition-colors shadow-2xs"
              >
                Esegui Fit & Calcola {fitNumPointsN} Punti in [{fitRangeC}, {fitRangeD}]
              </button>
            </form>

            {/* Quick Signal Processing & Calculus Tools */}
            <div className="p-2.5 bg-white border border-slate-300 rounded space-y-2">
              <div className="font-semibold text-slate-800 text-[11px] flex items-center justify-between">
                <span>Filtro Segnale, Smoothing & Calcolo</span>
                <Sparkles className="w-3.5 h-3.5 text-emerald-600" />
              </div>

              {onOpenFilterSmoothModal && (
                <button
                  type="button"
                  onClick={onOpenFilterSmoothModal}
                  className="w-full py-1.5 px-2.5 text-xs font-bold text-white bg-emerald-700 rounded hover:bg-emerald-800 transition-colors flex items-center justify-center gap-1.5 shadow-2xs"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  <span>Apri Suite Filtro & Smoothing Segnale...</span>
                </button>
              )}

              {activeSeriesObj && (
                <div className="space-y-1.5 pt-1 border-t border-slate-100">
                  <div className="flex items-center justify-between gap-1.5">
                    <button
                      type="button"
                      onClick={() =>
                        onApplySmoothing(
                          activeSeriesObj.sheetId,
                          activeSeriesObj.yColId,
                          smoothWindow
                        )
                      }
                      className="flex-1 py-1 px-2 text-[10.5px] font-medium text-slate-800 bg-slate-100 border border-slate-300 rounded hover:bg-slate-200"
                    >
                      Smooth Rapido (W={smoothWindow})
                    </button>
                    <input
                      type="number"
                      min={3}
                      max={25}
                      step={2}
                      value={smoothWindow}
                      onChange={(e) => setSmoothWindow(Number(e.target.value))}
                      className="w-12 px-1 py-0.5 text-right font-mono text-[10.5px] border border-slate-300 rounded"
                      title="Dimensione finestra di media mobile"
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-1.5">
                    <button
                      type="button"
                      onClick={() =>
                        onApplyDerivative(
                          activeSeriesObj.sheetId,
                          activeSeriesObj.xColId,
                          activeSeriesObj.yColId
                        )
                      }
                      className="py-1 px-2 text-[10.5px] font-medium text-slate-800 bg-slate-100 border border-slate-300 rounded hover:bg-slate-200"
                    >
                      Derivata dy/dx
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        onApplyIntegration(
                          activeSeriesObj.sheetId,
                          activeSeriesObj.xColId,
                          activeSeriesObj.yColId
                        )
                      }
                      className="py-1 px-2 text-[10.5px] font-medium text-slate-800 bg-slate-100 border border-slate-300 rounded hover:bg-slate-200"
                    >
                      Integrale ∫ y dx
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Active Fits Results Box */}
            <div className="space-y-2">
              <div className="font-semibold text-slate-800">
                Fit Calcolati sul Grafico ({graph.fits.length})
              </div>

              {graph.fits.map((fit) => (
                <div key={fit.id} className="p-2.5 bg-white border border-slate-300 rounded space-y-2">
                  <div className="flex items-center justify-between gap-1">
                    <span className="font-semibold text-slate-900 truncate">{fit.name}</span>
                    <button
                      type="button"
                      onClick={() => onDeleteFit(fit.id)}
                      className="p-0.5 text-slate-400 hover:text-red-700"
                      title="Elimina fit"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  <div className="p-1.5 bg-slate-50 border border-slate-200 rounded font-mono text-[10.5px] text-red-700 break-all">
                    {fit.equationText}
                  </div>

                  <div className="grid grid-cols-2 gap-1 font-mono text-[10.5px]">
                    <div className="p-1 bg-emerald-50 border border-emerald-200 rounded text-emerald-900 flex items-center justify-between">
                      <span>R²:</span>
                      <strong>{fit.rSquared.toFixed(5)}</strong>
                    </div>
                    <div className="p-1 bg-slate-50 border border-slate-200 rounded text-slate-800 flex items-center justify-between">
                      <span>RMSE:</span>
                      <strong>{formatSciNumber(fit.rmse, 3)}</strong>
                    </div>
                  </div>

                  {fit.outputSheetId && (
                    <button
                      type="button"
                      onClick={() => onOpenFitSheet(fit.outputSheetId!)}
                      className="w-full flex items-center justify-center gap-1 py-1 px-2 text-[11px] font-medium text-blue-800 bg-blue-50 border border-blue-200 rounded hover:bg-blue-100"
                    >
                      <FileSpreadsheet className="w-3.5 h-3.5" />
                      <span>Apri Tabella con {fit.numPointsN} Punti</span>
                    </button>
                  )}
                </div>
              ))}
            </div>
          </>
        )}

        {/* TAB 3: ASSI & SCALA (LINEAR / LOG10) */}
        {activeTab === 'axes' && (
          <div className="space-y-3">
            {/* Titles */}
            <div className="p-2.5 bg-white border border-slate-300 rounded space-y-2">
              <div className="font-semibold text-slate-800">Titoli del Grafico</div>
              <div>
                <label className="block text-[10.5px] text-slate-600 mb-0.5">Titolo</label>
                <input
                  type="text"
                  value={graph.title}
                  onChange={(e) => onUpdateGraph({ ...graph, title: e.target.value })}
                  className="w-full px-2 py-1 text-xs border border-slate-300 rounded"
                />
              </div>
              <div>
                <label className="block text-[10.5px] text-slate-600 mb-0.5">Sottotitolo</label>
                <input
                  type="text"
                  value={graph.subtitle}
                  onChange={(e) => onUpdateGraph({ ...graph, subtitle: e.target.value })}
                  className="w-full px-2 py-1 text-xs border border-slate-300 rounded"
                />
              </div>
            </div>

            {/* X Axis & Scale Type */}
            <div className="p-2.5 bg-white border border-slate-300 rounded space-y-2">
              <div className="font-semibold text-slate-800 flex items-center justify-between">
                <span>Asse X</span>
                <select
                  value={graph.xAxis.scaleType || 'linear'}
                  onChange={(e) =>
                    onUpdateGraph({
                      ...graph,
                      xAxis: { ...graph.xAxis, scaleType: e.target.value as 'linear' | 'log10' },
                    })
                  }
                  className="px-1.5 py-0.5 text-[10.5px] font-semibold bg-slate-100 border border-slate-300 rounded"
                >
                  <option value="linear">Scala Lineare</option>
                  <option value="log10">Scala Logaritmica (Log10)</option>
                </select>
              </div>

              <label className="flex items-center gap-1.5 cursor-pointer">
                <input
                  type="checkbox"
                  checked={graph.xAxis.autoLabel}
                  onChange={(e) =>
                    onUpdateGraph({
                      ...graph,
                      xAxis: { ...graph.xAxis, autoLabel: e.target.checked },
                    })
                  }
                  className="rounded border-slate-300 text-blue-700"
                />
                <span className="text-[11px] text-slate-700">Etichetta automatica da colonne X</span>
              </label>

              {!graph.xAxis.autoLabel && (
                <input
                  type="text"
                  value={graph.xAxis.label}
                  onChange={(e) =>
                    onUpdateGraph({
                      ...graph,
                      xAxis: { ...graph.xAxis, label: e.target.value },
                    })
                  }
                  placeholder="Etichetta asse X..."
                  className="w-full px-2 py-1 text-xs border border-slate-300 rounded"
                />
              )}

              <div className="grid grid-cols-2 gap-1.5">
                <div>
                  <label className="block text-[10px] text-slate-500">X Min (Auto)</label>
                  <input
                    type="text"
                    value={graph.xAxis.min === 'auto' ? '' : String(graph.xAxis.min)}
                    placeholder="Auto"
                    onChange={(e) => {
                      const v = e.target.value.trim();
                      onUpdateGraph({
                        ...graph,
                        xAxis: { ...graph.xAxis, min: v === '' ? 'auto' : Number(v) },
                      });
                    }}
                    className="w-full px-1.5 py-0.5 font-mono text-xs border border-slate-300 rounded"
                  />
                </div>
                <div>
                  <label className="block text-[10px] text-slate-500">X Max (Auto)</label>
                  <input
                    type="text"
                    value={graph.xAxis.max === 'auto' ? '' : String(graph.xAxis.max)}
                    placeholder="Auto"
                    onChange={(e) => {
                      const v = e.target.value.trim();
                      onUpdateGraph({
                        ...graph,
                        xAxis: { ...graph.xAxis, max: v === '' ? 'auto' : Number(v) },
                      });
                    }}
                    className="w-full px-1.5 py-0.5 font-mono text-xs border border-slate-300 rounded"
                  />
                </div>
              </div>
            </div>

            {/* Y Axis & Scale Type */}
            <div className="p-2.5 bg-white border border-slate-300 rounded space-y-2">
              <div className="font-semibold text-slate-800 flex items-center justify-between">
                <span>Asse Y</span>
                <select
                  value={graph.yAxis.scaleType || 'linear'}
                  onChange={(e) =>
                    onUpdateGraph({
                      ...graph,
                      yAxis: { ...graph.yAxis, scaleType: e.target.value as 'linear' | 'log10' },
                    })
                  }
                  className="px-1.5 py-0.5 text-[10.5px] font-semibold bg-slate-100 border border-slate-300 rounded"
                >
                  <option value="linear">Scala Lineare</option>
                  <option value="log10">Scala Logaritmica (Log10)</option>
                </select>
              </div>

              <label className="flex items-center gap-1.5 cursor-pointer">
                <input
                  type="checkbox"
                  checked={graph.yAxis.autoLabel}
                  onChange={(e) =>
                    onUpdateGraph({
                      ...graph,
                      yAxis: { ...graph.yAxis, autoLabel: e.target.checked },
                    })
                  }
                  className="rounded border-slate-300 text-blue-700"
                />
                <span className="text-[11px] text-slate-700">Etichetta automatica da colonne Y</span>
              </label>

              {!graph.yAxis.autoLabel && (
                <input
                  type="text"
                  value={graph.yAxis.label}
                  onChange={(e) =>
                    onUpdateGraph({
                      ...graph,
                      yAxis: { ...graph.yAxis, label: e.target.value },
                    })
                  }
                  placeholder="Etichetta asse Y..."
                  className="w-full px-2 py-1 text-xs border border-slate-300 rounded"
                />
              )}

              <div className="grid grid-cols-2 gap-1.5">
                <div>
                  <label className="block text-[10px] text-slate-500">Y Min (Auto)</label>
                  <input
                    type="text"
                    value={graph.yAxis.min === 'auto' ? '' : String(graph.yAxis.min)}
                    placeholder="Auto"
                    onChange={(e) => {
                      const v = e.target.value.trim();
                      onUpdateGraph({
                        ...graph,
                        yAxis: { ...graph.yAxis, min: v === '' ? 'auto' : Number(v) },
                      });
                    }}
                    className="w-full px-1.5 py-0.5 font-mono text-xs border border-slate-300 rounded"
                  />
                </div>
                <div>
                  <label className="block text-[10px] text-slate-500">Y Max (Auto)</label>
                  <input
                    type="text"
                    value={graph.yAxis.max === 'auto' ? '' : String(graph.yAxis.max)}
                    placeholder="Auto"
                    onChange={(e) => {
                      const v = e.target.value.trim();
                      onUpdateGraph({
                        ...graph,
                        yAxis: { ...graph.yAxis, max: v === '' ? 'auto' : Number(v) },
                      });
                    }}
                    className="w-full px-1.5 py-0.5 font-mono text-xs border border-slate-300 rounded"
                  />
                </div>
              </div>
            </div>

            {/* Grids and Ticks */}
            <div className="p-2.5 bg-white border border-slate-300 rounded space-y-2">
              <div className="font-semibold text-slate-800">Griglia e Ticchette (Origin Ticks)</div>
              <div className="flex items-center justify-between">
                <span className="text-[11px] text-slate-700">Direzione Ticchette</span>
                <select
                  value={graph.xAxis.tickDirection || 'both'}
                  onChange={(e) => {
                    const dir = e.target.value as 'in' | 'out' | 'both';
                    onUpdateGraph({
                      ...graph,
                      xAxis: { ...graph.xAxis, tickDirection: dir },
                      yAxis: { ...graph.yAxis, tickDirection: dir },
                    });
                  }}
                  className="px-1.5 py-0.5 text-xs bg-slate-50 border border-slate-300 rounded"
                >
                  <option value="both">Entrambi (In & Out)</option>
                  <option value="in">Solo Interni (Inward)</option>
                  <option value="out">Solo Esterni (Outward)</option>
                </select>
              </div>

              <label className="flex items-center gap-1.5 cursor-pointer">
                <input
                  type="checkbox"
                  checked={graph.xAxis.showMajorGrid}
                  onChange={(e) =>
                    onUpdateGraph({
                      ...graph,
                      xAxis: { ...graph.xAxis, showMajorGrid: e.target.checked },
                      yAxis: { ...graph.yAxis, showMajorGrid: e.target.checked },
                    })
                  }
                  className="rounded border-slate-300 text-blue-700"
                />
                <span className="text-[11px] text-slate-700">Mostra griglia principale</span>
              </label>

              <label className="flex items-center gap-1.5 cursor-pointer">
                <input
                  type="checkbox"
                  checked={graph.showLegend}
                  onChange={(e) => onUpdateGraph({ ...graph, showLegend: e.target.checked })}
                  className="rounded border-slate-300 text-blue-700"
                />
                <span className="text-[11px] text-slate-700">Mostra Legenda</span>
              </label>
            </div>
          </div>
        )}

        {/* TAB 4: BAD POINTS & MASKING */}
        {activeTab === 'badpoints' && (
          <div className="space-y-3">
            <div className="p-2.5 bg-red-50 border border-red-200 rounded space-y-2">
              <div className="font-semibold text-red-900">Remove Bad Data Points</div>
              <p className="text-[11px] text-red-800 leading-relaxed">
                Come in OriginLab, attiva il mirino e clicca direttamente sui punti errati sul grafico per rimuoverli o mascherarli istantaneamente.
              </p>

              <button
                type="button"
                onClick={() =>
                  onChangeToolMode(toolMode === 'remove_bad' ? 'pointer' : 'remove_bad')
                }
                className={`w-full py-2 px-3 text-xs font-bold rounded flex items-center justify-center gap-1.5 transition-colors ${
                  toolMode === 'remove_bad'
                    ? 'bg-slate-900 text-white'
                    : 'bg-red-600 text-white hover:bg-red-700 shadow-2xs'
                }`}
              >
                <Target className="w-4 h-4" />
                <span>
                  {toolMode === 'remove_bad' ? 'Disattiva Mirino' : 'Attiva Mirino Remove Bad Points'}
                </span>
              </button>
            </div>

            <div className="p-2.5 bg-white border border-slate-300 rounded space-y-2">
              <div className="font-semibold text-slate-800">Azione di Rimozione</div>
              <label className="flex items-start gap-2 cursor-pointer">
                <input
                  type="radio"
                  name="badmode"
                  checked={badPointRemovalMode === 'cell-only'}
                  onChange={() => onChangeBadPointRemovalMode('cell-only')}
                  className="mt-0.5 text-blue-700"
                />
                <div>
                  <div className="font-medium text-slate-800">Svuota solo il valore Y cliccato</div>
                  <div className="text-[10px] text-slate-500">
                    Consigliato: non tocca le altre curve presenti nella stessa riga.
                  </div>
                </div>
              </label>

              <label className="flex items-start gap-2 cursor-pointer pt-1">
                <input
                  type="radio"
                  name="badmode"
                  checked={badPointRemovalMode === 'entire-row'}
                  onChange={() => onChangeBadPointRemovalMode('entire-row')}
                  className="mt-0.5 text-blue-700"
                />
                <div>
                  <div className="font-medium text-slate-800">Elimina l&apos;intera riga della tabella</div>
                  <div className="text-[10px] text-slate-500">Cancella X e tutte le Y di quella riga.</div>
                </div>
              </label>
            </div>

            {/* Removed points log */}
            <div className="space-y-1.5">
              <div className="font-semibold text-slate-800">
                Punti Rimossi ({removedPointsLog.length})
              </div>
              {removedPointsLog.length === 0 ? (
                <div className="p-3 text-center text-slate-500 border border-dashed border-slate-300 rounded bg-white">
                  Nessun punto rimosso finora.
                </div>
              ) : (
                removedPointsLog.map((rec) => (
                  <div
                    key={rec.id}
                    className="p-2 bg-white border border-slate-300 rounded flex items-center justify-between gap-1.5"
                  >
                    <div className="min-w-0">
                      <div className="font-semibold text-slate-800 truncate">
                        {rec.seriesLabel} · Riga #{Number(rec.rowIndex) + 1}
                      </div>
                      <div className="font-mono text-[10px] text-slate-500">
                        X = {formatSciNumber(rec.xValue, 3)} · Y = {formatSciNumber(rec.yValue, 3)}
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => onRestoreRemovedPoint(rec.id)}
                      className="px-2 py-1 text-[10.5px] font-medium text-blue-800 bg-blue-50 border border-blue-200 rounded hover:bg-blue-100 shrink-0"
                    >
                      Ripristina
                    </button>
                  </div>
                ))
              )}
            </div>
          </div>
        )}
      </div>
    </aside>
  );
};
