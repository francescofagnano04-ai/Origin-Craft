import React, { useMemo, useRef, useState, useEffect } from 'react';
import Plotly from 'plotly.js-dist-min';
import {
  MousePointer,
  Crosshair,
  Target,
  ZoomIn,
  RotateCcw,
  Download,
  Sliders,
  Grid,
  EyeOff,
  Type,
  TrendingUp,
  Sparkles,
  Move,
  Maximize2,
  Minimize2,
  Plus,
  Cpu,
  Layers,
} from 'lucide-react';
import {
  ConnectionType,
  FitModelType,
  GraphAnnotation,
  GraphWindow,
  HoveredDataPoint,
  LineStyle,
  MarkerShape,
  PlotType,
  ToolMode,
  WorkbookSheet,
} from '../types/origin';
import {
  formatCellForDisplay,
  formatSciNumber,
  resolveErrorAndLabelColumnsForY,
  resolveXColumnForY,
} from '../utils/mathEngine';
import { extractWebGLSeriesTypedArrays } from '../utils/binaryColumnStore';

interface GraphCanvasProps {
  graph: GraphWindow;
  allGraphs: GraphWindow[];
  activeGraphId: string;
  onSelectGraph: (id: string) => void;
  sheets: WorkbookSheet[];
  toolMode: ToolMode;
  onChangeToolMode: (mode: ToolMode) => void;
  onUpdateGraph: (updated: GraphWindow) => void;
  onRemoveBadPoint: (params: {
    sheetId: string;
    rowIndex: number;
    xColId: string;
    yColId: string;
    seriesLabel: string;
    xValue: number;
    yValue: number;
  }) => void;
  onRemoveBadPointsBatch: (
    points: {
      sheetId: string;
      rowIndex: number;
      xColId: string;
      yColId: string;
      seriesLabel: string;
      xValue: number;
      yValue: number;
    }[]
  ) => void;
  onMaskPoint: (sheetId: string, yColId: string, rowIndex: number) => void;
  onUndoLastRemovedPoint?: () => void;
  canUndoRemovedPoint?: boolean;
  onHoverPoint: (pt: HoveredDataPoint | null) => void;
  onOpenPlotDetails?: () => void;
  onOpenPlotDetailsModal?: () => void;
  onOpenFilterSmoothModal?: () => void;
  onQuickFitActiveSeries?: (model: FitModelType) => void;
  onSelectSeriesInInspector: (seriesId: string) => void;
  isMaximized?: boolean;
  onToggleMaximize?: () => void;
  hideGraphTabs?: boolean;
  onRenameGraph?: (graphId: string, newName: string) => void;
  onDeleteGraph?: (graphId: string) => void;
  onAddGraph?: () => void;
}

export const PLOT_TYPE_LABELS: { id: PlotType; label: string; short: string }[] = [
  { id: 'line_symbol', label: 'Linea + Simboli (ScatterGL)', short: 'L+S' },
  { id: 'line', label: 'Solo Linea (ScatterGL)', short: 'Linea' },
  { id: 'scatter', label: 'Solo Punti / Scatter (ScatterGL)', short: 'Punti' },
  { id: 'spline', label: 'Curva Spline Smussata (WebGL)', short: 'Spline' },
  { id: 'step', label: 'Grafico a Gradino (StepGL)', short: 'Step' },
  { id: 'area', label: 'Area Sottesa (AreaGL)', short: 'Area' },
  { id: 'bar', label: 'Colonne / Barre (Bar)', short: 'Barre' },
  { id: 'stem', label: 'Impulsi / Stem (DropGL)', short: 'Stem' },
];

function mapMarkerShapeToPlotlySymbol(shape: MarkerShape): string {
  switch (shape) {
    case 'circle':
      return 'circle';
    case 'square':
      return 'square';
    case 'triangle_up':
      return 'triangle-up';
    case 'triangle_down':
      return 'triangle-down';
    case 'diamond':
      return 'diamond';
    case 'cross':
      return 'cross';
    case 'star':
      return 'star';
    case 'hexagon':
      return 'hexagon';
    default:
      return 'circle';
  }
}

function mapLineStyleToPlotlyDash(style: LineStyle): 'solid' | 'dash' | 'dot' | 'dashdot' {
  switch (style) {
    case 'dashed':
      return 'dash';
    case 'dotted':
      return 'dot';
    case 'dash_dot':
      return 'dashdot';
    default:
      return 'solid';
  }
}

function hexToRgba(hex: string, alpha: number): string {
  const cleaned = hex.replace('#', '');
  if (cleaned.length !== 6) return `rgba(37, 99, 235, ${alpha})`;
  const r = parseInt(cleaned.slice(0, 2), 16);
  const g = parseInt(cleaned.slice(2, 4), 16);
  const b = parseInt(cleaned.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

/**
 * Evaluates Natural Cubic Spline on a Float64Array to generate a smooth dense Float64Array for scattergl
 */
function buildSmoothSplineFloat64(
  xValid: Float64Array,
  yValid: Float64Array
): { xs: Float64Array; ys: Float64Array } {
  const n = xValid.length;
  if (n < 3 || n > 2500) {
    return { xs: xValid, ys: yValid };
  }
  // Sort by X for spline evaluation
  const idx = Array.from({ length: n }, (_, i) => i).sort((a, b) => xValid[a] - xValid[b]);
  const x = new Float64Array(n);
  const y = new Float64Array(n);
  for (let i = 0; i < n; i++) {
    x[i] = xValid[idx[i]];
    y[i] = yValid[idx[i]];
  }

  const denseCount = Math.min(800, n * 6);
  const xs = new Float64Array(denseCount);
  const ys = new Float64Array(denseCount);
  const xMin = x[0];
  const xMax = x[n - 1];
  if (xMax <= xMin) return { xs: xValid, ys: yValid };

  // Catmull-Rom C1 smooth interpolation in Float64Array
  let seg = 0;
  for (let i = 0; i < denseCount; i++) {
    const xv = xMin + (i * (xMax - xMin)) / (denseCount - 1);
    xs[i] = xv;
    while (seg < n - 2 && xv > x[seg + 1]) {
      seg++;
    }
    const x0 = x[Math.max(0, seg - 1)];
    const y0 = y[Math.max(0, seg - 1)];
    const x1 = x[seg];
    const y1 = y[seg];
    const x2 = x[seg + 1];
    const y2 = y[seg + 1];
    const x3 = x[Math.min(n - 1, seg + 2)];
    const y3 = y[Math.min(n - 1, seg + 2)];

    const dt = x2 - x1;
    if (dt <= 1e-12) {
      ys[i] = y1;
      continue;
    }
    const u = Math.max(0, Math.min(1, (xv - x1) / dt));
    const m1 = ((y2 - y0) / Math.max(1e-12, x2 - x0)) * dt;
    const m2 = ((y3 - y1) / Math.max(1e-12, x3 - x1)) * dt;
    const u2 = u * u;
    const u3 = u2 * u;
    ys[i] =
      (2 * u3 - 3 * u2 + 1) * y1 +
      (u3 - 2 * u2 + u) * m1 +
      (-2 * u3 + 3 * u2) * y2 +
      (u3 - u2) * m2;
  }

  return { xs, ys };
}

export const GraphCanvas: React.FC<GraphCanvasProps> = ({
  graph,
  allGraphs,
  activeGraphId,
  onSelectGraph,
  sheets,
  toolMode,
  onChangeToolMode,
  onUpdateGraph,
  onRemoveBadPoint,
  onRemoveBadPointsBatch,
  onMaskPoint,
  onHoverPoint,
  onOpenPlotDetails,
  onOpenPlotDetailsModal,
  onOpenFilterSmoothModal,
  onQuickFitActiveSeries,
  onSelectSeriesInInspector,
  isMaximized,
  onToggleMaximize,
  hideGraphTabs,
  onRenameGraph,
  onDeleteGraph,
  onAddGraph,
}) => {
  const plotContainerRef = useRef<HTMLDivElement | null>(null);
  const isPlotInitializedRef = useRef<boolean>(false);

  const [hoveredPt, setHoveredPt] = useState<HoveredDataPoint | null>(null);
  const [renamingGraphId, setRenamingGraphId] = useState<string | null>(null);
  const [renameGraphValue, setRenameGraphValue] = useState<string>('');
  const [useHeatmapGL, setUseHeatmapGL] = useState<boolean>(false);

  // Draggable Fit Report Box overlay (rendered on top of WebGL canvas so it's 100% crisp & movable)
  const [draggingFitBox, setDraggingFitBox] = useState<{
    startMouseX: number;
    startMouseY: number;
    initialOffsetX: number;
    initialOffsetY: number;
  } | null>(null);

  // Latest callbacks ref so Plotly event listeners always invoke fresh handlers without re-binding WebGL context
  const callbacksRef = useRef({
    toolMode,
    graph,
    onChangeToolMode,
    onUpdateGraph,
    onRemoveBadPoint,
    onRemoveBadPointsBatch,
    onMaskPoint,
    onHoverPoint,
    onSelectSeriesInInspector,
  });
  callbacksRef.current = {
    toolMode,
    graph,
    onChangeToolMode,
    onUpdateGraph,
    onRemoveBadPoint,
    onRemoveBadPointsBatch,
    onMaskPoint,
    onHoverPoint,
    onSelectSeriesInInspector,
  };

  /**
   * RULE 1 & RULE 3: Extract contiguous Float64Array & Int32Array buffers directly for WebGL
   */
  const resolvedWebGLSeries = useMemo(() => {
    return graph.series
      .filter((s) => s.visible)
      .map((s) => {
        const sheet = sheets.find((sh) => sh.id === s.sheetId);
        if (!sheet) return null;
        const yCol = sheet.columns.find((c) => c.id === s.yColId);
        if (!yCol) return null;
        const xCol =
          resolveXColumnForY(sheet, yCol) || sheet.columns.find((c) => c.id === s.xColId);
        if (!xCol) return null;

        const companions = resolveErrorAndLabelColumnsForY(sheet, yCol);
        const yErrCol = s.yErrColId
          ? sheet.columns.find((c) => c.id === s.yErrColId)
          : companions.yErrCol;
        const xErrCol = s.xErrColId
          ? sheet.columns.find((c) => c.id === s.xErrColId)
          : companions.xErrCol;
        const labelCol = s.labelColId
          ? sheet.columns.find((c) => c.id === s.labelColId)
          : companions.labelCol;

        const typed = extractWebGLSeriesTypedArrays(xCol, yCol, xErrCol, yErrCol);

        let pointLabels: string[] | undefined;
        if (labelCol && typed.validRowIndices.length <= 2000) {
          pointLabels = Array.from(typed.validRowIndices, (r) =>
            formatCellForDisplay(labelCol, r)
          );
        }

        const displayLabel =
          s.autoLabel && yCol
            ? `${yCol.longName || yCol.shortName}${yCol.unit ? ` (${yCol.unit})` : ''}`
            : s.label;

        return {
          series: s,
          displayLabel,
          sheet,
          xCol,
          yCol,
          xErrCol,
          yErrCol,
          labelCol,
          pointLabels,
          ...typed,
        };
      })
      .filter((item): item is NonNullable<typeof item> => item !== null);
  }, [graph.series, sheets]);

  const totalGpuPointsCount = useMemo(() => {
    let sum = 0;
    for (const item of resolvedWebGLSeries) {
      sum += item.xValid.length + item.xMasked.length;
    }
    for (const f of graph.fits) {
      if (f.visible) sum += f.generatedPoints.length;
    }
    return sum;
  }, [resolvedWebGLSeries, graph.fits]);

  // Dynamic Axis Labels computation
  const effectiveXLabel = useMemo(() => {
    if (!graph.xAxis.autoLabel && graph.xAxis.label.trim()) return graph.xAxis.label;
    const xLabels = new Set<string>();
    for (const item of resolvedWebGLSeries) {
      const name = item.xCol.longName.trim() || item.xCol.shortName;
      const unit = item.xCol.unit.trim() ? ` (${item.xCol.unit.trim()})` : '';
      xLabels.add(`${name}${unit}`);
    }
    return xLabels.size > 0 ? Array.from(xLabels).join(' / ') : graph.xAxis.label || 'Asse X';
  }, [graph.xAxis, resolvedWebGLSeries]);

  const effectiveYLabel = useMemo(() => {
    if (!graph.yAxis.autoLabel && graph.yAxis.label.trim()) return graph.yAxis.label;
    const yLabels = new Set<string>();
    for (const item of resolvedWebGLSeries) {
      const name = item.yCol.longName.trim() || item.yCol.shortName;
      const unit = item.yCol.unit.trim() ? ` (${item.yCol.unit.trim()})` : '';
      yLabels.add(`${name}${unit}`);
    }
    return yLabels.size > 0 ? Array.from(yLabels).join(' · ') : graph.yAxis.label || 'Asse Y';
  }, [graph.yAxis, resolvedWebGLSeries]);

  /**
   * Build Plotly WebGL Traces (`scattergl` & `heatmapgl`) passing Float64Array directly!
   */
  const { traces, traceMetaList } = useMemo(() => {
    const outTraces: Plotly.Data[] = [];
    const metaList: {
      kind: 'experimental' | 'masked' | 'fit' | 'stem' | 'heatmap';
      seriesId?: string;
      sheetId?: string;
      xColId?: string;
      yColId?: string;
      seriesLabel?: string;
      color?: string;
      rowIndices?: Int32Array;
      xTyped?: Float64Array;
      yTyped?: Float64Array;
    }[] = [];

    if (useHeatmapGL && resolvedWebGLSeries.length > 0) {
      const primary = resolvedWebGLSeries[0];
      const gridN = 45;
      const minX = Number.isFinite(primary.minX) ? primary.minX : 0;
      const maxX = Number.isFinite(primary.maxX) && primary.maxX > minX ? primary.maxX : minX + 10;
      const minY = Number.isFinite(primary.minY) ? primary.minY : 0;
      const maxY = Number.isFinite(primary.maxY) && primary.maxY > minY ? primary.maxY : minY + 10;

      const xGrid = new Float64Array(gridN);
      const yGrid = new Float64Array(gridN);
      const zMatrix: number[][] = Array.from({ length: gridN }, () => new Array(gridN).fill(0));

      for (let i = 0; i < gridN; i++) {
        xGrid[i] = minX + (i * (maxX - minX)) / (gridN - 1);
        yGrid[i] = minY + (i * (maxY - minY)) / (gridN - 1);
      }

      const len = primary.xValid.length;
      for (let k = 0; k < len; k++) {
        const gx = Math.min(
          gridN - 1,
          Math.max(0, Math.floor(((primary.xValid[k] - minX) / (maxX - minX)) * (gridN - 1)))
        );
        const gy = Math.min(
          gridN - 1,
          Math.max(0, Math.floor(((primary.yValid[k] - minY) / (maxY - minY)) * (gridN - 1)))
        );
        zMatrix[gy][gx] += 1;
      }

      outTraces.push({
        type: 'heatmapgl' as Plotly.PlotType,
        x: xGrid as unknown as Plotly.Datum[],
        y: yGrid as unknown as Plotly.Datum[],
        z: zMatrix,
        colorscale: 'Viridis',
        name: `Densità 2D (${primary.displayLabel})`,
        showscale: true,
      } as Plotly.Data);
      metaList.push({ kind: 'heatmap' });
    }

    for (const item of resolvedWebGLSeries) {
      const {
        series,
        displayLabel,
        xValid,
        yValid,
        xErrValid,
        yErrValid,
        validRowIndices,
        xMasked,
        yMasked,
        maskedRowIndices,
        pointLabels,
      } = item;

      // Optional Stem / DropLines trace in WebGL (NaN-separated vertical segments in a single Float64Array)
      if ((series.plotType === 'stem' || series.dropLines) && xValid.length > 0) {
        const segLen = xValid.length * 3;
        const stemX = new Float64Array(segLen);
        const stemY = new Float64Array(segLen);
        for (let i = 0; i < xValid.length; i++) {
          const base = i * 3;
          stemX[base] = xValid[i];
          stemY[base] = 0;
          stemX[base + 1] = xValid[i];
          stemY[base + 1] = yValid[i];
          stemX[base + 2] = Number.NaN;
          stemY[base + 2] = Number.NaN;
        }
        outTraces.push({
          type: 'scattergl',
          mode: 'lines',
          name: `${displayLabel} (Stem)`,
          showlegend: false,
          hoverinfo: 'skip',
          x: stemX as unknown as Plotly.Datum[],
          y: stemY as unknown as Plotly.Datum[],
          line: {
            color: series.color,
            width: series.plotType === 'stem' ? series.lineWidth || 1.75 : 1,
            dash: series.dropLines && series.plotType !== 'stem' ? 'dot' : 'solid',
          },
        });
        metaList.push({ kind: 'stem' });
      }

      // If spline plotType or connection === 'spline', add smooth spline curve trace in scattergl
      if (
        (series.plotType === 'spline' || series.connection === 'spline') &&
        xValid.length >= 3 &&
        series.plotType !== 'scatter' &&
        series.plotType !== 'bar'
      ) {
        const { xs, ys } = buildSmoothSplineFloat64(xValid, yValid);
        outTraces.push({
          type: 'scattergl',
          mode: 'lines',
          name: `${displayLabel} (Spline)`,
          showlegend: series.plotType === 'spline',
          hoverinfo: 'skip',
          x: xs as unknown as Plotly.Datum[],
          y: ys as unknown as Plotly.Datum[],
          line: {
            color: series.color,
            width: series.lineWidth,
            dash: mapLineStyleToPlotlyDash(series.lineStyle),
          },
          fill: series.fillArea || series.plotType === 'area' ? 'tozeroy' : 'none',
          fillcolor: hexToRgba(
            series.fillColor || series.color,
            series.fillOpacity ?? 0.2
          ),
        });
        metaList.push({ kind: 'stem' });
      }

      // Main Experimental Data Trace (ALWAYS scattergl except bar)
      if (series.plotType === 'bar') {
        outTraces.push({
          type: 'bar',
          name: displayLabel,
          x: xValid as unknown as Plotly.Datum[],
          y: yValid as unknown as Plotly.Datum[],
          marker: {
            color: series.color,
            opacity: 0.85,
          },
          error_y:
            (series.showErrorBars || Boolean(yErrValid)) && yErrValid
              ? {
                  type: 'data',
                  array: yErrValid as unknown as Plotly.Datum[],
                  visible: true,
                  color: series.errorBarColor || series.color,
                  width: series.errorBarCap ?? 4,
                }
              : undefined,
        });
      } else {
        const showLines =
          (series.plotType === 'line' ||
            series.plotType === 'line_symbol' ||
            series.plotType === 'step' ||
            series.plotType === 'area') &&
          series.connection !== 'none' &&
          series.connection !== 'spline';

        const showMarkers =
          series.plotType === 'scatter' ||
          series.plotType === 'line_symbol' ||
          series.plotType === 'stem' ||
          (series.connection === 'spline' && series.plotType !== 'spline');

        const modeParts: string[] = [];
        if (showLines) modeParts.push('lines');
        if (showMarkers || modeParts.length === 0) modeParts.push('markers');
        if (pointLabels && pointLabels.some((l) => l.trim() !== '')) {
          modeParts.push('text');
        }

        outTraces.push({
          type: 'scattergl',
          mode: modeParts.join('+') as
            | 'lines'
            | 'markers'
            | 'lines+markers'
            | 'Text'
            | 'lines+markers+text',
          name: displayLabel,
          // Pass Float64Array directly to WebGL buffer!
          x: xValid as unknown as Plotly.Datum[],
          y: yValid as unknown as Plotly.Datum[],
          text: pointLabels,
          textposition: 'top center',
          line: {
            color: series.color,
            width: series.lineWidth,
            dash: mapLineStyleToPlotlyDash(series.lineStyle),
            shape:
              series.plotType === 'step' || series.connection === 'step' ? 'hv' : 'linear',
          },
          marker: {
            symbol: mapMarkerShapeToPlotlySymbol(series.markerShape),
            size: Math.max(4, (series.markerSize || 4.5) * 1.85),
            color: series.symbolFill || series.color,
            line: {
              color: series.symbolBorder || '#ffffff',
              width: series.symbolBorderWidth ?? 1,
            },
          },
          fill: series.fillArea || series.plotType === 'area' ? 'tozeroy' : 'none',
          fillcolor: hexToRgba(
            series.fillColor || series.color,
            series.fillOpacity ?? 0.2
          ),
          error_x:
            (series.showErrorBars || Boolean(xErrValid)) && xErrValid
              ? {
                  type: 'data',
                  array: xErrValid as unknown as Plotly.Datum[],
                  visible: true,
                  color: series.errorBarColor || series.color,
                  width: series.errorBarCap ?? 4,
                }
              : undefined,
          error_y:
            (series.showErrorBars || Boolean(yErrValid)) && yErrValid
              ? {
                  type: 'data',
                  array: yErrValid as unknown as Plotly.Datum[],
                  visible: true,
                  color: series.errorBarColor || series.color,
                  width: series.errorBarCap ?? 4,
                }
              : undefined,
        } as Plotly.Data);
      }

      metaList.push({
        kind: 'experimental',
        seriesId: series.id,
        sheetId: series.sheetId,
        xColId: item.xCol.id,
        yColId: item.yCol.id,
        seriesLabel: displayLabel,
        color: series.color,
        rowIndices: validRowIndices,
        xTyped: xValid,
        yTyped: yValid,
      });

      // Masked Points Trace in scattergl (highlighted in crimson X so user can click to unmask)
      if (xMasked.length > 0) {
        outTraces.push({
          type: 'scattergl',
          mode: 'markers',
          name: `${displayLabel} (Mascherati)`,
          showlegend: false,
          x: xMasked as unknown as Plotly.Datum[],
          y: yMasked as unknown as Plotly.Datum[],
          marker: {
            symbol: 'x',
            size: 9,
            color: '#dc2626',
            line: { color: '#991b1b', width: 1.5 },
          },
        });
        metaList.push({
          kind: 'masked',
          seriesId: series.id,
          sheetId: series.sheetId,
          xColId: item.xCol.id,
          yColId: item.yCol.id,
          seriesLabel: displayLabel,
          color: '#dc2626',
          rowIndices: maskedRowIndices,
          xTyped: xMasked,
          yTyped: yMasked,
        });
      }
    }

    // Fitted Curves in scattergl with contiguous Float64Arrays
    for (const fit of graph.fits) {
      if (!fit.visible || fit.generatedPoints.length < 2) continue;
      const nPts = fit.generatedPoints.length;
      const fitXF64 = new Float64Array(nPts);
      const fitYF64 = new Float64Array(nPts);
      for (let i = 0; i < nPts; i++) {
        fitXF64[i] = fit.generatedPoints[i].x;
        fitYF64[i] = fit.generatedPoints[i].y;
      }

      outTraces.push({
        type: 'scattergl',
        mode: 'lines',
        name: `${fit.name} (R²=${fit.rSquared.toFixed(4)})`,
        x: fitXF64 as unknown as Plotly.Datum[],
        y: fitYF64 as unknown as Plotly.Datum[],
        line: {
          color: fit.color,
          width: fit.lineWidth || 2.5,
          dash: mapLineStyleToPlotlyDash(fit.lineStyle),
        },
      });
      metaList.push({ kind: 'fit' });
    }

    return { traces: outTraces, traceMetaList: metaList };
  }, [resolvedWebGLSeries, graph.fits, useHeatmapGL]);

  const traceMetaRef = useRef(traceMetaList);
  traceMetaRef.current = traceMetaList;

  // Build Plotly Layout
  const plotlyLayout = useMemo<Partial<Plotly.Layout>>(() => {
    const dragmode: Plotly.Layout['dragmode'] =
      toolMode === 'zoom'
        ? 'zoom'
        : toolMode === 'remove_bad' || toolMode === 'mask_point'
        ? 'select'
        : 'pan';

    const legendX = graph.legendPosition.includes('right') ? 0.98 : 0.02;
    const legendXAnchor = graph.legendPosition.includes('right') ? 'right' : 'left';
    const legendY = graph.legendPosition.includes('bottom') ? 0.02 : 0.98;
    const legendYAnchor = graph.legendPosition.includes('bottom') ? 'bottom' : 'top';

    const annotations: Record<string, unknown>[] = (graph.annotations || []).map((ann) => ({
      x: ann.x,
      y: ann.y,
      xref: 'x',
      yref: 'y',
      text: ann.text,
      showarrow: true,
      arrowhead: 2,
      ax: 0,
      ay: -22,
      font: {
        family: 'JetBrains Mono, monospace',
        size: ann.fontSize || 11,
        color: ann.color || '#0f172a',
      },
      bgcolor: ann.bgColor || '#ffffff',
      bordercolor: ann.borderColor || '#94a3b8',
      borderwidth: 1,
      borderpad: 4,
    }));

    return {
      title: {
        text: graph.subtitle
          ? `<b>${graph.title}</b><br><span style="font-size:11px;color:#64748b">${graph.subtitle}</span>`
          : `<b>${graph.title}</b>`,
        font: { family: 'Plus Jakarta Sans, sans-serif', size: 14, color: '#0f172a' },
        x: 0.5,
        xanchor: 'center',
      },
      dragmode,
      hovermode: 'closest',
      margin: { l: 68, r: 28, t: 54, b: 56 },
      paper_bgcolor: '#ffffff',
      plot_bgcolor: '#ffffff',
      showlegend: graph.showLegend,
      legend: {
        x: legendX,
        y: legendY,
        xanchor: legendXAnchor,
        yanchor: legendYAnchor,
        bgcolor: 'rgba(255,255,255,0.92)',
        bordercolor: '#cbd5e1',
        borderwidth: 1,
        font: { family: 'Plus Jakarta Sans, sans-serif', size: 11, color: '#0f172a' },
      },
      xaxis: {
        title: {
          text: effectiveXLabel,
          font: { family: 'Plus Jakarta Sans, sans-serif', size: 12, color: '#0f172a' },
        },
        type: graph.xAxis.scaleType === 'log10' ? 'log' : 'linear',
        autorange: graph.xAxis.min === 'auto' && graph.xAxis.max === 'auto',
        range:
          graph.xAxis.min !== 'auto' && graph.xAxis.max !== 'auto'
            ? [
                graph.xAxis.scaleType === 'log10'
                  ? Math.log10(Math.max(1e-9, Number(graph.xAxis.min)))
                  : Number(graph.xAxis.min),
                graph.xAxis.scaleType === 'log10'
                  ? Math.log10(Math.max(1e-9, Number(graph.xAxis.max)))
                  : Number(graph.xAxis.max),
              ]
            : undefined,
        showgrid: graph.xAxis.showMajorGrid,
        gridcolor: '#e2e8f0',
        nticks: graph.xAxis.tickCount || 7,
        ticks: graph.xAxis.tickDirection === 'in' ? 'inside' : 'outside',
        showline: true,
        linecolor: '#1e293b',
        linewidth: 1.5,
        mirror: true,
        zeroline: false,
        exponentformat: graph.xAxis.labelFormat === 'scientific' ? 'e' : 'none',
      },
      yaxis: {
        title: {
          text: effectiveYLabel,
          font: { family: 'Plus Jakarta Sans, sans-serif', size: 12, color: '#0f172a' },
        },
        type: graph.yAxis.scaleType === 'log10' ? 'log' : 'linear',
        autorange: graph.yAxis.min === 'auto' && graph.yAxis.max === 'auto',
        range:
          graph.yAxis.min !== 'auto' && graph.yAxis.max !== 'auto'
            ? [
                graph.yAxis.scaleType === 'log10'
                  ? Math.log10(Math.max(1e-9, Number(graph.yAxis.min)))
                  : Number(graph.yAxis.min),
                graph.yAxis.scaleType === 'log10'
                  ? Math.log10(Math.max(1e-9, Number(graph.yAxis.max)))
                  : Number(graph.yAxis.max),
              ]
            : undefined,
        showgrid: graph.yAxis.showMajorGrid,
        gridcolor: '#e2e8f0',
        nticks: graph.yAxis.tickCount || 6,
        ticks: graph.yAxis.tickDirection === 'in' ? 'inside' : 'outside',
        showline: true,
        linecolor: '#1e293b',
        linewidth: 1.5,
        mirror: true,
        zeroline: false,
        exponentformat: graph.yAxis.labelFormat === 'scientific' ? 'e' : 'none',
      },
      annotations,
    };
  }, [
    graph.title,
    graph.subtitle,
    graph.showLegend,
    graph.legendPosition,
    graph.xAxis,
    graph.yAxis,
    graph.annotations,
    effectiveXLabel,
    effectiveYLabel,
    toolMode,
  ]);

  /**
   * RULE 3: Update WebGL buffers in-place using `Plotly.react()` (never recreate with `newPlot`)
   */
  useEffect(() => {
    const el = plotContainerRef.current;
    if (!el) return;

    const config: Partial<Plotly.Config> = {
      responsive: true,
      displaylogo: false,
      displayModeBar: false,
      scrollZoom: true,
      doubleClick: 'reset',
    };

    Plotly.react(el, traces, plotlyLayout, config);

    if (!isPlotInitializedRef.current) {
      isPlotInitializedRef.current = true;
      const plotEl = el as unknown as Plotly.PlotlyHTMLElement;

      plotEl.on('plotly_hover', (eventData: Plotly.PlotMouseEvent) => {
        const pt = eventData.points?.[0];
        if (!pt) return;
        const meta = traceMetaRef.current[pt.curveNumber];
        if (!meta || !meta.rowIndices || !meta.xTyped || !meta.yTyped) return;

        const pointIdx = pt.pointIndex ?? pt.pointNumber;
        if (pointIdx === undefined || pointIdx < 0 || pointIdx >= meta.rowIndices.length) return;

        const rowIndex = meta.rowIndices[pointIdx];
        const xv = meta.xTyped[pointIdx];
        const yv = meta.yTyped[pointIdx];

        const hovered: HoveredDataPoint = {
          seriesId: meta.seriesId || '',
          seriesLabel: meta.seriesLabel || '',
          sheetId: meta.sheetId || '',
          rowIndex,
          xColId: meta.xColId || '',
          yColId: meta.yColId || '',
          x: xv,
          y: yv,
          svgX: 0,
          svgY: 0,
          color: meta.color || '#1d4ed8',
          isMasked: meta.kind === 'masked',
        };
        setHoveredPt(hovered);
        callbacksRef.current.onHoverPoint(hovered);
      });

      plotEl.on('plotly_unhover', () => {
        setHoveredPt(null);
        callbacksRef.current.onHoverPoint(null);
      });

      plotEl.on('plotly_click', (eventData: Plotly.PlotMouseEvent) => {
        const pt = eventData.points?.[0];
        if (!pt) return;
        const {
          toolMode: curMode,
          graph: curGraph,
          onRemoveBadPoint: removeFn,
          onMaskPoint: maskFn,
          onSelectSeriesInInspector: selectFn,
          onUpdateGraph: updateGraphFn,
          onChangeToolMode: changeModeFn,
        } = callbacksRef.current;

        const meta = traceMetaRef.current[pt.curveNumber];
        if (!meta || !meta.rowIndices || !meta.xTyped || !meta.yTyped) return;

        const pointIdx = pt.pointIndex ?? pt.pointNumber;
        if (pointIdx === undefined || pointIdx < 0 || pointIdx >= meta.rowIndices.length) return;

        const rowIndex = meta.rowIndices[pointIdx];
        const xv = meta.xTyped[pointIdx];
        const yv = meta.yTyped[pointIdx];

        if (curMode === 'add_text') {
          const text = prompt('Inserisci testo annotazione per il grafico:', 'Picco Risonanza');
          if (text && text.trim()) {
            const newAnn: GraphAnnotation = {
              id: `ann_${Date.now()}`,
              text: text.trim(),
              x: Number(xv.toPrecision(4)),
              y: Number(yv.toPrecision(4)),
              color: '#0f172a',
              bgColor: '#ffffff',
              fontSize: 11,
              borderColor: '#94a3b8',
            };
            updateGraphFn({
              ...curGraph,
              annotations: [...(curGraph.annotations || []), newAnn],
            });
          }
          changeModeFn('pointer');
          return;
        }

        if (curMode === 'remove_bad' && meta.sheetId && meta.xColId && meta.yColId) {
          removeFn({
            sheetId: meta.sheetId,
            rowIndex,
            xColId: meta.xColId,
            yColId: meta.yColId,
            seriesLabel: meta.seriesLabel || '',
            xValue: xv,
            yValue: yv,
          });
          setHoveredPt(null);
        } else if (curMode === 'mask_point' && meta.sheetId && meta.yColId) {
          maskFn(meta.sheetId, meta.yColId, rowIndex);
        } else if (meta.seriesId) {
          selectFn(meta.seriesId);
        }
      });

      plotEl.on('plotly_selected', (eventData: Plotly.PlotSelectionEvent) => {
        if (!eventData || !eventData.points || eventData.points.length === 0) return;
        const {
          toolMode: curMode,
          onRemoveBadPointsBatch: removeBatchFn,
          onMaskPoint: maskFn,
        } = callbacksRef.current;

        if (curMode !== 'remove_bad' && curMode !== 'mask_point') return;

        const batch: {
          sheetId: string;
          rowIndex: number;
          xColId: string;
          yColId: string;
          seriesLabel: string;
          xValue: number;
          yValue: number;
        }[] = [];

        for (const pt of eventData.points) {
          const meta = traceMetaRef.current[pt.curveNumber];
          if (!meta || !meta.rowIndices || !meta.xTyped || !meta.yTyped) continue;
          const pointIdx = pt.pointIndex ?? pt.pointNumber;
          if (pointIdx === undefined || pointIdx < 0 || pointIdx >= meta.rowIndices.length) continue;

          const rowIndex = meta.rowIndices[pointIdx];
          if (curMode === 'mask_point' && meta.sheetId && meta.yColId) {
            maskFn(meta.sheetId, meta.yColId, rowIndex);
          } else if (meta.sheetId && meta.xColId && meta.yColId) {
            batch.push({
              sheetId: meta.sheetId,
              rowIndex,
              xColId: meta.xColId,
              yColId: meta.yColId,
              seriesLabel: meta.seriesLabel || '',
              xValue: meta.xTyped[pointIdx],
              yValue: meta.yTyped[pointIdx],
            });
          }
        }

        if (batch.length > 0 && curMode === 'remove_bad') {
          removeBatchFn(batch);
        }
      });
    }
  }, [traces, plotlyLayout]);

  // Resize WebGL viewport smoothly when panels resize
  useEffect(() => {
    const el = plotContainerRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      if (plotContainerRef.current) {
        Plotly.Plots.resize(plotContainerRef.current);
      }
    });
    ro.observe(el);
    return () => {
      ro.disconnect();
    };
  }, []);

  const handleResetZoom = () => {
    onUpdateGraph({
      ...graph,
      xAxis: { ...graph.xAxis, min: 'auto', max: 'auto' },
      yAxis: { ...graph.yAxis, min: 'auto', max: 'auto' },
    });
    if (plotContainerRef.current) {
      Plotly.relayout(plotContainerRef.current, {
        'xaxis.autorange': true,
        'yaxis.autorange': true,
      } as unknown as Partial<Plotly.Layout>);
    }
  };

  const handleQuickChangeAllPlotType = (newPlotType: PlotType) => {
    const nextConn: ConnectionType =
      newPlotType === 'spline'
        ? 'spline'
        : newPlotType === 'step'
        ? 'step'
        : newPlotType === 'scatter' || newPlotType === 'bar' || newPlotType === 'stem'
        ? 'none'
        : 'linear';

    onUpdateGraph({
      ...graph,
      series: graph.series.map((s) => ({
        ...s,
        plotType: newPlotType,
        connection: nextConn,
        fillArea: newPlotType === 'area',
      })),
    });
  };

  /**
   * RULE 3: Vector SVG & High-Res PNG Export via `Plotly.downloadImage()` only when user clicks Export
   */
  const handleExportSVG = () => {
    if (!plotContainerRef.current) return;
    Plotly.downloadImage(plotContainerRef.current, {
      format: 'svg',
      width: 1100,
      height: 700,
      filename: graph.name || 'grafico_origin_webgl',
    });
  };

  const handleExportPNG = () => {
    if (!plotContainerRef.current) return;
    Plotly.downloadImage(plotContainerRef.current, {
      format: 'png',
      width: 1200,
      height: 760,
      scale: 2,
      filename: graph.name || 'grafico_origin_webgl',
    });
  };

  const primaryPlotType: PlotType = graph.series[0]?.plotType || 'line_symbol';
  const fitBoxPos = graph.fitBoxOffset || { x: 16, y: 58 };

  return (
    <div
      className="flex flex-col h-full bg-white border border-slate-300 rounded-xs shadow-xs overflow-hidden"
      onMouseMove={(e) => {
        if (draggingFitBox) {
          const dx = e.clientX - draggingFitBox.startMouseX;
          const dy = e.clientY - draggingFitBox.startMouseY;
          onUpdateGraph({
            ...graph,
            fitBoxOffset: {
              x: Math.max(8, draggingFitBox.initialOffsetX + dx),
              y: Math.max(8, draggingFitBox.initialOffsetY + dy),
            },
          });
        }
      }}
      onMouseUp={() => {
        if (draggingFitBox) setDraggingFitBox(null);
      }}
    >
      {/* Row 1: Graph Window Tabs + Interactive Tools + WebGL GPU Badge */}
      <div className="flex items-center justify-between gap-2 px-2 py-1 bg-[#e2e8f0] border-b border-slate-300 shrink-0">
        {/* Left: Graph Window Tabs */}
        <div className="flex items-center gap-1 overflow-x-auto min-w-0">
          {!hideGraphTabs ? (
            <>
              {allGraphs.map((g) => {
                const isActive = g.id === activeGraphId;
                return (
                  <div
                    key={g.id}
                    onClick={() => onSelectGraph(g.id)}
                    className={`group flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                      isActive
                        ? 'bg-white text-slate-900 border border-slate-300 font-semibold shadow-2xs'
                        : 'text-slate-600 hover:bg-slate-300/70'
                    }`}
                  >
                    {renamingGraphId === g.id ? (
                      <input
                        type="text"
                        value={renameGraphValue}
                        autoFocus
                        onClick={(e) => e.stopPropagation()}
                        onChange={(e) => setRenameGraphValue(e.target.value)}
                        onBlur={() => {
                          if (onRenameGraph && renameGraphValue.trim()) {
                            onRenameGraph(g.id, renameGraphValue.trim());
                          }
                          setRenamingGraphId(null);
                        }}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            if (onRenameGraph && renameGraphValue.trim()) {
                              onRenameGraph(g.id, renameGraphValue.trim());
                            }
                            setRenamingGraphId(null);
                          } else if (e.key === 'Escape') {
                            setRenamingGraphId(null);
                          }
                        }}
                        className="w-24 px-1 py-0 text-xs bg-blue-50 border border-blue-500 rounded focus:outline-none text-slate-900"
                      />
                    ) : (
                      <span
                        onDoubleClick={(e) => {
                          e.stopPropagation();
                          setRenamingGraphId(g.id);
                          setRenameGraphValue(g.name);
                        }}
                        title="Doppio clic per rinominare il grafico"
                      >
                        {g.name} ({g.series.filter((s) => s.visible).length})
                      </span>
                    )}
                    {onDeleteGraph && (
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          onDeleteGraph(g.id);
                        }}
                        className="opacity-0 group-hover:opacity-100 text-slate-400 hover:text-red-600 transition-opacity ml-0.5 font-bold"
                        title="Elimina grafico"
                      >
                        ×
                      </button>
                    )}
                  </div>
                );
              })}
              {onAddGraph && (
                <button
                  type="button"
                  onClick={onAddGraph}
                  className="flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-300/60 rounded transition-colors whitespace-nowrap shrink-0"
                  title="Nuova finestra grafico"
                >
                  <Plus className="w-3 h-3" />
                </button>
              )}
            </>
          ) : (
            <span className="px-2 py-0.5 text-xs font-bold text-slate-800 truncate">
              {graph.name} — {graph.title}
            </span>
          )}
        </div>

        {/* Center: Interactive Pointer / Bad Point / Mask / Zoom Tools */}
        <div className="flex items-center gap-0.5 bg-white/90 px-1 py-0.5 rounded border border-slate-300 shrink-0">
          <button
            type="button"
            onClick={() => onChangeToolMode('pointer')}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded transition-colors ${
              toolMode === 'pointer'
                ? 'bg-blue-700 text-white font-semibold'
                : 'text-slate-700 hover:bg-slate-100'
            }`}
            title="Puntatore Standard / Pan WebGL"
          >
            <MousePointer className="w-3 h-3" />
            <span className="hidden xl:inline">Puntatore</span>
          </button>

          <button
            type="button"
            onClick={() => onChangeToolMode('reader')}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded transition-colors ${
              toolMode === 'reader'
                ? 'bg-blue-700 text-white font-semibold'
                : 'text-slate-700 hover:bg-slate-100'
            }`}
            title="Data Reader (Lettura coordinate X, Y ad alta precisione)"
          >
            <Crosshair className="w-3 h-3" />
            <span className="hidden xl:inline">Reader</span>
          </button>

          <button
            type="button"
            onClick={() =>
              onChangeToolMode(toolMode === 'remove_bad' ? 'pointer' : 'remove_bad')
            }
            className={`flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold rounded transition-colors ${
              toolMode === 'remove_bad'
                ? 'bg-red-600 text-white shadow-2xs ring-1 ring-red-700'
                : 'text-red-700 bg-red-50 hover:bg-red-100'
            }`}
            title="Remove Bad Data Points: clicca sui punti errati o traccia un rettangolo WebGL per rimuoverli automaticamente anche dalla tabella"
          >
            <Target className="w-3 h-3" />
            <span>Remove Bad Points</span>
          </button>

          <button
            type="button"
            onClick={() =>
              onChangeToolMode(toolMode === 'mask_point' ? 'pointer' : 'mask_point')
            }
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded transition-colors ${
              toolMode === 'mask_point'
                ? 'bg-amber-600 text-white font-semibold'
                : 'text-slate-700 hover:bg-slate-100'
            }`}
            title="Mask Point: esclude/include un punto dai calcoli senza cancellarlo (Uint8Array mask)"
          >
            <EyeOff className="w-3 h-3" />
            <span className="hidden lg:inline">Mask</span>
          </button>

          <button
            type="button"
            onClick={() => onChangeToolMode(toolMode === 'zoom' ? 'pointer' : 'zoom')}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded transition-colors ${
              toolMode === 'zoom'
                ? 'bg-blue-700 text-white font-semibold'
                : 'text-slate-700 hover:bg-slate-100'
            }`}
            title="Zoom Regione Rettangolare WebGL"
          >
            <ZoomIn className="w-3 h-3" />
          </button>

          <button
            type="button"
            onClick={() => onChangeToolMode(toolMode === 'add_text' ? 'pointer' : 'add_text')}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded transition-colors ${
              toolMode === 'add_text'
                ? 'bg-blue-700 text-white font-semibold'
                : 'text-slate-700 hover:bg-slate-100'
            }`}
            title="Clicca su un punto del grafico per aggiungere un'annotazione di testo"
          >
            <Type className="w-3 h-3" />
          </button>

          <button
            type="button"
            onClick={handleResetZoom}
            className="flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-slate-700 hover:bg-slate-100 rounded transition-colors"
            title="Rescale to Show All (Ripristina scala automatica X e Y)"
          >
            <RotateCcw className="w-3 h-3" />
          </button>
        </div>

        {/* Right: GPU Badge + Plot Details + Export SVG/PNG + Maximize */}
        <div className="flex items-center gap-1 shrink-0">
          <span
            className="hidden 2xl:inline-flex items-center gap-1 px-1.5 py-0.5 text-[10px] font-mono font-semibold text-indigo-950 bg-indigo-100 border border-indigo-300 rounded"
            title="Regola 3 Attiva: Rendering GPU WebGL tramite Plotly scattergl / heatmapgl con Float64Array diretti"
          >
            <Cpu className="w-2.5 h-2.5 text-indigo-700" />
            <span>WebGL GPU · {totalGpuPointsCount.toLocaleString('it-IT')} pt</span>
          </span>

          <button
            type="button"
            onClick={() => setUseHeatmapGL((v) => !v)}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded border transition-colors whitespace-nowrap ${
              useHeatmapGL
                ? 'bg-purple-700 text-white border-purple-800'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="Attiva/Disattiva sovrapposizione matrice 2D HeatmapGL (WebGL)"
          >
            <Layers className="w-3 h-3" />
            <span className="hidden xl:inline">HeatmapGL</span>
          </button>

          <button
            type="button"
            onClick={() => (onOpenPlotDetails || onOpenPlotDetailsModal)?.()}
            className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold text-blue-900 bg-blue-50 border border-blue-300 rounded hover:bg-blue-100 transition-colors whitespace-nowrap"
            title="Apri finestra Plot Details (Spessore linee, simboli, colori, barre d'errore)"
          >
            <Sliders className="w-3 h-3 text-blue-700" />
            <span className="hidden md:inline">Plot Details</span>
          </button>

          <button
            type="button"
            onClick={handleExportPNG}
            className="flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 transition-colors"
            title="Esporta grafico in alta risoluzione .PNG"
          >
            <Download className="w-3 h-3 text-slate-500" />
            <span>PNG</span>
          </button>

          <button
            type="button"
            onClick={handleExportSVG}
            className="flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 transition-colors"
            title="Esporta grafico vettoriale .SVG (generato on-demand al click)"
          >
            <Download className="w-3 h-3 text-slate-500" />
            <span>SVG</span>
          </button>

          {onToggleMaximize && (
            <button
              type="button"
              onClick={onToggleMaximize}
              className="p-1 text-slate-600 hover:text-slate-900 hover:bg-slate-300/60 rounded transition-colors"
              title={isMaximized ? 'Ripristina vista divisa' : 'Ingrandisci grafico a tutto schermo'}
            >
              {isMaximized ? (
                <Minimize2 className="w-3.5 h-3.5" />
              ) : (
                <Maximize2 className="w-3.5 h-3.5" />
              )}
            </button>
          )}
        </div>
      </div>

      {/* Row 2: Quick Plot Type Switcher + Quick Analysis (Fit / Filter) */}
      <div className="flex flex-wrap items-center justify-between gap-2 px-2.5 py-1 bg-[#f8fafc] border-b border-slate-300 text-xs shrink-0">
        <div className="flex items-center gap-1 flex-wrap">
          <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mr-0.5">
            Tipo:
          </span>
          {PLOT_TYPE_LABELS.map((pt) => {
            const isActive = primaryPlotType === pt.id;
            return (
              <button
                key={pt.id}
                type="button"
                onClick={() => handleQuickChangeAllPlotType(pt.id)}
                className={`px-2 py-0.5 text-[11px] font-medium rounded border transition-colors ${
                  isActive
                    ? 'bg-blue-700 text-white border-blue-800 font-semibold shadow-2xs'
                    : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'
                }`}
                title={pt.label}
              >
                {pt.short}
              </button>
            );
          })}
        </div>

        <div className="flex items-center gap-1 flex-wrap">
          <button
            type="button"
            onClick={() => onQuickFitActiveSeries?.('linear')}
            className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-blue-50 hover:text-blue-800 hover:border-blue-300 transition-colors"
            title="Esegui Fit Lineare rapido sulla serie attiva"
          >
            <TrendingUp className="w-3 h-3 text-red-600" />
            <span>Fit Lineare</span>
          </button>

          <button
            type="button"
            onClick={() => onQuickFitActiveSeries?.('spline')}
            className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-blue-50 hover:text-blue-800 hover:border-blue-300 transition-colors"
            title="Esegui Interpolazione Spline Cubica sulla serie attiva"
          >
            <TrendingUp className="w-3 h-3 text-purple-600" />
            <span>Fit Spline</span>
          </button>

          {onOpenFilterSmoothModal && (
            <button
              type="button"
              onClick={onOpenFilterSmoothModal}
              className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold text-emerald-900 bg-emerald-50 border border-emerald-300 rounded hover:bg-emerald-100 transition-colors"
              title="Applica Filtro Segnale o Smoothing (Savitzky-Golay, Media Mobile, Mediano, Low-Pass)"
            >
              <Sparkles className="w-3 h-3 text-emerald-700" />
              <span>Filtro & Smooth</span>
            </button>
          )}

          <button
            type="button"
            onClick={() =>
              onUpdateGraph({
                ...graph,
                xAxis: { ...graph.xAxis, showMajorGrid: !graph.xAxis.showMajorGrid },
                yAxis: { ...graph.yAxis, showMajorGrid: !graph.yAxis.showMajorGrid },
              })
            }
            className={`p-1 rounded border transition-colors ${
              graph.xAxis.showMajorGrid
                ? 'bg-blue-50 text-blue-700 border-blue-200'
                : 'bg-white text-slate-500 border-slate-300'
            }`}
            title="Mostra/Nascondi Griglia Principale"
          >
            <Grid className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Active Tool Banner when in Remove Bad Data Points or Mask Mode */}
      {(toolMode === 'remove_bad' || toolMode === 'mask_point') && (
        <div
          className={`flex items-center justify-between px-3 py-1 text-xs border-b shrink-0 ${
            toolMode === 'remove_bad'
              ? 'bg-red-600 text-white border-red-700'
              : 'bg-amber-600 text-white border-amber-700'
          }`}
        >
          <div className="flex items-center gap-2">
            <Target className="w-3.5 h-3.5 shrink-0" />
            <span>
              {toolMode === 'remove_bad' ? (
                <>
                  <strong>STRUMENTO REMOVE BAD DATA POINTS ATTIVO (WebGL):</strong> Clicca su un
                  punto errato o trascina un riquadro di selezione per eliminarlo da grafico,
                  buffer Float64Array e curve di fit.
                </>
              ) : (
                <>
                  <strong>STRUMENTO MASK POINT ATTIVO (WebGL):</strong> Clicca o seleziona un
                  riquadro di punti per escluderli temporaneamente dai calcoli (Uint8Array=1).
                </>
              )}
            </span>
          </div>
          <button
            type="button"
            onClick={() => onChangeToolMode('pointer')}
            className="px-2 py-0.5 text-[11px] font-semibold bg-white/20 hover:bg-white/30 rounded"
          >
            Esci (ESC)
          </button>
        </div>
      )}

      {/* Main WebGL GPU Canvas Area */}
      <div className="flex-1 relative bg-white overflow-hidden min-h-0">
        {/* Plotly WebGL Container (`scattergl` & `heatmapgl`) */}
        <div ref={plotContainerRef} className="w-full h-full" />

        {/* Draggable Fit Equation & Statistical Report Box Overlay */}
        {graph.showFitReportBox !== false && graph.fits.some((f) => f.visible) && (
          <div
            style={{
              left: `${fitBoxPos.x}px`,
              top: `${fitBoxPos.y}px`,
            }}
            onMouseDown={(e) => {
              e.stopPropagation();
              setDraggingFitBox({
                startMouseX: e.clientX,
                startMouseY: e.clientY,
                initialOffsetX: fitBoxPos.x,
                initialOffsetY: fitBoxPos.y,
              });
            }}
            className="absolute z-20 w-72 bg-amber-50/95 border border-amber-400 rounded shadow-md px-3 py-2 cursor-move select-none text-xs"
            title="Trascina per spostare il riquadro dei parametri di Curve Fitting"
          >
            {(() => {
              const activeFit = graph.fits.filter((f) => f.visible).slice(-1)[0];
              if (!activeFit) return null;
              return (
                <>
                  <div className="flex items-center justify-between border-b border-amber-300 pb-1 mb-1">
                    <span className="font-bold text-amber-950 text-[11px] truncate">
                      {activeFit.name} [{activeFit.rangeC} → {activeFit.rangeD}] (n=
                      {activeFit.numPointsN})
                    </span>
                    <span className="text-[9.5px] font-mono text-amber-700 flex items-center gap-0.5">
                      <Move className="w-2.5 h-2.5" /> sposta
                    </span>
                  </div>
                  <div className="font-mono text-[10.5px] font-semibold text-red-700 truncate">
                    {activeFit.equationText}
                  </div>
                  <div className="font-mono text-[10px] text-slate-700 mt-1 flex items-center justify-between">
                    <span>R² = {activeFit.rSquared.toFixed(5)}</span>
                    <span>RMSE = {formatSciNumber(activeFit.rmse, 4)}</span>
                    <span>Punti = {activeFit.sourceValidPointsCount}</span>
                  </div>
                </>
              );
            })()}
          </div>
        )}

        {/* Bottom HUD Status Bar for Data Reader / Coordinates */}
        <div className="absolute bottom-2 left-3 right-3 pointer-events-none flex items-center justify-between text-[11px] font-mono z-20">
          {hoveredPt ? (
            <div className="bg-slate-900/95 text-white px-2.5 py-1 rounded shadow-md flex items-center gap-3 border border-slate-700">
              <span className="flex items-center gap-1.5 font-semibold text-amber-300">
                <span
                  className="w-2 h-2 rounded-full inline-block"
                  style={{ backgroundColor: hoveredPt.color }}
                />
                {hoveredPt.seriesLabel}
              </span>
              <span>Riga #{hoveredPt.rowIndex + 1}</span>
              <span>X = { hoveredPt.x }</span>
              <span>Y = { hoveredPt.y }</span>
              {hoveredPt.isMasked && <span className="text-red-400 font-bold">[MASCHERATO]</span>}
              {toolMode === 'remove_bad' && (
                <span className="text-red-400 font-bold animate-pulse">
                  [CLICCA PER RIMUOVERE PUNTO ERRATO]
                </span>
              )}
            </div>
          ) : (
            <div className="bg-white/90 text-slate-600 px-2 py-0.5 rounded border border-slate-300 shadow-2xs">
              WebGL GPU Ready ({totalGpuPointsCount.toLocaleString('it-IT')} punti in Float64Array)
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
