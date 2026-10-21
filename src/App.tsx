import React, { useState, useRef, useCallback, useEffect } from 'react';
import {
  Columns,
  Rows,
  LayoutGrid,
  FileSpreadsheet,
  LineChart,
  FolderTree,
  SlidersHorizontal,
  Download,
  Upload,
  RotateCcw,
  Plus,
  Target,
  TrendingUp,
  Sliders,
  Sparkles,
  AppWindow,
  Maximize2,
  Minimize2,
  Minus,
} from 'lucide-react';
import {
  BadPointRemovalMode,
  ColumnStatistics,
  ConnectionType,
  DockLayoutMode,
  FitModelType,
  FittedCurve,
  GraphWindow,
  HoveredDataPoint,
  LineStyle,
  MDIWindowFrame,
  PlotSeries,
  PlotType,
  RemovedPointRecord,
  ToolMode,
  WorkbookSheet,
} from './types/origin';
import {
  applyAutoFillToColumn,
  applyMaskingToSheet,
  applySignalFilterAndSmooth,
  computeColumnStatistics,
  computeCurveFit,
  createEmptySheet,
  createInitialWorkspace,
  differentiateSignal,
  evaluateColumnFormula,
  getColumnShortName,
  integrateSignal,
  rebindSheetMultiXAssociations,
  recalculateSheetFormulas,
  resolveErrorAndLabelColumnsForY,
  resolveXColumnForY,
  serializeSheetToText,
  smoothSignal,
  sortEntireSheetByColumn,
} from './utils/mathEngine';
import {
  exportWorkbookSheetFile,
  OriginProjectBundle,
} from './utils/importExportEngine';
import {
  attachBinaryBuffersToColumn,
  generateHighPerformanceDatasetSheet,
  getOrSyncColumnBuffers,
  readCellNumericO1,
  toggleColumnCellMaskBinary,
  updateColumnCellBinary,
} from './utils/binaryColumnStore';
import {
  exportOriginCloneBinaryZip,
  importOriginCloneBinaryZip,
  loadWorkspaceFromIndexedDB,
  saveWorkspaceToIndexedDB,
} from './utils/binaryPersistence';
import {
  applySignalFilterInWorker,
  computeCurveFitInWorker,
} from './utils/workerClient';
import { TopMenuBar } from './components/TopMenuBar';
import { WorkbookView, ActiveSubPanel } from './components/WorkbookView';
import { GraphCanvas } from './components/GraphCanvas';
import { InspectorPanel } from './components/InspectorPanel';
import { ImportModal } from './components/ImportModal';
import { StatisticsModal } from './components/StatisticsModal';
import { PlotDetailsModal } from './components/PlotDetailsModal';
import { FilterSmoothModal } from './components/FilterSmoothModal';

const SERIES_COLORS = [
  '#2563EB',
  '#0D9488',
  '#D97706',
  '#7C3AED',
  '#16A34A',
  '#DB2777',
  '#0F172A',
  '#DC2626',
];

export default function App() {
  const initialData = useRef(createInitialWorkspace());
  const [sheets, setSheets] = useState<WorkbookSheet[]>(initialData.current.sheets);
  const [graphs, setGraphs] = useState<GraphWindow[]>(initialData.current.graphs);

  const [activeSheetId, setActiveSheetId] = useState<string>(initialData.current.sheets[0].id);
  const [activeGraphId, setActiveGraphId] = useState<string>(initialData.current.graphs[0].id);
  const [selectedSeriesId, setSelectedSeriesId] = useState<string | null>(
    initialData.current.graphs[0].series[0]?.id || null
  );

  // Spatial layout & Resizable split state
  const [dockLayout, setDockLayout] = useState<DockLayoutMode>('split-horizontal');
  const [previousDockLayout, setPreviousDockLayout] = useState<DockLayoutMode>('split-horizontal');
  const [splitRatio, setSplitRatio] = useState<number>(44);
  const [isDraggingSplitter, setIsDraggingSplitter] = useState<boolean>(false);
  const centralStageRef = useRef<HTMLDivElement | null>(null);

  // OriginLab MDI Free Windows state
  const [mdiWindows, setMdiWindows] = useState<MDIWindowFrame[]>(() => [
    {
      id: `mdi_sheet_${initialData.current.sheets[0].id}`,
      kind: 'sheet',
      targetId: initialData.current.sheets[0].id,
      x: 16,
      y: 16,
      w: 500,
      h: 540,
      zIndex: 10,
      minimized: false,
      maximized: false,
    },
    {
      id: `mdi_graph_${initialData.current.graphs[0].id}`,
      kind: 'graph',
      targetId: initialData.current.graphs[0].id,
      x: 532,
      y: 16,
      w: 640,
      h: 540,
      zIndex: 11,
      minimized: false,
      maximized: false,
    },
  ]);
  const [topZIndex, setTopZIndex] = useState<number>(20);
  const [mdiDragState, setMdiDragState] = useState<{
    winId: string;
    mode: 'move' | 'resize';
    startX: number;
    startY: number;
    origX: number;
    origY: number;
    origW: number;
    origH: number;
  } | null>(null);

  const [showLeftExplorer, setShowLeftExplorer] = useState<boolean>(true);
  const [showRightInspector, setShowRightInspector] = useState<boolean>(true);
  const [inspectorTab, setInspectorTab] = useState<'series' | 'fitting' | 'axes' | 'badpoints'>(
    'fitting'
  );

  const [toolMode, setToolMode] = useState<ToolMode>('pointer');
  const [badPointRemovalMode, setBadPointRemovalMode] = useState<BadPointRemovalMode>('cell-only');
  const [removedPointsLog, setRemovedPointsLog] = useState<RemovedPointRecord[]>([]);
  const [hoveredGraphPoint, setHoveredGraphPoint] = useState<HoveredDataPoint | null>(null);

  // Modals & Import Wizard state
  const [isImportModalOpen, setIsImportModalOpen] = useState<boolean>(false);
  const [importClipboardMode, setImportClipboardMode] = useState<boolean>(false);
  const [droppedFilesForImport, setDroppedFilesForImport] = useState<File[]>([]);
  const [isGlobalDragOver, setIsGlobalDragOver] = useState<boolean>(false);

  const [isStatsModalOpen, setIsStatsModalOpen] = useState<boolean>(false);
  const [activeStatsData, setActiveStatsData] = useState<ColumnStatistics[]>([]);
  const [isPlotDetailsOpen, setIsPlotDetailsOpen] = useState<boolean>(false);
  const [isFilterSmoothOpen, setIsFilterSmoothOpen] = useState<boolean>(false);
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const [explorerRenaming, setExplorerRenaming] = useState<{
    kind: 'sheet' | 'graph';
    id: string;
    value: string;
  } | null>(null);
  const [externalSubPanelRequest, setExternalSubPanelRequest] = useState<{
    panel: ActiveSubPanel;
    nonce: number;
  } | null>(null);

  const projectInputRef = useRef<HTMLInputElement | null>(null);

  const triggerToast = useCallback((msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage((prev) => (prev === msg ? null : prev));
    }, 3800);
  }, []);

  const activeSheet = sheets.find((s) => s.id === activeSheetId) || sheets[0];
  const activeGraph = graphs.find((g) => g.id === activeGraphId) || graphs[0];

  // RULE 5: Load persisted binary workspace from IndexedDB (Dexie.js) on initial mount
  const hasLoadedIndexedDBRef = useRef<boolean>(false);
  useEffect(() => {
    if (hasLoadedIndexedDBRef.current) return;
    hasLoadedIndexedDBRef.current = true;
    loadWorkspaceFromIndexedDB().then((saved) => {
      if (saved && saved.sheets.length > 0 && saved.graphs.length > 0) {
        setSheets(saved.sheets);
        setGraphs(saved.graphs);
        setActiveSheetId(saved.activeSheetId || saved.sheets[0].id);
        setActiveGraphId(saved.activeGraphId || saved.graphs[0].id);
        setDockLayout(saved.dockLayout || 'split-horizontal');
        setSplitRatio(saved.splitRatio || 48);
        if (saved.mdiWindows.length > 0) setMdiWindows(saved.mdiWindows);
        setRemovedPointsLog(saved.removedPointsLog || []);
      }
    });
  }, []);

  // RULE 5: Debounced auto-save of binary Float64Array & Uint8Array buffers into IndexedDB (no localStorage!)
  useEffect(() => {
    if (!hasLoadedIndexedDBRef.current) return;
    const timer = setTimeout(() => {
      saveWorkspaceToIndexedDB({
        activeSheetId,
        activeGraphId,
        dockLayout,
        splitRatio,
        sheets,
        graphs,
        mdiWindows,
        removedPointsLog,
      });
    }, 650);
    return () => clearTimeout(timer);
  }, [
    activeSheetId,
    activeGraphId,
    dockLayout,
    splitRatio,
    sheets,
    graphs,
    mdiWindows,
    removedPointsLog,
  ]);

  // Keep MDI windows synced whenever sheets or graphs are added/removed
  useEffect(() => {
    setMdiWindows((prev) => {
      const next = [...prev];
      let changed = false;

      sheets.forEach((s, idx) => {
        if (!next.some((w) => w.kind === 'sheet' && w.targetId === s.id)) {
          changed = true;
          next.push({
            id: `mdi_sheet_${s.id}`,
            kind: 'sheet',
            targetId: s.id,
            x: 28 + ((idx * 32) % 180),
            y: 28 + ((idx * 28) % 140),
            w: 490,
            h: 480,
            zIndex: topZIndex + idx + 1,
            minimized: false,
            maximized: false,
          });
        }
      });

      graphs.forEach((g, idx) => {
        if (!next.some((w) => w.kind === 'graph' && w.targetId === g.id)) {
          changed = true;
          next.push({
            id: `mdi_graph_${g.id}`,
            kind: 'graph',
            targetId: g.id,
            x: 520 + ((idx * 32) % 160),
            y: 28 + ((idx * 28) % 140),
            w: 620,
            h: 500,
            zIndex: topZIndex + sheets.length + idx + 2,
            minimized: false,
            maximized: false,
          });
        }
      });

      const validSheetIds = new Set(sheets.map((s) => s.id));
      const validGraphIds = new Set(graphs.map((g) => g.id));
      const filtered = next.filter(
        (w) =>
          (w.kind === 'sheet' && validSheetIds.has(w.targetId)) ||
          (w.kind === 'graph' && validGraphIds.has(w.targetId))
      );
      if (filtered.length !== next.length) changed = true;

      return changed ? filtered : prev;
    });
  }, [sheets, graphs]);

  // Handle splitter drag in split-horizontal / split-vertical modes
  useEffect(() => {
    if (!isDraggingSplitter) return;

    const handleMouseMove = (e: MouseEvent) => {
      if (!centralStageRef.current) return;
      const rect = centralStageRef.current.getBoundingClientRect();
      if (dockLayout === 'split-horizontal') {
        const pct = ((e.clientX - rect.left) / rect.width) * 100;
        setSplitRatio(Math.max(24, Math.min(76, Math.round(pct))));
      } else if (dockLayout === 'split-vertical') {
        const pct = ((e.clientY - rect.top) / rect.height) * 100;
        setSplitRatio(Math.max(25, Math.min(75, Math.round(pct))));
      }
    };

    const handleMouseUp = () => {
      setIsDraggingSplitter(false);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isDraggingSplitter, dockLayout]);

  // Handle MDI window move/resize
  useEffect(() => {
    if (!mdiDragState) return;

    const handleMouseMove = (e: MouseEvent) => {
      const dx = e.clientX - mdiDragState.startX;
      const dy = e.clientY - mdiDragState.startY;

      setMdiWindows((prev) =>
        prev.map((w) => {
          if (w.id !== mdiDragState.winId) return w;
          if (mdiDragState.mode === 'move') {
            return {
              ...w,
              x: Math.max(0, mdiDragState.origX + dx),
              y: Math.max(0, mdiDragState.origY + dy),
            };
          } else {
            return {
              ...w,
              w: Math.max(340, mdiDragState.origW + dx),
              h: Math.max(260, mdiDragState.origH + dy),
            };
          }
        })
      );
    };

    const handleMouseUp = () => {
      setMdiDragState(null);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [mdiDragState]);

  const bringMdiWindowToFront = (winId: string) => {
    setTopZIndex((prevZ) => {
      const nextZ = prevZ + 1;
      setMdiWindows((prev) =>
        prev.map((w) => (w.id === winId ? { ...w, zIndex: nextZ, minimized: false } : w))
      );
      return nextZ;
    });
  };

  const handleTileMdiWindows = () => {
    if (!centralStageRef.current) return;
    const rect = centralStageRef.current.getBoundingClientRect();
    const visibleWins = mdiWindows.filter((w) => !w.minimized);
    if (visibleWins.length === 0) return;

    const cols = visibleWins.length === 1 ? 1 : visibleWins.length <= 4 ? 2 : 3;
    const rows = Math.ceil(visibleWins.length / cols);
    const pad = 10;
    const cellW = Math.max(360, Math.floor((rect.width - pad * (cols + 1)) / cols));
    const cellH = Math.max(280, Math.floor((rect.height - 46 - pad * (rows + 1)) / rows));

    setMdiWindows((prev) => {
      let vIdx = 0;
      return prev.map((w) => {
        if (w.minimized) return w;
        const c = vIdx % cols;
        const r = Math.floor(vIdx / cols);
        vIdx++;
        return {
          ...w,
          maximized: false,
          x: pad + c * (cellW + pad),
          y: pad + r * (cellH + pad),
          w: cellW,
          h: cellH,
        };
      });
    });
    triggerToast('Finestre affiancate automaticamente in griglia');
  };

  const handleCascadeMdiWindows = () => {
    setMdiWindows((prev) =>
      prev.map((w, idx) => ({
        ...w,
        minimized: false,
        maximized: false,
        x: 20 + idx * 36,
        y: 18 + idx * 32,
        w: w.kind === 'graph' ? 640 : 520,
        h: 490,
        zIndex: 20 + idx,
      }))
    );
    setTopZIndex(30 + mdiWindows.length);
    triggerToast('Finestre disposte a cascata');
  };

  /**
   * Recomputes all calculated formula columns (Set Column Values & Error Propagation)
   * and all active FittedCurves that depend on modified sheets
   */
  const syncFitsWithUpdatedSheets = useCallback(
    (nextSheets: WorkbookSheet[], currentGraphs: GraphWindow[]) => {
      // First: cascade automatic recalculation on any calculated columns whose source columns changed
      let updatedSheets = nextSheets.map((sh) => recalculateSheetFormulas(sh));

      const updatedGraphs = currentGraphs.map((g) => {
        const updatedFits = g.fits.map((fit) => {
          if (!fit.autoUpdateOnDataChange) return fit;
          const srcSheet = updatedSheets.find((s) => s.id === fit.sheetId);
          if (!srcSheet) return fit;
          const xCol = srcSheet.columns.find((c) => c.id === fit.xColId);
          const yCol = srcSheet.columns.find((c) => c.id === fit.yColId);
          if (!xCol || !yCol) return fit;

          const maskedSet = new Set(yCol.maskedIndices || []);
          const validPts: { x: number; y: number }[] = [];
          const len = Math.min(xCol.values.length, yCol.values.length);
          for (let i = 0; i < len; i++) {
            if (maskedSet.has(i)) continue;
            const xv = xCol.values[i];
            const yv = yCol.values[i];
            if (xv !== null && yv !== null && Number.isFinite(xv) && Number.isFinite(yv)) {
              validPts.push({ x: xv, y: yv });
            }
          }

          const calc = computeCurveFit({
            points: validPts,
            model: fit.model,
            polyDegree: fit.polyDegree,
            rangeC: fit.rangeC,
            rangeD: fit.rangeD,
            numPointsN: fit.numPointsN,
          });

          if (fit.outputSheetId) {
            updatedSheets = updatedSheets.map((sh) => {
              if (sh.id !== fit.outputSheetId) return sh;
              return {
                ...sh,
                columns: [
                  {
                    ...sh.columns[0],
                    longName: `X Fittata [${fit.rangeC} → ${fit.rangeD}]`,
                    values: calc.generatedPoints.map((p) => p.x),
                  },
                  {
                    ...sh.columns[1],
                    longName: `Y Fit (n=${fit.numPointsN})`,
                    values: calc.generatedPoints.map((p) => p.y),
                  },
                ],
              };
            });
          }

          return {
            ...fit,
            equationText: calc.equationText,
            coefficients: calc.coefficients,
            rSquared: calc.rSquared,
            rmse: calc.rmse,
            sourceDomainMin: calc.sourceDomainMin,
            sourceDomainMax: calc.sourceDomainMax,
            sourceValidPointsCount: calc.sourceValidPointsCount,
            generatedPoints: calc.generatedPoints,
          };
        });

        return { ...g, fits: updatedFits };
      });

      setSheets(updatedSheets);
      setGraphs(updatedGraphs);
    },
    []
  );

  const handleUpdateSheet = (updated: WorkbookSheet) => {
    const nextSheets = sheets.map((s) => (s.id === updated.id ? updated : s));
    syncFitsWithUpdatedSheets(nextSheets, graphs);
  };

  const handleAddSheet = () => {
    const newSheet = createEmptySheet(`Foglio${sheets.length + 1}`, 30);
    setSheets((prev) => [...prev, newSheet]);
    setActiveSheetId(newSheet.id);
    triggerToast(`Creato nuovo foglio ${newSheet.name}`);
  };

  const handleDuplicateSheet = (sheetId: string) => {
    const src = sheets.find((s) => s.id === sheetId);
    if (!src) return;
    const newSheetId = `sheet_dup_${Date.now()}`;
    const colIdMap = new Map<string, string>();
    src.columns.forEach((c, idx) => {
      colIdMap.set(c.id, `col_dup_${Date.now()}_${idx}`);
    });
    const clonedColumns = src.columns.map((c) => {
      const { float64, maskUint8 } = getOrSyncColumnBuffers(c);
      const cloned = attachBinaryBuffersToColumn(
        {
          ...c,
          id: colIdMap.get(c.id) || c.id,
          linkedXColId: c.linkedXColId ? colIdMap.get(c.linkedXColId) : undefined,
          linkedErrColId: c.linkedErrColId ? colIdMap.get(c.linkedErrColId) : undefined,
          linkedXErrColId: c.linkedXErrColId ? colIdMap.get(c.linkedXErrColId) : undefined,
          rawStrings: c.rawStrings ? [...c.rawStrings] : undefined,
        },
        new Float64Array(float64),
        new Uint8Array(maskUint8)
      );
      return cloned;
    });
    const dupSheet: WorkbookSheet = {
      ...src,
      id: newSheetId,
      name: `${src.name}_Copia`,
      columns: clonedColumns,
    };
    setSheets((prev) => [...prev, dupSheet]);
    setActiveSheetId(dupSheet.id);
    triggerToast(`Duplicato foglio "${src.name}" in "${dupSheet.name}"`);
  };

  const handleCreateMultipleSheets = (newSheets: WorkbookSheet[], summary: string) => {
    if (newSheets.length === 0) return;
    setSheets((prev) => [...prev, ...newSheets]);
    setActiveSheetId(newSheets[0].id);
    triggerToast(summary || `Creati ${newSheets.length} nuovi fogli separati nel progetto`);
  };

  const handleDeleteSheet = (sheetId: string) => {
    const target = sheets.find((s) => s.id === sheetId);
    if (sheets.length <= 1) {
      const freshSheet = createEmptySheet('Foglio1', 30);
      setSheets([freshSheet]);
      setActiveSheetId(freshSheet.id);
      triggerToast(`Eliminato "${target?.name || 'Foglio'}" e creato nuovo foglio vuoto`);
      return;
    }
    const nextSheets = sheets.filter((s) => s.id !== sheetId);
    setSheets(nextSheets);
    if (activeSheetId === sheetId && nextSheets[0]) {
      setActiveSheetId(nextSheets[0].id);
    }
    triggerToast(`Eliminato foglio "${target?.name || 'Foglio'}"`);
  };

  const handleRenameSheet = (sheetId: string, newName: string) => {
    const trimmed = newName.trim();
    if (!trimmed) return;
    setSheets((prev) => prev.map((s) => (s.id === sheetId ? { ...s, name: trimmed } : s)));
    triggerToast(`Foglio rinominato in "${trimmed}"`);
  };

  const handleAddGraph = () => {
    const newGraph: GraphWindow = {
      id: `graph_${Date.now()}`,
      name: `Graph${graphs.length + 1}`,
      title: `Grafico Scientifico ${graphs.length + 1}`,
      subtitle: '',
      xAxis: {
        label: 'Asse X',
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
        label: 'Asse Y',
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
      showFitReportBox: true,
      annotations: [],
      series: [],
      fits: [],
    };
    setGraphs((prev) => [...prev, newGraph]);
    setActiveGraphId(newGraph.id);
    triggerToast(`Creata nuova finestra grafico ${newGraph.name}`);
  };

  const handleDeleteGraph = (graphId: string) => {
    const target = graphs.find((g) => g.id === graphId);
    if (graphs.length <= 1) {
      const freshGraph: GraphWindow = {
        id: `graph_${Date.now()}`,
        name: 'Graph1',
        title: 'Nuovo Grafico Scientifico',
        subtitle: '',
        xAxis: {
          label: 'Asse X',
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
          label: 'Asse Y',
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
        showFitReportBox: true,
        annotations: [],
        series: [],
        fits: [],
      };
      setGraphs([freshGraph]);
      setActiveGraphId(freshGraph.id);
      triggerToast(`Eliminato "${target?.name || 'Grafico'}" e creata nuova finestra vuota`);
      return;
    }
    const nextGraphs = graphs.filter((g) => g.id !== graphId);
    setGraphs(nextGraphs);
    if (activeGraphId === graphId && nextGraphs[0]) {
      setActiveGraphId(nextGraphs[0].id);
    }
    triggerToast(`Eliminato grafico "${target?.name || 'Grafico'}"`);
  };

  const handleRenameGraph = (graphId: string, newName: string) => {
    const trimmed = newName.trim();
    if (!trimmed) return;
    setGraphs((prev) => prev.map((g) => (g.id === graphId ? { ...g, name: trimmed } : g)));
    triggerToast(`Grafico rinominato in "${trimmed}"`);
  };

  const handleUpdateGraph = (updated: GraphWindow) => {
    setGraphs((prev) => prev.map((g) => (g.id === updated.id ? updated : g)));
  };

  const handlePlotColumnsToActiveGraph = (
    sheetId: string,
    yColIds: string[],
    plotType: PlotType,
    connection: ConnectionType
  ) => {
    const sheet = sheets.find((s) => s.id === sheetId);
    if (!sheet || !activeGraph) return;

    const existingYColIds = new Set(
      activeGraph.series.filter((s) => s.sheetId === sheetId).map((s) => s.yColId)
    );

    const newSeriesList: PlotSeries[] = [];
    yColIds.forEach((yId, idx) => {
      const yCol = sheet.columns.find((c) => c.id === yId);
      if (!yCol) return;
      const xCol = resolveXColumnForY(sheet, yCol);
      if (!xCol) return;
      const companions = resolveErrorAndLabelColumnsForY(sheet, yCol);
      const yErrId = yCol.linkedErrColId || companions.yErrCol?.id;
      const xErrId = yCol.linkedXErrColId || companions.xErrCol?.id;
      const lblId = companions.labelCol?.id;

      if (!existingYColIds.has(yId)) {
        const color =
          SERIES_COLORS[
            (activeGraph.series.length + newSeriesList.length + idx) % SERIES_COLORS.length
          ];
        newSeriesList.push({
          id: `series_${Date.now()}_${idx}_${Math.random().toString(36).slice(2, 5)}`,
          sheetId,
          xColId: xCol.id,
          yColId: yCol.id,
          xErrColId: xErrId,
          yErrColId: yErrId,
          labelColId: lblId,
          label: `${yCol.longName || yCol.shortName}${yCol.unit ? ` (${yCol.unit})` : ''}`,
          autoLabel: true,
          plotType,
          color,
          lineWidth: 1.75,
          lineStyle: 'solid',
          connection,
          markerShape: 'circle',
          markerSize: 4.5,
          symbolFill: color,
          symbolBorder: '#ffffff',
          symbolBorderWidth: 1,
          fillArea: plotType === 'area',
          fillColor: color,
          fillOpacity: 0.2,
          dropLines: false,
          showErrorBars: Boolean(yErrId || xErrId),
          errorBarCap: 4,
          errorBarColor: color,
          visible: true,
        });
      }
    });

    const updatedSeries = activeGraph.series.map((s) =>
      s.sheetId === sheetId && yColIds.includes(s.yColId) ? { ...s, plotType, connection } : s
    );

    handleUpdateGraph({
      ...activeGraph,
      series: [...updatedSeries, ...newSeriesList],
    });

    if (dockLayout === 'sheet-only') {
      setDockLayout('split-horizontal');
    }
    triggerToast(
      newSeriesList.length > 0
        ? `Aggiunte ${newSeriesList.length} serie a ${activeGraph.name}`
        : `Aggiornato tipo di grafico (${plotType}) su ${activeGraph.name}`
    );
  };

  const handlePlotColumnsToNewGraph = (
    sheetId: string,
    yColIds: string[],
    plotType: PlotType,
    connection: ConnectionType
  ) => {
    const sheet = sheets.find((s) => s.id === sheetId);
    if (!sheet) return;

    const series: PlotSeries[] = [];
    yColIds.forEach((yId, idx) => {
      const yCol = sheet.columns.find((c) => c.id === yId);
      if (!yCol) return;
      const xCol = resolveXColumnForY(sheet, yCol);
      if (!xCol) return;
      const companions = resolveErrorAndLabelColumnsForY(sheet, yCol);
      const yErrId = yCol.linkedErrColId || companions.yErrCol?.id;
      const xErrId = yCol.linkedXErrColId || companions.xErrCol?.id;
      const lblId = companions.labelCol?.id;

      const color = SERIES_COLORS[idx % SERIES_COLORS.length];
      series.push({
        id: `series_${Date.now()}_${idx}`,
        sheetId,
        xColId: xCol.id,
        yColId: yCol.id,
        xErrColId: xErrId,
        yErrColId: yErrId,
        labelColId: lblId,
        label: `${yCol.longName || yCol.shortName}${yCol.unit ? ` (${yCol.unit})` : ''}`,
        autoLabel: true,
        plotType,
        color,
        lineWidth: 1.75,
        lineStyle: 'solid',
        connection,
        markerShape: idx % 2 === 0 ? 'circle' : 'square',
        markerSize: 4.5,
        symbolFill: color,
        symbolBorder: '#ffffff',
        symbolBorderWidth: 1,
        fillArea: plotType === 'area',
        fillColor: color,
        fillOpacity: 0.2,
        dropLines: false,
        showErrorBars: Boolean(yErrId || xErrId),
        errorBarCap: 4,
        errorBarColor: color,
        visible: true,
      });
    });

    const newGraph: GraphWindow = {
      id: `graph_${Date.now()}`,
      name: `Graph${graphs.length + 1}_${sheet.name}`,
      title: `Analisi Dati — ${sheet.name}`,
      subtitle: `${series.length} serie tracciate`,
      xAxis: {
        label: 'Asse X',
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
        label: 'Asse Y',
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
      showFitReportBox: true,
      annotations: [],
      series,
      fits: [],
    };

    setGraphs((prev) => [...prev, newGraph]);
    setActiveGraphId(newGraph.id);
    if (dockLayout === 'sheet-only') {
      setDockLayout('split-horizontal');
    }
    triggerToast(`Creato nuovo grafico ${newGraph.name}`);
  };

  /**
   * "Remove Bad Data Points" handler
   */
  const handleRemoveBadPointsBatch = (
    pointsToRemove: {
      sheetId: string;
      rowIndex: number;
      xColId: string;
      yColId: string;
      seriesLabel: string;
      xValue: number;
      yValue: number;
    }[]
  ) => {
    if (pointsToRemove.length === 0) return;

    const newRecords: RemovedPointRecord[] = [];
    const sortedPts = [...pointsToRemove].sort((a, b) => b.rowIndex - a.rowIndex);

    let nextSheets = [...sheets];

    for (const pt of sortedPts) {
      const targetSheet = nextSheets.find((s) => s.id === pt.sheetId);
      if (!targetSheet) continue;

      const fullRowSnapshot = targetSheet.columns.map((c) =>
        readCellNumericO1(c, pt.rowIndex)
      );

      newRecords.push({
        id: `rem_${Date.now()}_${pt.rowIndex}_${Math.random().toString(36).slice(2, 6)}`,
        timestamp: new Date().toLocaleTimeString('it-IT'),
        sheetId: pt.sheetId,
        sheetName: targetSheet.name,
        rowIndex: pt.rowIndex,
        xColId: pt.xColId,
        yColId: pt.yColId,
        seriesLabel: pt.seriesLabel,
        xValue: pt.xValue,
        yValue: pt.yValue,
        removalMode: badPointRemovalMode,
        fullRowSnapshot,
      });

      nextSheets = nextSheets.map((sh) => {
        if (sh.id !== pt.sheetId) return sh;
        if (badPointRemovalMode === 'entire-row') {
          return {
            ...sh,
            columns: sh.columns.map((col) => {
              const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
              if (pt.rowIndex >= length) return col;
              const newLen = Math.max(0, length - 1);
              const nextF64 = new Float64Array(newLen);
              const nextU8 = new Uint8Array(newLen);
              let dst = 0;
              for (let i = 0; i < length; i++) {
                if (i === pt.rowIndex) continue;
                nextF64[dst] = float64[i];
                nextU8[dst] = maskUint8[i];
                dst++;
              }
              return attachBinaryBuffersToColumn(col, nextF64, nextU8);
            }),
          };
        } else if (badPointRemovalMode === 'mask-only') {
          return {
            ...sh,
            columns: sh.columns.map((col) =>
              col.id === pt.yColId ? toggleColumnCellMaskBinary(col, pt.rowIndex) : col
            ),
          };
        } else {
          return {
            ...sh,
            columns: sh.columns.map((col) =>
              col.id === pt.yColId ? updateColumnCellBinary(col, pt.rowIndex, null) : col
            ),
          };
        }
      });
    }

    setRemovedPointsLog((prev) => [...newRecords, ...prev]);
    syncFitsWithUpdatedSheets(nextSheets, graphs);
    triggerToast(
      pointsToRemove.length === 1
        ? `Punto rimosso da buffer Float64Array (Riga #${pointsToRemove[0].rowIndex + 1}) e curve ricalcolate`
        : `Rimossi ${pointsToRemove.length} punti dal buffer binario e curve ricalcolate`
    );
  };

  const handleMaskPoint = (sheetId: string, colId: string, rowIndex: number) => {
    const nextSheets = sheets.map((sh) => {
      if (sh.id !== sheetId) return sh;
      return {
        ...sh,
        columns: sh.columns.map((c) =>
          c.id === colId ? toggleColumnCellMaskBinary(c, rowIndex) : c
        ),
      };
    });
    syncFitsWithUpdatedSheets(nextSheets, graphs);
    triggerToast(`Modificato stato maschera Uint8Array per riga #${rowIndex + 1}`);
  };

  const handleRestoreRemovedPoint = (recordId: string) => {
    const rec = removedPointsLog.find((r) => r.id === recordId);
    if (!rec) return;

    const rowIdx = Number(rec.rowIndex);
    const nextSheets = sheets.map((sh) => {
      if (sh.id !== rec.sheetId) return sh;
      if (rec.removalMode === 'entire-row' && rec.fullRowSnapshot) {
        return {
          ...sh,
          columns: sh.columns.map((col, cIdx) => {
            const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
            const newLen = Math.max(length + 1, rowIdx + 1);
            const nextF64 = new Float64Array(newLen).fill(Number.NaN);
            const nextU8 = new Uint8Array(newLen).fill(2);
            for (let i = 0; i < rowIdx && i < length; i++) {
              nextF64[i] = float64[i];
              nextU8[i] = maskUint8[i];
            }
            const restoredVal = rec.fullRowSnapshot?.[cIdx] ?? null;
            nextF64[rowIdx] = restoredVal !== null ? restoredVal : Number.NaN;
            nextU8[rowIdx] = restoredVal !== null ? 0 : 2;
            for (let i = rowIdx; i < length; i++) {
              nextF64[i + 1] = float64[i];
              nextU8[i + 1] = maskUint8[i];
            }
            return attachBinaryBuffersToColumn(col, nextF64, nextU8);
          }),
        };
      } else {
        return {
          ...sh,
          columns: sh.columns.map((col) => {
            if (col.id !== rec.yColId) return col;
            return updateColumnCellBinary(col, rowIdx, rec.yValue);
          }),
        };
      }
    });

    setRemovedPointsLog((prev) => prev.filter((r) => r.id !== recordId));
    syncFitsWithUpdatedSheets(nextSheets, graphs);
    triggerToast(`Ripristinato punto alla Riga #${rowIdx + 1} (${rec.seriesLabel})`);
  };

  /**
   * RULE 2: Run Curve Fitting across [C, D] with n points inside math.worker.ts on Float64Array buffers
   */
  const handleRunCurveFit = async (params: {
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
  }) => {
    if (!activeGraph) return;
    const series = activeGraph.series.find((s) => s.id === params.sourceSeriesId);
    if (!series) return;
    const srcSheet = sheets.find((s) => s.id === series.sheetId);
    if (!srcSheet) return;

    const yCol = srcSheet.columns.find((c) => c.id === series.yColId);
    if (!yCol) return;
    const xCol =
      resolveXColumnForY(srcSheet, yCol) || srcSheet.columns.find((c) => c.id === series.xColId);
    if (!xCol) return;

    let calc: Awaited<ReturnType<typeof computeCurveFitInWorker>>;
    try {
      calc = await computeCurveFitInWorker({
        xCol,
        yCol,
        model: params.model,
        polyDegree: params.polyDegree,
        rangeC: params.rangeC,
        rangeD: params.rangeD,
        numPointsN: params.numPointsN,
      });
    } catch {
      const maskedSet = new Set(yCol.maskedIndices || []);
      const validPts: { x: number; y: number }[] = [];
      const len = Math.min(xCol.values.length, yCol.values.length);
      for (let i = 0; i < len; i++) {
        if (maskedSet.has(i)) continue;
        const xv = xCol.values[i];
        const yv = yCol.values[i];
        if (xv !== null && yv !== null && Number.isFinite(xv) && Number.isFinite(yv)) {
          validPts.push({ x: xv, y: yv });
        }
      }
      const fallback = computeCurveFit({
        points: validPts,
        model: params.model,
        polyDegree: params.polyDegree,
        rangeC: params.rangeC,
        rangeD: params.rangeD,
        numPointsN: params.numPointsN,
      });
      const fx = new Float64Array(fallback.generatedPoints.map((p) => p.x));
      const fy = new Float64Array(fallback.generatedPoints.map((p) => p.y));
      calc = { ...fallback, fitX: fx, fitY: fy, engineUsed: 'native-float64-worker' };
    }

    const fitId = `fit_${Date.now()}`;
    const modelLabel =
      params.model === 'linear'
        ? 'Fit Lineare'
        : params.model === 'polynomial'
        ? `Fit Polinomiale (g=${params.polyDegree})`
        : params.model === 'gaussian'
        ? 'Fit Gaussiano Picco'
        : params.model === 'exponential'
        ? 'Fit Esponenziale'
        : 'Spline Cubica C²';

    let outputSheetId: string | undefined;
    if (params.createOutputSheet) {
      outputSheetId = `sheet_fit_${Date.now()}`;
      const colFX = attachBinaryBuffersToColumn(
        {
          id: `col_fx_${Date.now()}`,
          shortName: 'A',
          longName: `${xCol.longName || 'X'} Fittata [${params.rangeC} → ${params.rangeD}]`,
          unit: xCol.unit,
          role: 'X',
          values: [],
        },
        calc.fitX
      );
      const colFY = attachBinaryBuffersToColumn(
        {
          id: `col_fy_${Date.now()}`,
          shortName: 'B',
          longName: `${modelLabel} - ${yCol.longName || 'Y'} (n=${params.numPointsN})`,
          unit: yCol.unit,
          role: 'Y',
          linkedXColId: colFX.id,
          values: [],
        },
        calc.fitY
      );
      const fitSheet: WorkbookSheet = {
        id: outputSheetId,
        name: `Fit_${yCol.shortName}_${params.numPointsN}pt`,
        isFitOutput: true,
        sourceFitId: fitId,
        columns: [colFX, colFY],
      };
      setSheets((prev) => [...prev, fitSheet]);
    }

    const newFit: FittedCurve = {
      id: fitId,
      name: `${modelLabel} (${yCol.longName || yCol.shortName})`,
      sourceSeriesId: series.id,
      sheetId: srcSheet.id,
      xColId: xCol.id,
      yColId: yCol.id,
      model: params.model,
      polyDegree: params.polyDegree,
      rangeC: params.rangeC,
      rangeD: params.rangeD,
      numPointsN: params.numPointsN,
      autoUpdateOnDataChange: true,
      outputSheetId,
      equationText: calc.equationText,
      coefficients: calc.coefficients,
      rSquared: calc.rSquared,
      rmse: calc.rmse,
      sourceDomainMin: calc.sourceDomainMin,
      sourceDomainMax: calc.sourceDomainMax,
      sourceValidPointsCount: calc.sourceValidPointsCount,
      generatedPoints: calc.generatedPoints,
      color: params.color,
      lineWidth: params.lineWidth,
      lineStyle: params.lineStyle,
      visible: true,
    };

    handleUpdateGraph({
      ...activeGraph,
      fits: [newFit, ...activeGraph.fits],
    });

    triggerToast(
      `Calcolato ${modelLabel}: generati ${params.numPointsN} punti tra C=${params.rangeC} e D=${params.rangeD}`
    );
  };

  /**
   * Quick fit on active series directly from TopMenuBar
   */
  const handleQuickFitActiveSeries = (model: FitModelType) => {
    if (!activeGraph || activeGraph.series.length === 0) {
      triggerToast('Nessuna serie attiva sul grafico per eseguire il fit');
      return;
    }
    const targetSeries =
      activeGraph.series.find((s) => s.id === selectedSeriesId) || activeGraph.series[0];
    const srcSheet = sheets.find((s) => s.id === targetSeries.sheetId);
    const xCol = srcSheet?.columns.find((c) => c.id === targetSeries.xColId);
    const validX = (xCol?.values || []).filter(
      (v): v is number => v !== null && Number.isFinite(v)
    );
    const cMin = validX.length >= 2 ? Math.min(...validX) : 0;
    const dMax = validX.length >= 2 ? Math.max(...validX) : 10;

    handleRunCurveFit({
      sourceSeriesId: targetSeries.id,
      model,
      polyDegree: 3,
      rangeC: Number(cMin.toFixed(4)),
      rangeD: Number(dMax.toFixed(4)),
      numPointsN: 200,
      color: '#DC2626',
      lineWidth: 2.25,
      lineStyle: 'solid',
      createOutputSheet: true,
    });
  };

  /**
   * Apply Signal Filter & Smoothing from FilterSmoothModal
   */
  const handleApplyFilterSmoothResult = (params: {
    sheetId: string;
    sourceColId: string;
    filteredValues: (number | null)[];
    methodLabel: string;
    outputMode: 'new-col-and-plot' | 'new-col-only' | 'overwrite';
    color: string;
  }) => {
    const sheet = sheets.find((s) => s.id === params.sheetId);
    if (!sheet) return;
    const srcCol = sheet.columns.find((c) => c.id === params.sourceColId);
    if (!srcCol) return;

    if (params.outputMode === 'overwrite') {
      const updatedCols = sheet.columns.map((c) =>
        c.id === srcCol.id
          ? {
              ...c,
              values: params.filteredValues,
              comment: `${c.comment ? `${c.comment} · ` : ''}${params.methodLabel}`,
            }
          : c
      );
      const nextSheets = sheets.map((sh) =>
        sh.id === sheet.id ? { ...sh, columns: updatedCols } : sh
      );
      syncFitsWithUpdatedSheets(nextSheets, graphs);
      triggerToast(`Sovrascritta colonna ${srcCol.shortName} con ${params.methodLabel}`);
      return;
    }

    const newColIdx = sheet.columns.length;
    const newShortName = getColumnShortName(newColIdx);
    const xCol = resolveXColumnForY(sheet, srcCol);
    const newColId = `col_filt_${Date.now()}`;
    const newCol: WorkbookSheet['columns'][0] = {
      id: newColId,
      shortName: newShortName,
      longName: `${srcCol.longName || srcCol.shortName} (${params.methodLabel})`,
      unit: srcCol.unit,
      comment: `Filtrato da ${srcCol.shortName}(Y)`,
      role: 'Y',
      linkedXColId: xCol?.id,
      values: params.filteredValues,
    };

    const nextSheets = sheets.map((sh) =>
      sh.id === sheet.id ? { ...sh, columns: [...sh.columns, newCol] } : sh
    );
    setSheets(nextSheets);

    if (params.outputMode === 'new-col-and-plot' && activeGraph && xCol) {
      const newSeries: PlotSeries = {
        id: `series_filt_${Date.now()}`,
        sheetId: sheet.id,
        xColId: xCol.id,
        yColId: newColId,
        label: `${newCol.longName}${newCol.unit ? ` (${newCol.unit})` : ''}`,
        autoLabel: true,
        plotType: 'line',
        color: params.color,
        lineWidth: 2.5,
        lineStyle: 'solid',
        connection: 'linear',
        markerShape: 'circle',
        markerSize: 3.5,
        symbolFill: params.color,
        symbolBorder: '#ffffff',
        symbolBorderWidth: 1,
        fillArea: false,
        fillColor: params.color,
        fillOpacity: 0.2,
        dropLines: false,
        showErrorBars: false,
        errorBarCap: 4,
        errorBarColor: colorOrDefault(params.color),
        visible: true,
      };
      handleUpdateGraph({
        ...activeGraph,
        series: [...activeGraph.series, newSeries],
      });
      triggerToast(`Aggiunta curva filtrata "${params.methodLabel}" a tabella e grafico`);
    } else {
      triggerToast(`Creata nuova colonna filtrata ${newShortName}(Y): ${params.methodLabel}`);
    }
  };

  function colorOrDefault(c: string) {
    return c || '#059669';
  }

  /**
   * Quick Signal processing functions: Smoothing, Derivative, Integration
   */
  const handleApplySmoothing = (sheetId: string, colId: string, windowSize: number) => {
    const sheet = sheets.find((s) => s.id === sheetId);
    if (!sheet) return;
    const col = sheet.columns.find((c) => c.id === colId);
    if (!col) return;

    const smoothed = smoothSignal(col.values, windowSize, 'gaussian');
    const newCol: WorkbookSheet['columns'][0] = {
      id: `col_smooth_${Date.now()}`,
      shortName: getColumnShortName(sheet.columns.length),
      longName: `${col.longName} (Levigato W=${windowSize})`,
      unit: col.unit,
      role: 'Y',
      linkedXColId: col.linkedXColId,
      values: smoothed,
    };

    const nextSheets = sheets.map((sh) =>
      sh.id === sheetId ? { ...sh, columns: [...sh.columns, newCol] } : sh
    );
    setSheets(nextSheets);
    triggerToast(`Aggiunta colonna levigata ${newCol.longName}`);
  };

  const handleApplyDerivative = (sheetId: string, xColId: string, yColId: string) => {
    const sheet = sheets.find((s) => s.id === sheetId);
    if (!sheet) return;
    const xCol = sheet.columns.find((c) => c.id === xColId);
    const yCol = sheet.columns.find((c) => c.id === yColId);
    if (!xCol || !yCol) return;

    const deriv = differentiateSignal(xCol.values, yCol.values);
    const newCol: WorkbookSheet['columns'][0] = {
      id: `col_deriv_${Date.now()}`,
      shortName: getColumnShortName(sheet.columns.length),
      longName: `d(${yCol.shortName})/d(${xCol.shortName})`,
      unit: yCol.unit && xCol.unit ? `${yCol.unit}/${xCol.unit}` : '',
      role: 'Y',
      linkedXColId: xCol.id,
      values: deriv,
    };

    const nextSheets = sheets.map((sh) =>
      sh.id === sheetId ? { ...sh, columns: [...sh.columns, newCol] } : sh
    );
    setSheets(nextSheets);
    triggerToast(`Calcolata derivata numerica: ${newCol.longName}`);
  };

  const handleApplyIntegration = (sheetId: string, xColId: string, yColId: string) => {
    const sheet = sheets.find((s) => s.id === sheetId);
    if (!sheet) return;
    const xCol = sheet.columns.find((c) => c.id === xColId);
    const yCol = sheet.columns.find((c) => c.id === yColId);
    if (!xCol || !yCol) return;

    const integ = integrateSignal(xCol.values, yCol.values);
    const newCol: WorkbookSheet['columns'][0] = {
      id: `col_integ_${Date.now()}`,
      shortName: getColumnShortName(sheet.columns.length),
      longName: `Integrale ∫(${yCol.shortName})d(${xCol.shortName})`,
      unit: yCol.unit && xCol.unit ? `${yCol.unit}·${xCol.unit}` : '',
      role: 'Y',
      linkedXColId: xCol.id,
      values: integ,
    };

    const nextSheets = sheets.map((sh) =>
      sh.id === sheetId ? { ...sh, columns: [...sh.columns, newCol] } : sh
    );
    setSheets(nextSheets);
    triggerToast(`Calcolato integrale numerico: ${newCol.longName}`);
  };

  /**
   * Tabular Data Export: CSV (comma or semicolon), TXT/TSV, DAT, or Excel (.xlsx)
   */
  const handleExportSheet = (
    sheet: WorkbookSheet,
    format: 'txt' | 'csv' | 'csv-semi' | 'dat' | 'xlsx'
  ) => {
    exportWorkbookSheetFile(sheet, format, '.');
    triggerToast(`Esportato foglio "${sheet.name}" in formato .${format === 'csv-semi' ? 'csv' : format}`);
  };

  const handleCopyActiveSheetToClipboard = () => {
    if (!activeSheet) return;
    const tsv = serializeSheetToText(activeSheet, 'txt');
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(tsv).then(() => {
        triggerToast(`Foglio "${activeSheet.name}" copiato negli appunti (TSV per Excel)`);
      });
    }
  };

  const handleExportFitTxt = (fit: FittedCurve) => {
    const header = [
      `# OriginStudio Lab - Curve Fitting Export`,
      `# Modello: ${fit.name}`,
      `# Equazione: ${fit.equationText}`,
      `# R^2 = ${fit.rSquared.toFixed(6)} | RMSE = ${fit.rmse}`,
      `# Dominio selezionato [C, D] = [${fit.rangeC}, ${fit.rangeD}] | Numero valori n = ${fit.numPointsN}`,
      `X_Fit\tY_Fit`,
    ].join('\n');
    const rows = fit.generatedPoints.map((p) => `${p.x}\t${p.y}`).join('\n');
    const blob = new Blob([`${header}\n${rows}`], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${fit.name.replace(/\s+/g, '_')}_${fit.numPointsN}pt.txt`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    triggerToast(
      `Scaricato file .TXT con ${fit.numPointsN} valori fittati in [${fit.rangeC}, ${fit.rangeD}]`
    );
  };

  /**
   * Full OriginProjectBundle (.opj.json) Save & Load
   * Preserves sheets, graphs, fits, MDI window positions, and removed points history
   */
  const handleSaveProjectBundle = () => {
    const bundle: OriginProjectBundle = {
      format: 'OriginStudioLab_Project_v2',
      savedAt: new Date().toISOString(),
      activeSheetId,
      activeGraphId,
      dockLayout,
      splitRatio,
      sheets,
      graphs,
      mdiWindows,
      removedPointsLog,
    };
    const payload = JSON.stringify(bundle, null, 2);
    const blob = new Blob([payload], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `OriginStudio_Progetto.opj.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    triggerToast('Progetto completo salvato in .opj.json (fogli, grafici, fit e cronologia)');
  };

  const loadProjectFromJsonText = useCallback(
    (jsonText: string) => {
      try {
        const parsed = JSON.parse(jsonText);
        if (Array.isArray(parsed.sheets) && Array.isArray(parsed.graphs)) {
          setSheets(parsed.sheets);
          setGraphs(parsed.graphs);
          setRemovedPointsLog(
            Array.isArray(parsed.removedPointsLog) ? parsed.removedPointsLog : []
          );
          if (Array.isArray(parsed.mdiWindows) && parsed.mdiWindows.length > 0) {
            setMdiWindows(parsed.mdiWindows);
          }
          if (parsed.dockLayout) {
            setDockLayout(parsed.dockLayout);
          }
          if (typeof parsed.splitRatio === 'number') {
            setSplitRatio(parsed.splitRatio);
          }
          if (parsed.activeSheetId) {
            setActiveSheetId(parsed.activeSheetId);
          } else if (parsed.sheets[0]) {
            setActiveSheetId(parsed.sheets[0].id);
          }
          if (parsed.activeGraphId) {
            setActiveGraphId(parsed.activeGraphId);
          } else if (parsed.graphs[0]) {
            setActiveGraphId(parsed.graphs[0].id);
          }
          triggerToast(
            `Progetto caricato (${parsed.sheets.length} fogli, ${parsed.graphs.length} grafici)`
          );
        } else {
          triggerToast('Formato file progetto non valido');
        }
      } catch {
        triggerToast('Impossibile leggere il file progetto .opj.json');
      }
    },
    [triggerToast]
  );

  const handleSaveBinaryOriginClone = async () => {
    try {
      await exportOriginCloneBinaryZip({
        activeSheetId,
        activeGraphId,
        dockLayout,
        splitRatio,
        sheets,
        graphs,
        mdiWindows,
        removedPointsLog,
      });
      triggerToast('Archivio binario .originclone (Float64Array + Uint8Array) salvato!');
    } catch {
      triggerToast('Errore durante il salvataggio dell’archivio binario .originclone');
    }
  };

  const handleLoadProjectJson = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const lowerName = file.name.toLowerCase();
    if (lowerName.endsWith('.originclone') || lowerName.endsWith('.zip')) {
      importOriginCloneBinaryZip(file)
        .then((restored) => {
          setSheets(restored.sheets);
          setGraphs(restored.graphs);
          setActiveSheetId(restored.activeSheetId || restored.sheets[0]?.id);
          setActiveGraphId(restored.activeGraphId || restored.graphs[0]?.id);
          setRemovedPointsLog(restored.removedPointsLog);
          setDockLayout(restored.dockLayout);
          setSplitRatio(restored.splitRatio);
          if (restored.mdiWindows.length > 0) {
            setMdiWindows(restored.mdiWindows);
          }
          triggerToast(`Archivio binario "${file.name}" caricato (Float64Array)!`);
        })
        .catch(() => {
          triggerToast('Errore nel caricamento dell’archivio .originclone');
        });
      e.target.value = '';
      return;
    }

    const reader = new FileReader();
    reader.onload = (ev) => {
      loadProjectFromJsonText(String(ev.target?.result || ''));
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const handleResetDemo = () => {
    const fresh = createInitialWorkspace();
    setSheets(fresh.sheets);
    setGraphs(fresh.graphs);
    setActiveSheetId(fresh.sheets[0].id);
    setActiveGraphId(fresh.graphs[0].id);
    setRemovedPointsLog([]);
    setToolMode('pointer');
    triggerToast('Ripristinato dataset scientifico calibrato (100 punti + outlier)');
  };

  const handleGenerateHighPerfDataset = (pointsCount: number) => {
    const labelK =
      pointsCount >= 1_000_000
        ? `${pointsCount / 1_000_000}M`
        : `${Math.round(pointsCount / 1000)}k`;
    const hpSheet = generateHighPerformanceDatasetSheet(pointsCount, `Segnale_${labelK}_WebGL`);
    setSheets((prev) => [...prev, hpSheet]);
    setActiveSheetId(hpSheet.id);

    const xCol = hpSheet.columns[0];
    const yCol = hpSheet.columns[1];
    const yErrCol = hpSheet.columns[2];

    if (activeGraph && xCol && yCol) {
      const color = SERIES_COLORS[activeGraph.series.length % SERIES_COLORS.length];
      const hpSeries: PlotSeries = {
        id: `series_hp_${Date.now()}`,
        sheetId: hpSheet.id,
        xColId: xCol.id,
        yColId: yCol.id,
        yErrColId: yErrCol?.id,
        label: `${yCol.longName} (${pointsCount.toLocaleString('it-IT')} pt)`,
        autoLabel: true,
        plotType: 'line',
        color,
        lineWidth: 1.5,
        lineStyle: 'solid',
        connection: 'linear',
        markerShape: 'circle',
        markerSize: 3,
        symbolFill: color,
        symbolBorder: '#ffffff',
        symbolBorderWidth: 0.5,
        fillArea: false,
        fillColor: color,
        fillOpacity: 0.15,
        dropLines: false,
        showErrorBars: false,
        errorBarCap: 3,
        errorBarColor: color,
        visible: true,
      };
      handleUpdateGraph({
        ...activeGraph,
        series: [...activeGraph.series, hpSeries],
      });
      setSelectedSeriesId(hpSeries.id);
    }

    triggerToast(
      `Generato dataset Float64Array da ${pointsCount.toLocaleString('it-IT')} punti su GPU WebGL (scattergl)!`
    );
  };

  const handleSwitchDockLayout = (nextMode: DockLayoutMode) => {
    if (nextMode !== 'sheet-only' && nextMode !== 'graph-only') {
      setPreviousDockLayout(nextMode);
    }
    setDockLayout(nextMode);
  };

  /**
   * Handle batch of imported sheets from ImportModal (single or multi-file / Excel)
   */
  const handleImportSheetsBatch = (importedSheets: WorkbookSheet[], plotImmediately: boolean) => {
    if (importedSheets.length === 0) return;

    setSheets((prev) => [...prev, ...importedSheets]);
    const firstImported = importedSheets[0];
    setActiveSheetId(firstImported.id);

    if (plotImmediately && activeGraph) {
      const addedSeries: PlotSeries[] = [];
      let colorOffset = activeGraph.series.length;

      for (const sh of importedSheets) {
        const yCols = sh.columns.filter((c) => c.role === 'Y');
        yCols.forEach((yCol, idx) => {
          const xCol = resolveXColumnForY(sh, yCol);
          if (!xCol) return;
          const color = SERIES_COLORS[(colorOffset + idx) % SERIES_COLORS.length];
          addedSeries.push({
            id: `series_imp_${Date.now()}_${colorOffset}_${idx}`,
            sheetId: sh.id,
            xColId: xCol.id,
            yColId: yCol.id,
            yErrColId: yCol.linkedErrColId,
            label: `${yCol.longName || yCol.shortName}${yCol.unit ? ` (${yCol.unit})` : ''}`,
            autoLabel: true,
            plotType: 'line_symbol',
            color,
            lineWidth: 1.75,
            lineStyle: 'solid',
            connection: 'linear',
            markerShape: 'circle',
            markerSize: 4.5,
            symbolFill: color,
            symbolBorder: '#ffffff',
            symbolBorderWidth: 1,
            fillArea: false,
            fillColor: color,
            fillOpacity: 0.2,
            dropLines: false,
            showErrorBars: !!yCol.linkedErrColId,
            errorBarCap: 4,
            errorBarColor: color,
            visible: true,
          });
        });
        colorOffset += yCols.length;
      }

      if (addedSeries.length > 0) {
        handleUpdateGraph({
          ...activeGraph,
          series: [...activeGraph.series, ...addedSeries],
        });
      }
    }

    triggerToast(
      importedSheets.length === 1
        ? `Importato foglio "${firstImported.name}" (${firstImported.columns.length} colonne)`
        : `Importati ${importedSheets.length} fogli dati nel progetto`
    );
  };

  /**
   * Global Drag & Drop handler on application window (.csv, .txt, .dat, .xlsx, .opj.json)
   */
  const handleWindowDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    if (e.dataTransfer?.types?.includes('Files')) {
      e.preventDefault();
      if (!isGlobalDragOver) setIsGlobalDragOver(true);
    }
  };

  const handleWindowDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    if (e.currentTarget === e.target) {
      setIsGlobalDragOver(false);
    }
  };

  const handleWindowDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsGlobalDragOver(false);

    const files = Array.from(e.dataTransfer?.files || []);
    if (files.length === 0) return;

    // If a single .originclone / .zip binary archive is dropped, load it directly
    if (
      files.length === 1 &&
      (files[0].name.toLowerCase().endsWith('.originclone') ||
        files[0].name.toLowerCase().endsWith('.zip'))
    ) {
      importOriginCloneBinaryZip(files[0])
        .then((restored) => {
          setSheets(restored.sheets);
          setGraphs(restored.graphs);
          setActiveSheetId(restored.activeSheetId || restored.sheets[0]?.id);
          setActiveGraphId(restored.activeGraphId || restored.graphs[0]?.id);
          setRemovedPointsLog(restored.removedPointsLog);
          setDockLayout(restored.dockLayout);
          setSplitRatio(restored.splitRatio);
          if (restored.mdiWindows.length > 0) {
            setMdiWindows(restored.mdiWindows);
          }
          triggerToast(`Archivio binario "${files[0].name}" caricato (Float64Array)!`);
        })
        .catch(() => {
          triggerToast('Errore nel caricamento dell’archivio .originclone.');
        });
      return;
    }

    // If a single .json / .opj.json project file is dropped, load the project directly
    if (
      files.length === 1 &&
      (files[0].name.toLowerCase().endsWith('.json') ||
        files[0].name.toLowerCase().endsWith('.opj'))
    ) {
      const reader = new FileReader();
      reader.onload = (ev) => {
        loadProjectFromJsonText(String(ev.target?.result || ''));
      };
      reader.readAsText(files[0]);
      return;
    }

    // Otherwise open the Import Wizard with all dropped data files (.csv, .txt, .dat, .xlsx)
    setDroppedFilesForImport(files);
    setImportClipboardMode(false);
    setIsImportModalOpen(true);
  };

  return (
    <div
      onDragOver={handleWindowDragOver}
      onDragLeave={handleWindowDragLeave}
      onDrop={handleWindowDrop}
      className="flex flex-col w-screen h-screen overflow-hidden bg-[#e2e8f0] text-slate-900 font-sans relative"
    >
      {/* Global Drag & Drop Overlay */}
      {isGlobalDragOver && (
        <div
          onDragOver={(e) => e.preventDefault()}
          onDragLeave={() => setIsGlobalDragOver(false)}
          onDrop={handleWindowDrop}
          className="fixed inset-0 z-[500] bg-blue-900/60 backdrop-blur-xs flex items-center justify-center p-8"
        >
          <div className="bg-white border-4 border-dashed border-blue-600 rounded-xl shadow-2xl px-12 py-10 text-center max-w-lg pointer-events-none">
            <Upload className="w-12 h-12 text-blue-700 mx-auto mb-3 animate-bounce" />
            <h3 className="text-lg font-bold text-slate-900">
              Rilascia i file per importarli in OriginStudio Lab
            </h3>
            <p className="text-xs text-slate-600 mt-1.5">
              Supporta file singoli o multipli <strong>.CSV, .TXT, .TSV, .DAT</strong>, fogli{' '}
              <strong>Excel (.XLSX / .XLS)</strong> e progetti completi{' '}
              <strong>.ORIGINCLONE / .OPJ.JSON</strong>.
            </p>
          </div>
        </div>
      )}

      {/* 1. PERSISTENT ORIGINLAB DROPDOWN TOP MENU BAR */}
      <TopMenuBar
        onOpenImportWizard={(clipboardMode = false) => {
          setImportClipboardMode(clipboardMode);
          setIsImportModalOpen(true);
        }}
        onNewSheet={handleAddSheet}
        onExportActiveSheet={(fmt) => {
          if (activeSheet) handleExportSheet(activeSheet, fmt);
        }}
        onCopyActiveSheetToClipboard={handleCopyActiveSheetToClipboard}
        onSaveProject={handleSaveProjectBundle}
        onSaveBinaryOriginClone={handleSaveBinaryOriginClone}
        onOpenProjectFile={() => projectInputRef.current?.click()}
        onResetDemo={handleResetDemo}
        onGenerateHighPerfDataset={handleGenerateHighPerfDataset}
        onAddColumnRole={(role) => {
          if (!activeSheet) return;
          const idx = activeSheet.columns.length;
          const shortName = getColumnShortName(idx);
          const firstX = activeSheet.columns.find((c) => c.role === 'X');
          const rowCount = Math.max(20, ...activeSheet.columns.map((c) => c.values.length));
          const newCol: WorkbookSheet['columns'][0] = {
            id: `col_${Date.now()}`,
            shortName,
            longName:
              role === 'X'
                ? `Asse X (${shortName})`
                : role === 'Y_ERR'
                ? `Errore Y (${shortName})`
                : role === 'X_ERR'
                ? `Errore X (${shortName})`
                : role === 'LABEL'
                ? `Etichetta (${shortName})`
                : role === 'IGNORE'
                ? `Note (${shortName})`
                : `Segnale Y (${shortName})`,
            unit: '',
            comment: '',
            role,
            dataType: role === 'LABEL' ? 'text' : 'numeric',
            linkedXColId: role !== 'X' ? firstX?.id : undefined,
            values: Array(rowCount).fill(null),
            rawStrings: role === 'LABEL' ? Array(rowCount).fill(null) : undefined,
          };
          const rebound = rebindSheetMultiXAssociations({
            ...activeSheet,
            columns: [...activeSheet.columns, newCol],
          });
          handleUpdateSheet(rebound);
          triggerToast(`Aggiunta colonna ${shortName}(${role}) a ${activeSheet.name}`);
        }}
        onAddRows={(count) => {
          if (!activeSheet) return;
          const updatedCols = activeSheet.columns.map((c) => ({
            ...c,
            values: [...c.values, ...Array(count).fill(null)],
            rawStrings: c.rawStrings ? [...c.rawStrings, ...Array(count).fill(null)] : undefined,
          }));
          handleUpdateSheet({ ...activeSheet, columns: updatedCols });
          triggerToast(`Aggiunte ${count} righe a ${activeSheet.name}`);
        }}
        onQuickSortSheet={(direction) => {
          if (!activeSheet || !activeSheet.columns[0]) return;
          const xCol = activeSheet.columns.find((c) => c.role === 'X') || activeSheet.columns[0];
          const sorted = sortEntireSheetByColumn(activeSheet, xCol.id, direction);
          handleUpdateSheet(sorted);
          triggerToast(
            `Foglio ${activeSheet.name} ordinato (${
              direction === 'asc' ? 'Crescente ↑' : 'Decrescente ↓'
            }) per ${xCol.shortName}`
          );
        }}
        onQuickAutoFillX={(mode) => {
          if (!activeSheet || !activeSheet.columns[0]) return;
          const rowCount = Math.max(20, ...activeSheet.columns.map((c) => c.values.length));
          if (mode === 'row_numbers') {
            const xCol = activeSheet.columns.find((c) => c.role === 'X') || activeSheet.columns[0];
            const filled = applyAutoFillToColumn(xCol, { mode: 'row_numbers', rowCount });
            handleUpdateSheet({
              ...activeSheet,
              columns: activeSheet.columns.map((c) => (c.id === xCol.id ? filled : c)),
            });
            triggerToast(`Colonna ${xCol.shortName} riempita con numeri di riga 1..${rowCount}`);
          } else {
            const yCol =
              activeSheet.columns.find((c) => c.role === 'Y') || activeSheet.columns[0];
            const filled = applyAutoFillToColumn(yCol, {
              mode: 'random_normal',
              rowCount,
              mean: 0,
              stdDev: 1,
            });
            handleUpdateSheet({
              ...activeSheet,
              columns: activeSheet.columns.map((c) => (c.id === yCol.id ? filled : c)),
            });
            triggerToast(`Colonna ${yCol.shortName} riempita con sequenza casuale normale`);
          }
        }}
        onUnmaskAllActiveSheet={() => {
          if (!activeSheet) return;
          const unmasked = applyMaskingToSheet(activeSheet, {
            targetColId: 'ALL_Y',
            mode: 'unmask_all',
          });
          const allUnmasked = {
            ...unmasked,
            columns: unmasked.columns.map((c) => ({ ...c, maskedIndices: [] })),
          };
          handleUpdateSheet(allUnmasked);
          triggerToast(`Tutti i punti mascherati di ${activeSheet.name} sono stati ripristinati`);
        }}
        canUndoRemovedPoint={removedPointsLog.length > 0}
        onUndoRemovedPoint={() => {
          if (removedPointsLog[0]) handleRestoreRemovedPoint(removedPointsLog[0].id);
        }}
        showLeftExplorer={showLeftExplorer}
        onToggleLeftExplorer={() => setShowLeftExplorer((v) => !v)}
        showRightInspector={showRightInspector}
        onToggleRightInspector={() => setShowRightInspector((v) => !v)}
        showLegend={Boolean(activeGraph?.showLegend)}
        onToggleLegend={() => {
          if (activeGraph) {
            handleUpdateGraph({ ...activeGraph, showLegend: !activeGraph.showLegend });
          }
        }}
        showFitReportBox={activeGraph?.showFitReportBox !== false}
        onToggleFitReportBox={() => {
          if (activeGraph) {
            handleUpdateGraph({
              ...activeGraph,
              showFitReportBox: activeGraph.showFitReportBox === false ? true : false,
            });
          }
        }}
        showMajorGrid={Boolean(activeGraph?.xAxis.showMajorGrid)}
        onToggleMajorGrid={() => {
          if (activeGraph) {
            const nextGrid = !activeGraph.xAxis.showMajorGrid;
            handleUpdateGraph({
              ...activeGraph,
              xAxis: { ...activeGraph.xAxis, showMajorGrid: nextGrid },
              yAxis: { ...activeGraph.yAxis, showMajorGrid: nextGrid },
            });
          }
        }}
        onRescaleGraph={() => {
          if (activeGraph) {
            handleUpdateGraph({
              ...activeGraph,
              xAxis: { ...activeGraph.xAxis, min: 'auto', max: 'auto' },
              yAxis: { ...activeGraph.yAxis, min: 'auto', max: 'auto' },
            });
            triggerToast('Assi X e Y riscalati automaticamente ai dati');
          }
        }}
        onQuickChangePlotType={(plotType) => {
          if (!activeGraph || activeGraph.series.length === 0) return;
          const connection: ConnectionType =
            plotType === 'spline' ? 'spline' : plotType === 'step' ? 'step' : 'linear';
          handleUpdateGraph({
            ...activeGraph,
            series: activeGraph.series.map((s) => ({
              ...s,
              plotType,
              connection,
              fillArea: plotType === 'area' ? true : s.fillArea,
            })),
          });
          triggerToast(`Tipo grafico impostato su "${plotType}"`);
        }}
        onPlotSelectedToNewGraph={() => {
          if (!activeSheet) return;
          const yIds = activeSheet.columns.filter((c) => c.role === 'Y').map((c) => c.id);
          if (yIds.length > 0) {
            handlePlotColumnsToNewGraph(activeSheet.id, yIds, 'line_symbol', 'linear');
          }
        }}
        onOpenPlotDetails={() => setIsPlotDetailsOpen(true)}
        onOpenCurveFitInspector={() => {
          setShowRightInspector(true);
          setInspectorTab('fitting');
        }}
        onQuickFitActiveSeries={handleQuickFitActiveSeries}
        onRunDerivativeOnActive={() => {
          if (!activeSheet) return;
          const yCol = activeSheet.columns.find((c) => c.role === 'Y');
          if (!yCol) return;
          const xCol = resolveXColumnForY(activeSheet, yCol);
          if (xCol) handleApplyDerivative(activeSheet.id, xCol.id, yCol.id);
        }}
        onRunIntegralOnActive={() => {
          if (!activeSheet) return;
          const yCol = activeSheet.columns.find((c) => c.role === 'Y');
          if (!yCol) return;
          const xCol = resolveXColumnForY(activeSheet, yCol);
          if (xCol) handleApplyIntegration(activeSheet.id, xCol.id, yCol.id);
        }}
        onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
        onQuickFilterActive={(method) => {
          if (!activeSheet) return;
          const yCol = activeSheet.columns.find((c) => c.role === 'Y');
          if (!yCol) return;
          const labelMap: Record<string, string> = {
            savitzky_golay: 'Savitzky-Golay (W=7, g=2)',
            moving_average: 'Media Mobile (W=7)',
            gaussian: 'Gaussiano (W=7)',
            median: 'Mediana Anti-Spike (W=7)',
            outlier_zscore: 'Filtro Outlier MAD (2.0σ)',
          };
          applySignalFilterInWorker({
            col: yCol,
            method,
            windowSize: 7,
            polyOrder: 2,
            cutoffRatio: 0.22,
            outlierSigma: 2.0,
          })
            .then((res) => {
              const filteredValues: (number | null)[] = new Array(res.float64.length);
              for (let i = 0; i < res.float64.length; i++) {
                filteredValues[i] = Number.isFinite(res.float64[i]) ? res.float64[i] : null;
              }
              handleApplyFilterSmoothResult({
                sheetId: activeSheet.id,
                sourceColId: yCol.id,
                filteredValues,
                methodLabel: labelMap[method] || method,
                outputMode: 'new-col-and-plot',
                color: '#059669',
              });
            })
            .catch(() => {
              const fallback = applySignalFilterAndSmooth({
                values: yCol.values,
                method,
                windowSize: 7,
                polyOrder: 2,
                cutoffAlpha: 0.22,
                zScoreThreshold: 2.0,
              });
              handleApplyFilterSmoothResult({
                sheetId: activeSheet.id,
                sourceColId: yCol.id,
                filteredValues: fallback.filtered,
                methodLabel: labelMap[method] || method,
                outputMode: 'new-col-and-plot',
                color: '#059669',
              });
            });
        }}
        onOpenColumnStats={() => {
          if (activeSheet) {
            const stats = activeSheet.columns
              .filter((c) => c.role === 'X' || c.role === 'Y' || c.role === 'Y_ERR')
              .map(computeColumnStatistics);
            setActiveStatsData(stats);
            setIsStatsModalOpen(true);
          }
        }}
        onDuplicateActiveSheet={() => {
          if (activeSheet) handleDuplicateSheet(activeSheet.id);
        }}
        onOpenWorkbookSubPanel={(panel) => {
          if (dockLayout === 'graph-only') {
            setDockLayout(previousDockLayout);
          }
          setExternalSubPanelRequest({ panel, nonce: Date.now() });
        }}
        onRecalculateActiveSheetFormulas={() => {
          if (!activeSheet) return;
          const recalc = recalculateSheetFormulas(activeSheet);
          handleUpdateSheet(recalc);
          triggerToast(`Ricalcolate tutte le colonne formula F(x) di ${activeSheet.name}`);
        }}
        onQuickCreateCalculatedColumn={(preset) => {
          if (!activeSheet) return;
          const xCol = activeSheet.columns.find((c) => c.role === 'X') || activeSheet.columns[0];
          const yCol =
            activeSheet.columns.find((c) => c.role === 'Y') ||
            activeSheet.columns[1] ||
            activeSheet.columns[0];
          if (!xCol || !yCol) return;

          const expr =
            preset === 'norm'
              ? `norm(col(${yCol.shortName}))`
              : preset === 'zscore'
              ? `zscore(col(${yCol.shortName}))`
              : `col(${xCol.shortName}) * 2 + sin(col(${yCol.shortName}))`;

          const label =
            preset === 'norm'
              ? `${yCol.longName || yCol.shortName} [Norm 0..1]`
              : preset === 'zscore'
              ? `${yCol.longName || yCol.shortName} [Z-Score]`
              : `2·${xCol.shortName} + sin(${yCol.shortName})`;

          const newShort = getColumnShortName(activeSheet.columns.length);
          const computed = evaluateColumnFormula(expr, yCol, activeSheet.columns);
          if (!computed) return;

          const f64 = new Float64Array(computed.length);
          const u8 = new Uint8Array(computed.length);
          for (let i = 0; i < computed.length; i++) {
            const v = computed[i];
            if (v !== null && Number.isFinite(v)) {
              f64[i] = v;
              u8[i] = 0;
            } else {
              f64[i] = Number.NaN;
              u8[i] = 2;
            }
          }

          const newCol = attachBinaryBuffersToColumn(
            {
              id: `col_calc_${Date.now()}`,
              shortName: newShort,
              longName: label,
              unit: preset === 'norm' || preset === 'zscore' ? '' : yCol.unit,
              comment: `F(x): ${expr}`,
              formula: expr,
              role: 'Y',
              dataType: 'numeric',
              linkedXColId: xCol.id,
              values: computed,
            },
            f64,
            u8
          );

          const updated = recalculateSheetFormulas(
            rebindSheetMultiXAssociations({
              ...activeSheet,
              columns: [...activeSheet.columns, newCol],
            })
          );
          handleUpdateSheet(updated);
          triggerToast(`Creata colonna calcolata ${newShort}(Y): ${expr}`);
        }}
        toolMode={toolMode}
        onChangeToolMode={(mode) => {
          setToolMode(mode);
          if (mode === 'remove_bad') {
            setShowRightInspector(true);
            setInspectorTab('badpoints');
          }
        }}
        dockLayout={dockLayout}
        onChangeDockLayout={handleSwitchDockLayout}
        onTileMdiWindows={handleTileMdiWindows}
        onCascadeMdiWindows={handleCascadeMdiWindows}
        toastMessage={toastMessage}
      />

      {/* Hidden Project File Input (.originclone / .opj.json / .json) */}
      <input
        ref={projectInputRef}
        type="file"
        accept=".originclone,.zip,.json,.opj.json"
        onChange={handleLoadProjectJson}
        className="hidden"
      />

      {/* 2. COMPACT ORIGINLAB WORKSPACE & SPATIAL WINDOW MANAGER TOOLBAR */}
      <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-1 bg-[#f1f5f9] border-b border-slate-300 text-xs shrink-0 select-none">
        {/* Left: Spatial Layout Switcher & Splitter Controls */}
        <div className="flex items-center gap-1.5 flex-wrap">
          <button
            type="button"
            onClick={() => setShowLeftExplorer((v) => !v)}
            className={`flex items-center gap-1 px-2 py-0.5 rounded border text-[11px] font-medium transition-colors ${
              showLeftExplorer
                ? 'bg-white text-slate-900 border-slate-300 shadow-2xs font-semibold'
                : 'bg-transparent text-slate-600 border-transparent hover:bg-slate-200'
            }`}
            title="Mostra / Nascondi Project Explorer laterale"
          >
            <FolderTree className="w-3.5 h-3.5 text-blue-700" />
            <span>Explorer</span>
          </button>

          <span className="text-slate-300" aria-hidden="true">
            |
          </span>

          {/* Spatial Window Management Modes */}
          <div className="flex items-center gap-0.5 bg-slate-200/80 p-0.5 rounded border border-slate-300">
            <button
              type="button"
              onClick={() => handleSwitchDockLayout('split-horizontal')}
              className={`flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
                dockLayout === 'split-horizontal'
                  ? 'bg-white text-slate-900 font-semibold shadow-2xs'
                  : 'text-slate-700 hover:text-slate-900'
              }`}
              title="Affiancata ridimensionabile (Tabella a sinistra | Grafico a destra)"
            >
              <Columns className="w-3 h-3 text-blue-700" />
              <span>Affiancata</span>
            </button>

            <button
              type="button"
              onClick={() => handleSwitchDockLayout('split-vertical')}
              className={`flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
                dockLayout === 'split-vertical'
                  ? 'bg-white text-slate-900 font-semibold shadow-2xs'
                  : 'text-slate-700 hover:text-slate-900'
              }`}
              title="Impilata ridimensionabile (Grafico sopra | Tabella sotto)"
            >
              <Rows className="w-3 h-3 text-blue-700" />
              <span>Impilata</span>
            </button>

            <button
              type="button"
              onClick={() => handleSwitchDockLayout('mdi-free')}
              className={`flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
                dockLayout === 'mdi-free'
                  ? 'bg-blue-700 text-white font-semibold shadow-2xs'
                  : 'text-slate-700 hover:text-slate-900'
              }`}
              title="Finestre Libere MDI in stile OriginLab Desktop (Sposta, Ridimensiona, Affianca, Cascata)"
            >
              <AppWindow className="w-3 h-3" />
              <span>Finestre Libere (MDI)</span>
            </button>

            <button
              type="button"
              onClick={() => handleSwitchDockLayout('multi-grid')}
              className={`flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
                dockLayout === 'multi-grid'
                  ? 'bg-white text-slate-900 font-semibold shadow-2xs'
                  : 'text-slate-700 hover:text-slate-900'
              }`}
              title="Mosaico Multi-Pannello automatico"
            >
              <LayoutGrid className="w-3 h-3 text-blue-700" />
              <span>Mosaico</span>
            </button>

            <button
              type="button"
              onClick={() => handleSwitchDockLayout('sheet-only')}
              className={`flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
                dockLayout === 'sheet-only'
                  ? 'bg-white text-slate-900 font-semibold shadow-2xs'
                  : 'text-slate-700 hover:text-slate-900'
              }`}
              title="Massimizza solo la Tabella Dati"
            >
              <FileSpreadsheet className="w-3 h-3 text-blue-700" />
              <span>Solo Tabella</span>
            </button>

            <button
              type="button"
              onClick={() => handleSwitchDockLayout('graph-only')}
              className={`flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
                dockLayout === 'graph-only'
                  ? 'bg-white text-slate-900 font-semibold shadow-2xs'
                  : 'text-slate-700 hover:text-slate-900'
              }`}
              title="Massimizza solo il Grafico"
            >
              <LineChart className="w-3 h-3 text-blue-700" />
              <span>Solo Grafico</span>
            </button>
          </div>

          {/* Quick Split Ratio Presets when in Split Mode */}
          {(dockLayout === 'split-horizontal' || dockLayout === 'split-vertical') && (
            <div className="hidden xl:flex items-center gap-1 text-[10.5px] text-slate-600 bg-white px-1.5 py-0.5 rounded border border-slate-300">
              <span>Proporzione:</span>
              {[35, 45, 50, 60].map((r) => (
                <button
                  key={r}
                  type="button"
                  onClick={() => setSplitRatio(r)}
                  className={`px-1 rounded font-mono ${
                    Math.abs(splitRatio - r) <= 2
                      ? 'bg-blue-100 text-blue-900 font-bold'
                      : 'hover:bg-slate-100'
                  }`}
                >
                  {r}/{100 - r}
                </button>
              ))}
            </div>
          )}

          {/* MDI Window Arrangement Actions when in MDI mode */}
          {dockLayout === 'mdi-free' && (
            <div className="flex items-center gap-1 bg-blue-50 px-1.5 py-0.5 rounded border border-blue-200">
              <button
                type="button"
                onClick={handleTileMdiWindows}
                className="px-1.5 py-0.5 text-[10.5px] font-semibold text-blue-900 hover:bg-blue-100 rounded"
                title="Disponi tutte le finestre aperte in griglia ordinata senza sovrapposizioni"
              >
                Affianca in Griglia
              </button>
              <button
                type="button"
                onClick={handleCascadeMdiWindows}
                className="px-1.5 py-0.5 text-[10.5px] font-semibold text-blue-900 hover:bg-blue-100 rounded"
                title="Disponi finestre a cascata (Window > Cascade)"
              >
                Cascata
              </button>
            </div>
          )}
        </div>

        {/* Right: Project Save/Load, Demo Reset & Inspector Toggle */}
        <div className="flex items-center gap-1.5">
          <button
            type="button"
            onClick={handleSaveProjectBundle}
            className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 transition-colors"
            title="Salva progetto completo in .opj.json (fogli, grafici, fit e cronologia)"
          >
            <Download className="w-3 h-3" />
            <span className="hidden sm:inline">Salva Progetto</span>
          </button>

          <button
            type="button"
            onClick={() => projectInputRef.current?.click()}
            className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 transition-colors"
            title="Apri progetto .opj.json"
          >
            <Upload className="w-3 h-3" />
            <span className="hidden sm:inline">Apri Progetto</span>
          </button>

          <button
            type="button"
            onClick={handleResetDemo}
            className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium text-slate-700 hover:text-slate-900 hover:bg-slate-200 rounded transition-colors"
            title="Ripristina dataset sperimentale di calibrazione"
          >
            <RotateCcw className="w-3 h-3" />
            <span className="hidden md:inline">Reset Demo</span>
          </button>

          <span className="text-slate-300" aria-hidden="true">
            |
          </span>

          <button
            type="button"
            onClick={() => setShowRightInspector((v) => !v)}
            className={`flex items-center gap-1 px-2 py-0.5 rounded border text-[11px] font-medium transition-colors ${
              showRightInspector
                ? 'bg-white text-slate-900 border-slate-300 shadow-2xs font-semibold'
                : 'bg-transparent text-slate-600 border-transparent hover:bg-slate-200'
            }`}
            title="Mostra / Nascondi Pannello Ispettore destro"
          >
            <SlidersHorizontal className="w-3.5 h-3.5 text-blue-700" />
            <span>Ispettore</span>
          </button>
        </div>
      </div>

      {/* 3. MAIN WORKBENCH: Left Project Explorer + Central Flexible Workspace + Right Inspector */}
      <div className="flex-1 flex min-h-0 overflow-hidden">
        {/* Left Project Explorer */}
        {showLeftExplorer && (
          <aside className="w-52 shrink-0 bg-[#f8fafc] border-r border-slate-300 flex flex-col select-none overflow-y-auto p-2 space-y-3 text-xs">
            {/* Workbooks Tree */}
            <div>
              <div className="flex items-center justify-between px-1 mb-1 text-slate-700 font-bold text-[11px]">
                <span>Fogli Dati ({sheets.length})</span>
                <button
                  type="button"
                  onClick={handleAddSheet}
                  className="p-0.5 text-slate-500 hover:text-blue-700 rounded"
                  title="Nuovo foglio dati"
                >
                  <Plus className="w-3.5 h-3.5" />
                </button>
              </div>
              <div className="space-y-0.5">
                {sheets.map((s) => {
                  const isRenaming =
                    explorerRenaming?.kind === 'sheet' && explorerRenaming.id === s.id;
                  return (
                    <div
                      key={s.id}
                      onClick={() => {
                        setActiveSheetId(s.id);
                        if (dockLayout === 'graph-only') setDockLayout(previousDockLayout);
                        if (dockLayout === 'mdi-free') bringMdiWindowToFront(`mdi_sheet_${s.id}`);
                      }}
                      onDoubleClick={(e) => {
                        e.stopPropagation();
                        setExplorerRenaming({ kind: 'sheet', id: s.id, value: s.name });
                      }}
                      title="Clic per selezionare · Doppio clic per rinominare il foglio"
                      className={`group w-full flex items-center justify-between gap-1 px-2 py-1 rounded text-left cursor-pointer transition-colors ${
                        s.id === activeSheetId
                          ? 'bg-white text-slate-900 font-bold border border-slate-300 shadow-2xs'
                          : 'text-slate-700 hover:bg-slate-200/60'
                      }`}
                    >
                      <div className="flex items-center gap-1.5 min-w-0 flex-1">
                        <FileSpreadsheet
                          className={`w-3.5 h-3.5 shrink-0 ${
                            s.isFitOutput ? 'text-red-600' : 'text-blue-700'
                          }`}
                        />
                        {isRenaming ? (
                          <input
                            type="text"
                            value={explorerRenaming.value}
                            autoFocus
                            onClick={(e) => e.stopPropagation()}
                            onChange={(e) =>
                              setExplorerRenaming({
                                ...explorerRenaming,
                                value: e.target.value,
                              })
                            }
                            onBlur={() => {
                              if (explorerRenaming.value.trim()) {
                                handleRenameSheet(s.id, explorerRenaming.value);
                              }
                              setExplorerRenaming(null);
                            }}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                if (explorerRenaming.value.trim()) {
                                  handleRenameSheet(s.id, explorerRenaming.value);
                                }
                                setExplorerRenaming(null);
                              } else if (e.key === 'Escape') {
                                setExplorerRenaming(null);
                              }
                            }}
                            className="w-full px-1 py-0 text-[11px] font-semibold bg-white border border-blue-600 rounded text-slate-900 focus:outline-none"
                          />
                        ) : (
                          <span className="truncate">{s.name}</span>
                        )}
                      </div>
                      <div className="flex items-center gap-1 shrink-0">
                        <span className="font-mono text-[10px] text-slate-400">
                          {s.columns.length}c
                        </span>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            handleDeleteSheet(s.id);
                          }}
                          className="w-4 h-4 flex items-center justify-center rounded text-slate-400 hover:text-red-600 hover:bg-red-50 font-bold leading-none transition-colors"
                          title={`Elimina foglio "${s.name}"`}
                        >
                          ×
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Graphs Tree */}
            <div>
              <div className="flex items-center justify-between px-1 mb-1 text-slate-700 font-bold text-[11px]">
                <span>Finestre Grafico ({graphs.length})</span>
                <button
                  type="button"
                  onClick={handleAddGraph}
                  className="p-0.5 text-slate-500 hover:text-teal-700 rounded"
                  title="Nuova finestra grafico"
                >
                  <Plus className="w-3.5 h-3.5" />
                </button>
              </div>
              <div className="space-y-0.5">
                {graphs.map((g) => {
                  const isRenaming =
                    explorerRenaming?.kind === 'graph' && explorerRenaming.id === g.id;
                  return (
                    <div
                      key={g.id}
                      onClick={() => {
                        setActiveGraphId(g.id);
                        if (dockLayout === 'sheet-only') setDockLayout(previousDockLayout);
                        if (dockLayout === 'mdi-free') bringMdiWindowToFront(`mdi_graph_${g.id}`);
                      }}
                      onDoubleClick={(e) => {
                        e.stopPropagation();
                        setExplorerRenaming({ kind: 'graph', id: g.id, value: g.name });
                      }}
                      title="Clic per selezionare · Doppio clic per rinominare il grafico"
                      className={`group w-full flex items-center justify-between gap-1 px-2 py-1 rounded text-left cursor-pointer transition-colors ${
                        g.id === activeGraphId
                          ? 'bg-white text-slate-900 font-bold border border-slate-300 shadow-2xs'
                          : 'text-slate-700 hover:bg-slate-200/60'
                      }`}
                    >
                      <div className="flex items-center gap-1.5 min-w-0 flex-1">
                        <LineChart className="w-3.5 h-3.5 text-teal-700 shrink-0" />
                        {isRenaming ? (
                          <input
                            type="text"
                            value={explorerRenaming.value}
                            autoFocus
                            onClick={(e) => e.stopPropagation()}
                            onChange={(e) =>
                              setExplorerRenaming({
                                ...explorerRenaming,
                                value: e.target.value,
                              })
                            }
                            onBlur={() => {
                              if (explorerRenaming.value.trim()) {
                                handleRenameGraph(g.id, explorerRenaming.value);
                              }
                              setExplorerRenaming(null);
                            }}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                if (explorerRenaming.value.trim()) {
                                  handleRenameGraph(g.id, explorerRenaming.value);
                                }
                                setExplorerRenaming(null);
                              } else if (e.key === 'Escape') {
                                setExplorerRenaming(null);
                              }
                            }}
                            className="w-full px-1 py-0 text-[11px] font-semibold bg-white border border-blue-600 rounded text-slate-900 focus:outline-none"
                          />
                        ) : (
                          <span className="truncate">{g.name}</span>
                        )}
                      </div>
                      <div className="flex items-center gap-1 shrink-0">
                        <span className="font-mono text-[10px] text-slate-400">
                          {g.series.length}s
                        </span>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            handleDeleteGraph(g.id);
                          }}
                          className="w-4 h-4 flex items-center justify-center rounded text-slate-400 hover:text-red-600 hover:bg-red-50 font-bold leading-none transition-colors"
                          title={`Elimina grafico "${g.name}"`}
                        >
                          ×
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Origin Scientific Quick Tool Card */}
            <div className="p-2 bg-white border border-slate-300 rounded space-y-1.5">
              <div className="font-bold text-slate-800 text-[10.5px]">Accesso Rapido</div>

              <button
                type="button"
                onClick={() => {
                  setImportClipboardMode(false);
                  setIsImportModalOpen(true);
                }}
                className="w-full flex items-center gap-1.5 px-2 py-1 rounded text-[11px] font-semibold bg-blue-700 text-white hover:bg-blue-800 transition-colors"
              >
                <Upload className="w-3.5 h-3.5 shrink-0" />
                <span className="truncate">Importa File / Excel</span>
              </button>

              <button
                type="button"
                onClick={() => setIsFilterSmoothOpen(true)}
                className="w-full flex items-center gap-1.5 px-2 py-1 rounded text-[11px] font-semibold bg-emerald-50 text-emerald-900 border border-emerald-200 hover:bg-emerald-100 transition-colors"
              >
                <Sparkles className="w-3.5 h-3.5 shrink-0 text-emerald-700" />
                <span className="truncate">Filtro & Smoothing</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  setToolMode((prev) => (prev === 'remove_bad' ? 'pointer' : 'remove_bad'));
                  setShowRightInspector(true);
                  setInspectorTab('badpoints');
                }}
                className={`w-full flex items-center gap-1.5 px-2 py-1 rounded text-[11px] font-semibold transition-colors ${
                  toolMode === 'remove_bad'
                    ? 'bg-red-600 text-white'
                    : 'bg-red-50 text-red-800 hover:bg-red-100'
                }`}
              >
                <Target className="w-3.5 h-3.5 shrink-0" />
                <span className="truncate">Remove Bad Points</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  setShowRightInspector(true);
                  setInspectorTab('fitting');
                }}
                className="w-full flex items-center gap-1.5 px-2 py-1 rounded text-[11px] font-semibold bg-blue-50 text-blue-800 hover:bg-blue-100 transition-colors"
              >
                <TrendingUp className="w-3.5 h-3.5 shrink-0" />
                <span className="truncate">Curve Fit [C, D] (n pt)</span>
              </button>

              <button
                type="button"
                onClick={() => setIsPlotDetailsOpen(true)}
                className="w-full flex items-center gap-1.5 px-2 py-1 rounded text-[11px] font-semibold bg-slate-100 text-slate-800 hover:bg-slate-200 transition-colors"
              >
                <Sliders className="w-3.5 h-3.5 shrink-0 text-blue-700" />
                <span className="truncate">Tipi Grafico & Stile</span>
              </button>
            </div>
          </aside>
        )}

        {/* Central Spatial Stage */}
        <main
          ref={centralStageRef}
          className="flex-1 min-w-0 h-full p-1.5 overflow-hidden bg-[#cbd5e1]/50 relative"
        >
          {/* MODE 1: RESIZABLE HORIZONTAL SPLIT (Affiancata) */}
          {dockLayout === 'split-horizontal' && (
            <div className="flex flex-col lg:flex-row h-full w-full min-h-0">
              <div
                style={{ width: `${splitRatio}%` }}
                className="h-full min-h-0 min-w-[260px] max-lg:!w-full"
              >
                <WorkbookView
                  sheet={activeSheet}
                  allSheets={sheets}
                  activeSheetId={activeSheet.id}
                  onSelectSheet={setActiveSheetId}
                  onAddSheet={handleAddSheet}
                  onCreateMultipleSheets={handleCreateMultipleSheets}
                  onDuplicateSheet={handleDuplicateSheet}
                  onDeleteSheet={handleDeleteSheet}
                  onRenameSheet={handleRenameSheet}
                  onUpdateSheet={handleUpdateSheet}
                  onPlotColumnsToActiveGraph={handlePlotColumnsToActiveGraph}
                  onPlotColumnsToNewGraph={handlePlotColumnsToNewGraph}
                  onExportSheet={handleExportSheet}
                  highlightedRowIndex={hoveredGraphPoint?.rowIndex ?? null}
                  highlightedSheetId={hoveredGraphPoint?.sheetId ?? null}
                  onOpenStatsModal={(stats) => {
                    setActiveStatsData(stats);
                    setIsStatsModalOpen(true);
                  }}
                  onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                  isMaximized={false}
                  onToggleMaximize={() => handleSwitchDockLayout('sheet-only')}
                  onSmartPasteNotification={triggerToast}
                  externalSubPanelRequest={externalSubPanelRequest}
                />
              </div>

              {/* Draggable Vertical Splitter Handle */}
              <div
                onMouseDown={() => setIsDraggingSplitter(true)}
                className="hidden lg:flex w-2 shrink-0 cursor-col-resize items-center justify-center group hover:bg-blue-400/40 transition-colors select-none"
                title="Trascina per ridimensionare lo spazio tra Tabella e Grafico"
              >
                <div className="w-0.5 h-10 bg-slate-400 group-hover:bg-blue-700 rounded-full" />
              </div>

              <div
                style={{ width: `${100 - splitRatio}%` }}
                className="h-full min-h-0 min-w-[300px] max-lg:!w-full"
              >
                <GraphCanvas
                  graph={activeGraph}
                  allGraphs={graphs}
                  activeGraphId={activeGraph.id}
                  sheets={sheets}
                  toolMode={toolMode}
                  onChangeToolMode={setToolMode}
                  onSelectGraph={setActiveGraphId}
                  onUpdateGraph={handleUpdateGraph}
                  onRemoveBadPoint={(pt) => handleRemoveBadPointsBatch([pt])}
                  onRemoveBadPointsBatch={handleRemoveBadPointsBatch}
                  onMaskPoint={handleMaskPoint}
                  onUndoLastRemovedPoint={() => {
                    if (removedPointsLog[0]) handleRestoreRemovedPoint(removedPointsLog[0].id);
                  }}
                  canUndoRemovedPoint={removedPointsLog.length > 0}
                  onHoverPoint={setHoveredGraphPoint}
                  onSelectSeriesInInspector={(sId) => {
                    setSelectedSeriesId(sId);
                    setShowRightInspector(true);
                  }}
                  onOpenPlotDetailsModal={() => setIsPlotDetailsOpen(true)}
                  onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                  onQuickFitActiveSeries={handleQuickFitActiveSeries}
                  isMaximized={false}
                  onToggleMaximize={() => handleSwitchDockLayout('graph-only')}
                  onRenameGraph={handleRenameGraph}
                  onDeleteGraph={handleDeleteGraph}
                  onAddGraph={handleAddGraph}
                />
              </div>
            </div>
          )}

          {/* MODE 2: RESIZABLE VERTICAL SPLIT (Impilata) */}
          {dockLayout === 'split-vertical' && (
            <div className="flex flex-col h-full w-full min-h-0">
              <div style={{ height: `${splitRatio}%` }} className="w-full min-h-[200px]">
                <GraphCanvas
                  graph={activeGraph}
                  allGraphs={graphs}
                  activeGraphId={activeGraph.id}
                  sheets={sheets}
                  toolMode={toolMode}
                  onChangeToolMode={setToolMode}
                  onSelectGraph={setActiveGraphId}
                  onUpdateGraph={handleUpdateGraph}
                  onRemoveBadPoint={(pt) => handleRemoveBadPointsBatch([pt])}
                  onRemoveBadPointsBatch={handleRemoveBadPointsBatch}
                  onMaskPoint={handleMaskPoint}
                  onUndoLastRemovedPoint={() => {
                    if (removedPointsLog[0]) handleRestoreRemovedPoint(removedPointsLog[0].id);
                  }}
                  canUndoRemovedPoint={removedPointsLog.length > 0}
                  onHoverPoint={setHoveredGraphPoint}
                  onSelectSeriesInInspector={(sId) => {
                    setSelectedSeriesId(sId);
                    setShowRightInspector(true);
                  }}
                  onOpenPlotDetailsModal={() => setIsPlotDetailsOpen(true)}
                  onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                  onQuickFitActiveSeries={handleQuickFitActiveSeries}
                  isMaximized={false}
                  onToggleMaximize={() => handleSwitchDockLayout('graph-only')}
                  onRenameGraph={handleRenameGraph}
                  onDeleteGraph={handleDeleteGraph}
                  onAddGraph={handleAddGraph}
                />
              </div>

              {/* Draggable Horizontal Splitter Handle */}
              <div
                onMouseDown={() => setIsDraggingSplitter(true)}
                className="h-2 w-full shrink-0 cursor-row-resize flex items-center justify-center group hover:bg-blue-400/40 transition-colors select-none"
                title="Trascina per ridimensionare l'altezza tra Grafico e Tabella"
              >
                <div className="h-0.5 w-12 bg-slate-400 group-hover:bg-blue-700 rounded-full" />
              </div>

              <div style={{ height: `${100 - splitRatio}%` }} className="w-full min-h-[180px]">
                <WorkbookView
                  sheet={activeSheet}
                  allSheets={sheets}
                  activeSheetId={activeSheet.id}
                  onSelectSheet={setActiveSheetId}
                  onAddSheet={handleAddSheet}
                  onCreateMultipleSheets={handleCreateMultipleSheets}
                  onDuplicateSheet={handleDuplicateSheet}
                  onDeleteSheet={handleDeleteSheet}
                  onRenameSheet={handleRenameSheet}
                  onUpdateSheet={handleUpdateSheet}
                  onPlotColumnsToActiveGraph={handlePlotColumnsToActiveGraph}
                  onPlotColumnsToNewGraph={handlePlotColumnsToNewGraph}
                  onExportSheet={handleExportSheet}
                  highlightedRowIndex={hoveredGraphPoint?.rowIndex ?? null}
                  highlightedSheetId={hoveredGraphPoint?.sheetId ?? null}
                  onOpenStatsModal={(stats) => {
                    setActiveStatsData(stats);
                    setIsStatsModalOpen(true);
                  }}
                  onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                  isMaximized={false}
                  onToggleMaximize={() => handleSwitchDockLayout('sheet-only')}
                  onSmartPasteNotification={triggerToast}
                  externalSubPanelRequest={externalSubPanelRequest}
                />
              </div>
            </div>
          )}

          {/* MODE 3: ORIGINLAB MDI FREE DESKTOP WORKSPACE (Finestre Libere) */}
          {dockLayout === 'mdi-free' && (
            <div className="w-full h-full relative overflow-auto bg-[#cbd5e1]/70 rounded-xs border border-slate-300">
              {mdiWindows
                .filter((w) => !w.minimized)
                .map((win) => {
                  const isSheet = win.kind === 'sheet';
                  const targetSheet = isSheet
                    ? sheets.find((s) => s.id === win.targetId)
                    : undefined;
                  const targetGraph = !isSheet
                    ? graphs.find((g) => g.id === win.targetId)
                    : undefined;
                  if (isSheet && !targetSheet) return null;
                  if (!isSheet && !targetGraph) return null;

                  const isFocused =
                    (isSheet && activeSheetId === win.targetId) ||
                    (!isSheet && activeGraphId === win.targetId);

                  const winStyle: React.CSSProperties = win.maximized
                    ? {
                        left: 4,
                        top: 4,
                        width: 'calc(100% - 8px)',
                        height: 'calc(100% - 44px)',
                        zIndex: win.zIndex + 100,
                      }
                    : {
                        left: win.x,
                        top: win.y,
                        width: win.w,
                        height: win.h,
                        zIndex: win.zIndex,
                      };

                  return (
                    <div
                      key={win.id}
                      style={winStyle}
                      onMouseDown={() => {
                        bringMdiWindowToFront(win.id);
                        if (isSheet) setActiveSheetId(win.targetId);
                        else setActiveGraphId(win.targetId);
                      }}
                      className={`absolute flex flex-col bg-white rounded shadow-lg border-2 overflow-hidden transition-shadow ${
                        isFocused ? 'border-blue-700 ring-1 ring-blue-700/30' : 'border-slate-400'
                      }`}
                    >
                      {/* OriginLab MDI Window Title Bar */}
                      <div
                        onMouseDown={(e) => {
                          if (win.maximized) return;
                          e.stopPropagation();
                          bringMdiWindowToFront(win.id);
                          if (isSheet) setActiveSheetId(win.targetId);
                          else setActiveGraphId(win.targetId);
                          setMdiDragState({
                            winId: win.id,
                            mode: 'move',
                            startX: e.clientX,
                            startY: e.clientY,
                            origX: win.x,
                            origY: win.y,
                            origW: win.w,
                            origH: win.h,
                          });
                        }}
                        className={`flex items-center justify-between px-2.5 py-1 cursor-move select-none shrink-0 ${
                          isFocused
                            ? 'bg-gradient-to-r from-blue-800 to-blue-700 text-white'
                            : 'bg-slate-300 text-slate-800'
                        }`}
                      >
                        <div className="flex items-center gap-1.5 text-xs font-bold truncate">
                          {isSheet ? (
                            <FileSpreadsheet className="w-3.5 h-3.5 shrink-0" />
                          ) : (
                            <LineChart className="w-3.5 h-3.5 shrink-0" />
                          )}
                          <span className="truncate">
                            {isSheet
                              ? `[Workbook] ${targetSheet!.name}`
                              : `[Graph] ${targetGraph!.name} — ${targetGraph!.title}`}
                          </span>
                        </div>

                        <div
                          className="flex items-center gap-1 shrink-0"
                          onMouseDown={(e) => e.stopPropagation()}
                        >
                          <button
                            type="button"
                            onClick={() =>
                              setMdiWindows((prev) =>
                                prev.map((w) =>
                                  w.id === win.id ? { ...w, minimized: true, maximized: false } : w
                                )
                              )
                            }
                            className="p-0.5 rounded hover:bg-white/20"
                            title="Riduci finestra a icona nella barra inferiore"
                          >
                            <Minus className="w-3 h-3" />
                          </button>
                          <button
                            type="button"
                            onClick={() =>
                              setMdiWindows((prev) =>
                                prev.map((w) =>
                                  w.id === win.id ? { ...w, maximized: !w.maximized } : w
                                )
                              )
                            }
                            className="p-0.5 rounded hover:bg-white/20"
                            title={win.maximized ? 'Ripristina dimensioni' : 'Massimizza finestra'}
                          >
                            {win.maximized ? (
                              <Minimize2 className="w-3 h-3" />
                            ) : (
                              <Maximize2 className="w-3 h-3" />
                            )}
                          </button>
                        </div>
                      </div>

                      {/* Window Body */}
                      <div className="flex-1 min-h-0 overflow-hidden relative">
                        {isSheet && targetSheet ? (
                          <WorkbookView
                            sheet={targetSheet}
                            allSheets={sheets}
                            activeSheetId={targetSheet.id}
                            onSelectSheet={setActiveSheetId}
                            onAddSheet={handleAddSheet}
                            onCreateMultipleSheets={handleCreateMultipleSheets}
                            onDuplicateSheet={handleDuplicateSheet}
                            onDeleteSheet={handleDeleteSheet}
                            onRenameSheet={handleRenameSheet}
                            onUpdateSheet={handleUpdateSheet}
                            onPlotColumnsToActiveGraph={handlePlotColumnsToActiveGraph}
                            onPlotColumnsToNewGraph={handlePlotColumnsToNewGraph}
                            onExportSheet={handleExportSheet}
                            highlightedRowIndex={hoveredGraphPoint?.rowIndex ?? null}
                            highlightedSheetId={hoveredGraphPoint?.sheetId ?? null}
                            onOpenStatsModal={(stats) => {
                              setActiveStatsData(stats);
                              setIsStatsModalOpen(true);
                            }}
                            onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                            hideSheetTabs
                            onSmartPasteNotification={triggerToast}
                            externalSubPanelRequest={
                              targetSheet.id === activeSheet.id ? externalSubPanelRequest : null
                            }
                          />
                        ) : (
                          targetGraph && (
                            <GraphCanvas
                              graph={targetGraph}
                              allGraphs={graphs}
                              activeGraphId={targetGraph.id}
                              sheets={sheets}
                              toolMode={toolMode}
                              onChangeToolMode={setToolMode}
                              onSelectGraph={setActiveGraphId}
                              onUpdateGraph={handleUpdateGraph}
                              onRemoveBadPoint={(pt) => handleRemoveBadPointsBatch([pt])}
                              onRemoveBadPointsBatch={handleRemoveBadPointsBatch}
                              onMaskPoint={handleMaskPoint}
                              onUndoLastRemovedPoint={() => {
                                if (removedPointsLog[0])
                                  handleRestoreRemovedPoint(removedPointsLog[0].id);
                              }}
                              canUndoRemovedPoint={removedPointsLog.length > 0}
                              onHoverPoint={setHoveredGraphPoint}
                              onSelectSeriesInInspector={(sId) => {
                                setSelectedSeriesId(sId);
                                setShowRightInspector(true);
                              }}
                              onOpenPlotDetailsModal={() => setIsPlotDetailsOpen(true)}
                              onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                              onQuickFitActiveSeries={handleQuickFitActiveSeries}
                              hideGraphTabs
                            />
                          )
                        )}

                        {/* Bottom-Right Corner Resize Grip */}
                        {!win.maximized && (
                          <div
                            onMouseDown={(e) => {
                              e.stopPropagation();
                              bringMdiWindowToFront(win.id);
                              setMdiDragState({
                                winId: win.id,
                                mode: 'resize',
                                startX: e.clientX,
                                startY: e.clientY,
                                origX: win.x,
                                origY: win.y,
                                origW: win.w,
                                origH: win.h,
                              });
                            }}
                            className="absolute bottom-0 right-0 w-4 h-4 cursor-nwse-resize z-40 bg-slate-300/80 hover:bg-blue-600 transition-colors flex items-end justify-end p-0.5 rounded-tl"
                            title="Trascina l'angolo per ridimensionare la finestra"
                          >
                            <div className="w-2 h-2 border-r-2 border-b-2 border-slate-700" />
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}

              {/* Minimized MDI Windows Taskbar at Bottom */}
              <div className="absolute bottom-1 left-1 right-1 h-8 bg-slate-800/90 text-white rounded px-2 flex items-center justify-between gap-2 z-[200] text-xs">
                <div className="flex items-center gap-1.5 overflow-x-auto">
                  <span className="text-[10.5px] font-semibold text-slate-300 mr-1">
                    Finestre MDI ({mdiWindows.length}):
                  </span>
                  {mdiWindows.map((w) => {
                    const label =
                      w.kind === 'sheet'
                        ? sheets.find((s) => s.id === w.targetId)?.name || 'Foglio'
                        : graphs.find((g) => g.id === w.targetId)?.name || 'Grafico';
                    return (
                      <button
                        key={w.id}
                        type="button"
                        onClick={() => bringMdiWindowToFront(w.id)}
                        className={`flex items-center gap-1 px-2 py-0.5 rounded text-[11px] transition-colors ${
                          w.minimized
                            ? 'bg-amber-500/20 text-amber-200 border border-amber-400/40'
                            : 'bg-slate-700 text-white hover:bg-blue-600'
                        }`}
                      >
                        {w.kind === 'sheet' ? (
                          <FileSpreadsheet className="w-3 h-3 text-blue-300" />
                        ) : (
                          <LineChart className="w-3 h-3 text-teal-300" />
                        )}
                        <span>{label}</span>
                        {w.minimized && <span className="text-[9px] uppercase">(Icona)</span>}
                      </button>
                    );
                  })}
                </div>

                <div className="flex items-center gap-1 shrink-0">
                  <button
                    type="button"
                    onClick={handleTileMdiWindows}
                    className="px-2 py-0.5 bg-blue-600 hover:bg-blue-500 rounded text-[10.5px] font-semibold"
                  >
                    Affianca Tutte
                  </button>
                  <button
                    type="button"
                    onClick={handleCascadeMdiWindows}
                    className="px-2 py-0.5 bg-slate-700 hover:bg-slate-600 rounded text-[10.5px]"
                  >
                    Cascata
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* MODE 4: MULTI-GRID (Mosaico) */}
          {dockLayout === 'multi-grid' && (
            <div className="grid grid-cols-1 lg:grid-cols-2 grid-rows-2 gap-1.5 h-full">
              <div className="h-full min-h-0">
                <WorkbookView
                  sheet={sheets[0]}
                  allSheets={sheets}
                  activeSheetId={sheets[0].id}
                  onSelectSheet={setActiveSheetId}
                  onAddSheet={handleAddSheet}
                  onDeleteSheet={handleDeleteSheet}
                  onRenameSheet={handleRenameSheet}
                  onUpdateSheet={handleUpdateSheet}
                  onPlotColumnsToActiveGraph={handlePlotColumnsToActiveGraph}
                  onPlotColumnsToNewGraph={handlePlotColumnsToNewGraph}
                  onExportSheet={handleExportSheet}
                  highlightedRowIndex={hoveredGraphPoint?.rowIndex ?? null}
                  highlightedSheetId={hoveredGraphPoint?.sheetId ?? null}
                  onOpenStatsModal={(stats) => {
                    setActiveStatsData(stats);
                    setIsStatsModalOpen(true);
                  }}
                  onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                />
              </div>

              <div className="h-full min-h-0 row-span-2">
                <GraphCanvas
                  graph={activeGraph}
                  allGraphs={graphs}
                  activeGraphId={activeGraph.id}
                  sheets={sheets}
                  toolMode={toolMode}
                  onChangeToolMode={setToolMode}
                  onSelectGraph={setActiveGraphId}
                  onUpdateGraph={handleUpdateGraph}
                  onRemoveBadPoint={(pt) => handleRemoveBadPointsBatch([pt])}
                  onRemoveBadPointsBatch={handleRemoveBadPointsBatch}
                  onMaskPoint={handleMaskPoint}
                  onUndoLastRemovedPoint={() => {
                    if (removedPointsLog[0]) handleRestoreRemovedPoint(removedPointsLog[0].id);
                  }}
                  canUndoRemovedPoint={removedPointsLog.length > 0}
                  onHoverPoint={setHoveredGraphPoint}
                  onSelectSeriesInInspector={(sId) => {
                    setSelectedSeriesId(sId);
                    setShowRightInspector(true);
                  }}
                  onOpenPlotDetailsModal={() => setIsPlotDetailsOpen(true)}
                  onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                  onQuickFitActiveSeries={handleQuickFitActiveSeries}
                  onRenameGraph={handleRenameGraph}
                  onDeleteGraph={handleDeleteGraph}
                  onAddGraph={handleAddGraph}
                />
              </div>

              <div className="h-full min-h-0">
                <WorkbookView
                  sheet={sheets[1] || sheets[0]}
                  allSheets={sheets}
                  activeSheetId={(sheets[1] || sheets[0]).id}
                  onSelectSheet={setActiveSheetId}
                  onAddSheet={handleAddSheet}
                  onDeleteSheet={handleDeleteSheet}
                  onRenameSheet={handleRenameSheet}
                  onUpdateSheet={handleUpdateSheet}
                  onPlotColumnsToActiveGraph={handlePlotColumnsToActiveGraph}
                  onPlotColumnsToNewGraph={handlePlotColumnsToNewGraph}
                  onExportSheet={handleExportSheet}
                  highlightedRowIndex={hoveredGraphPoint?.rowIndex ?? null}
                  highlightedSheetId={hoveredGraphPoint?.sheetId ?? null}
                  onOpenStatsModal={(stats) => {
                    setActiveStatsData(stats);
                    setIsStatsModalOpen(true);
                  }}
                  onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                />
              </div>
            </div>
          )}

          {/* MODE 5: SHEET ONLY (Maximized Table) */}
          {dockLayout === 'sheet-only' && (
            <div className="h-full">
              <WorkbookView
                sheet={activeSheet}
                allSheets={sheets}
                activeSheetId={activeSheet.id}
                onSelectSheet={setActiveSheetId}
                onAddSheet={handleAddSheet}
                onCreateMultipleSheets={handleCreateMultipleSheets}
                onDuplicateSheet={handleDuplicateSheet}
                onDeleteSheet={handleDeleteSheet}
                onRenameSheet={handleRenameSheet}
                onUpdateSheet={handleUpdateSheet}
                onPlotColumnsToActiveGraph={handlePlotColumnsToActiveGraph}
                onPlotColumnsToNewGraph={handlePlotColumnsToNewGraph}
                onExportSheet={handleExportSheet}
                highlightedRowIndex={hoveredGraphPoint?.rowIndex ?? null}
                highlightedSheetId={hoveredGraphPoint?.sheetId ?? null}
                onOpenStatsModal={(stats) => {
                  setActiveStatsData(stats);
                  setIsStatsModalOpen(true);
                }}
                onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                isMaximized
                onToggleMaximize={() => setDockLayout(previousDockLayout)}
                onSmartPasteNotification={triggerToast}
                externalSubPanelRequest={externalSubPanelRequest}
              />
            </div>
          )}

          {/* MODE 6: GRAPH ONLY (Maximized Graph) */}
          {dockLayout === 'graph-only' && (
            <div className="h-full">
              <GraphCanvas
                graph={activeGraph}
                allGraphs={graphs}
                activeGraphId={activeGraph.id}
                sheets={sheets}
                toolMode={toolMode}
                onChangeToolMode={setToolMode}
                onSelectGraph={setActiveGraphId}
                onUpdateGraph={handleUpdateGraph}
                onRemoveBadPoint={(pt) => handleRemoveBadPointsBatch([pt])}
                onRemoveBadPointsBatch={handleRemoveBadPointsBatch}
                onMaskPoint={handleMaskPoint}
                onUndoLastRemovedPoint={() => {
                  if (removedPointsLog[0]) handleRestoreRemovedPoint(removedPointsLog[0].id);
                }}
                canUndoRemovedPoint={removedPointsLog.length > 0}
                onHoverPoint={setHoveredGraphPoint}
                onSelectSeriesInInspector={(sId) => {
                  setSelectedSeriesId(sId);
                  setShowRightInspector(true);
                }}
                onOpenPlotDetailsModal={() => setIsPlotDetailsOpen(true)}
                onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
                onQuickFitActiveSeries={handleQuickFitActiveSeries}
                isMaximized
                onToggleMaximize={() => setDockLayout(previousDockLayout)}
                onRenameGraph={handleRenameGraph}
                onDeleteGraph={handleDeleteGraph}
                onAddGraph={handleAddGraph}
              />
            </div>
          )}
        </main>

        {/* Right Contextual Inspector */}
        {showRightInspector && (
          <InspectorPanel
            activeTab={inspectorTab}
            onChangeTab={setInspectorTab}
            graph={activeGraph}
            sheets={sheets}
            selectedSeriesId={selectedSeriesId}
            onSelectSeries={setSelectedSeriesId}
            onUpdateGraph={handleUpdateGraph}
            onAddSeriesToGraph={({ sheetId, xColId, yColId, plotType, connection, color }) => {
              const sh = sheets.find((s) => s.id === sheetId);
              const yCol = sh?.columns.find((c) => c.id === yColId);
              const newSeries: PlotSeries = {
                id: `series_${Date.now()}_${Math.random().toString(36).slice(2, 5)}`,
                sheetId,
                xColId,
                yColId,
                yErrColId: yCol?.linkedErrColId,
                label: yCol
                  ? `${yCol.longName || yCol.shortName}${yCol.unit ? ` (${yCol.unit})` : ''}`
                  : 'Nuova Serie',
                autoLabel: true,
                plotType,
                color,
                lineWidth: 1.75,
                lineStyle: 'solid',
                connection,
                markerShape: 'circle',
                markerSize: 4.5,
                symbolFill: color,
                symbolBorder: '#ffffff',
                symbolBorderWidth: 1,
                fillArea: plotType === 'area',
                fillColor: color,
                fillOpacity: 0.2,
                dropLines: false,
                showErrorBars: !!yCol?.linkedErrColId,
                errorBarCap: 4,
                errorBarColor: color,
                visible: true,
              };
              handleUpdateGraph({
                ...activeGraph,
                series: [...activeGraph.series, newSeries],
              });
              triggerToast(`Aggiunta serie ${newSeries.label} al grafico`);
            }}
            onRunCurveFit={handleRunCurveFit}
            onDeleteFit={(fitId) => {
              handleUpdateGraph({
                ...activeGraph,
                fits: activeGraph.fits.filter((f) => f.id !== fitId),
              });
            }}
            onOpenFitSheet={(sheetId) => {
              setActiveSheetId(sheetId);
              if (dockLayout === 'graph-only') {
                setDockLayout(previousDockLayout);
              } else if (dockLayout === 'mdi-free') {
                bringMdiWindowToFront(`mdi_sheet_${sheetId}`);
              }
            }}
            onExportFitTxt={handleExportFitTxt}
            toolMode={toolMode}
            onChangeToolMode={setToolMode}
            badPointRemovalMode={badPointRemovalMode}
            onChangeBadPointRemovalMode={setBadPointRemovalMode}
            removedPointsLog={removedPointsLog}
            onRestoreRemovedPoint={handleRestoreRemovedPoint}
            onApplySmoothing={handleApplySmoothing}
            onApplyDerivative={handleApplyDerivative}
            onApplyIntegration={handleApplyIntegration}
            onOpenFilterSmoothModal={() => setIsFilterSmoothOpen(true)}
          />
        )}
      </div>

      {/* Advanced OriginLab Import Wizard Modal (CSV/TSV, Excel .xlsx, Multi-file, Smart Clipboard) */}
      <ImportModal
        isOpen={isImportModalOpen}
        onClose={() => {
          setIsImportModalOpen(false);
          setImportClipboardMode(false);
        }}
        onImportSheets={handleImportSheetsBatch}
        initialDroppedFiles={droppedFilesForImport}
        onClearDroppedFiles={() => setDroppedFilesForImport([])}
        initialClipboardMode={importClipboardMode}
      />

      {/* Origin Statistics on Columns Modal */}
      <StatisticsModal
        isOpen={isStatsModalOpen}
        onClose={() => setIsStatsModalOpen(false)}
        stats={activeStatsData}
        onCreateStatsSheet={(newSheet) => {
          setSheets((prev) => [...prev, newSheet]);
          setActiveSheetId(newSheet.id);
          triggerToast('Creato nuovo foglio con statistiche descrittive');
        }}
      />

      {/* Origin Plot Details Modal */}
      <PlotDetailsModal
        isOpen={isPlotDetailsOpen}
        onClose={() => setIsPlotDetailsOpen(false)}
        graph={activeGraph}
        sheets={sheets}
        onUpdateGraph={handleUpdateGraph}
      />

      {/* Origin Signal Filtering & Smoothing Suite Modal */}
      <FilterSmoothModal
        isOpen={isFilterSmoothOpen}
        onClose={() => setIsFilterSmoothOpen(false)}
        sheets={sheets}
        activeSheetId={activeSheet.id}
        onApplyFilterResult={handleApplyFilterSmoothResult}
      />
    </div>
  );
}
