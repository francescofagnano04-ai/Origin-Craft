import React, { useState, useEffect, useRef } from 'react';
import {
  Activity,
  ChevronDown,
  Upload,
  FileSpreadsheet,
  ClipboardPaste,
  Download,
  RotateCcw,
  Plus,
  Columns,
  Rows,
  AppWindow,
  LayoutGrid,
  LineChart,
  TrendingUp,
  Sparkles,
  Sigma,
  Target,
  Crosshair,
  MousePointer,
  EyeOff,
  ZoomIn,
  Type,
  Sliders,
  FolderTree,
  SlidersHorizontal,
  Undo2,
  Check,
  Copy,
  Calculator,
  Search,
  Filter,
  GitFork,
  Hash,
  RefreshCw,
} from 'lucide-react';
import { DockLayoutMode, FitModelType, PlotType, ToolMode } from '../types/origin';
import { ActiveSubPanel } from './WorkbookView';

interface TopMenuBarProps {
  onOpenImportWizard: (clipboardMode?: boolean) => void;
  onNewSheet: () => void;
  onExportActiveSheet: (format: 'txt' | 'csv' | 'csv-semi' | 'dat' | 'xlsx') => void;
  onCopyActiveSheetToClipboard: () => void;
  onSaveProject: () => void;
  onSaveBinaryOriginClone?: () => void;
  onOpenProjectFile: () => void;
  onResetDemo: () => void;
  onGenerateHighPerfDataset?: (pointsCount: number) => void;
  onAddColumnRole: (role: 'X' | 'Y' | 'X_ERR' | 'Y_ERR' | 'LABEL' | 'IGNORE') => void;
  onAddRows: (count: number) => void;
  onQuickSortSheet?: (direction: 'asc' | 'desc') => void;
  onQuickAutoFillX?: (mode: 'row_numbers' | 'linear_step' | 'random_normal') => void;
  onUnmaskAllActiveSheet?: () => void;
  canUndoRemovedPoint: boolean;
  onUndoRemovedPoint: () => void;
  showLeftExplorer: boolean;
  onToggleLeftExplorer: () => void;
  showRightInspector: boolean;
  onToggleRightInspector: () => void;
  showLegend: boolean;
  onToggleLegend: () => void;
  showFitReportBox: boolean;
  onToggleFitReportBox: () => void;
  showMajorGrid: boolean;
  onToggleMajorGrid: () => void;
  onRescaleGraph: () => void;
  onQuickChangePlotType: (plotType: PlotType) => void;
  onPlotSelectedToNewGraph: () => void;
  onOpenPlotDetails: () => void;
  onOpenCurveFitInspector: () => void;
  onQuickFitActiveSeries: (model: FitModelType) => void;
  onRunDerivativeOnActive: () => void;
  onRunIntegralOnActive: () => void;
  onOpenFilterSmoothModal: () => void;
  onQuickFilterActive: (
    method: 'savitzky_golay' | 'moving_average' | 'gaussian' | 'median' | 'outlier_zscore'
  ) => void;
  onOpenColumnStats: () => void;
  onDuplicateActiveSheet?: () => void;
  onOpenWorkbookSubPanel?: (panel: ActiveSubPanel) => void;
  onRecalculateActiveSheetFormulas?: () => void;
  onQuickCreateCalculatedColumn?: (preset: 'norm' | 'zscore' | 'linear_transform') => void;
  toolMode: ToolMode;
  onChangeToolMode: (mode: ToolMode) => void;
  dockLayout: DockLayoutMode;
  onChangeDockLayout: (mode: DockLayoutMode) => void;
  onTileMdiWindows: () => void;
  onCascadeMdiWindows: () => void;
  toastMessage: string | null;
}

type MenuId =
  | 'file'
  | 'edit'
  | 'column_math'
  | 'view'
  | 'plot'
  | 'analysis'
  | 'signal'
  | 'stats'
  | 'tools'
  | 'window'
  | null;

export const TopMenuBar: React.FC<TopMenuBarProps> = ({
  onOpenImportWizard,
  onNewSheet,
  onExportActiveSheet,
  onCopyActiveSheetToClipboard,
  onSaveProject,
  onSaveBinaryOriginClone,
  onOpenProjectFile,
  onResetDemo,
  onGenerateHighPerfDataset,
  onAddColumnRole,
  onAddRows,
  onQuickSortSheet,
  onQuickAutoFillX,
  onUnmaskAllActiveSheet,
  canUndoRemovedPoint,
  onUndoRemovedPoint,
  showLeftExplorer,
  onToggleLeftExplorer,
  showRightInspector,
  onToggleRightInspector,
  showLegend,
  onToggleLegend,
  showFitReportBox,
  onToggleFitReportBox,
  showMajorGrid,
  onToggleMajorGrid,
  onRescaleGraph,
  onQuickChangePlotType,
  onPlotSelectedToNewGraph,
  onOpenPlotDetails,
  onOpenCurveFitInspector,
  onQuickFitActiveSeries,
  onRunDerivativeOnActive,
  onRunIntegralOnActive,
  onOpenFilterSmoothModal,
  onQuickFilterActive,
  onOpenColumnStats,
  onDuplicateActiveSheet,
  onOpenWorkbookSubPanel,
  onRecalculateActiveSheetFormulas,
  onQuickCreateCalculatedColumn,
  toolMode,
  onChangeToolMode,
  dockLayout,
  onChangeDockLayout,
  onTileMdiWindows,
  onCascadeMdiWindows,
  toastMessage,
}) => {
  const [openMenu, setOpenMenu] = useState<MenuId>(null);
  const menuBarRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (menuBarRef.current && !menuBarRef.current.contains(e.target as Node)) {
        setOpenMenu(null);
      }
    };
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpenMenu(null);
    };
    window.addEventListener('mousedown', handleClickOutside);
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('mousedown', handleClickOutside);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, []);

  const triggerAndClose = (fn: () => void) => {
    fn();
    setOpenMenu(null);
  };

  const renderMenuButton = (id: Exclude<MenuId, null>, label: string) => {
    const isOpen = openMenu === id;
    return (
      <button
        type="button"
        onClick={() => setOpenMenu(isOpen ? null : id)}
        onMouseEnter={() => {
          if (openMenu !== null) setOpenMenu(id);
        }}
        className={`flex items-center gap-1 px-2.5 py-1 rounded text-xs font-medium transition-colors select-none whitespace-nowrap ${
          isOpen
            ? 'bg-blue-700 text-white font-semibold shadow-2xs'
            : 'text-slate-800 hover:bg-slate-300/70'
        }`}
      >
        <span>{label}</span>
        <ChevronDown className="w-3 h-3 opacity-75" />
      </button>
    );
  };

  return (
    <header
      ref={menuBarRef}
      className="relative z-40 flex items-center justify-between gap-2 px-3 py-1 bg-[#dbe4ee] border-b border-slate-300 shrink-0 select-none text-xs"
    >
      {/* Left: Brand + Classic OriginLab Dropdown Menus */}
      <div className="flex items-center gap-1 flex-wrap">
        <div
          onClick={() => onChangeDockLayout('split-horizontal')}
          className="flex items-center gap-1.5 pr-3 mr-1 border-r border-slate-300 cursor-pointer text-slate-900 font-bold text-sm whitespace-nowrap"
        >
          <Activity className="w-4 h-4 text-blue-800" />
          <span>OriginStudio Lab</span>
        </div>

        {/* 1. FILE MENU */}
        <div className="relative">
          {renderMenuButton('file', 'File')}
          {openMenu === 'file' && (
            <div className="absolute left-0 top-full mt-1 w-72 bg-white border border-slate-300 rounded shadow-xl py-1 z-50">
              <button
                type="button"
                onClick={() => triggerAndClose(onNewSheet)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Plus className="w-3.5 h-3.5 text-blue-700" />
                  <span>Nuovo Foglio Dati (Workbook)</span>
                </span>
                <span className="text-[10px] font-mono text-slate-400">Ctrl+N</span>
              </button>

              {onDuplicateActiveSheet && (
                <button
                  type="button"
                  onClick={() => triggerAndClose(onDuplicateActiveSheet)}
                  className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                >
                  <span className="flex items-center gap-2">
                    <Copy className="w-3.5 h-3.5 text-indigo-700" />
                    <span>Duplica Scheda Foglio Attivo (Multi-Foglio)</span>
                  </span>
                </button>
              )}

              <div className="my-1 border-t border-slate-200" />

              <button
                type="button"
                onClick={() => triggerAndClose(() => onOpenImportWizard(false))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left font-semibold text-blue-900"
              >
                <span className="flex items-center gap-2">
                  <Upload className="w-3.5 h-3.5 text-blue-700" />
                  <span>Importa CSV / TSV / TXT / DAT...</span>
                </span>
                <span className="text-[10px] font-mono text-slate-400">Ctrl+I</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onOpenImportWizard(false))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-emerald-900 font-medium"
              >
                <span className="flex items-center gap-2">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700" />
                  <span>Importa da Excel (.XLSX / .XLS)...</span>
                </span>
                <span className="text-[10px] font-mono text-emerald-700">Multi-Foglio</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onOpenImportWizard(false))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Columns className="w-3.5 h-3.5 text-indigo-700" />
                  <span>Importazione Multi-File (Accoda/Separa)...</span>
                </span>
              </button>

              <div className="my-1 border-t border-slate-200" />
              <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                Esportazione Dati Tabellari
              </div>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onExportActiveSheet('txt'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Download className="w-3.5 h-3.5 text-slate-600" />
                  <span>Esporta Tabella in .TXT (Tabulazioni)</span>
                </span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onExportActiveSheet('csv'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Download className="w-3.5 h-3.5 text-slate-600" />
                  <span>Esporta Tabella in .CSV (Virgola)</span>
                </span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onExportActiveSheet('csv-semi'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Download className="w-3.5 h-3.5 text-slate-600" />
                  <span>Esporta Tabella in .CSV (Punto e virgola ;)</span>
                </span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onExportActiveSheet('xlsx'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-emerald-900 font-medium"
              >
                <span className="flex items-center gap-2">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700" />
                  <span>Esporta Tabella in Excel (.XLSX)</span>
                </span>
              </button>

              <div className="my-1 border-t border-slate-200" />
              <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                Progetto Binario (.ORIGINCLONE / .OPJ)
              </div>

              {onSaveBinaryOriginClone && (
                <button
                  type="button"
                  onClick={() => triggerAndClose(onSaveBinaryOriginClone)}
                  className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-indigo-50 text-left font-semibold text-indigo-950"
                >
                  <span className="flex items-center gap-2">
                    <Download className="w-3.5 h-3.5 text-indigo-700" />
                    <span>Salva Archivio Binario (.originclone ZIP)</span>
                  </span>
                  <span className="text-[10px] font-mono text-indigo-700">Float64</span>
                </button>
              )}

              <button
                type="button"
                onClick={() => triggerAndClose(onSaveProject)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left font-medium text-slate-900"
              >
                <span className="flex items-center gap-2">
                  <Download className="w-3.5 h-3.5 text-blue-700" />
                  <span>Salva Progetto JSON (.opj.json)</span>
                </span>
                <span className="text-[10px] font-mono text-slate-400">Ctrl+S</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(onOpenProjectFile)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Upload className="w-3.5 h-3.5 text-blue-700" />
                  <span>Apri Progetto (.originclone / .opj.json)...</span>
                </span>
                <span className="text-[10px] font-mono text-slate-400">Ctrl+O</span>
              </button>

              {onGenerateHighPerfDataset && (
                <>
                  <div className="my-1 border-t border-slate-200" />
                  <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-emerald-700">
                    Test Hardware GPU WebGL & Float64Array
                  </div>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onGenerateHighPerfDataset(100000))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-emerald-50 text-left font-semibold text-emerald-950"
                  >
                    <span className="flex items-center gap-2">
                      <Sparkles className="w-3.5 h-3.5 text-emerald-700" />
                      <span>Simula 100.000 Punti (WebGL @ 60 FPS)</span>
                    </span>
                  </button>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onGenerateHighPerfDataset(500000))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-emerald-50 text-left font-medium text-emerald-900"
                  >
                    <span className="flex items-center gap-2">
                      <Sparkles className="w-3.5 h-3.5 text-emerald-700" />
                      <span>Simula 500.000 Punti (Stress Test 1M Valori)</span>
                    </span>
                  </button>
                </>
              )}

              <div className="my-1 border-t border-slate-200" />

              <button
                type="button"
                onClick={() => triggerAndClose(onResetDemo)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-red-50 text-left text-slate-700"
              >
                <span className="flex items-center gap-2">
                  <RotateCcw className="w-3.5 h-3.5 text-slate-500" />
                  <span>Ripristina Progetto Demo Calibrato</span>
                </span>
              </button>
            </div>
          )}
        </div>

        {/* 2. EDIT / MODIFICA MENU */}
        <div className="relative">
          {renderMenuButton('edit', 'Modifica')}
          {openMenu === 'edit' && (
            <div className="absolute left-0 top-full mt-1 w-72 bg-white border border-slate-300 rounded shadow-xl py-1 z-50">
              <button
                type="button"
                onClick={() => triggerAndClose(() => onOpenImportWizard(true))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left font-semibold text-blue-900"
              >
                <span className="flex items-center gap-2">
                  <ClipboardPaste className="w-3.5 h-3.5 text-blue-700" />
                  <span>Incolla Tabella Smart da Appunti...</span>
                </span>
                <span className="text-[10px] font-mono text-slate-400">Excel/PDF</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(onCopyActiveSheetToClipboard)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Copy className="w-3.5 h-3.5 text-slate-600" />
                  <span>Copia Tabella Attiva negli Appunti</span>
                </span>
                <span className="text-[10px] font-mono text-slate-400">Ctrl+C</span>
              </button>

              <div className="my-1 border-t border-slate-200" />
              <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                Ruoli di Colonna Dedicati & Righe
              </div>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onAddColumnRole('X'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Columns className="w-3.5 h-3.5 text-amber-700" />
                  <span>Aggiungi Colonna Asse X (Multi-X)</span>
                </span>
                <span className="text-[10px] font-mono text-amber-700">X</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onAddColumnRole('Y'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Columns className="w-3.5 h-3.5 text-blue-700" />
                  <span>Aggiungi Colonna Serie Y</span>
                </span>
                <span className="text-[10px] font-mono text-blue-700">Y</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onAddColumnRole('X_ERR'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Activity className="w-3.5 h-3.5 text-indigo-700" />
                  <span>Aggiungi Colonna Incertezza Orizzontale (ΔX)</span>
                </span>
                <span className="text-[10px] font-mono text-indigo-700">xErr±</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onAddColumnRole('Y_ERR'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Activity className="w-3.5 h-3.5 text-purple-700" />
                  <span>Aggiungi Colonna Incertezza Verticale (ΔY)</span>
                </span>
                <span className="text-[10px] font-mono text-purple-700">yErr±</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onAddColumnRole('LABEL'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Type className="w-3.5 h-3.5 text-teal-700" />
                  <span>Aggiungi Colonna Etichetta Punti (L)</span>
                </span>
                <span className="text-[10px] font-mono text-teal-700">Label</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onAddRows(10))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <Rows className="w-3.5 h-3.5 text-slate-600" />
                  <span>Aggiungi 10 Righe Vuote al Foglio</span>
                </span>
              </button>

              {onQuickSortSheet && (
                <>
                  <div className="my-1 border-t border-slate-200" />
                  <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                    Ordinamento, Auto-Fill & Masking
                  </div>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onQuickSortSheet('asc'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                  >
                    <span>Ordina Foglio per Asse X (Crescente ↑)</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onQuickSortSheet('desc'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                  >
                    <span>Ordina Foglio per Asse X (Decrescente ↓)</span>
                  </button>
                </>
              )}

              {onQuickAutoFillX && (
                <>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onQuickAutoFillX('row_numbers'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                  >
                    <span>Auto-Fill Colonna X con Numeri di Riga (1..N)</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onQuickAutoFillX('linear_step'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                  >
                    <span>Auto-Fill Nuova Colonna con Rumore Casuale</span>
                  </button>
                </>
              )}

              {onUnmaskAllActiveSheet && (
                <button
                  type="button"
                  onClick={() => triggerAndClose(onUnmaskAllActiveSheet)}
                  className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-red-50 text-left text-red-900"
                >
                  <span className="flex items-center gap-2">
                    <EyeOff className="w-3.5 h-3.5 text-red-700" />
                    <span>Smaschera Tutti i Punti del Foglio Attivo</span>
                  </span>
                </button>
              )}

              {onOpenWorkbookSubPanel && (
                <>
                  <div className="my-1 border-t border-slate-200" />
                  <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                    Ricerca, Filtro Righe & Split Foglio
                  </div>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onOpenWorkbookSubPanel('find_replace'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-amber-50 text-left font-medium text-amber-950"
                  >
                    <span className="flex items-center gap-2">
                      <Search className="w-3.5 h-3.5 text-amber-700" />
                      <span>Trova e Sostituisci (Pulisci -9999 / NaN)...</span>
                    </span>
                  </button>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onOpenWorkbookSubPanel('row_filter'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-cyan-50 text-left font-medium text-cyan-950"
                  >
                    <span className="flex items-center gap-2">
                      <Filter className="w-3.5 h-3.5 text-cyan-700" />
                      <span>Filtro Condizionale Righe (es. Col(B) &gt; 0)...</span>
                    </span>
                  </button>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onOpenWorkbookSubPanel('split_sheet'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-purple-50 text-left font-medium text-purple-950"
                  >
                    <span className="flex items-center gap-2">
                      <GitFork className="w-3.5 h-3.5 text-purple-700" />
                      <span>Divisione Foglio (Split per Gruppo/Categoria)...</span>
                    </span>
                  </button>
                </>
              )}

              <div className="my-1 border-t border-slate-200" />

              <button
                type="button"
                disabled={!canUndoRemovedPoint}
                onClick={() => triggerAndClose(onUndoRemovedPoint)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-amber-50 text-left text-amber-900 disabled:opacity-40"
              >
                <span className="flex items-center gap-2">
                  <Undo2 className="w-3.5 h-3.5 text-amber-700" />
                  <span>Ripristina Ultimo Punto Errato Rimosso</span>
                </span>
                <span className="text-[10px] font-mono">Ctrl+Z</span>
              </button>
            </div>
          )}
        </div>

        {/* 2B. COLONNA & MATEMATICA MENU (Batch 3 Scientific & Tabular Manipulation) */}
        <div className="relative">
          {renderMenuButton('column_math', 'Colonna & Matematica')}
          {openMenu === 'column_math' && (
            <div className="absolute left-0 top-full mt-1 w-80 bg-white border border-slate-300 rounded shadow-xl py-1 z-50">
              <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-indigo-700">
                Motore Vettoriale & Propagazione Errori
              </div>

              {onOpenWorkbookSubPanel && (
                <>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onOpenWorkbookSubPanel('formula_calc'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-indigo-50 text-left font-semibold text-indigo-950"
                  >
                    <span className="flex items-center gap-2">
                      <Calculator className="w-3.5 h-3.5 text-indigo-700" />
                      <span>Set Column Values: col(C) = col(A)*2 + sin(col(B))...</span>
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onOpenWorkbookSubPanel('error_prop'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-rose-50 text-left font-semibold text-rose-950"
                  >
                    <span className="flex items-center gap-2">
                      <Sigma className="w-3.5 h-3.5 text-rose-700" />
                      <span>Propagazione Elementare Errori (ΔZ per A±B, A·B, A/B, Aⁿ)...</span>
                    </span>
                  </button>
                </>
              )}

              {onRecalculateActiveSheetFormulas && (
                <button
                  type="button"
                  onClick={() => triggerAndClose(onRecalculateActiveSheetFormulas)}
                  className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                >
                  <span className="flex items-center gap-2">
                    <RefreshCw className="w-3.5 h-3.5 text-blue-700" />
                    <span>Ricalcola Ora Tutte le Formule F(x) del Foglio</span>
                  </span>
                  <span className="text-[10px] font-mono text-emerald-700">Auto-Recalc ON</span>
                </button>
              )}

              {onQuickCreateCalculatedColumn && (
                <>
                  <div className="my-1 border-t border-slate-200" />
                  <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                    Colonne Calcolate Rapide (con Auto-Recalc)
                  </div>
                  <button
                    type="button"
                    onClick={() =>
                      triggerAndClose(() => onQuickCreateCalculatedColumn('linear_transform'))
                    }
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                  >
                    <span>Crea Colonna: col(A) * 2 + sin(col(B))</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onQuickCreateCalculatedColumn('norm'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                  >
                    <span>Crea Colonna: Normalizzata [0, 1] → norm(col(B))</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onQuickCreateCalculatedColumn('zscore'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                  >
                    <span>Crea Colonna: Standardizzata Z-Score → zscore(col(B))</span>
                  </button>
                </>
              )}

              {onOpenWorkbookSubPanel && (
                <>
                  <div className="my-1 border-t border-slate-200" />
                  <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                    Formattazione, Filtro Righe & Divisione Foglio
                  </div>

                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onOpenWorkbookSubPanel('num_format'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-teal-50 text-left font-medium text-teal-950"
                  >
                    <span className="flex items-center gap-2">
                      <Hash className="w-3.5 h-3.5 text-teal-700" />
                      <span>Formattazione Numerica (Standard / 1.23×10⁻⁴ / Eng)...</span>
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onOpenWorkbookSubPanel('find_replace'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-amber-50 text-left text-slate-800"
                  >
                    <span className="flex items-center gap-2">
                      <Search className="w-3.5 h-3.5 text-amber-700" />
                      <span>Trova e Sostituisci (es. -9999 → NaN / Maschera)...</span>
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onOpenWorkbookSubPanel('row_filter'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-cyan-50 text-left text-slate-800"
                  >
                    <span className="flex items-center gap-2">
                      <Filter className="w-3.5 h-3.5 text-cyan-700" />
                      <span>Filtro Condizionale Righe (es. Col(B) &gt; 0)...</span>
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => triggerAndClose(() => onOpenWorkbookSubPanel('split_sheet'))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-purple-50 text-left text-slate-800"
                  >
                    <span className="flex items-center gap-2">
                      <GitFork className="w-3.5 h-3.5 text-purple-700" />
                      <span>Divisione Foglio (Split in Fogli o Colonne)...</span>
                    </span>
                  </button>
                </>
              )}
            </div>
          )}
        </div>

        {/* 3. VIEW / VISUALIZZA MENU */}
        <div className="relative">
          {renderMenuButton('view', 'Visualizza')}
          {openMenu === 'view' && (
            <div className="absolute left-0 top-full mt-1 w-64 bg-white border border-slate-300 rounded shadow-xl py-1 z-50">
              <button
                type="button"
                onClick={() => triggerAndClose(onToggleLeftExplorer)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <FolderTree className="w-3.5 h-3.5 text-blue-700" />
                  <span>Project Explorer (Sinistra)</span>
                </span>
                {showLeftExplorer && <Check className="w-3.5 h-3.5 text-blue-700" />}
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(onToggleRightInspector)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span className="flex items-center gap-2">
                  <SlidersHorizontal className="w-3.5 h-3.5 text-blue-700" />
                  <span>Pannello Ispettore (Destra)</span>
                </span>
                {showRightInspector && <Check className="w-3.5 h-3.5 text-blue-700" />}
              </button>

              <div className="my-1 border-t border-slate-200" />

              <button
                type="button"
                onClick={() => triggerAndClose(onToggleLegend)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Mostra Legenda sul Grafico</span>
                {showLegend && <Check className="w-3.5 h-3.5 text-blue-700" />}
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(onToggleFitReportBox)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Mostra Box Report Equazione Fit</span>
                {showFitReportBox && <Check className="w-3.5 h-3.5 text-blue-700" />}
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(onToggleMajorGrid)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Mostra Griglia Principale Assi</span>
                {showMajorGrid && <Check className="w-3.5 h-3.5 text-blue-700" />}
              </button>

              <div className="my-1 border-t border-slate-200" />

              <button
                type="button"
                onClick={() => triggerAndClose(onRescaleGraph)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left font-medium text-blue-900"
              >
                <span className="flex items-center gap-2">
                  <RotateCcw className="w-3.5 h-3.5 text-blue-700" />
                  <span>Rescale to Show All (Auto-Scala)</span>
                </span>
              </button>
            </div>
          )}
        </div>

        {/* 4. PLOT / TRACCIA MENU */}
        <div className="relative">
          {renderMenuButton('plot', 'Traccia (Plot)')}
          {openMenu === 'plot' && (
            <div className="absolute left-0 top-full mt-1 w-72 bg-white border border-slate-300 rounded shadow-xl py-1 z-50">
              <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                8 Tipi di Grafico OriginLab
              </div>
              {[
                { id: 'line_symbol' as PlotType, label: 'Linea + Simboli (Line + Symbol)' },
                { id: 'line' as PlotType, label: 'Solo Linea Continua (Line)' },
                { id: 'scatter' as PlotType, label: 'Solo Punti / Simboli (Scatter)' },
                { id: 'spline' as PlotType, label: 'Curva Spline Cubica Smussata' },
                { id: 'step' as PlotType, label: 'Grafico a Gradino (Step Plot)' },
                { id: 'area' as PlotType, label: 'Area Sottesa (Filled Area)' },
                { id: 'bar' as PlotType, label: 'Colonne / Istogramma (Bar)' },
                { id: 'stem' as PlotType, label: 'Impulsi Verticali (Stem / Lollipop)' },
              ].map((pt) => (
                <button
                  key={pt.id}
                  type="button"
                  onClick={() => triggerAndClose(() => onQuickChangePlotType(pt.id))}
                  className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                >
                  <span className="flex items-center gap-2">
                    <LineChart className="w-3.5 h-3.5 text-blue-700" />
                    <span>{pt.label}</span>
                  </span>
                </button>
              ))}

              <div className="my-1 border-t border-slate-200" />

              <button
                type="button"
                onClick={() => triggerAndClose(onPlotSelectedToNewGraph)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left font-medium text-blue-900"
              >
                <span className="flex items-center gap-2">
                  <Plus className="w-3.5 h-3.5 text-blue-700" />
                  <span>Crea Nuova Finestra Grafico da Tabella</span>
                </span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(onOpenPlotDetails)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left font-semibold text-slate-900"
              >
                <span className="flex items-center gap-2">
                  <Sliders className="w-3.5 h-3.5 text-blue-700" />
                  <span>Format: Plot Details (Spessori, Colori, Simboli)...</span>
                </span>
              </button>
            </div>
          )}
        </div>

        {/* 5. ANALYSIS / ANALISI & FIT MENU */}
        <div className="relative">
          {renderMenuButton('analysis', 'Analisi & Fit')}
          {openMenu === 'analysis' && (
            <div className="absolute left-0 top-full mt-1 w-76 bg-white border border-slate-300 rounded shadow-xl py-1 z-50">
              <button
                type="button"
                onClick={() => triggerAndClose(onOpenCurveFitInspector)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-red-50 text-left font-semibold text-red-900"
              >
                <span className="flex items-center gap-2">
                  <TrendingUp className="w-3.5 h-3.5 text-red-600" />
                  <span>Pannello Curve Fitting su [C, D] con n Punti...</span>
                </span>
              </button>

              <div className="my-1 border-t border-slate-200" />
              <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                Fit Rapido sulla Serie Attiva
              </div>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onQuickFitActiveSeries('linear'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Fit Lineare (y = a + b·x)</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onQuickFitActiveSeries('polynomial'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Fit Polinomiale (Grado N)</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onQuickFitActiveSeries('gaussian'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Fit Picco Gaussiano (y₀ + A·exp)</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onQuickFitActiveSeries('exponential'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Fit Decadimento/Crescita Esponenziale</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onQuickFitActiveSeries('spline'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Interpolazione Spline Cubica Naturale C²</span>
              </button>

              <div className="my-1 border-t border-slate-200" />
              <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                Calcolo Differenziale & Integrale
              </div>

              <button
                type="button"
                onClick={() => triggerAndClose(onRunDerivativeOnActive)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Calcola Derivata Numerica dy/dx</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(onRunIntegralOnActive)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Calcola Integrale Cumulativo ∫ y dx</span>
              </button>
            </div>
          )}
        </div>

        {/* 6. SIGNAL PROCESSING / SEGNALE & FILTRI MENU */}
        <div className="relative">
          {renderMenuButton('signal', 'Segnale & Filtri')}
          {openMenu === 'signal' && (
            <div className="absolute left-0 top-full mt-1 w-72 bg-white border border-slate-300 rounded shadow-xl py-1 z-50">
              <button
                type="button"
                onClick={() => triggerAndClose(onOpenFilterSmoothModal)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-emerald-50 text-left font-semibold text-emerald-950"
              >
                <span className="flex items-center gap-2">
                  <Sparkles className="w-3.5 h-3.5 text-emerald-700" />
                  <span>Suite Filtro & Smoothing (Anteprima Live)...</span>
                </span>
              </button>

              <div className="my-1 border-t border-slate-200" />
              <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                Applicazione Rapida sulla Serie Attiva
              </div>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onQuickFilterActive('savitzky_golay'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Smoothing Savitzky-Golay (W=7, Polinomio g=2)</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onQuickFilterActive('moving_average'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Media Mobile Uniforme (Finestra W=5)</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onQuickFilterActive('gaussian'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Smoothing Kernel Gaussiano (W=7)</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onQuickFilterActive('median'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Filtro Mediano Anti-Spike Impulsivo (W=5)</span>
              </button>

              <button
                type="button"
                onClick={() => triggerAndClose(() => onQuickFilterActive('outlier_zscore'))}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Rimozione Automatica Outlier MAD (2.0σ)</span>
              </button>
            </div>
          )}
        </div>

        {/* 7. STATISTICS / STATISTICHE MENU */}
        <div className="relative">
          {renderMenuButton('stats', 'Statistiche')}
          {openMenu === 'stats' && (
            <div className="absolute left-0 top-full mt-1 w-68 bg-white border border-slate-300 rounded shadow-xl py-1 z-50">
              <button
                type="button"
                onClick={() => triggerAndClose(onOpenColumnStats)}
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left font-semibold text-slate-900"
              >
                <span className="flex items-center gap-2">
                  <Sigma className="w-3.5 h-3.5 text-teal-700" />
                  <span>Statistiche Descrittive sulle Colonne...</span>
                </span>
              </button>
            </div>
          )}
        </div>

        {/* 8. TOOLS / STRUMENTI PUNTATORE MENU */}
        <div className="relative">
          {renderMenuButton('tools', 'Strumenti')}
          {openMenu === 'tools' && (
            <div className="absolute left-0 top-full mt-1 w-68 bg-white border border-slate-300 rounded shadow-xl py-1 z-50">
              {[
                { id: 'pointer' as ToolMode, label: 'Puntatore Selezione Standard', icon: MousePointer },
                { id: 'reader' as ToolMode, label: 'Data Reader (Lettura Coordinate X, Y)', icon: Crosshair },
                {
                  id: 'remove_bad' as ToolMode,
                  label: 'Remove Bad Data Points (Elimina Punti Errati)',
                  icon: Target,
                },
                { id: 'mask_point' as ToolMode, label: 'Mask Point (Maschera Outlier)', icon: EyeOff },
                { id: 'zoom' as ToolMode, label: 'Zoom Regione Rettangolare', icon: ZoomIn },
                { id: 'add_text' as ToolMode, label: 'Aggiungi Annotazione di Testo', icon: Type },
              ].map((t) => {
                const Icon = t.icon;
                return (
                  <button
                    key={t.id}
                    type="button"
                    onClick={() => triggerAndClose(() => onChangeToolMode(t.id))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                  >
                    <span className="flex items-center gap-2">
                      <Icon className="w-3.5 h-3.5 text-blue-700" />
                      <span>{t.label}</span>
                    </span>
                    {toolMode === t.id && <Check className="w-3.5 h-3.5 text-blue-700" />}
                  </button>
                );
              })}
            </div>
          )}
        </div>

        {/* 9. WINDOW / FINESTRE MENU */}
        <div className="relative">
          {renderMenuButton('window', 'Finestre')}
          {openMenu === 'window' && (
            <div className="absolute left-0 top-full mt-1 w-68 bg-white border border-slate-300 rounded shadow-xl py-1 z-50">
              <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                Disposizione Spaziale Area di Lavoro
              </div>

              {[
                {
                  id: 'split-horizontal' as DockLayoutMode,
                  label: 'Affiancata Ridimensionabile (Tabella | Grafico)',
                  icon: Columns,
                },
                {
                  id: 'split-vertical' as DockLayoutMode,
                  label: 'Impilata Ridimensionabile (Grafico / Tabella)',
                  icon: Rows,
                },
                {
                  id: 'mdi-free' as DockLayoutMode,
                  label: 'Finestre Libere MDI (Stile Origin Desktop)',
                  icon: AppWindow,
                },
                {
                  id: 'multi-grid' as DockLayoutMode,
                  label: 'Mosaico Multi-Pannello',
                  icon: LayoutGrid,
                },
                {
                  id: 'sheet-only' as DockLayoutMode,
                  label: 'Solo Tabella Dati (Schermo Intero)',
                  icon: FileSpreadsheet,
                },
                {
                  id: 'graph-only' as DockLayoutMode,
                  label: 'Solo Grafico (Schermo Intero)',
                  icon: LineChart,
                },
              ].map((m) => {
                const Icon = m.icon;
                return (
                  <button
                    key={m.id}
                    type="button"
                    onClick={() => triggerAndClose(() => onChangeDockLayout(m.id))}
                    className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
                  >
                    <span className="flex items-center gap-2">
                      <Icon className="w-3.5 h-3.5 text-blue-700" />
                      <span>{m.label}</span>
                    </span>
                    {dockLayout === m.id && <Check className="w-3.5 h-3.5 text-blue-700" />}
                  </button>
                );
              })}

              <div className="my-1 border-t border-slate-200" />
              <div className="px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                Gestione Finestre MDI
              </div>

              <button
                type="button"
                onClick={() =>
                  triggerAndClose(() => {
                    onChangeDockLayout('mdi-free');
                    setTimeout(onTileMdiWindows, 30);
                  })
                }
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Affianca Tutte le Finestre in Griglia (Tile)</span>
              </button>

              <button
                type="button"
                onClick={() =>
                  triggerAndClose(() => {
                    onChangeDockLayout('mdi-free');
                    setTimeout(onCascadeMdiWindows, 30);
                  })
                }
                className="w-full flex items-center justify-between px-3 py-1.5 hover:bg-blue-50 text-left text-slate-800"
              >
                <span>Disponi Finestre a Cascata (Cascade)</span>
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Right: Toast feedback + Compact Workspace Switcher + Import Shortcut */}
      <div className="flex items-center gap-1.5 shrink-0">
        {toastMessage && (
          <span className="px-2 py-0.5 text-[11px] font-semibold text-emerald-950 bg-emerald-100 border border-emerald-300 rounded max-w-xs truncate">
            {toastMessage}
          </span>
        )}

        <div className="hidden xl:flex items-center gap-0.5 bg-white/80 p-0.5 rounded border border-slate-300">
          <button
            type="button"
            onClick={() => onChangeDockLayout('split-horizontal')}
            className={`px-1.5 py-0.5 rounded text-[10.5px] font-medium ${
              dockLayout === 'split-horizontal'
                ? 'bg-blue-700 text-white font-semibold'
                : 'text-slate-700 hover:bg-slate-200'
            }`}
            title="Vista Affiancata"
          >
            Affiancata
          </button>
          <button
            type="button"
            onClick={() => onChangeDockLayout('mdi-free')}
            className={`px-1.5 py-0.5 rounded text-[10.5px] font-medium ${
              dockLayout === 'mdi-free'
                ? 'bg-blue-700 text-white font-semibold'
                : 'text-slate-700 hover:bg-slate-200'
            }`}
            title="Finestre Libere MDI"
          >
            MDI Libere
          </button>
        </div>

        <button
          type="button"
          onClick={() => onOpenImportWizard(false)}
          className="px-2.5 py-1 text-xs font-semibold text-white bg-blue-700 rounded hover:bg-blue-800 transition-colors shadow-2xs flex items-center gap-1 whitespace-nowrap"
          title="Importa file .CSV, .TXT, .DAT, .XLSX o incolla da Appunti"
        >
          <Upload className="w-3.5 h-3.5" />
          <span>Importa Dati</span>
        </button>
      </div>
    </header>
  );
};
