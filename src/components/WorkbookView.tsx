import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';
import {
  Plus,
  Trash2,
  LineChart,
  CheckSquare,
  Square,
  FileSpreadsheet,
  Rows,
  Calculator,
  Activity,
  EyeOff,
  Eye,
  Sigma,
  Sparkles,
  Maximize2,
  Minimize2,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Wand2,
  GripVertical,
  ChevronLeft,
  ChevronRight,
  Tag,
  Cpu,
  Search,
  Filter,
  Split,
  Hash,
  Copy,
  RefreshCw,
} from 'lucide-react';
import {
  ColumnDataType,
  ColumnRole,
  ColumnStatistics,
  ConnectionType,
  DataColumn,
  NumericFormatMode,
  NumericPrecisionMode,
  PlotType,
  RowFilterOperator,
  WorkbookSheet,
} from '../types/origin';
import {
  AutoFillMode,
  ErrorPropagationOp,
  FindMatchMode,
  MaskOperationMode,
  ReplaceActionMode,
  SplitWorksheetMode,
  applyAutoFillToColumn,
  applyElementaryErrorPropagation,
  applyMaskingToSheet,
  computeColumnStatistics,
  computeQuickSelectionStats,
  detectColumnDataType,
  evaluateColumnFormula,
  evaluateSheetRowFilter,
  extractFilteredRowsToNewSheet,
  findAndReplaceInSheet,
  formatNumericValueWithColumnSpec,
  formatSciNumber,
  getColumnShortName,
  getOriginColumnRoleBadge,
  moveColumnInSheet,
  parseCellInputForColumn,
  parseFormulaTargetColumnAndExpr,
  rebindSheetMultiXAssociations,
  recalculateSheetFormulas,
  resolveErrorAndLabelColumnsForY,
  resolveXColumnForY,
  sortEntireSheetByColumn,
  sortSingleColumn,
  splitWorksheetByColumn,
} from '../utils/mathEngine';
import {
  attachBinaryBuffersToColumn,
  countColumnMaskedCells,
  getColumnRowCount,
  getOrSyncColumnBuffers,
  getSheetRowCount,
  isCellMaskedO1,
  readCellNumericO1,
  toggleColumnCellMaskBinary,
  updateColumnCellBinary,
} from '../utils/binaryColumnStore';
import { evaluateFormulaInWorker } from '../utils/workerClient';
import { PLOT_TYPE_LABELS } from './GraphCanvas';

export type ActiveSubPanel =
  | 'none'
  | 'sort'
  | 'autofill'
  | 'masking'
  | 'num_format'
  | 'find_replace'
  | 'row_filter'
  | 'split_sheet'
  | 'formula_calc'
  | 'error_prop';

interface WorkbookViewProps {
  sheet: WorkbookSheet;
  allSheets: WorkbookSheet[];
  activeSheetId: string;
  onSelectSheet: (sheetId: string) => void;
  onAddSheet: () => void;
  onDuplicateSheet?: (sheetId: string) => void;
  onCreateMultipleSheets?: (newSheets: WorkbookSheet[], summary: string) => void;
  onDeleteSheet: (sheetId: string) => void;
  onRenameSheet: (sheetId: string, newName: string) => void;
  onUpdateSheet: (updated: WorkbookSheet) => void;
  onPlotColumnsToActiveGraph: (
    sheetId: string,
    yColIds: string[],
    plotType: PlotType,
    connection: ConnectionType
  ) => void;
  onPlotColumnsToNewGraph: (
    sheetId: string,
    yColIds: string[],
    plotType: PlotType,
    connection: ConnectionType
  ) => void;
  onExportSheet: (
    sheet: WorkbookSheet,
    format: 'txt' | 'csv' | 'csv-semi' | 'dat' | 'xlsx'
  ) => void;
  highlightedRowIndex: number | null;
  highlightedSheetId: string | null;
  onOpenStatsModal: (stats: ColumnStatistics[]) => void;
  onOpenFilterSmoothModal?: () => void;
  onSmartPasteNotification?: (msg: string) => void;
  externalSubPanelRequest?: { panel: ActiveSubPanel; nonce: number } | null;
  isMaximized?: boolean;
  onToggleMaximize?: () => void;
  hideSheetTabs?: boolean;
}

/**
 * RULE 4: Total DOM Virtualization Constants
 * Never render more than ~45 rows in the DOM, even for 1,000,000 rows!
 */
const VIRTUAL_ROW_HEIGHT = 24;
const VIRTUAL_OVERSCAN = 12;

export const WorkbookView: React.FC<WorkbookViewProps> = ({
  sheet,
  allSheets,
  activeSheetId,
  onSelectSheet,
  onAddSheet,
  onDuplicateSheet,
  onCreateMultipleSheets,
  onDeleteSheet,
  onRenameSheet,
  onUpdateSheet,
  onPlotColumnsToActiveGraph,
  onPlotColumnsToNewGraph,
  highlightedRowIndex,
  highlightedSheetId,
  onOpenStatsModal,
  onOpenFilterSmoothModal,
  onSmartPasteNotification,
  externalSubPanelRequest,
  isMaximized,
  onToggleMaximize,
  hideSheetTabs,
}) => {
  const [selectedYColIds, setSelectedYColIds] = useState<string[]>(() =>
    sheet.columns.filter((c) => c.role === 'Y').map((c) => c.id)
  );
  const [quickPlotType, setQuickPlotType] = useState<PlotType>('line_symbol');
  const [quickConnection, setQuickConnection] = useState<ConnectionType>('linear');
  const [editingSheetNameId, setEditingSheetNameId] = useState<string | null>(null);
  const [formulaEditingColId, setFormulaEditingColId] = useState<string | null>(null);
  const [tempFormula, setTempFormula] = useState<string>('');
  const [focusedCellKey, setFocusedCellKey] = useState<string | null>(null);

  // 2D Cell Range Selection for Quick Selection Statistics in Status Bar
  const [cellSelection, setCellSelection] = useState<{
    startColIdx: number;
    endColIdx: number;
    startRowIdx: number;
    endRowIdx: number;
  } | null>(null);
  const isDraggingCellsRef = useRef<boolean>(false);

  // Column Drag & Drop state
  const [draggedColIndex, setDraggedColIndex] = useState<number | null>(null);
  const [dragOverColIndex, setDragOverColIndex] = useState<number | null>(null);

  // Sub-panels for Batch 2 & Batch 3 tools
  const [activeSubPanel, setActiveSubPanel] = useState<ActiveSubPanel>('none');

  useEffect(() => {
    if (externalSubPanelRequest && externalSubPanelRequest.panel !== 'none') {
      setActiveSubPanel(externalSubPanelRequest.panel);
    }
  }, [externalSubPanelRequest]);

  useEffect(() => {
    const stopCellDrag = () => {
      isDraggingCellsRef.current = false;
    };
    window.addEventListener('mouseup', stopCellDrag);
    return () => window.removeEventListener('mouseup', stopCellDrag);
  }, []);

  // Sort panel state
  const [sortColId, setSortColId] = useState<string>(() => sheet.columns[0]?.id || '');
  const [sortScope, setSortScope] = useState<'entire_sheet' | 'single_col'>('entire_sheet');
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('asc');

  // Auto-Fill panel state
  const [fillColId, setFillColId] = useState<string>(() => sheet.columns[0]?.id || '');
  const [fillMode, setFillMode] = useState<AutoFillMode>('linear_step');
  const [fillRows, setFillRows] = useState<number>(() => Math.max(20, getSheetRowCount(sheet)));
  const [fillStart, setFillStart] = useState<number>(0);
  const [fillStep, setFillStep] = useState<number>(0.5);
  const [fillStop, setFillStop] = useState<number>(10);
  const [fillMin, setFillMin] = useState<number>(0);
  const [fillMax, setFillMax] = useState<number>(1);
  const [fillMean, setFillMean] = useState<number>(0);
  const [fillStd, setFillStd] = useState<number>(1);

  // Range / Conditional Masking panel state
  const [maskTargetColId, setMaskTargetColId] = useState<string>('ALL_Y');
  const [maskMode, setMaskMode] = useState<MaskOperationMode>('mask_range');
  const [maskStartRow, setMaskStartRow] = useState<number>(1);
  const [maskEndRow, setMaskEndRow] = useState<number>(5);
  const [maskThreshold, setMaskThreshold] = useState<number>(0);
  const [maskMinThresh, setMaskMinThresh] = useState<number>(0);
  const [maskMaxThresh, setMaskMaxThresh] = useState<number>(100);

  // Numeric Formatting panel state
  const [fmtTargetColId, setFmtTargetColId] = useState<string>('ALL_NUMERIC');
  const [fmtMode, setFmtMode] = useState<NumericFormatMode>('scientific');
  const [fmtPrecisionMode, setFmtPrecisionMode] = useState<NumericPrecisionMode>('decimals');
  const [fmtPrecision, setFmtPrecision] = useState<number>(3);

  // Find & Replace panel state
  const [findTargetColId, setFindTargetColId] = useState<string>('ALL');
  const [findMatchMode, setFindMatchMode] = useState<FindMatchMode>('exact');
  const [findQuery, setFindQuery] = useState<string>('-9999');
  const [findMin, setFindMin] = useState<number>(-1000);
  const [findMax, setFindMax] = useState<number>(1000);
  const [replaceAction, setReplaceAction] = useState<ReplaceActionMode>('replace_nan');
  const [replaceWith, setReplaceWith] = useState<string>('0');

  // Row Filter panel state
  const [filterColId, setFilterColId] = useState<string>(() => sheet.columns[1]?.id || sheet.columns[0]?.id || '');
  const [filterOp, setFilterOp] = useState<RowFilterOperator>('gt');
  const [filterVal1, setFilterVal1] = useState<string>('0');
  const [filterVal2, setFilterVal2] = useState<string>('10');
  const [filterCustomExpr, setFilterCustomExpr] = useState<string>('col(B) > 0');

  // Split Worksheet panel state
  const [splitMode, setSplitMode] = useState<SplitWorksheetMode>('by_category_sheets');
  const [splitGroupColId, setSplitGroupColId] = useState<string>(() => sheet.columns[0]?.id || '');
  const [splitChunkRows, setSplitChunkRows] = useState<number>(10);

  // Set Column Values (Vector Formula Engine) panel state
  const [calcTargetColId, setCalcTargetColId] = useState<string>(() => sheet.columns[sheet.columns.length - 1]?.id || 'NEW_COL');
  const [calcExpr, setCalcExpr] = useState<string>('col(A) * 2 + sin(col(B))');

  // Elementary Error Propagation panel state
  const [errPropOp, setErrPropOp] = useState<ErrorPropagationOp>('mul');
  const [errPropColAId, setErrPropColAId] = useState<string>(() => sheet.columns[0]?.id || '');
  const [errPropErrAMode, setErrPropErrAMode] = useState<'column' | 'constant'>('constant');
  const [errPropColErrAId, setErrPropColErrAId] = useState<string>(() => sheet.columns.find((c) => c.role === 'Y_ERR' || c.role === 'X_ERR')?.id || '');
  const [errPropConstA, setErrPropConstA] = useState<number>(0.05);
  const [errPropColBId, setErrPropColBId] = useState<string>(() => sheet.columns[Math.min(1, sheet.columns.length - 1)]?.id || '');
  const [errPropErrBMode, setErrPropErrBMode] = useState<'column' | 'constant'>('constant');
  const [errPropColErrBId, setErrPropColErrBId] = useState<string>(() => sheet.columns.find((c) => c.role === 'Y_ERR' || c.role === 'X_ERR')?.id || '');
  const [errPropConstB, setErrPropConstB] = useState<number>(0.1);
  const [errPropExpN, setErrPropExpN] = useState<number>(2);
  const [errPropCoeffA, setErrPropCoeffA] = useState<number>(1);
  const [errPropCoeffB, setErrPropCoeffB] = useState<number>(1);

  // Virtual Scroller State (Rule 4: requestAnimationFrame throttled viewport slice)
  const scrollContainerRef = useRef<HTMLDivElement | null>(null);
  const rafRef = useRef<number | null>(null);
  const [scrollTop, setScrollTop] = useState<number>(0);
  const [viewportHeight, setViewportHeight] = useState<number>(480);

  const handleScroll = useCallback((e: React.UIEvent<HTMLDivElement>) => {
    const top = e.currentTarget.scrollTop;
    const height = e.currentTarget.clientHeight;
    if (rafRef.current !== null) {
      cancelAnimationFrame(rafRef.current);
    }
    rafRef.current = requestAnimationFrame(() => {
      setScrollTop(top);
      if (height > 0) setViewportHeight(height);
      rafRef.current = null;
    });
  }, []);

  useEffect(() => {
    const el = scrollContainerRef.current;
    if (!el) return;
    if (el.clientHeight > 0) {
      setViewportHeight(el.clientHeight);
    }
    const ro = new ResizeObserver((entries) => {
      for (const entry of entries) {
        if (entry.contentRect.height > 0) {
          setViewportHeight(entry.contentRect.height);
        }
      }
    });
    ro.observe(el);
    return () => {
      ro.disconnect();
      if (rafRef.current !== null) cancelAnimationFrame(rafRef.current);
    };
  }, []);

  // Auto-scroll virtual grid when a point is hovered/highlighted on the graph
  useEffect(() => {
    if (
      highlightedSheetId === sheet.id &&
      highlightedRowIndex !== null &&
      scrollContainerRef.current
    ) {
      const targetY = highlightedRowIndex * VIRTUAL_ROW_HEIGHT;
      const el = scrollContainerRef.current;
      if (targetY < el.scrollTop || targetY > el.scrollTop + el.clientHeight - 160) {
        el.scrollTop = Math.max(0, targetY - el.clientHeight / 3);
      }
    }
  }, [highlightedRowIndex, highlightedSheetId, sheet.id]);

  useEffect(() => {
    setSelectedYColIds(sheet.columns.filter((c) => c.role === 'Y').map((c) => c.id));
    if (!sheet.columns.some((c) => c.id === sortColId)) {
      setSortColId(sheet.columns[0]?.id || '');
    }
    if (!sheet.columns.some((c) => c.id === fillColId)) {
      setFillColId(sheet.columns[0]?.id || '');
    }
    if (!sheet.columns.some((c) => c.id === filterColId)) {
      setFilterColId(sheet.columns[1]?.id || sheet.columns[0]?.id || '');
    }
    if (!sheet.columns.some((c) => c.id === splitGroupColId)) {
      setSplitGroupColId(sheet.columns[0]?.id || '');
    }
    if (calcTargetColId !== 'NEW_COL' && !sheet.columns.some((c) => c.id === calcTargetColId)) {
      setCalcTargetColId(sheet.columns[sheet.columns.length - 1]?.id || 'NEW_COL');
    }
    if (!sheet.columns.some((c) => c.id === errPropColAId)) {
      setErrPropColAId(sheet.columns[0]?.id || '');
    }
    if (!sheet.columns.some((c) => c.id === errPropColBId)) {
      setErrPropColBId(sheet.columns[Math.min(1, sheet.columns.length - 1)]?.id || '');
    }
  }, [sheet.id, sheet.columns.length]);

  const maxRows = Math.max(20, getSheetRowCount(sheet));
  const xColumns = sheet.columns.filter((c) => c.role === 'X');
  const yErrColumns = sheet.columns.filter((c) => c.role === 'Y_ERR');
  const xErrColumns = sheet.columns.filter((c) => c.role === 'X_ERR');

  // Evaluate Row Filter (O(rows) over Float64Array, returning matching row indices)
  const rowFilterEval = useMemo(
    () => evaluateSheetRowFilter(sheet, maxRows),
    [sheet, maxRows]
  );
  const filteredRowIndices = rowFilterEval.matchingRows;
  const effectiveRowCount = filteredRowIndices ? filteredRowIndices.length : maxRows;

  // Compute Virtual Slice [startIndex, endIndex] over effectiveRowCount
  const visibleRowCount = Math.ceil(viewportHeight / VIRTUAL_ROW_HEIGHT) + VIRTUAL_OVERSCAN * 2;
  const startIndex = Math.max(0, Math.floor(scrollTop / VIRTUAL_ROW_HEIGHT) - VIRTUAL_OVERSCAN);
  const endIndex = Math.min(Math.max(0, effectiveRowCount - 1), startIndex + visibleRowCount);
  const topSpacerHeight = startIndex * VIRTUAL_ROW_HEIGHT;
  const bottomSpacerHeight = Math.max(0, (effectiveRowCount - 1 - endIndex) * VIRTUAL_ROW_HEIGHT);

  // Real-time Quick Statistics on Selection (Status Bar)
  const quickStats = useMemo(
    () => computeQuickSelectionStats(sheet, cellSelection),
    [sheet, cellSelection]
  );

  const handleRoleChange = (colId: string, newRole: ColumnRole) => {
    const updatedCols = sheet.columns.map((col) => {
      if (col.id !== colId) return col;
      const nextDataType: ColumnDataType =
        newRole === 'LABEL'
          ? 'text'
          : col.dataType === 'text'
          ? 'numeric'
          : col.dataType || 'numeric';
      return { ...col, role: newRole, dataType: nextDataType };
    });
    const rebound = rebindSheetMultiXAssociations({ ...sheet, columns: updatedCols });
    onUpdateSheet(rebound);

    if (newRole === 'Y' && !selectedYColIds.includes(colId)) {
      setSelectedYColIds((prev) => [...prev, colId]);
    } else if (newRole !== 'Y') {
      setSelectedYColIds((prev) => prev.filter((id) => id !== colId));
    }
  };

  const handleDataTypeChange = (colId: string, newType: ColumnDataType) => {
    const updatedCols = sheet.columns.map((col) => {
      if (col.id !== colId) return col;
      if (newType === 'integer') {
        const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
        const roundedF64 = new Float64Array(length);
        for (let i = 0; i < length; i++) {
          roundedF64[i] = Number.isFinite(float64[i]) ? Math.round(float64[i]) : Number.NaN;
        }
        return attachBinaryBuffersToColumn({ ...col, dataType: newType }, roundedF64, maskUint8);
      }
      return { ...col, dataType: newType };
    });
    onUpdateSheet({ ...sheet, columns: updatedCols });
  };

  const handleLinkedXChange = (yColId: string, linkedXId: string) => {
    const updatedCols = sheet.columns.map((col) =>
      col.id === yColId ? { ...col, linkedXColId: linkedXId } : col
    );
    onUpdateSheet({ ...sheet, columns: updatedCols });
  };

  const handleLinkedErrChange = (yColId: string, errColId: string) => {
    const updatedCols = sheet.columns.map((col) =>
      col.id === yColId ? { ...col, linkedErrColId: errColId || undefined } : col
    );
    onUpdateSheet({ ...sheet, columns: updatedCols });
  };

  const handleLinkedXErrChange = (yColId: string, xErrColId: string) => {
    const updatedCols = sheet.columns.map((col) =>
      col.id === yColId ? { ...col, linkedXErrColId: xErrColId || undefined } : col
    );
    onUpdateSheet({ ...sheet, columns: updatedCols });
  };

  const handleMetaChange = (
    colId: string,
    field: 'longName' | 'unit' | 'comment' | 'formula',
    value: string
  ) => {
    const updatedCols = sheet.columns.map((col) =>
      col.id === colId ? { ...col, [field]: value } : col
    );
    onUpdateSheet({ ...sheet, columns: updatedCols });
  };

  /**
   * Rule 2: Offloads Set Column Values (F(x)=) evaluation to math.worker.ts using Float64Array buffers.
   * Supports both `col(A) * 2 + sin(col(B))` AND full assignment `col(C) = col(A) * 2 + sin(col(B))`!
   */
  const handleApplyFormula = async (colId: string, formulaStr: string) => {
    const { explicitTargetRef, rhsFormula } = parseFormulaTargetColumnAndExpr(formulaStr);
    let targetCol = sheet.columns.find((c) => c.id === colId);

    // If user typed `col(C) = ...` and column C exists, route to column C!
    if (explicitTargetRef) {
      const matchedByRef = sheet.columns.find(
        (c) =>
          c.shortName.toUpperCase() === explicitTargetRef.toUpperCase() ||
          c.longName.trim().toUpperCase() === explicitTargetRef.toUpperCase()
      );
      if (matchedByRef) {
        targetCol = matchedByRef;
      }
    }

    if (!targetCol) return;
    setFormulaEditingColId(null);

    try {
      const { float64, maskUint8 } = await evaluateFormulaInWorker(
        rhsFormula,
        targetCol,
        sheet.columns
      );
      const updatedCols = sheet.columns.map((col) =>
        col.id === targetCol!.id
          ? attachBinaryBuffersToColumn({ ...col, formula: rhsFormula }, float64, maskUint8)
          : col
      );
      const recalculated = recalculateSheetFormulas({ ...sheet, columns: updatedCols });
      onUpdateSheet(recalculated);
      onSmartPasteNotification?.(
        `Formula F(x) = ${rhsFormula} calcolata su ${float64.length.toLocaleString('it-IT')} righe (${targetCol.shortName})`
      );
    } catch {
      const computedValues = evaluateColumnFormula(rhsFormula, targetCol, sheet.columns);
      const updatedCols = sheet.columns.map((col) =>
        col.id === targetCol!.id ? { ...col, formula: rhsFormula, values: computedValues } : col
      );
      const recalculated = recalculateSheetFormulas({ ...sheet, columns: updatedCols });
      onUpdateSheet(recalculated);
    }
  };

  /**
   * Rule 1: Direct O(1) mutation in Float64Array & Uint8Array binary column buffers
   * + Instant Automatic Recalculation of dependent F(x) columns!
   */
  const handleCellChange = (colId: string, rowIndex: number, rawVal: string) => {
    const updatedCols = sheet.columns.map((col) => {
      if (col.id !== colId) return col;
      const dt = col.dataType || (col.role === 'LABEL' ? 'text' : 'numeric');
      const { numVal, strVal } = parseCellInputForColumn(rawVal, dt, col.role);
      return updateColumnCellBinary(col, rowIndex, numVal, strVal);
    });

    const recalculated = recalculateSheetFormulas({ ...sheet, columns: updatedCols });
    onUpdateSheet(recalculated);
  };

  /**
   * Smart 2D Clipboard Paste directly into Float64Array & Uint8Array buffers
   */
  const handleSmartCellPaste = (
    e: React.ClipboardEvent<HTMLInputElement>,
    startColIdx: number,
    startRowIdx: number
  ) => {
    const pastedText = e.clipboardData.getData('text/plain');
    if (!pastedText) return;

    const lines = pastedText
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .split('\n')
      .map((l) => l.trim())
      .filter((l) => l.length > 0);

    if (lines.length <= 1 && !pastedText.includes('\t') && !pastedText.includes(';')) {
      return;
    }

    e.preventDefault();

    const hasTabs = lines.some((l) => l.includes('\t'));
    const hasSemi = !hasTabs && lines.some((l) => l.includes(';'));
    const tokenize = (line: string) =>
      hasTabs
        ? line.split('\t').map((s) => s.trim())
        : hasSemi
        ? line.split(';').map((s) => s.trim())
        : line.split(/\s+/).map((s) => s.trim());

    const rowsTokens = lines.map(tokenize);
    const maxPastedCols = Math.max(1, ...rowsTokens.map((r) => r.length));

    const parseNum = (tok: string): number | null => {
      if (!tok || tok === '--' || tok.toLowerCase() === 'nan') return null;
      const cleaned = tok.replace(',', '.');
      const v = Number(cleaned);
      return Number.isFinite(v) ? v : null;
    };

    const firstRow = rowsTokens[0];
    const firstRowNums = firstRow.filter((t) => parseNum(t) !== null).length;
    const isFirstRowHeader = rowsTokens.length > 1 && firstRowNums < firstRow.length / 2;
    const dataTokenRows = isFirstRowHeader ? rowsTokens.slice(1) : rowsTokens;

    const nextCols = [...sheet.columns];
    const defaultX = nextCols.find((c) => c.role === 'X');
    const targetMaxRow = Math.max(maxRows, startRowIdx + dataTokenRows.length);

    while (nextCols.length < startColIdx + maxPastedCols) {
      const idx = nextCols.length;
      const shortName = getColumnShortName(idx);
      const f64 = new Float64Array(targetMaxRow).fill(Number.NaN);
      const u8 = new Uint8Array(targetMaxRow).fill(2);
      nextCols.push(
        attachBinaryBuffersToColumn(
          {
            id: `col_paste_${Date.now()}_${idx}`,
            shortName,
            longName: `Serie ${shortName}`,
            unit: '',
            comment: '',
            role: 'Y',
            dataType: 'numeric',
            linkedXColId: defaultX?.id,
            values: [],
          },
          f64,
          u8
        )
      );
    }

    if (isFirstRowHeader) {
      for (let c = 0; c < firstRow.length; c++) {
        const targetCol = nextCols[startColIdx + c];
        if (targetCol && firstRow[c]) {
          const bracketMatch = firstRow[c].match(/^(.+?)\s*[\(\[]([^)\]]+)[\)\]]$/);
          if (bracketMatch) {
            targetCol.longName = bracketMatch[1].trim();
            targetCol.unit = bracketMatch[2].trim();
          } else {
            targetCol.longName = firstRow[c];
          }
        }
      }
    }

    for (let c = 0; c < maxPastedCols; c++) {
      const colIdx = startColIdx + c;
      const targetCol = nextCols[colIdx];
      if (!targetCol) continue;

      const colTokens = dataTokenRows.map((r) => r[c] ?? '');
      const detectedType = detectColumnDataType(colTokens);
      const curBuf = getOrSyncColumnBuffers(targetCol);
      const newLen = Math.max(curBuf.length, startRowIdx + dataTokenRows.length);

      const nextF64 = new Float64Array(newLen).fill(Number.NaN);
      const nextU8 = new Uint8Array(newLen).fill(2);
      nextF64.set(curBuf.float64, 0);
      nextU8.set(curBuf.maskUint8, 0);

      let nextRawStrings = targetCol.rawStrings ? [...targetCol.rawStrings] : undefined;

      for (let r = 0; r < dataTokenRows.length; r++) {
        const destRow = startRowIdx + r;
        const tok = dataTokenRows[r][c] || '';
        const { numVal, strVal } = parseCellInputForColumn(tok, detectedType, targetCol.role);
        if (numVal !== null && Number.isFinite(numVal)) {
          nextF64[destRow] = numVal;
          nextU8[destRow] = 0;
        } else {
          nextF64[destRow] = Number.NaN;
          nextU8[destRow] = 2;
        }
        if (strVal !== null || nextRawStrings) {
          if (!nextRawStrings) nextRawStrings = new Array(newLen).fill(null);
          while (nextRawStrings.length < newLen) nextRawStrings.push(null);
          nextRawStrings[destRow] = strVal;
        }
      }

      nextCols[colIdx] = attachBinaryBuffersToColumn(
        { ...targetCol, dataType: detectedType, rawStrings: nextRawStrings },
        nextF64,
        nextU8
      );
    }

    const rebound = rebindSheetMultiXAssociations({ ...sheet, columns: nextCols });
    onUpdateSheet(recalculateSheetFormulas(rebound));
    onSmartPasteNotification?.(
      `Incollati ${dataTokenRows.length}×${maxPastedCols} valori in buffer Float64Array (formule ricalcolate)`
    );
  };

  const handleToggleMaskCell = (colId: string, rowIndex: number) => {
    const updatedCols = sheet.columns.map((col) =>
      col.id === colId ? toggleColumnCellMaskBinary(col, rowIndex) : col
    );
    onUpdateSheet(recalculateSheetFormulas({ ...sheet, columns: updatedCols }));
  };

  const handleAddColumn = (role: ColumnRole = 'Y', insertAtIndex?: number) => {
    const nextIdx = insertAtIndex !== undefined ? insertAtIndex : sheet.columns.length;
    const shortName = getColumnShortName(nextIdx);
    const defaultX = sheet.columns.find((c) => c.role === 'X');
    const f64 = new Float64Array(maxRows).fill(Number.NaN);
    const u8 = new Uint8Array(maxRows).fill(2);

    const newCol = attachBinaryBuffersToColumn(
      {
        id: `col_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        shortName,
        longName:
          role === 'X'
            ? `Asse X`
            : role === 'Y_ERR'
            ? `Errore Y (±)`
            : role === 'X_ERR'
            ? `Errore X (±)`
            : role === 'LABEL'
            ? `Etichetta`
            : role === 'IGNORE'
            ? `Note`
            : `Serie ${shortName}`,
        unit: '',
        comment: '',
        role,
        dataType: role === 'LABEL' ? 'text' : 'numeric',
        linkedXColId: role === 'Y' && defaultX ? defaultX.id : undefined,
        values: [],
        rawStrings: role === 'LABEL' ? new Array(maxRows).fill(null) : undefined,
      },
      f64,
      u8
    );

    const nextCols = [...sheet.columns];
    if (insertAtIndex !== undefined && insertAtIndex >= 0 && insertAtIndex <= nextCols.length) {
      nextCols.splice(insertAtIndex, 0, newCol);
    } else {
      nextCols.push(newCol);
    }

    const rebound = rebindSheetMultiXAssociations({ ...sheet, columns: nextCols });
    onUpdateSheet(rebound);
    if (role === 'Y') {
      setSelectedYColIds((prev) => [...prev, newCol.id]);
    }
  };

  const handleDeleteColumn = (colId: string) => {
    if (sheet.columns.length <= 1) return;
    const filtered = sheet.columns.filter((c) => c.id !== colId);
    const rebound = rebindSheetMultiXAssociations({ ...sheet, columns: filtered });
    onUpdateSheet(rebound);
    setSelectedYColIds((prev) => prev.filter((id) => id !== colId));
  };

  const handleMoveColumn = (fromIndex: number, toIndex: number) => {
    const movedSheet = moveColumnInSheet(sheet, fromIndex, toIndex);
    onUpdateSheet(movedSheet);
    onSmartPasteNotification?.(
      `Colonna spostata in posizione ${getColumnShortName(toIndex)} (associazioni Multi-X aggiornate)`
    );
  };

  const handleAddRows = (count = 10) => {
    const updatedCols = sheet.columns.map((col) => {
      const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
      const newLen = length + count;
      const nextF64 = new Float64Array(newLen).fill(Number.NaN);
      const nextU8 = new Uint8Array(newLen).fill(2);
      nextF64.set(float64, 0);
      nextU8.set(maskUint8, 0);
      return attachBinaryBuffersToColumn(
        {
          ...col,
          rawStrings: col.rawStrings
            ? [...col.rawStrings, ...new Array(count).fill(null)]
            : undefined,
        },
        nextF64,
        nextU8
      );
    });
    onUpdateSheet({ ...sheet, columns: updatedCols });
  };

  const handleInsertRowAt = (rowIndex: number) => {
    const updatedCols = sheet.columns.map((col) => {
      const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
      const newLen = Math.max(length + 1, rowIndex + 1);
      const nextF64 = new Float64Array(newLen).fill(Number.NaN);
      const nextU8 = new Uint8Array(newLen).fill(2);

      for (let i = 0; i < rowIndex && i < length; i++) {
        nextF64[i] = float64[i];
        nextU8[i] = maskUint8[i];
      }
      for (let i = rowIndex; i < length; i++) {
        nextF64[i + 1] = float64[i];
        nextU8[i + 1] = maskUint8[i];
      }

      let nextStrs = col.rawStrings ? [...col.rawStrings] : undefined;
      if (nextStrs) {
        while (nextStrs.length < rowIndex) nextStrs.push(null);
        nextStrs.splice(rowIndex, 0, null);
      }

      return attachBinaryBuffersToColumn({ ...col, rawStrings: nextStrs }, nextF64, nextU8);
    });
    onUpdateSheet({ ...sheet, columns: updatedCols });
  };

  const handleDeleteRow = (rowIndex: number) => {
    const updatedCols = sheet.columns.map((col) => {
      const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
      if (rowIndex >= length) return col;
      const newLen = Math.max(0, length - 1);
      const nextF64 = new Float64Array(newLen);
      const nextU8 = new Uint8Array(newLen);

      let dst = 0;
      for (let i = 0; i < length; i++) {
        if (i === rowIndex) continue;
        nextF64[dst] = float64[i];
        nextU8[dst] = maskUint8[i];
        dst++;
      }

      const nextStrs = col.rawStrings ? [...col.rawStrings] : undefined;
      if (nextStrs && rowIndex < nextStrs.length) {
        nextStrs.splice(rowIndex, 1);
      }
      return attachBinaryBuffersToColumn({ ...col, rawStrings: nextStrs }, nextF64, nextU8);
    });
    onUpdateSheet({ ...sheet, columns: updatedCols });
  };

  const handleQuickSortColumn = (
    colId: string,
    direction: 'asc' | 'desc',
    entireSheet: boolean
  ) => {
    const col = sheet.columns.find((c) => c.id === colId);
    const sorted = entireSheet
      ? sortEntireSheetByColumn(sheet, colId, direction)
      : sortSingleColumn(sheet, colId, direction);
    onUpdateSheet(recalculateSheetFormulas(sorted));
    onSmartPasteNotification?.(
      entireSheet
        ? `Intero foglio ordinato (${direction === 'asc' ? 'Crescente ↑' : 'Decrescente ↓'}) per colonna ${col?.shortName || ''}`
        : `Colonna ${col?.shortName || ''} ordinata (${direction === 'asc' ? 'Crescente ↑' : 'Decrescente ↓'})`
    );
  };

  const handleExecuteAutoFill = () => {
    const target = sheet.columns.find((c) => c.id === fillColId);
    if (!target) return;
    const filledCol = applyAutoFillToColumn(target, {
      mode: fillMode,
      rowCount: fillRows,
      startValue: fillStart,
      stepValue: fillStep,
      stopValue: fillStop,
      minVal: fillMin,
      maxVal: fillMax,
      mean: fillMean,
      stdDev: fillStd,
    });
    const updatedCols = sheet.columns.map((c) => (c.id === fillColId ? filledCol : c));
    onUpdateSheet(recalculateSheetFormulas({ ...sheet, columns: updatedCols }));
    onSmartPasteNotification?.(
      `Riempimento rapido (${fillRows.toLocaleString('it-IT')} valori Float64) completato su colonna ${target.shortName}`
    );
  };

  const handleExecuteMaskOperation = () => {
    const updated = applyMaskingToSheet(sheet, {
      targetColId: maskTargetColId,
      mode: maskMode,
      startRow1: maskStartRow,
      endRow1: maskEndRow,
      threshold: maskThreshold,
      minThreshold: maskMinThresh,
      maxThreshold: maskMaxThresh,
    });
    onUpdateSheet(recalculateSheetFormulas(updated));
    onSmartPasteNotification?.('Mascheramento dati aggiornato sul foglio attivo');
  };

  const handleApplyNumericFormat = () => {
    const updatedCols = sheet.columns.map((col) => {
      const isMatch =
        fmtTargetColId === 'ALL_NUMERIC'
          ? col.dataType !== 'text' && col.dataType !== 'datetime' && col.role !== 'LABEL'
          : col.id === fmtTargetColId;
      if (!isMatch) return col;
      return {
        ...col,
        numericFormat: fmtMode,
        numericPrecision: fmtPrecision,
        numericPrecisionMode: fmtPrecisionMode,
      };
    });
    onUpdateSheet({ ...sheet, columns: updatedCols });
    onSmartPasteNotification?.(
      `Formattazione numerica (${fmtMode}, ${fmtPrecision} ${
        fmtPrecisionMode === 'sigfigs' ? 'cifre signif.' : 'decimali'
      }) applicata`
    );
  };

  const handleSingleColFormatChange = (
    colId: string,
    numericFormat: NumericFormatMode,
    numericPrecision?: number
  ) => {
    const updatedCols = sheet.columns.map((col) =>
      col.id === colId
        ? {
            ...col,
            numericFormat,
            numericPrecision: numericPrecision ?? col.numericPrecision ?? 3,
          }
        : col
    );
    onUpdateSheet({ ...sheet, columns: updatedCols });
  };

  const handleExecuteFindReplace = () => {
    const res = findAndReplaceInSheet(sheet, {
      targetColId: findTargetColId,
      matchMode: findMatchMode,
      findQuery,
      findMin,
      findMax,
      replaceAction,
      replaceWith,
    });
    onUpdateSheet(res.updatedSheet);
    onSmartPasteNotification?.(
      res.matchCount > 0
        ? `Trova & Sostituisci: ${res.matchCount} celle aggiornate su ${res.affectedColsCount} colonne`
        : `Nessuna occorrenza trovata per i criteri specificati`
    );
  };

  const handleApplyRowFilter = (enabled: boolean) => {
    onUpdateSheet({
      ...sheet,
      rowFilter: {
        enabled,
        colId: filterColId,
        operator: filterOp,
        value1: filterVal1,
        value2: filterVal2,
        customExpression: filterCustomExpr,
      },
    });
    onSmartPasteNotification?.(
      enabled ? 'Filtro righe condizionale attivato' : 'Filtro righe disattivato (mostrate tutte le righe)'
    );
  };

  const handleExtractFilteredToNewSheet = () => {
    if (!filteredRowIndices || filteredRowIndices.length === 0) return;
    const newSh = extractFilteredRowsToNewSheet(sheet, filteredRowIndices);
    onCreateMultipleSheets?.(
      [newSh],
      `Estratte ${filteredRowIndices.length} righe filtrate nel nuovo foglio "${newSh.name}"`
    );
  };

  const handleExecuteSplitWorksheet = () => {
    const res = splitWorksheetByColumn(sheet, {
      mode: splitMode,
      groupColId: splitGroupColId,
      chunkSizeRows: splitChunkRows,
    });
    if (res.createdSheets.length > 0) {
      onCreateMultipleSheets?.(res.createdSheets, res.summaryText);
    } else {
      onSmartPasteNotification?.(res.summaryText);
    }
  };

  const handleExecuteSetColumnValuesPanel = async () => {
    const { explicitTargetRef, rhsFormula } = parseFormulaTargetColumnAndExpr(calcExpr);
    let targetColId = calcTargetColId;

    if (explicitTargetRef) {
      const found = sheet.columns.find(
        (c) =>
          c.shortName.toUpperCase() === explicitTargetRef.toUpperCase() ||
          c.longName.trim().toUpperCase() === explicitTargetRef.toUpperCase()
      );
      if (found) {
        targetColId = found.id;
      } else {
        targetColId = 'NEW_COL';
      }
    }

    if (targetColId === 'NEW_COL') {
      const nextIdx = sheet.columns.length;
      const shortName = explicitTargetRef
        ? explicitTargetRef.toUpperCase()
        : getColumnShortName(nextIdx);
      const defaultX = sheet.columns.find((c) => c.role === 'X');
      const f64 = new Float64Array(maxRows).fill(Number.NaN);
      const u8 = new Uint8Array(maxRows).fill(2);
      const newCol = attachBinaryBuffersToColumn(
        {
          id: `col_calc_${Date.now()}`,
          shortName,
          longName: `Calcolata ${shortName}`,
          unit: '',
          comment: `F(x) = ${rhsFormula}`,
          formula: rhsFormula,
          role: 'Y',
          dataType: 'numeric',
          linkedXColId: defaultX?.id,
          values: [],
        },
        f64,
        u8
      );
      const withNew = rebindSheetMultiXAssociations({
        ...sheet,
        columns: [...sheet.columns, newCol],
      });
      const recalculated = recalculateSheetFormulas(withNew);
      onUpdateSheet(recalculated);
      setSelectedYColIds((prev) => [...prev, newCol.id]);
      onSmartPasteNotification?.(
        `Creata colonna ${newCol.shortName} con formula vettoriale F(x) = ${rhsFormula}`
      );
      return;
    }

    await handleApplyFormula(targetColId, rhsFormula);
  };

  const handleExecuteErrorPropagation = () => {
    const res = applyElementaryErrorPropagation(sheet, {
      operation: errPropOp,
      colAId: errPropColAId,
      errAMode: errPropErrAMode,
      colErrAId: errPropColErrAId,
      constErrA: errPropConstA,
      colBId: errPropColBId,
      errBMode: errPropErrBMode,
      colErrBId: errPropColErrBId,
      constErrB: errPropConstB,
      exponentN: errPropExpN,
      coeffA: errPropCoeffA,
      coeffB: errPropCoeffB,
      createZColumn: true,
    });
    onUpdateSheet(res.updatedSheet);
    onSmartPasteNotification?.(
      `Propagazione errori completata: create Col(${res.zColShortName}) [Z] e Col(${res.errZColShortName}) [ΔZ ±]`
    );
  };

  const toggleSelectYCol = (colId: string) => {
    setSelectedYColIds((prev) =>
      prev.includes(colId) ? prev.filter((id) => id !== colId) : [...prev, colId]
    );
  };

  const handleShowColumnStatistics = () => {
    const stats = sheet.columns
      .filter(
        (c) =>
          c.role === 'X' ||
          c.role === 'Y' ||
          c.role === 'Y_ERR' ||
          c.role === 'X_ERR'
      )
      .map(computeColumnStatistics);
    onOpenStatsModal(stats);
  };

  const activeYColsForPlot = selectedYColIds.filter((id) =>
    sheet.columns.some((c) => c.id === id && (c.role === 'Y' || c.role === 'Y_ERR'))
  );

  const handlePlotTypeSelectChange = (pt: PlotType) => {
    setQuickPlotType(pt);
    if (pt === 'spline') setQuickConnection('spline');
    else if (pt === 'step') setQuickConnection('step');
    else if (pt === 'scatter' || pt === 'bar' || pt === 'stem') setQuickConnection('none');
    else setQuickConnection('linear');
  };

  const totalMaskedCount = sheet.columns.reduce(
    (acc, c) => acc + countColumnMaskedCells(c),
    0
  );

  /**
   * Formats a cell in O(1) directly from the Float64Array / rawStrings without scanning the column.
   * Respects the column's numericFormat ('auto' | 'decimal' | 'scientific' | 'engineering') & precision!
   */
  const formatVirtualCellO1 = (col: DataColumn, rowIndex: number, isFocused: boolean): string => {
    const dt = col.dataType || (col.role === 'LABEL' ? 'text' : 'numeric');
    if (
      (dt === 'text' || dt === 'datetime' || col.role === 'LABEL') &&
      col.rawStrings &&
      rowIndex < col.rawStrings.length &&
      col.rawStrings[rowIndex] !== null &&
      col.rawStrings[rowIndex] !== undefined
    ) {
      return col.rawStrings[rowIndex]!;
    }
    const numVal = readCellNumericO1(col, rowIndex);
    if (numVal === null || !Number.isFinite(numVal)) return '';
    if (dt === 'integer') return String(Math.round(numVal));
    if (isFocused) {
      return String(numVal);
    }
    if (col.numericFormat && col.numericFormat !== 'auto') {
      return formatNumericValueWithColumnSpec(numVal, col, true);
    }
    if (col.numericPrecision !== undefined) {
      return formatNumericValueWithColumnSpec(numVal, col, false);
    }
    return String(numVal);
  };

  return (
    <div className="flex flex-col h-full bg-white border border-slate-300 rounded-xs shadow-xs overflow-hidden">
      {/* Row 1: Sheet Tabs & Table Column/Row Tools */}
      <div className="flex items-center justify-between gap-2 px-2 py-1 bg-[#e2e8f0] border-b border-slate-300 shrink-0">
        {/* Sheet Tabs */}
        <div className="flex items-center gap-1 overflow-x-auto min-w-0">
          {!hideSheetTabs ? (
            <>
              {allSheets.map((s) => {
                const isActive = s.id === activeSheetId;
                return (
                  <div
                    key={s.id}
                    onClick={() => onSelectSheet(s.id)}
                    className={`group flex items-center gap-1 px-2 py-0.5 text-xs font-medium rounded border cursor-pointer transition-colors whitespace-nowrap shrink-0 ${
                      isActive
                        ? 'bg-white text-slate-900 border-slate-300 font-semibold shadow-2xs'
                        : 'bg-slate-200/70 text-slate-600 border-transparent hover:bg-slate-200 hover:text-slate-900'
                    }`}
                  >
                    <FileSpreadsheet
                      className={`w-3 h-3 ${s.isFitOutput ? 'text-red-600' : 'text-blue-700'}`}
                    />
                    {editingSheetNameId === s.id ? (
                      <input
                        type="text"
                        defaultValue={s.name}
                        autoFocus
                        onBlur={(e) => {
                          onRenameSheet(s.id, e.target.value.trim() || s.name);
                          setEditingSheetNameId(null);
                        }}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            onRenameSheet(
                              s.id,
                              (e.target as HTMLInputElement).value.trim() || s.name
                            );
                            setEditingSheetNameId(null);
                          } else if (e.key === 'Escape') {
                            setEditingSheetNameId(null);
                          }
                        }}
                        className="w-20 px-1 py-0 text-xs bg-blue-50 border border-blue-400 rounded focus:outline-none"
                      />
                    ) : (
                      <span
                        onDoubleClick={() => setEditingSheetNameId(s.id)}
                        title="Doppio clic per rinominare il foglio"
                      >
                        {s.name}
                      </span>
                    )}
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        onDeleteSheet(s.id);
                      }}
                      className="opacity-0 group-hover:opacity-100 text-slate-400 hover:text-red-600 transition-opacity ml-0.5"
                      title="Elimina foglio"
                    >
                      ×
                    </button>
                  </div>
                );
              })}
              <button
                type="button"
                onClick={onAddSheet}
                className="flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-300/60 rounded transition-colors whitespace-nowrap shrink-0"
                title="Nuovo foglio dati vuoto (Gestione Multi-Foglio)"
              >
                <Plus className="w-3 h-3" />
                <span>Foglio</span>
              </button>
              {onDuplicateSheet && (
                <button
                  type="button"
                  onClick={() => onDuplicateSheet(sheet.id)}
                  className="flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-300/60 rounded transition-colors whitespace-nowrap shrink-0"
                  title="Duplica il foglio attivo in una nuova scheda"
                >
                  <Copy className="w-2.5 h-2.5" />
                  <span>Duplica</span>
                </button>
              )}
              {allSheets.length > 2 && (
                <select
                  value={activeSheetId}
                  onChange={(e) => onSelectSheet(e.target.value)}
                  className="px-1.5 py-0.5 text-[11px] bg-white border border-slate-300 rounded text-slate-700 font-medium cursor-pointer shrink-0"
                  title="Passa rapidamente a un foglio del progetto"
                >
                  {allSheets.map((s, idx) => (
                    <option key={s.id} value={s.id}>
                      #{idx + 1} · {s.name}
                    </option>
                  ))}
                </select>
              )}
            </>
          ) : (
            <span className="px-2 py-0.5 text-xs font-bold text-slate-800 truncate">
              {sheet.name} ({sheet.columns.length} col · {maxRows.toLocaleString('it-IT')} righe)
            </span>
          )}
        </div>

        {/* Table Quick Actions: Dynamic Columns, Rows, Auto-Fill, Sort, Masking, Find, Filter, Split, F(x), ΔZ */}
        <div className="flex items-center gap-1 shrink-0 flex-wrap">
          <span
            className="hidden 2xl:inline-flex items-center gap-1 px-1.5 py-0.5 text-[10px] font-mono font-semibold text-emerald-900 bg-emerald-100/90 border border-emerald-300 rounded"
            title="Regola 1 & 4 Attive: Buffer Float64Array + Uint8Array e Virtualizzazione DOM"
          >
            <Cpu className="w-2.5 h-2.5 text-emerald-700" />
            <span>
              Float64 · {effectiveRowCount.toLocaleString('it-IT')}/{maxRows.toLocaleString('it-IT')} righe
            </span>
          </span>

          <button
            type="button"
            onClick={() => handleAddColumn('X')}
            className="flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-amber-900 bg-amber-50 border border-amber-300 rounded hover:bg-amber-100 transition-colors whitespace-nowrap"
            title="Aggiungi nuova colonna Asse X (Supporto Multi-X: X1, Y1, X2, Y2)"
          >
            <Plus className="w-2.5 h-2.5 text-amber-700" />
            <span>Col(X)</span>
          </button>

          <button
            type="button"
            onClick={() => handleAddColumn('Y')}
            className="flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-blue-900 bg-blue-50 border border-blue-300 rounded hover:bg-blue-100 transition-colors whitespace-nowrap"
            title="Aggiungi nuova colonna Serie Y"
          >
            <Plus className="w-2.5 h-2.5 text-blue-700" />
            <span>Col(Y)</span>
          </button>

          <button
            type="button"
            onClick={() => handleAddColumn('Y_ERR')}
            className="flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-purple-900 bg-purple-50 border border-purple-300 rounded hover:bg-purple-100 transition-colors whitespace-nowrap hidden md:flex"
            title="Aggiungi colonna errore ΔY (yErr ±)"
          >
            <Activity className="w-2.5 h-2.5 text-purple-700" />
            <span>+ΔY</span>
          </button>

          <button
            type="button"
            onClick={() => handleAddRows(10)}
            className="flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 transition-colors whitespace-nowrap"
            title="Aggiungi 10 righe"
          >
            <Rows className="w-3 h-3 text-slate-600" />
            <span>+10</span>
          </button>

          {/* Toggle Numeric Format Subpanel */}
          <button
            type="button"
            onClick={() => setActiveSubPanel((p) => (p === 'num_format' ? 'none' : 'num_format'))}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded border transition-colors whitespace-nowrap ${
              activeSubPanel === 'num_format'
                ? 'bg-sky-700 text-white border-sky-800'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="Formattazione Numerica: Standard, Scientifica (1.23×10⁻⁴), Ingegneristica (10³ⁿ) e cifre decimali/significative"
          >
            <Hash className="w-3 h-3" />
            <span>Formato</span>
          </button>

          {/* Toggle Find & Replace Subpanel */}
          <button
            type="button"
            onClick={() => setActiveSubPanel((p) => (p === 'find_replace' ? 'none' : 'find_replace'))}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded border transition-colors whitespace-nowrap ${
              activeSubPanel === 'find_replace'
                ? 'bg-amber-700 text-white border-amber-800'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="Trova e Sostituisci (Pulizia codici errore strumentali es. -9999 → NaN)"
          >
            <Search className="w-3 h-3" />
            <span>Trova</span>
          </button>

          {/* Toggle Row Filter Subpanel */}
          <button
            type="button"
            onClick={() => setActiveSubPanel((p) => (p === 'row_filter' ? 'none' : 'row_filter'))}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded border transition-colors whitespace-nowrap ${
              activeSubPanel === 'row_filter'
                ? 'bg-teal-700 text-white border-teal-800'
                : sheet.rowFilter?.enabled
                ? 'bg-teal-50 text-teal-900 border-teal-400 font-semibold'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="Filtro Righe Condizionale (es. mostra solo dove Col(B) > 0)"
          >
            <Filter className="w-3 h-3" />
            <span>
              Filtro{sheet.rowFilter?.enabled ? ` (${effectiveRowCount})` : ''}
            </span>
          </button>

          {/* Toggle Split Sheet Subpanel */}
          <button
            type="button"
            onClick={() => setActiveSubPanel((p) => (p === 'split_sheet' ? 'none' : 'split_sheet'))}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded border transition-colors whitespace-nowrap ${
              activeSubPanel === 'split_sheet'
                ? 'bg-violet-700 text-white border-violet-800'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="Divisione Foglio (Split): separa il dataset in più fogli o colonne in base a una variabile categorica"
          >
            <Split className="w-3 h-3" />
            <span>Split</span>
          </button>

          {/* Toggle Set Column Values F(x) Subpanel */}
          <button
            type="button"
            onClick={() => setActiveSubPanel((p) => (p === 'formula_calc' ? 'none' : 'formula_calc'))}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-semibold rounded border transition-colors whitespace-nowrap ${
              activeSubPanel === 'formula_calc'
                ? 'bg-emerald-700 text-white border-emerald-800'
                : 'bg-emerald-50 text-emerald-900 border-emerald-300 hover:bg-emerald-100'
            }`}
            title="Motore di Calcolo Vettoriale per Colonna (Set Column Values: col(C) = col(A)*2 + sin(col(B)))"
          >
            <Calculator className="w-3 h-3" />
            <span>F(x)</span>
          </button>

          {/* Toggle Error Propagation ΔZ Subpanel */}
          <button
            type="button"
            onClick={() => setActiveSubPanel((p) => (p === 'error_prop' ? 'none' : 'error_prop'))}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-semibold rounded border transition-colors whitespace-nowrap ${
              activeSubPanel === 'error_prop'
                ? 'bg-purple-700 text-white border-purple-800'
                : 'bg-purple-50 text-purple-900 border-purple-300 hover:bg-purple-100'
            }`}
            title="Propagazione Elementare degli Errori (Calcolo automatico ΔZ per A±B, A·B, A/B, Aⁿ)"
          >
            <Activity className="w-3 h-3" />
            <span>Prop.ΔZ</span>
          </button>

          {/* Toggle Sort Subpanel */}
          <button
            type="button"
            onClick={() => setActiveSubPanel((p) => (p === 'sort' ? 'none' : 'sort'))}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded border transition-colors whitespace-nowrap ${
              activeSubPanel === 'sort'
                ? 'bg-blue-700 text-white border-blue-800'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="Ordinamento (Singola colonna o intero foglio per colonna chiave)"
          >
            <ArrowUpDown className="w-3 h-3" />
            <span>Ordina</span>
          </button>

          {/* Toggle Auto-Fill Subpanel */}
          <button
            type="button"
            onClick={() => setActiveSubPanel((p) => (p === 'autofill' ? 'none' : 'autofill'))}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded border transition-colors whitespace-nowrap ${
              activeSubPanel === 'autofill'
                ? 'bg-indigo-700 text-white border-indigo-800'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="Riempimento Rapido (Auto-Fill: sequenze lineari, step temporali, numeri casuali fino a 1.000.000 righe)"
          >
            <Wand2 className="w-3 h-3" />
            <span>Auto-Fill</span>
          </button>

          {/* Toggle Masking Subpanel */}
          <button
            type="button"
            onClick={() => setActiveSubPanel((p) => (p === 'masking' ? 'none' : 'masking'))}
            className={`flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium rounded border transition-colors whitespace-nowrap ${
              activeSubPanel === 'masking'
                ? 'bg-red-700 text-white border-red-800'
                : totalMaskedCount > 0
                ? 'bg-red-50 text-red-800 border-red-300 hover:bg-red-100'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="Mascheramento Dati (Escludi celle o intervalli da fit e grafico senza cancellarli)"
          >
            <EyeOff className="w-3 h-3" />
            <span>Mask{totalMaskedCount > 0 ? ` (${totalMaskedCount})` : ''}</span>
          </button>

          <button
            type="button"
            onClick={handleShowColumnStatistics}
            className="flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 transition-colors whitespace-nowrap hidden sm:flex"
            title="Statistiche descrittive sulle colonne"
          >
            <Sigma className="w-3 h-3 text-teal-700" />
            <span className="hidden xl:inline">Stat</span>
          </button>

          {onToggleMaximize && (
            <button
              type="button"
              onClick={onToggleMaximize}
              className="p-1 text-slate-600 hover:text-slate-900 hover:bg-slate-300/60 rounded transition-colors"
              title={isMaximized ? 'Ripristina vista divisa' : 'Ingrandisci tabella a tutto schermo'}
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

      {/* Collapsible Sub-Panels for Sort, Auto-Fill, and Masking */}
      {activeSubPanel === 'sort' && (
        <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-1.5 bg-blue-50/90 border-b border-blue-200 text-xs shrink-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-bold text-blue-950 flex items-center gap-1">
              <ArrowUpDown className="w-3.5 h-3.5 text-blue-700" />
              Ordinamento Dati:
            </span>
            <label className="flex items-center gap-1">
              <span className="text-slate-600">Colonna Chiave:</span>
              <select
                value={sortColId}
                onChange={(e) => setSortColId(e.target.value)}
                className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
              >
                {sheet.columns.map((c, idx) => (
                  <option key={c.id} value={c.id}>
                    {c.shortName}({getOriginColumnRoleBadge(sheet, idx)}) - {c.longName || c.shortName}
                  </option>
                ))}
              </select>
            </label>

            <label className="flex items-center gap-1">
              <span className="text-slate-600">Ambito:</span>
              <select
                value={sortScope}
                onChange={(e) => setSortScope(e.target.value as 'entire_sheet' | 'single_col')}
                className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs font-medium"
              >
                <option value="entire_sheet">Intero Foglio (mantieni righe allineate)</option>
                <option value="single_col">Solo la Singola Colonna Selezionata</option>
              </select>
            </label>

            <label className="flex items-center gap-1">
              <span className="text-slate-600">Ordine:</span>
              <select
                value={sortDir}
                onChange={(e) => setSortDir(e.target.value as 'asc' | 'desc')}
                className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs"
              >
                <option value="asc">Crescente (A → Z / Min → Max)</option>
                <option value="desc">Decrescente (Z → A / Max → Min)</option>
              </select>
            </label>
          </div>

          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={() =>
                handleQuickSortColumn(sortColId, sortDir, sortScope === 'entire_sheet')
              }
              className="px-2.5 py-0.5 bg-blue-700 text-white font-semibold rounded hover:bg-blue-800 transition-colors"
            >
              Applica Ordinamento
            </button>
            <button
              type="button"
              onClick={() => setActiveSubPanel('none')}
              className="px-1.5 py-0.5 text-slate-500 hover:text-slate-800"
            >
              Chiudi
            </button>
          </div>
        </div>
      )}

      {activeSubPanel === 'autofill' && (
        <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-1.5 bg-indigo-50/90 border-b border-indigo-200 text-xs shrink-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-bold text-indigo-950 flex items-center gap-1">
              <Wand2 className="w-3.5 h-3.5 text-indigo-700" />
              Auto-Fill Colonna:
            </span>

            <select
              value={fillColId}
              onChange={(e) => setFillColId(e.target.value)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
            >
              {sheet.columns.map((c, idx) => (
                <option key={c.id} value={c.id}>
                  {c.shortName}({getOriginColumnRoleBadge(sheet, idx)}) - {c.longName || c.shortName}
                </option>
              ))}
            </select>

            <select
              value={fillMode}
              onChange={(e) => setFillMode(e.target.value as AutoFillMode)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs font-medium"
            >
              <option value="row_numbers">Numeri di Riga (1, 2, 3, ..., N)</option>
              <option value="linear_step">Step Lineare / Temporale Costante (Start, ΔStep)</option>
              <option value="linspace">Intervallo Equispaziato [Start → Stop] in N punti</option>
              <option value="random_uniform">Numeri Casuali Uniformi [Min, Max]</option>
              <option value="random_normal">Numeri Casuali Normali Gaussiani (μ, σ)</option>
            </select>

            <label className="flex items-center gap-1">
              <span className="text-slate-600">N Righe:</span>
              <input
                type="number"
                min={1}
                max={1000000}
                value={fillRows}
                onChange={(e) =>
                  setFillRows(Math.max(1, Math.min(1000000, Number(e.target.value) || 20)))
                }
                className="w-20 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
              />
            </label>

            {fillMode === 'linear_step' && (
              <>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Inizio:</span>
                  <input
                    type="number"
                    step="any"
                    value={fillStart}
                    onChange={(e) => setFillStart(Number(e.target.value) || 0)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Step Δ:</span>
                  <input
                    type="number"
                    step="any"
                    value={fillStep}
                    onChange={(e) => setFillStep(Number(e.target.value) || 1)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
              </>
            )}

            {fillMode === 'linspace' && (
              <>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Da:</span>
                  <input
                    type="number"
                    step="any"
                    value={fillStart}
                    onChange={(e) => setFillStart(Number(e.target.value) || 0)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">A:</span>
                  <input
                    type="number"
                    step="any"
                    value={fillStop}
                    onChange={(e) => setFillStop(Number(e.target.value) || 10)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
              </>
            )}

            {fillMode === 'random_uniform' && (
              <>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Min:</span>
                  <input
                    type="number"
                    step="any"
                    value={fillMin}
                    onChange={(e) => setFillMin(Number(e.target.value) || 0)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Max:</span>
                  <input
                    type="number"
                    step="any"
                    value={fillMax}
                    onChange={(e) => setFillMax(Number(e.target.value) || 1)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
              </>
            )}

            {fillMode === 'random_normal' && (
              <>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Media μ:</span>
                  <input
                    type="number"
                    step="any"
                    value={fillMean}
                    onChange={(e) => setFillMean(Number(e.target.value) || 0)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Dev.Std σ:</span>
                  <input
                    type="number"
                    step="any"
                    value={fillStd}
                    onChange={(e) => setFillStd(Number(e.target.value) || 1)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
              </>
            )}
          </div>

          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={handleExecuteAutoFill}
              className="px-2.5 py-0.5 bg-indigo-700 text-white font-semibold rounded hover:bg-indigo-800 transition-colors"
            >
              Genera Sequenza
            </button>
            <button
              type="button"
              onClick={() => setActiveSubPanel('none')}
              className="px-1.5 py-0.5 text-slate-500 hover:text-slate-800"
            >
              Chiudi
            </button>
          </div>
        </div>
      )}

      {activeSubPanel === 'masking' && (
        <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-1.5 bg-red-50/90 border-b border-red-200 text-xs shrink-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-bold text-red-950 flex items-center gap-1">
              <EyeOff className="w-3.5 h-3.5 text-red-700" />
              Mascheramento Dati (Uint8Array Mask):
            </span>

            <select
              value={maskTargetColId}
              onChange={(e) => setMaskTargetColId(e.target.value)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
            >
              <option value="ALL_Y">Tutte le Colonne Serie Y</option>
              {sheet.columns.map((c, idx) => (
                <option key={c.id} value={c.id}>
                  {c.shortName}({getOriginColumnRoleBadge(sheet, idx)}) - {c.longName || c.shortName}
                </option>
              ))}
            </select>

            <select
              value={maskMode}
              onChange={(e) => setMaskMode(e.target.value as MaskOperationMode)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs font-medium"
            >
              <option value="mask_range">Maschera Intervallo di Righe [Da..A]</option>
              <option value="unmask_range">Smaschera Intervallo di Righe [Da..A]</option>
              <option value="mask_greater">Maschera Valori &gt; Soglia</option>
              <option value="mask_less">Maschera Valori &lt; Soglia</option>
              <option value="mask_outside">Maschera Fuori Intervallo [Min, Max]</option>
              <option value="invert_mask">Inverti Maschera Attuale</option>
              <option value="unmask_all">Smaschera Tutti i Punti</option>
            </select>

            {(maskMode === 'mask_range' || maskMode === 'unmask_range') && (
              <>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Da Riga:</span>
                  <input
                    type="number"
                    min={1}
                    value={maskStartRow}
                    onChange={(e) => setMaskStartRow(Math.max(1, Number(e.target.value) || 1))}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">A Riga:</span>
                  <input
                    type="number"
                    min={1}
                    value={maskEndRow}
                    onChange={(e) => setMaskEndRow(Math.max(1, Number(e.target.value) || 5))}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
              </>
            )}

            {(maskMode === 'mask_greater' || maskMode === 'mask_less') && (
              <label className="flex items-center gap-1">
                <span className="text-slate-600">Soglia:</span>
                <input
                  type="number"
                  step="any"
                  value={maskThreshold}
                  onChange={(e) => setMaskThreshold(Number(e.target.value) || 0)}
                  className="w-20 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                />
              </label>
            )}

            {maskMode === 'mask_outside' && (
              <>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Min:</span>
                  <input
                    type="number"
                    step="any"
                    value={maskMinThresh}
                    onChange={(e) => setMaskMinThresh(Number(e.target.value) || 0)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Max:</span>
                  <input
                    type="number"
                    step="any"
                    value={maskMaxThresh}
                    onChange={(e) => setMaskMaxThresh(Number(e.target.value) || 100)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
              </>
            )}
          </div>

          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={handleExecuteMaskOperation}
              className="px-2.5 py-0.5 bg-red-700 text-white font-semibold rounded hover:bg-red-800 transition-colors"
            >
              Applica Maschera
            </button>
            <button
              type="button"
              onClick={() => setActiveSubPanel('none')}
              className="px-1.5 py-0.5 text-slate-500 hover:text-slate-800"
            >
              Chiudi
            </button>
          </div>
        </div>
      )}

      {/* SUBPANEL: NUMERIC FORMATTING */}
      {activeSubPanel === 'num_format' && (
        <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-1.5 bg-sky-50/95 border-b border-sky-200 text-xs shrink-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-bold text-sky-950 flex items-center gap-1">
              <Hash className="w-3.5 h-3.5 text-sky-700" />
              Formattazione Numerica:
            </span>

            <select
              value={fmtTargetColId}
              onChange={(e) => setFmtTargetColId(e.target.value)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
            >
              <option value="ALL_NUMERIC">Tutte le Colonne Numeriche</option>
              {sheet.columns.map((c, idx) => (
                <option key={c.id} value={c.id}>
                  {c.shortName}({getOriginColumnRoleBadge(sheet, idx)}) - {c.longName || c.shortName}
                </option>
              ))}
            </select>

            <select
              value={fmtMode}
              onChange={(e) => setFmtMode(e.target.value as NumericFormatMode)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs font-medium"
            >
              <option value="auto">Standard / Automatico (es. 0.00123)</option>
              <option value="decimal">Decimale Fisso (es. 12.340)</option>
              <option value="scientific">Scientifica (1.23 × 10⁻⁴)</option>
              <option value="engineering">Ingegneristica (Esponente 10³ⁿ: 123.4 × 10⁻⁶)</option>
            </select>

            <select
              value={fmtPrecisionMode}
              onChange={(e) => setFmtPrecisionMode(e.target.value as NumericPrecisionMode)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs"
            >
              <option value="decimals">Cifre Decimali</option>
              <option value="sigfigs">Cifre Significative</option>
            </select>

            <label className="flex items-center gap-1">
              <span className="text-slate-600">Cifre:</span>
              <input
                type="number"
                min={0}
                max={10}
                value={fmtPrecision}
                onChange={(e) =>
                  setFmtPrecision(Math.max(0, Math.min(10, Number(e.target.value) || 0)))
                }
                className="w-14 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
              />
            </label>
          </div>

          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={handleApplyNumericFormat}
              className="px-2.5 py-0.5 bg-sky-700 text-white font-semibold rounded hover:bg-sky-800 transition-colors"
            >
              Applica Formato
            </button>
            <button
              type="button"
              onClick={() => setActiveSubPanel('none')}
              className="px-1.5 py-0.5 text-slate-500 hover:text-slate-800"
            >
              Chiudi
            </button>
          </div>
        </div>
      )}

      {/* SUBPANEL: FIND & REPLACE (INSTRUMENTAL ERROR CLEANUP) */}
      {activeSubPanel === 'find_replace' && (
        <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-1.5 bg-amber-50/95 border-b border-amber-200 text-xs shrink-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-bold text-amber-950 flex items-center gap-1">
              <Search className="w-3.5 h-3.5 text-amber-700" />
              Trova e Sostituisci:
            </span>

            <select
              value={findTargetColId}
              onChange={(e) => setFindTargetColId(e.target.value)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
            >
              <option value="ALL">Intero Foglio</option>
              {sheet.columns.map((c, idx) => (
                <option key={c.id} value={c.id}>
                  {c.shortName}({getOriginColumnRoleBadge(sheet, idx)}) - {c.longName || c.shortName}
                </option>
              ))}
            </select>

            <select
              value={findMatchMode}
              onChange={(e) => setFindMatchMode(e.target.value as FindMatchMode)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs font-medium"
            >
              <option value="exact">Valore o Stringa Esatta (es. -9999)</option>
              <option value="instrument_errors">
                Codici Errore Strumentali (-9999, -999, 9999, ERR, OVERFLOW)
              </option>
              <option value="contains">Contiene Sottostringa</option>
              <option value="lt">Valori &lt; Soglia</option>
              <option value="gt">Valori &gt; Soglia</option>
              <option value="outside_range">Valori Fuori Range [Min, Max]</option>
            </select>

            {findMatchMode !== 'instrument_errors' && findMatchMode !== 'outside_range' && (
              <label className="flex items-center gap-1">
                <span className="text-slate-600">Cerca:</span>
                <input
                  type="text"
                  value={findQuery}
                  onChange={(e) => setFindQuery(e.target.value)}
                  placeholder="es. -9999"
                  className="w-24 px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                />
              </label>
            )}

            {findMatchMode === 'outside_range' && (
              <>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Min:</span>
                  <input
                    type="number"
                    step="any"
                    value={findMin}
                    onChange={(e) => setFindMin(Number(e.target.value) || 0)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">Max:</span>
                  <input
                    type="number"
                    step="any"
                    value={findMax}
                    onChange={(e) => setFindMax(Number(e.target.value) || 100)}
                    className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
              </>
            )}

            <select
              value={replaceAction}
              onChange={(e) => setReplaceAction(e.target.value as ReplaceActionMode)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs font-medium"
            >
              <option value="replace_nan">Sostituisci con NaN / Cella Vuota (--)</option>
              <option value="replace_value">Sostituisci con Nuovo Valore / Testo</option>
              <option value="mask_matched">Maschera Celle Trovate (Uint8=1)</option>
            </select>

            {replaceAction === 'replace_value' && (
              <label className="flex items-center gap-1">
                <span className="text-slate-600">Con:</span>
                <input
                  type="text"
                  value={replaceWith}
                  onChange={(e) => setReplaceWith(e.target.value)}
                  placeholder="Nuovo valore..."
                  className="w-20 px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                />
              </label>
            )}
          </div>

          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={handleExecuteFindReplace}
              className="px-2.5 py-0.5 bg-amber-700 text-white font-semibold rounded hover:bg-amber-800 transition-colors"
            >
              Sostituisci Tutto
            </button>
            <button
              type="button"
              onClick={() => setActiveSubPanel('none')}
              className="px-1.5 py-0.5 text-slate-500 hover:text-slate-800"
            >
              Chiudi
            </button>
          </div>
        </div>
      )}

      {/* SUBPANEL: CONDITIONAL ROW FILTER */}
      {activeSubPanel === 'row_filter' && (
        <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-1.5 bg-teal-50/95 border-b border-teal-200 text-xs shrink-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-bold text-teal-950 flex items-center gap-1">
              <Filter className="w-3.5 h-3.5 text-teal-700" />
              Filtro Righe Condizionale:
            </span>

            {filterOp !== 'custom_expr' && (
              <select
                value={filterColId}
                onChange={(e) => setFilterColId(e.target.value)}
                className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
              >
                {sheet.columns.map((c, idx) => (
                  <option key={c.id} value={c.id}>
                    Col({c.shortName}) [{getOriginColumnRoleBadge(sheet, idx)}] - {c.longName || c.shortName}
                  </option>
                ))}
              </select>
            )}

            <select
              value={filterOp}
              onChange={(e) => setFilterOp(e.target.value as RowFilterOperator)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs font-medium"
            >
              <option value="gt">&gt; Maggiore di</option>
              <option value="gte">&ge; Maggiore o uguale a</option>
              <option value="lt">&lt; Minore di</option>
              <option value="lte">&le; Minore o uguale a</option>
              <option value="eq">= Uguale a</option>
              <option value="neq">&ne; Diverso da</option>
              <option value="between">Compreso tra [Min .. Max]</option>
              <option value="contains">Contiene Testo</option>
              <option value="not_nan">Solo Celle Valide (Non-NaN)</option>
              <option value="unmasked">Solo Celle Non Mascherate</option>
              <option value="custom_expr">Espressione Vettoriale (es. col(B) &gt; 0)</option>
            </select>

            {filterOp === 'custom_expr' ? (
              <input
                type="text"
                value={filterCustomExpr}
                onChange={(e) => setFilterCustomExpr(e.target.value)}
                placeholder="es. col(B) > 0 && col(A) <= 10"
                className="w-56 px-2 py-0.5 bg-white border border-teal-400 rounded font-mono text-xs"
              />
            ) : (
              filterOp !== 'not_nan' &&
              filterOp !== 'unmasked' && (
                <>
                  <input
                    type="text"
                    value={filterVal1}
                    onChange={(e) => setFilterVal1(e.target.value)}
                    placeholder="Valore..."
                    className="w-20 px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                  {filterOp === 'between' && (
                    <>
                      <span className="text-slate-500">e</span>
                      <input
                        type="text"
                        value={filterVal2}
                        onChange={(e) => setFilterVal2(e.target.value)}
                        placeholder="Max..."
                        className="w-20 px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                      />
                    </>
                  )}
                </>
              )
            )}

            {sheet.rowFilter?.enabled && (
              <span className="px-1.5 py-0.5 bg-teal-200/80 text-teal-950 font-mono font-semibold rounded text-[11px]">
                Visibili: {effectiveRowCount} / {maxRows} righe
              </span>
            )}
          </div>

          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={() => handleApplyRowFilter(true)}
              className="px-2.5 py-0.5 bg-teal-700 text-white font-semibold rounded hover:bg-teal-800 transition-colors"
            >
              Applica Filtro
            </button>
            {sheet.rowFilter?.enabled && (
              <>
                <button
                  type="button"
                  onClick={handleExtractFilteredToNewSheet}
                  className="px-2 py-0.5 bg-white border border-teal-600 text-teal-900 font-semibold rounded hover:bg-teal-100 transition-colors"
                  title="Crea un nuovo foglio contenente solo le righe filtrate"
                >
                  Estrai in Nuovo Foglio
                </button>
                <button
                  type="button"
                  onClick={() => handleApplyRowFilter(false)}
                  className="px-2 py-0.5 bg-red-50 border border-red-300 text-red-800 font-medium rounded hover:bg-red-100"
                >
                  Reset Filtro
                </button>
              </>
            )}
            <button
              type="button"
              onClick={() => setActiveSubPanel('none')}
              className="px-1.5 py-0.5 text-slate-500 hover:text-slate-800"
            >
              Chiudi
            </button>
          </div>
        </div>
      )}

      {/* SUBPANEL: SPLIT WORKSHEET */}
      {activeSubPanel === 'split_sheet' && (
        <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-1.5 bg-violet-50/95 border-b border-violet-200 text-xs shrink-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-bold text-violet-950 flex items-center gap-1">
              <Split className="w-3.5 h-3.5 text-violet-700" />
              Divisione Foglio (Split):
            </span>

            <select
              value={splitMode}
              onChange={(e) => setSplitMode(e.target.value as SplitWorksheetMode)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs font-medium"
            >
              <option value="by_category_sheets">
                Separa in Più Fogli Distinti per Valore Categorico della Colonna
              </option>
              <option value="by_category_columns">
                Separa in Colonne Affiancate Multi-X (Unstack per Categoria)
              </option>
              <option value="by_row_chunks">
                Separa in Fogli Ogni N Righe (Blocchi Sequenziali)
              </option>
            </select>

            {splitMode === 'by_row_chunks' ? (
              <label className="flex items-center gap-1">
                <span className="text-slate-600">Righe per blocco:</span>
                <input
                  type="number"
                  min={2}
                  max={500000}
                  value={splitChunkRows}
                  onChange={(e) => setSplitChunkRows(Math.max(2, Number(e.target.value) || 10))}
                  className="w-20 px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                />
              </label>
            ) : (
              <label className="flex items-center gap-1">
                <span className="text-slate-600">Colonna Categorica / Gruppo:</span>
                <select
                  value={splitGroupColId}
                  onChange={(e) => setSplitGroupColId(e.target.value)}
                  className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                >
                  {sheet.columns.map((c, idx) => (
                    <option key={c.id} value={c.id}>
                      {c.shortName}({getOriginColumnRoleBadge(sheet, idx)}) - {c.longName || c.shortName}
                    </option>
                  ))}
                </select>
              </label>
            )}
          </div>

          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={handleExecuteSplitWorksheet}
              className="px-2.5 py-0.5 bg-violet-700 text-white font-semibold rounded hover:bg-violet-800 transition-colors"
            >
              Esegui Divisione (Split)
            </button>
            <button
              type="button"
              onClick={() => setActiveSubPanel('none')}
              className="px-1.5 py-0.5 text-slate-500 hover:text-slate-800"
            >
              Chiudi
            </button>
          </div>
        </div>
      )}

      {/* SUBPANEL: SET COLUMN VALUES (VECTORIAL FORMULA ENGINE & AUTO-RECALC) */}
      {activeSubPanel === 'formula_calc' && (
        <div className="flex flex-col gap-1.5 px-3 py-1.5 bg-emerald-50/95 border-b border-emerald-200 text-xs shrink-0">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2 flex-wrap flex-1">
              <span className="font-bold text-emerald-950 flex items-center gap-1">
                <Calculator className="w-3.5 h-3.5 text-emerald-700" />
                Set Column Values F(x):
              </span>

              <select
                value={calcTargetColId}
                onChange={(e) => {
                  const id = e.target.value;
                  setCalcTargetColId(id);
                  const found = sheet.columns.find((c) => c.id === id);
                  if (found?.formula) setCalcExpr(found.formula);
                }}
                className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs font-semibold"
              >
                <option value="NEW_COL">+ Crea Nuova Colonna Calcolata</option>
                {sheet.columns.map((c, idx) => (
                  <option key={c.id} value={c.id}>
                    Col({c.shortName}) [{getOriginColumnRoleBadge(sheet, idx)}] - {c.longName || c.shortName}
                  </option>
                ))}
              </select>

              <span className="font-mono font-bold text-emerald-800">=</span>

              <input
                type="text"
                value={calcExpr}
                onChange={(e) => setCalcExpr(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleExecuteSetColumnValuesPanel();
                }}
                placeholder="es. col(A) * 2 + sin(col(B)) oppure col(C) = col(A)^2"
                className="flex-1 min-w-[240px] px-2 py-0.5 bg-white border border-emerald-500 rounded font-mono text-xs text-slate-900"
              />
            </div>

            <div className="flex items-center gap-1.5">
              <label
                className="flex items-center gap-1 px-2 py-0.5 bg-white border border-emerald-300 rounded text-[11px] font-medium text-emerald-900 cursor-pointer"
                title="Ricalcolo Automatico: aggiorna istantaneamente le colonne con formula e i grafici quando i dati di origine cambiano"
              >
                <input
                  type="checkbox"
                  checked={sheet.autoRecalculate !== false}
                  onChange={(e) =>
                    onUpdateSheet(
                      recalculateSheetFormulas({
                        ...sheet,
                        autoRecalculate: e.target.checked,
                      })
                    )
                  }
                />
                <RefreshCw className="w-3 h-3 text-emerald-700" />
                <span>Auto-Recalc</span>
              </label>

              <button
                type="button"
                onClick={handleExecuteSetColumnValuesPanel}
                className="px-2.5 py-0.5 bg-emerald-700 text-white font-semibold rounded hover:bg-emerald-800 transition-colors"
              >
                Calcola Vettore
              </button>
              <button
                type="button"
                onClick={() => setActiveSubPanel('none')}
                className="px-1.5 py-0.5 text-slate-500 hover:text-slate-800"
              >
                Chiudi
              </button>
            </div>
          </div>

          <div className="flex items-center gap-1 flex-wrap text-[10.5px]">
            <span className="text-slate-500 font-medium">Inserisci rapido:</span>
            {sheet.columns.map((c) => (
              <button
                key={c.id}
                type="button"
                onClick={() => setCalcExpr((prev) => `${prev} col(${c.shortName})`.trim())}
                className="px-1.5 py-0 bg-white border border-slate-300 rounded font-mono text-slate-700 hover:bg-emerald-100"
              >
                col({c.shortName})
              </button>
            ))}
            {[
              { label: 'sin()', snippet: 'sin(col(A))' },
              { label: 'cos()', snippet: 'cos(col(A))' },
              { label: 'exp()', snippet: 'exp(-col(A))' },
              { label: 'ln()', snippet: 'ln(col(A))' },
              { label: 'sqrt()', snippet: 'sqrt(col(A))' },
              { label: 'mean()', snippet: 'col(B) - mean(col(B))' },
              { label: 'norm[0,1]', snippet: 'norm(col(B))' },
              { label: 'zscore()', snippet: 'zscore(col(B))' },
              { label: 'cumsum()', snippet: 'cumsum(col(B))' },
            ].map((fn) => (
              <button
                key={fn.label}
                type="button"
                onClick={() => setCalcExpr(fn.snippet)}
                className="px-1.5 py-0 bg-emerald-100/80 border border-emerald-300 rounded font-mono text-emerald-900 hover:bg-emerald-200"
              >
                {fn.label}
              </button>
            ))}
          </div>
        </div>
      )}

      {/* SUBPANEL: ELEMENTARY ERROR PROPAGATION (ΔZ) */}
      {activeSubPanel === 'error_prop' && (
        <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-1.5 bg-purple-50/95 border-b border-purple-200 text-xs shrink-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-bold text-purple-950 flex items-center gap-1">
              <Activity className="w-3.5 h-3.5 text-purple-700" />
              Propagazione Errori (Z ± ΔZ):
            </span>

            <select
              value={errPropOp}
              onChange={(e) => setErrPropOp(e.target.value as ErrorPropagationOp)}
              className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-xs font-semibold text-purple-950"
            >
              <option value="add">Somma: Z = A + B → ΔZ = √(ΔA² + ΔB²)</option>
              <option value="sub">Sottrazione: Z = A − B → ΔZ = √(ΔA² + ΔB²)</option>
              <option value="mul">Prodotto: Z = A · B → ΔZ = |Z|·√((ΔA/A)² + (ΔB/B)²)</option>
              <option value="div">Rapporto: Z = A / B → ΔZ = |Z|·√((ΔA/A)² + (ΔB/B)²)</option>
              <option value="pow">Potenza: Z = Aⁿ → ΔZ = |n·Aⁿ⁻¹|·ΔA</option>
              <option value="linear_comb">Comb. Lineare: Z = k₁A + k₂B</option>
            </select>

            <label className="flex items-center gap-1">
              <span className="text-slate-600 font-semibold">A:</span>
              <select
                value={errPropColAId}
                onChange={(e) => setErrPropColAId(e.target.value)}
                className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
              >
                {sheet.columns.map((c) => (
                  <option key={c.id} value={c.id}>
                    Col({c.shortName}) {c.longName}
                  </option>
                ))}
              </select>
            </label>

            <label className="flex items-center gap-1">
              <span className="text-slate-600">ΔA:</span>
              <select
                value={errPropErrAMode}
                onChange={(e) => setErrPropErrAMode(e.target.value as 'column' | 'constant')}
                className="px-1 py-0.5 bg-white border border-slate-300 rounded text-xs"
              >
                <option value="constant">Valore Costante</option>
                <option value="column">Da Colonna</option>
              </select>
              {errPropErrAMode === 'column' ? (
                <select
                  value={errPropColErrAId}
                  onChange={(e) => setErrPropColErrAId(e.target.value)}
                  className="px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                >
                  {sheet.columns.map((c) => (
                    <option key={c.id} value={c.id}>
                      Col({c.shortName})
                    </option>
                  ))}
                </select>
              ) : (
                <input
                  type="number"
                  step="any"
                  value={errPropConstA}
                  onChange={(e) => setErrPropConstA(Number(e.target.value) || 0)}
                  className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                />
              )}
            </label>

            {errPropOp === 'pow' ? (
              <label className="flex items-center gap-1">
                <span className="text-slate-600 font-semibold">Esponente n:</span>
                <input
                  type="number"
                  step="any"
                  value={errPropExpN}
                  onChange={(e) => setErrPropExpN(Number(e.target.value) || 1)}
                  className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                />
              </label>
            ) : (
              <>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600 font-semibold">B:</span>
                  <select
                    value={errPropColBId}
                    onChange={(e) => setErrPropColBId(e.target.value)}
                    className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  >
                    {sheet.columns.map((c) => (
                      <option key={c.id} value={c.id}>
                        Col({c.shortName}) {c.longName}
                      </option>
                    ))}
                  </select>
                </label>

                <label className="flex items-center gap-1">
                  <span className="text-slate-600">ΔB:</span>
                  <select
                    value={errPropErrBMode}
                    onChange={(e) => setErrPropErrBMode(e.target.value as 'column' | 'constant')}
                    className="px-1 py-0.5 bg-white border border-slate-300 rounded text-xs"
                  >
                    <option value="constant">Valore Costante</option>
                    <option value="column">Da Colonna</option>
                  </select>
                  {errPropErrBMode === 'column' ? (
                    <select
                      value={errPropColErrBId}
                      onChange={(e) => setErrPropColErrBId(e.target.value)}
                      className="px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                    >
                      {sheet.columns.map((c) => (
                        <option key={c.id} value={c.id}>
                          Col({c.shortName})
                        </option>
                      ))}
                    </select>
                  ) : (
                    <input
                      type="number"
                      step="any"
                      value={errPropConstB}
                      onChange={(e) => setErrPropConstB(Number(e.target.value) || 0)}
                      className="w-16 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                    />
                  )}
                </label>
              </>
            )}

            {errPropOp === 'linear_comb' && (
              <>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">k₁:</span>
                  <input
                    type="number"
                    step="any"
                    value={errPropCoeffA}
                    onChange={(e) => setErrPropCoeffA(Number(e.target.value) || 1)}
                    className="w-14 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
                <label className="flex items-center gap-1">
                  <span className="text-slate-600">k₂:</span>
                  <input
                    type="number"
                    step="any"
                    value={errPropCoeffB}
                    onChange={(e) => setErrPropCoeffB(Number(e.target.value) || 1)}
                    className="w-14 px-1 py-0.5 bg-white border border-slate-300 rounded font-mono text-xs"
                  />
                </label>
              </>
            )}
          </div>

          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={handleExecuteErrorPropagation}
              className="px-2.5 py-0.5 bg-purple-700 text-white font-semibold rounded hover:bg-purple-800 transition-colors"
            >
              Genera Colonne Z e ΔZ (±)
            </button>
            <button
              type="button"
              onClick={() => setActiveSubPanel('none')}
              className="px-1.5 py-0.5 text-slate-500 hover:text-slate-800"
            >
              Chiudi
            </button>
          </div>
        </div>
      )}

      {/* Row 2: Plot Type Selection & Plot to Graph Bar */}
      <div className="flex flex-wrap items-center justify-between gap-2 px-2.5 py-1 bg-[#f8fafc] border-b border-slate-300 text-xs shrink-0">
        <div className="flex items-center gap-2 flex-wrap">
          <label className="flex items-center gap-1.5 text-[11px] font-semibold text-slate-700">
            <span>Tipo Grafico:</span>
            <select
              value={quickPlotType}
              onChange={(e) => handlePlotTypeSelectChange(e.target.value as PlotType)}
              className="px-2 py-0.5 text-[11px] font-semibold bg-white border border-slate-300 rounded text-blue-900 focus:outline-none cursor-pointer"
            >
              {PLOT_TYPE_LABELS.map((pt) => (
                <option key={pt.id} value={pt.id}>
                  {pt.label}
                </option>
              ))}
            </select>
          </label>

          <label className="flex items-center gap-1 text-[11px] text-slate-600">
            <span>Connessione:</span>
            <select
              value={quickConnection}
              onChange={(e) => setQuickConnection(e.target.value as ConnectionType)}
              className="px-1.5 py-0.5 text-[11px] bg-white border border-slate-300 rounded text-slate-800 focus:outline-none cursor-pointer"
            >
              <option value="linear">Lineare (Spezzata)</option>
              <option value="spline">Spline Cubica</option>
              <option value="step">Step (Gradino)</option>
              <option value="none">Nessuna (Solo Simboli)</option>
            </select>
          </label>

          {xColumns.length > 1 && (
            <span className="px-1.5 py-0.5 text-[10px] font-mono font-semibold bg-amber-100 text-amber-900 border border-amber-300 rounded">
              Multi-X Attivo ({xColumns.length} Assi X: associazione automatica a sinistra)
            </span>
          )}
        </div>

        <div className="flex items-center gap-1.5">
          <button
            type="button"
            disabled={activeYColsForPlot.length === 0}
            onClick={() =>
              onPlotColumnsToActiveGraph(
                sheet.id,
                activeYColsForPlot,
                quickPlotType,
                quickConnection
              )
            }
            className="flex items-center gap-1 px-2.5 py-0.5 text-[11px] font-semibold text-white bg-blue-700 rounded hover:bg-blue-800 disabled:opacity-40 transition-colors whitespace-nowrap shadow-2xs"
          >
            <LineChart className="w-3 h-3" />
            <span>Traccia sul Grafico Attivo ({activeYColsForPlot.length})</span>
          </button>

          <button
            type="button"
            disabled={activeYColsForPlot.length === 0}
            onClick={() =>
              onPlotColumnsToNewGraph(sheet.id, activeYColsForPlot, quickPlotType, quickConnection)
            }
            className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium text-blue-800 bg-blue-50 border border-blue-300 rounded hover:bg-blue-100 disabled:opacity-40 transition-colors whitespace-nowrap"
          >
            <Plus className="w-3 h-3" />
            <span>Nuovo Grafico</span>
          </button>
        </div>
      </div>

      {/* RULE 4: Total DOM Virtualized Data Grid with Pinned OriginLab Metadata Rows */}
      <div
        ref={scrollContainerRef}
        onScroll={handleScroll}
        className="flex-1 overflow-auto relative min-h-0"
      >
        <table className="w-full border-collapse text-left select-none text-xs">
          <thead className="sticky top-0 z-20 bg-[#e2e8f0] border-b-2 border-slate-400 shadow-2xs">
            {/* Pinned Row 1: Designation (X1, Y1, X2, Y2, xEr±, yEr±, L, Ign), Role Selector, DataType, Drag & Drop, Quick Sort */}
            <tr className="border-b border-slate-300 bg-[#edf2f7]">
              <th className="w-16 px-1.5 py-1 text-center text-[10px] font-mono font-bold text-slate-600 bg-[#cbd5e1] border-r border-slate-300 sticky left-0 z-30">
                Col
              </th>
              {sheet.columns.map((col, colIdx) => {
                const isX = col.role === 'X';
                const isY = col.role === 'Y';
                const isYErr = col.role === 'Y_ERR';
                const isXErr = col.role === 'X_ERR';
                const isLabel = col.role === 'LABEL';
                const isSelected = selectedYColIds.includes(col.id);
                const roleBadge = getOriginColumnRoleBadge(sheet, colIdx);
                const resolvedX = isY ? resolveXColumnForY(sheet, col) : undefined;
                const companions = isY ? resolveErrorAndLabelColumnsForY(sheet, col) : undefined;
                const colDataType: ColumnDataType =
                  col.dataType || (col.role === 'LABEL' ? 'text' : 'numeric');
                const isDragTarget = dragOverColIndex === colIdx && draggedColIndex !== colIdx;

                return (
                  <th
                    key={col.id}
                    draggable
                    onDragStart={(e) => {
                      setDraggedColIndex(colIdx);
                      e.dataTransfer.effectAllowed = 'move';
                    }}
                    onDragOver={(e) => {
                      e.preventDefault();
                      e.dataTransfer.dropEffect = 'move';
                      if (dragOverColIndex !== colIdx) {
                        setDragOverColIndex(colIdx);
                      }
                    }}
                    onDragLeave={() => {
                      if (dragOverColIndex === colIdx) {
                        setDragOverColIndex(null);
                      }
                    }}
                    onDrop={(e) => {
                      e.preventDefault();
                      if (draggedColIndex !== null && draggedColIndex !== colIdx) {
                        handleMoveColumn(draggedColIndex, colIdx);
                      }
                      setDraggedColIndex(null);
                      setDragOverColIndex(null);
                    }}
                    onDragEnd={() => {
                      setDraggedColIndex(null);
                      setDragOverColIndex(null);
                    }}
                    className={`min-w-[185px] px-1.5 py-1 border-r border-slate-300 transition-colors ${
                      isDragTarget ? 'ring-2 ring-blue-600 bg-blue-200/80' : ''
                    } ${
                      isX
                        ? 'bg-amber-100/75 text-amber-950'
                        : isY
                        ? 'bg-blue-100/75 text-blue-950'
                        : isYErr
                        ? 'bg-purple-100/75 text-purple-950'
                        : isXErr
                        ? 'bg-indigo-100/75 text-indigo-950'
                        : isLabel
                        ? 'bg-teal-100/75 text-teal-950'
                        : 'bg-slate-200/80 text-slate-700'
                    }`}
                  >
                    <div className="flex items-center justify-between gap-1">
                      <div className="flex items-center gap-1 min-w-0">
                        <span
                          className="cursor-grab active:cursor-grabbing text-slate-400 hover:text-slate-700 shrink-0"
                          title="Trascina per riordinare la colonna (Drag & Drop)"
                        >
                          <GripVertical className="w-3.5 h-3.5" />
                        </span>

                        {(isY || isYErr) && (
                          <button
                            type="button"
                            onClick={() => toggleSelectYCol(col.id)}
                            className="text-blue-700 hover:text-blue-900 focus:outline-none shrink-0"
                            title="Seleziona colonna per tracciamento sul grafico"
                          >
                            {isSelected ? (
                              <CheckSquare className="w-3.5 h-3.5 text-blue-700" />
                            ) : (
                              <Square className="w-3.5 h-3.5 text-slate-400" />
                            )}
                          </button>
                        )}

                        <span
                          className="font-mono text-xs font-bold truncate"
                          title={`Colonna ${col.shortName} · Ruolo: ${roleBadge} · ${getColumnRowCount(col).toLocaleString('it-IT')} valori`}
                        >
                          {col.shortName}({roleBadge})
                        </span>
                      </div>

                      <div className="flex items-center gap-0.5 shrink-0">
                        <select
                          value={col.role}
                          onChange={(e) => handleRoleChange(col.id, e.target.value as ColumnRole)}
                          className="px-1 py-0.5 text-[10px] font-mono font-semibold bg-white border border-slate-300 rounded text-slate-800 cursor-pointer focus:outline-none"
                          title="Ruolo dedicato della colonna (X, Y, ΔX, ΔY, Etichetta L, Ignora)"
                        >
                          <option value="X">X (Asse X)</option>
                          <option value="Y">Y (Serie Y)</option>
                          <option value="X_ERR">ΔX (xErr ±)</option>
                          <option value="Y_ERR">ΔY (yErr ±)</option>
                          <option value="LABEL">L (Etichetta)</option>
                          <option value="IGNORE">Ignora</option>
                        </select>

                        {sheet.columns.length > 1 && (
                          <button
                            type="button"
                            onClick={() => handleDeleteColumn(col.id)}
                            className="p-0.5 text-slate-400 hover:text-red-700 transition-colors"
                            title="Elimina colonna"
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        )}
                      </div>
                    </div>

                    <div className="mt-1 flex items-center justify-between gap-1 text-[10px] font-normal">
                      <div className="flex items-center gap-0.5">
                        <select
                          value={colDataType}
                          onChange={(e) =>
                            handleDataTypeChange(col.id, e.target.value as ColumnDataType)
                          }
                          className="px-1 py-0 text-[9.5px] font-mono bg-white/90 border border-slate-300 rounded text-slate-700 cursor-pointer focus:outline-none"
                          title="Tipizzazione dati colonna (Float64, Intero, Testo, Data/Ora)"
                        >
                          <option value="numeric">Float64</option>
                          <option value="integer">Int</option>
                          <option value="text">Testo</option>
                          <option value="datetime">Data/Ora</option>
                        </select>

                        {colDataType === 'numeric' && (
                          <select
                            value={col.numericFormat || 'auto'}
                            onChange={(e) =>
                              handleSingleColFormatChange(
                                col.id,
                                e.target.value as NumericFormatMode
                              )
                            }
                            className="px-1 py-0 text-[9.5px] font-mono bg-white/90 border border-slate-300 rounded text-sky-900 cursor-pointer focus:outline-none"
                            title="Formato numerico colonna: Automatico, Decimale, Scientifico (1.23×10⁻⁴), Ingegneristico (10³ⁿ)"
                          >
                            <option value="auto">Auto</option>
                            <option value="decimal">Dec</option>
                            <option value="scientific">Sci 10ⁿ</option>
                            <option value="engineering">Eng 10³ⁿ</option>
                          </select>
                        )}
                      </div>

                      <div className="flex items-center gap-0.5">
                        <button
                          type="button"
                          onClick={() =>
                            setCellSelection({
                              startColIdx: colIdx,
                              endColIdx: colIdx,
                              startRowIdx: 0,
                              endRowIdx: Math.max(0, maxRows - 1),
                            })
                          }
                          className="px-1 py-0 bg-white/80 hover:bg-white border border-slate-300 rounded text-teal-800 font-mono text-[9px] font-bold"
                          title="Seleziona intera colonna per Statistiche Rapide nella barra di stato"
                        >
                          Σ
                        </button>
                        <button
                          type="button"
                          onClick={() => handleQuickSortColumn(col.id, 'asc', true)}
                          className="px-1 py-0 bg-white/80 hover:bg-white border border-slate-300 rounded text-slate-600 hover:text-blue-700"
                          title="Ordina foglio in modo Crescente (↑) secondo questa colonna"
                        >
                          <ArrowUp className="w-2.5 h-2.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => handleQuickSortColumn(col.id, 'desc', true)}
                          className="px-1 py-0 bg-white/80 hover:bg-white border border-slate-300 rounded text-slate-600 hover:text-blue-700"
                          title="Ordina foglio in modo Decrescente (↓) secondo questa colonna"
                        >
                          <ArrowDown className="w-2.5 h-2.5" />
                        </button>
                        <button
                          type="button"
                          disabled={colIdx === 0}
                          onClick={() => handleMoveColumn(colIdx, colIdx - 1)}
                          className="px-0.5 py-0 bg-white/80 hover:bg-white border border-slate-300 rounded text-slate-600 disabled:opacity-30"
                          title="Sposta colonna a sinistra"
                        >
                          <ChevronLeft className="w-2.5 h-2.5" />
                        </button>
                        <button
                          type="button"
                          disabled={colIdx === sheet.columns.length - 1}
                          onClick={() => handleMoveColumn(colIdx, colIdx + 1)}
                          className="px-0.5 py-0 bg-white/80 hover:bg-white border border-slate-300 rounded text-slate-600 disabled:opacity-30"
                          title="Sposta colonna a destra"
                        >
                          <ChevronRight className="w-2.5 h-2.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => handleAddColumn('Y', colIdx + 1)}
                          className="px-1 py-0 bg-white/80 hover:bg-white border border-slate-300 rounded text-blue-700 font-bold"
                          title="Inserisci nuova colonna a destra"
                        >
                          +
                        </button>
                      </div>
                    </div>

                    {isY && (xColumns.length > 1 || yErrColumns.length > 0 || xErrColumns.length > 0) && (
                      <div className="mt-1 flex items-center justify-between gap-1 text-[9.5px] text-slate-600 font-normal flex-wrap">
                        {xColumns.length > 1 && (
                          <div className="flex items-center gap-0.5" title="Asse X associato (Multi-X)">
                            <span>Asse X:</span>
                            <select
                              value={resolvedX?.id || ''}
                              onChange={(e) => handleLinkedXChange(col.id, e.target.value)}
                              className="px-1 py-0 bg-white border border-slate-300 rounded font-mono text-[9.5px]"
                            >
                              {xColumns.map((xc) => (
                                <option key={xc.id} value={xc.id}>
                                  {xc.shortName}(X)
                                </option>
                              ))}
                            </select>
                          </div>
                        )}
                        {yErrColumns.length > 0 && (
                          <div className="flex items-center gap-0.5" title="Colonna incertezza verticale ΔY">
                            <span>ΔY:</span>
                            <select
                              value={col.linkedErrColId || companions?.yErrCol?.id || ''}
                              onChange={(e) => handleLinkedErrChange(col.id, e.target.value)}
                              className="px-1 py-0 bg-white border border-slate-300 rounded font-mono text-[9.5px]"
                            >
                              <option value="">Auto/Nessuna</option>
                              {yErrColumns.map((ec) => (
                                <option key={ec.id} value={ec.id}>
                                  {ec.shortName}
                                </option>
                              ))}
                            </select>
                          </div>
                        )}
                        {xErrColumns.length > 0 && (
                          <div className="flex items-center gap-0.5" title="Colonna incertezza orizzontale ΔX">
                            <span>ΔX:</span>
                            <select
                              value={col.linkedXErrColId || companions?.xErrCol?.id || ''}
                              onChange={(e) => handleLinkedXErrChange(col.id, e.target.value)}
                              className="px-1 py-0 bg-white border border-slate-300 rounded font-mono text-[9.5px]"
                            >
                              <option value="">Auto/Nessuna</option>
                              {xErrColumns.map((ec) => (
                                <option key={ec.id} value={ec.id}>
                                  {ec.shortName}
                                </option>
                              ))}
                            </select>
                          </div>
                        )}
                      </div>
                    )}
                  </th>
                );
              })}
            </tr>

            {/* Pinned Row 2: Long Name */}
            <tr className="border-b border-slate-300 bg-white">
              <th className="px-1.5 py-0.5 text-right text-[10px] font-semibold text-slate-600 bg-[#edf2f7] border-r border-slate-300 sticky left-0 z-30">
                Long Name
              </th>
              {sheet.columns.map((col) => (
                <th key={col.id} className="p-0 border-r border-slate-300">
                  <input
                    type="text"
                    value={col.longName}
                    onChange={(e) => handleMetaChange(col.id, 'longName', e.target.value)}
                    placeholder="Nome grandezza..."
                    className="w-full px-2 py-0.5 text-xs font-semibold text-slate-800 bg-transparent focus:bg-blue-50/50 focus:outline-none"
                  />
                </th>
              ))}
            </tr>

            {/* Pinned Row 3: Units */}
            <tr className="border-b border-slate-300 bg-slate-50">
              <th className="px-1.5 py-0.5 text-right text-[10px] font-semibold text-slate-600 bg-[#edf2f7] border-r border-slate-300 sticky left-0 z-30">
                Units
              </th>
              {sheet.columns.map((col) => (
                <th key={col.id} className="p-0 border-r border-slate-300">
                  <input
                    type="text"
                    value={col.unit}
                    onChange={(e) => handleMetaChange(col.id, 'unit', e.target.value)}
                    placeholder="es. s, V, mA, nm..."
                    className="w-full px-2 py-0.5 text-[11px] font-mono text-slate-700 bg-transparent focus:bg-blue-50/50 focus:outline-none"
                  />
                </th>
              ))}
            </tr>

            {/* Pinned Row 4: Comments */}
            <tr className="border-b border-slate-300 bg-[#f8fafc]">
              <th className="px-1.5 py-0.5 text-right text-[10px] font-semibold text-slate-500 bg-[#edf2f7] border-r border-slate-300 sticky left-0 z-30">
                Comments
              </th>
              {sheet.columns.map((col) => (
                <th key={col.id} className="p-0 border-r border-slate-300">
                  <input
                    type="text"
                    value={col.comment || ''}
                    onChange={(e) => handleMetaChange(col.id, 'comment', e.target.value)}
                    placeholder="Note / commento..."
                    className="w-full px-2 py-0.5 text-[11px] italic font-normal text-slate-600 bg-transparent focus:bg-blue-50/50 focus:outline-none"
                  />
                </th>
              ))}
            </tr>

            {/* Pinned Row 5: F(x)= Formula */}
            <tr className="bg-[#f0fdf4] border-b-2 border-slate-400">
              <th className="px-1.5 py-0.5 text-right text-[10px] font-mono font-bold text-emerald-800 bg-emerald-100 border-r border-slate-300 sticky left-0 z-30">
                F(x)=
              </th>
              {sheet.columns.map((col) => (
                <th key={col.id} className="p-0 border-r border-slate-300 relative">
                  {formulaEditingColId === col.id ? (
                    <div className="flex items-center px-1">
                      <input
                        type="text"
                        autoFocus
                        value={tempFormula}
                        onChange={(e) => setTempFormula(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            handleApplyFormula(col.id, tempFormula);
                          } else if (e.key === 'Escape') {
                            setFormulaEditingColId(null);
                          }
                        }}
                        placeholder="es. col(A)*2 + 5"
                        className="w-full px-1.5 py-0.5 text-[11px] font-mono bg-white border border-emerald-500 rounded focus:outline-none text-slate-900"
                      />
                      <button
                        type="button"
                        onClick={() => handleApplyFormula(col.id, tempFormula)}
                        className="ml-1 px-1.5 py-0.5 text-[10px] font-bold text-white bg-emerald-600 rounded hover:bg-emerald-700"
                      >
                        =
                      </button>
                    </div>
                  ) : (
                    <div
                      onClick={() => {
                        setFormulaEditingColId(col.id);
                        setTempFormula(col.formula || '');
                      }}
                      className="px-2 py-0.5 text-[11px] font-mono text-emerald-700 cursor-pointer hover:bg-emerald-50 flex items-center justify-between"
                      title="Clicca per inserire formula matematica sulla colonna (eseguita nel Math Worker su Float64Array)"
                    >
                      <span className="truncate">{col.formula || 'F(x) ...'}</span>
                      <Calculator className="w-3 h-3 text-emerald-600 shrink-0 opacity-70" />
                    </div>
                  )}
                </th>
              ))}
            </tr>
          </thead>

          {/* RULE 4: Virtualized Table Body — Only renders rows in [startIndex, endIndex] */}
          <tbody className="divide-y divide-slate-200 font-mono text-xs tabular-nums bg-white">
            {topSpacerHeight > 0 && (
              <tr style={{ height: `${topSpacerHeight}px` }} aria-hidden="true">
                <td colSpan={sheet.columns.length + 1} className="p-0 border-0" />
              </tr>
            )}

            {Array.from(
              { length: Math.max(0, endIndex - startIndex + 1) },
              (_, offset) => startIndex + offset
            ).map((virtualIdx) => {
              const rowIndex = filteredRowIndices ? filteredRowIndices[virtualIdx] : virtualIdx;
              if (rowIndex === undefined) return null;

              const isRowHighlighted =
                highlightedSheetId === sheet.id && highlightedRowIndex === rowIndex;

              const selCMin = cellSelection
                ? Math.min(cellSelection.startColIdx, cellSelection.endColIdx)
                : -1;
              const selCMax = cellSelection
                ? Math.max(cellSelection.startColIdx, cellSelection.endColIdx)
                : -1;
              const selRMin = cellSelection
                ? Math.min(cellSelection.startRowIdx, cellSelection.endRowIdx)
                : -1;
              const selRMax = cellSelection
                ? Math.max(cellSelection.startRowIdx, cellSelection.endRowIdx)
                : -1;

              return (
                <tr
                  key={rowIndex}
                  style={{ height: `${VIRTUAL_ROW_HEIGHT}px` }}
                  className={`group transition-colors ${
                    isRowHighlighted
                      ? 'bg-red-100 ring-1 ring-red-500'
                      : virtualIdx % 2 === 0
                      ? 'bg-white'
                      : 'bg-slate-50/50'
                  }`}
                >
                  {/* Row Index Column (Click to select entire row for Quick Stats) */}
                  <td
                    onClick={() =>
                      setCellSelection({
                        startColIdx: 0,
                        endColIdx: Math.max(0, sheet.columns.length - 1),
                        startRowIdx: rowIndex,
                        endRowIdx: rowIndex,
                      })
                    }
                    className={`px-1 py-0 text-center text-[11px] border-r border-slate-300 sticky left-0 z-10 select-none cursor-pointer ${
                      isRowHighlighted
                        ? 'bg-red-600 text-white font-bold'
                        : rowIndex >= selRMin && rowIndex <= selRMax
                        ? 'bg-blue-200 text-blue-950 font-semibold'
                        : 'bg-[#edf2f7] text-slate-500 group-hover:text-slate-800'
                    }`}
                    title={`Riga #${rowIndex + 1} — Clicca per selezionare l'intera riga`}
                  >
                    <div className="flex items-center justify-between gap-0.5">
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleInsertRowAt(rowIndex);
                        }}
                        className="opacity-0 group-hover:opacity-100 text-blue-600 hover:text-blue-800 font-bold transition-opacity"
                        title={`Inserisci riga vuota sopra la riga ${rowIndex + 1}`}
                      >
                        +
                      </button>
                      <span>{rowIndex + 1}</span>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleDeleteRow(rowIndex);
                        }}
                        className="opacity-0 group-hover:opacity-100 text-slate-400 hover:text-red-700 transition-opacity"
                        title={`Elimina riga ${rowIndex + 1}`}
                      >
                        ×
                      </button>
                    </div>
                  </td>

                  {/* Cell Values read in O(1) directly from Float64Array & Uint8Array */}
                  {sheet.columns.map((col, colIdx) => {
                    const cellKey = `${col.id}_${rowIndex}`;
                    const isFocused = focusedCellKey === cellKey;
                    const displayStr = formatVirtualCellO1(col, rowIndex, isFocused);
                    const isMasked = isCellMaskedO1(col, rowIndex);
                    const isTextOrLabel =
                      col.dataType === 'text' ||
                      col.dataType === 'datetime' ||
                      col.role === 'LABEL';
                    const isSelectedCell =
                      colIdx >= selCMin &&
                      colIdx <= selCMax &&
                      rowIndex >= selRMin &&
                      rowIndex <= selRMax;

                    return (
                      <td
                        key={col.id}
                        onMouseDown={(e) => {
                          if (e.shiftKey && cellSelection) {
                            setCellSelection({
                              ...cellSelection,
                              endColIdx: colIdx,
                              endRowIdx: rowIndex,
                            });
                          } else {
                            isDraggingCellsRef.current = true;
                            setCellSelection({
                              startColIdx: colIdx,
                              endColIdx: colIdx,
                              startRowIdx: rowIndex,
                              endRowIdx: rowIndex,
                            });
                          }
                        }}
                        onMouseEnter={() => {
                          if (isDraggingCellsRef.current) {
                            setCellSelection((prev) =>
                              prev
                                ? {
                                    ...prev,
                                    endColIdx: colIdx,
                                    endRowIdx: rowIndex,
                                  }
                                : null
                            );
                          }
                        }}
                        className={`p-0 border-r border-slate-200 relative group/cell ${
                          isMasked
                            ? 'bg-red-50/95 text-red-700'
                            : isSelectedCell
                            ? 'bg-blue-100/75 ring-1 ring-inset ring-blue-400/60'
                            : ''
                        }`}
                      >
                        <div className="flex items-center w-full h-[23px]">
                          <input
                            type="text"
                            inputMode={isTextOrLabel ? 'text' : 'decimal'}
                            value={displayStr}
                            onFocus={() => setFocusedCellKey(cellKey)}
                            onBlur={() =>
                              setFocusedCellKey((prev) => (prev === cellKey ? null : prev))
                            }
                            onChange={(e) => handleCellChange(col.id, rowIndex, e.target.value)}
                            onPaste={(e) => handleSmartCellPaste(e, colIdx, rowIndex)}
                            placeholder="--"
                            className={`w-full px-2 py-0 ${
                              isTextOrLabel ? 'text-left' : 'text-right'
                            } font-mono text-xs tabular-nums placeholder:text-slate-300 bg-transparent focus:outline-none focus:bg-blue-50/50 ${
                              isMasked
                                ? 'line-through text-red-600 font-semibold'
                                : isTextOrLabel
                                ? 'text-teal-900 font-medium'
                                : col.formula
                                ? 'text-emerald-900 font-semibold'
                                : 'text-slate-800'
                            }`}
                          />
                          <button
                            type="button"
                            onClick={() => handleToggleMaskCell(col.id, rowIndex)}
                            className={`p-0.5 transition-opacity mr-1 ${
                              isMasked
                                ? 'opacity-100 text-red-600 hover:text-red-800'
                                : 'opacity-0 group-hover/cell:opacity-100 text-slate-400 hover:text-red-600'
                            }`}
                            title={
                              isMasked
                                ? 'Cella mascherata (Uint8=1) — Clicca per smascherare'
                                : 'Maschera cella (Uint8=1: escludi da fit e grafico senza cancellarla)'
                            }
                          >
                            {isMasked ? (
                              <Eye className="w-2.5 h-2.5" />
                            ) : (
                              <EyeOff className="w-2.5 h-2.5" />
                            )}
                          </button>
                        </div>
                      </td>
                    );
                  })}
                </tr>
              );
            })}

            {bottomSpacerHeight > 0 && (
              <tr style={{ height: `${bottomSpacerHeight}px` }} aria-hidden="true">
                <td colSpan={sheet.columns.length + 1} className="p-0 border-0" />
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* STATUS BAR: Quick Selection Statistics (Media, Dev.Std, Somma, Min, Max, N) */}
      <div className="flex flex-wrap items-center justify-between gap-2 px-2.5 py-1 bg-[#edf2f7] border-t border-slate-300 text-[11px] font-mono text-slate-700 shrink-0 select-none">
        <div className="flex items-center gap-1.5 flex-wrap">
          <span className="font-sans font-bold text-slate-800 flex items-center gap-1">
            <Sigma className="w-3.5 h-3.5 text-blue-700" />
            <span>{quickStats.selectionLabel}:</span>
          </span>

          <span
            className="px-1.5 py-0.5 bg-white border border-slate-300 rounded font-semibold text-slate-900"
            title="Conteggio celle numeriche valide nella selezione"
          >
            N = {quickStats.n.toLocaleString('it-IT')}
          </span>

          <span
            className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-blue-900 font-semibold"
            title="Media aritmetica delle celle selezionate"
          >
            Media (μ) = {quickStats.n > 0 ? formatSciNumber(quickStats.mean, 5) : '--'}
          </span>

          <span
            className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-indigo-900"
            title="Deviazione Standard campionaria (σ) delle celle selezionate"
          >
            Dev.Std (σ) = {quickStats.n > 1 ? formatSciNumber(quickStats.sd, 5) : '0'}
          </span>

          <span
            className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-emerald-900"
            title="Somma algebrica (Σ) delle celle selezionate"
          >
            Somma (Σ) = {quickStats.n > 0 ? formatSciNumber(quickStats.sum, 5) : '0'}
          </span>

          <span
            className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-slate-800"
            title="Valore minimo e massimo nella selezione"
          >
            Min = {quickStats.n > 0 ? formatSciNumber(quickStats.min, 5) : '--'} · Max ={' '}
            {quickStats.n > 0 ? formatSciNumber(quickStats.max, 5) : '--'}
          </span>

          {quickStats.maskedCount > 0 && (
            <span className="px-1.5 py-0.5 bg-red-50 border border-red-200 rounded text-red-700">
              Mask: {quickStats.maskedCount}
            </span>
          )}
        </div>

        <div className="flex items-center gap-1 font-sans">
          {cellSelection ? (
            <button
              type="button"
              onClick={() => setCellSelection(null)}
              className="px-1.5 py-0.5 text-[10.5px] bg-white border border-slate-300 rounded hover:bg-slate-100 text-slate-600"
              title="Mostra statistiche sull'intero foglio"
            >
              Tutto il Foglio
            </button>
          ) : (
            <span className="text-[10.5px] text-slate-500 hidden lg:inline">
              Trascina sulle celle o clicca Σ sulla colonna per statistiche su selezione
            </span>
          )}
        </div>
      </div>
    </div>
  );
};
