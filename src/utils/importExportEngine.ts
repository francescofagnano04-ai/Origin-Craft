import * as XLSX from 'xlsx';
import {
  DataColumn,
  DockLayoutMode,
  GraphWindow,
  MDIWindowFrame,
  RemovedPointRecord,
  WorkbookSheet,
} from '../types/origin';
import {
  createEmptySheet,
  detectColumnDataType,
  getColumnShortName,
  parseDateTimeTokenToNumeric,
  rebindSheetMultiXAssociations,
} from './mathEngine';

export type DelimiterOption = 'auto' | 'tab' | 'semicolon' | 'comma' | 'space' | 'custom';
export type DecimalSeparatorOption = 'auto' | 'dot' | 'comma';
export type MultiFileImportMode = 'separate-sheets' | 'append-rows' | 'append-columns';

export interface AdvancedParseOptions {
  delimiter: DelimiterOption;
  customDelimiter?: string;
  decimalSeparator: DecimalSeparatorOption;
  skipLines: number;
  skipCommentLines: boolean;
  extractHeaderLongName: boolean;
  extractHeaderUnits: boolean;
  extractHeaderComments: boolean;
}

export const DEFAULT_PARSE_OPTIONS: AdvancedParseOptions = {
  delimiter: 'auto',
  customDelimiter: '|',
  decimalSeparator: 'auto',
  skipLines: 0,
  skipCommentLines: true,
  extractHeaderLongName: true,
  extractHeaderUnits: true,
  extractHeaderComments: false,
};

export interface StagedImportFile {
  id: string;
  fileName: string;
  kind: 'text' | 'excel';
  rawText?: string;
  excelWorkbook?: XLSX.WorkBook;
  excelSheetNames?: string[];
  selectedExcelSheet?: string;
  importAllExcelSheets?: boolean;
}

export interface ParsedPreviewResult {
  sheet: WorkbookSheet;
  skippedHeaderPreview: string[];
  detectedDelimiterLabel: string;
  detectedDecimalLabel: string;
  totalRawLines: number;
}

/**
 * Extracts embedded unit from a header string such as "Tensione (V)", "Lunghezza [nm]", or "Corrente / mA"
 */
export function extractNameAndUnitFromHeader(rawHeader: string): {
  longName: string;
  unit: string;
} {
  const trimmed = rawHeader.trim();
  if (!trimmed) return { longName: '', unit: '' };

  // Pattern 1: "Name (unit)" or "Name [unit]"
  const bracketMatch = trimmed.match(/^(.+?)\s*[\(\[]([^)\]]+)[\)\]]$/);
  if (bracketMatch) {
    return {
      longName: bracketMatch[1].trim(),
      unit: bracketMatch[2].trim(),
    };
  }

  // Pattern 2: "Name / unit"
  const slashMatch = trimmed.match(/^(.+?)\s+\/\s+([a-zA-Z0-9°µΩ%\-\/\^]+)$/);
  if (slashMatch) {
    return {
      longName: slashMatch[1].trim(),
      unit: slashMatch[2].trim(),
    };
  }

  return { longName: trimmed, unit: '' };
}

/**
 * Splits a CSV/TSV line respecting quoted fields
 */
function splitLineWithQuotes(line: string, sep: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      inQuotes = !inQuotes;
    } else if (!inQuotes && line.startsWith(sep, i)) {
      result.push(current.trim());
      current = '';
      i += sep.length - 1;
    } else {
      current += ch;
    }
  }
  result.push(current.trim());
  return result.map((s) => s.replace(/^"|"$/g, '').trim());
}

/**
 * Advanced CSV / TSV / DAT / Clipboard Parser with configurable Delimiter, Decimal Separator,
 * Skip N Header Lines, and Automatic Header Metadata Extraction (Long Name, Units, Comments)
 */
export function parseTabularTextAdvanced(
  rawText: string,
  sheetName = 'Dati_Importati',
  options: AdvancedParseOptions = DEFAULT_PARSE_OPTIONS
): ParsedPreviewResult {
  const allLines = rawText
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n');

  const skipN = Math.max(0, Math.min(allLines.length, Math.floor(options.skipLines || 0)));
  const skippedHeaderPreview = allLines.slice(0, skipN);
  const linesAfterSkip = allLines.slice(skipN);

  const activeLines = linesAfterSkip
    .map((l) => l.trim())
    .filter((l) => {
      if (l.length === 0) return false;
      if (
        options.skipCommentLines &&
        (l.startsWith('#') || l.startsWith('//') || l.startsWith('%') || l.startsWith('!'))
      ) {
        return false;
      }
      return true;
    });

  if (activeLines.length === 0) {
    return {
      sheet: createEmptySheet(sheetName, 15),
      skippedHeaderPreview,
      detectedDelimiterLabel: 'Nessun dato',
      detectedDecimalLabel: '.',
      totalRawLines: allLines.length,
    };
  }

  // 1. Determine Delimiter
  const sampleLines = activeLines.slice(0, Math.min(15, activeLines.length));
  let splitFn: (line: string) => string[];
  let detectedDelimiterLabel = 'Automatico';
  let effectiveDelimiterType: 'tab' | 'semicolon' | 'comma' | 'space' | 'custom' = 'tab';

  if (options.delimiter === 'tab') {
    splitFn = (l) => splitLineWithQuotes(l, '\t');
    detectedDelimiterLabel = 'Tabulazione (\\t)';
    effectiveDelimiterType = 'tab';
  } else if (options.delimiter === 'semicolon') {
    splitFn = (l) => splitLineWithQuotes(l, ';');
    detectedDelimiterLabel = 'Punto e virgola (;)';
    effectiveDelimiterType = 'semicolon';
  } else if (options.delimiter === 'comma') {
    splitFn = (l) => splitLineWithQuotes(l, ',');
    detectedDelimiterLabel = 'Virgola (,)';
    effectiveDelimiterType = 'comma';
  } else if (options.delimiter === 'space') {
    splitFn = (l) => l.split(/\s+/).map((s) => s.trim());
    detectedDelimiterLabel = 'Spazi multipli';
    effectiveDelimiterType = 'space';
  } else if (options.delimiter === 'custom' && options.customDelimiter) {
    const customSep = options.customDelimiter;
    splitFn = (l) => splitLineWithQuotes(l, customSep);
    detectedDelimiterLabel = `Personalizzato (${customSep})`;
    effectiveDelimiterType = 'custom';
  } else {
    // Auto-detect delimiter from sample lines (including PDF / Webpage / Excel copy-paste)
    const tabCount = sampleLines.filter((l) => l.includes('\t')).length;
    const semiCount = sampleLines.filter((l) => l.includes(';')).length;
    const commaCount = sampleLines.filter((l) => l.includes(',')).length;

    if (tabCount >= sampleLines.length * 0.4) {
      splitFn = (l) => splitLineWithQuotes(l, '\t');
      detectedDelimiterLabel = 'Auto: Tabulazione (TSV / Excel)';
      effectiveDelimiterType = 'tab';
    } else if (semiCount >= sampleLines.length * 0.4) {
      splitFn = (l) => splitLineWithQuotes(l, ';');
      detectedDelimiterLabel = 'Auto: Punto e virgola (CSV Europeo)';
      effectiveDelimiterType = 'semicolon';
    } else if (commaCount >= sampleLines.length * 0.4) {
      // Check if comma is actually a decimal separator inside space-separated columns (e.g. "1,25  3,40")
      const looksLikeSpaceWithCommaDec = sampleLines.some((l) =>
        /[+-]?\d+,\d+\s+[+-]?\d+,\d+/.test(l)
      );
      if (looksLikeSpaceWithCommaDec) {
        splitFn = (l) => l.split(/\s+/).map((s) => s.trim());
        detectedDelimiterLabel = 'Auto: Spazio (con virgola decimale)';
        effectiveDelimiterType = 'space';
      } else {
        splitFn = (l) => splitLineWithQuotes(l, ',');
        detectedDelimiterLabel = 'Auto: Virgola (CSV Standard)';
        effectiveDelimiterType = 'comma';
      }
    } else {
      splitFn = (l) => l.split(/\s+/).map((s) => s.trim());
      detectedDelimiterLabel = 'Auto: Spazio / Tabella Testo';
      effectiveDelimiterType = 'space';
    }
  }

  // 2. Determine Decimal Separator
  let detectedDecimalLabel = 'Punto (.)';
  const parseNumericCell = (rawToken: string): number | null => {
    if (!rawToken) return null;
    const t = rawToken.trim().replace(/^"|"$/g, '');
    if (!t || t === '--' || t === '-' || t === 'NaN' || t === 'null' || t === 'N/A') return null;

    let normalized = t;
    if (options.decimalSeparator === 'comma') {
      // Remove thousand dots if present (e.g., 1.234,56 -> 1234.56)
      if (/^[+-]?\d{1,3}(\.\d{3})+(,\d+)$/.test(normalized)) {
        normalized = normalized.replace(/\./g, '');
      }
      normalized = normalized.replace(',', '.');
    } else if (options.decimalSeparator === 'dot') {
      // Remove thousand commas if present (e.g., 1,234.56 -> 1234.56)
      if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)$/.test(normalized)) {
        normalized = normalized.replace(/,/g, '');
      }
    } else {
      // Auto decimal separator
      if (effectiveDelimiterType !== 'comma') {
        if (/^[+-]?\d{1,3}(\.\d{3})+(,\d+)$/.test(normalized)) {
          normalized = normalized.replace(/\./g, '').replace(',', '.');
        } else if (normalized.includes(',') && !normalized.includes('.')) {
          normalized = normalized.replace(',', '.');
        }
      }
    }

    const num = Number(normalized);
    return Number.isFinite(num) ? num : null;
  };

  if (options.decimalSeparator === 'comma') {
    detectedDecimalLabel = 'Virgola (,)';
  } else if (options.decimalSeparator === 'dot') {
    detectedDecimalLabel = 'Punto (.)';
  } else {
    const hasCommaDecimals =
      effectiveDelimiterType !== 'comma' &&
      sampleLines.some((l) => /[+-]?\d+,\d+/.test(l));
    detectedDecimalLabel = hasCommaDecimals ? 'Auto: Virgola (,)' : 'Auto: Punto (.)';
  }

  const tokenizedRows = activeLines.map(splitFn);
  const maxCols = Math.max(1, ...tokenizedRows.map((r) => r.length));

  // 3. Detect / Extract Header Metadata (Long Name, Units, Comments)
  let currentRowCursor = 0;
  let headerLongNames: string[] = [];
  let headerUnits: string[] = [];
  let headerComments: string[] = [];

  const isRowMostlyText = (row: string[] | undefined): boolean => {
    if (!row || row.length === 0) return false;
    const nonEmpty = row.filter((cell) => cell.trim() !== '');
    if (nonEmpty.length === 0) return false;
    const numericCount = nonEmpty.filter((cell) => parseNumericCell(cell) !== null).length;
    return numericCount < nonEmpty.length / 2;
  };

  if (options.extractHeaderLongName && currentRowCursor < tokenizedRows.length) {
    const candidateRow = tokenizedRows[currentRowCursor];
    if (isRowMostlyText(candidateRow)) {
      headerLongNames = candidateRow;
      currentRowCursor++;
    }
  }

  if (options.extractHeaderUnits && currentRowCursor < tokenizedRows.length) {
    const candidateUnitRow = tokenizedRows[currentRowCursor];
    if (isRowMostlyText(candidateUnitRow)) {
      headerUnits = candidateUnitRow.map((u) => u.replace(/^[\[\(]|[\]\)]$/g, '').trim());
      currentRowCursor++;
    }
  }

  if (options.extractHeaderComments && currentRowCursor < tokenizedRows.length) {
    const candidateCommentRow = tokenizedRows[currentRowCursor];
    if (isRowMostlyText(candidateCommentRow)) {
      headerComments = candidateCommentRow;
      currentRowCursor++;
    }
  }

  const dataRows = tokenizedRows.slice(currentRowCursor);

  const columns: DataColumn[] = [];
  for (let colIdx = 0; colIdx < maxCols; colIdx++) {
    const shortName = getColumnShortName(colIdx);
    const rawHeaderCell = headerLongNames[colIdx] || '';
    const extracted = extractNameAndUnitFromHeader(rawHeaderCell);

    const longName = extracted.longName || `Colonna ${shortName}`;
    const unit = headerUnits[colIdx] || extracted.unit || '';
    const comment = headerComments[colIdx] || '';

    const rawTokens = dataRows.map((row) => (colIdx < row.length ? row[colIdx] : ''));
    const dataType = detectColumnDataType(rawTokens);

    const values: (number | null)[] = rawTokens.map((tok, rIdx) => {
      if (dataType === 'datetime') {
        const dtSec = parseDateTimeTokenToNumeric(tok || '');
        if (dtSec !== null) return dtSec;
      }
      const num = parseNumericCell(tok || '');
      if (num !== null) return dataType === 'integer' ? Math.round(num) : num;
      return dataType === 'text' ? rIdx + 1 : null;
    });

    const rawStrings: (string | null)[] | undefined =
      dataType === 'text' || dataType === 'datetime'
        ? rawTokens.map((t) => (t && t.trim() !== '' ? t.trim() : null))
        : undefined;

    let role: DataColumn['role'] = colIdx === 0 ? 'X' : 'Y';
    if (dataType === 'text' && colIdx > 0) {
      role = 'LABEL';
    } else if (/xerr|dx|delta\s*x|err\s*x/i.test(longName)) {
      role = 'X_ERR';
    } else if (
      /err|error|sigma|dev|incert/i.test(longName) ||
      /err|error|sigma/i.test(unit)
    ) {
      role = 'Y_ERR';
    } else if (/^x\d*$|^asse\s*x/i.test(longName)) {
      role = 'X';
    } else if (/label|etichetta|campione|nome/i.test(longName)) {
      role = 'LABEL';
    }

    columns.push({
      id: `col_${Date.now()}_${colIdx}_${Math.random().toString(36).slice(2, 6)}`,
      shortName,
      longName,
      unit,
      comment,
      role,
      dataType,
      values,
      rawStrings,
    });
  }

  const reboundSheet = rebindSheetMultiXAssociations({
    id: `sheet_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    name: sheetName,
    columns,
  });

  return {
    sheet: reboundSheet,
    skippedHeaderPreview,
    detectedDelimiterLabel,
    detectedDecimalLabel,
    totalRawLines: allLines.length,
  };
}

/**
 * Parses an Excel (.xlsx / .xls) worksheet into an OriginStudio Lab WorkbookSheet
 */
export function parseExcelSheetToWorkbook(
  workbook: XLSX.WorkBook,
  excelSheetName: string,
  options: AdvancedParseOptions = DEFAULT_PARSE_OPTIONS,
  customOutputName?: string
): ParsedPreviewResult {
  const worksheet = workbook.Sheets[excelSheetName];
  if (!worksheet) {
    return {
      sheet: createEmptySheet(customOutputName || excelSheetName, 15),
      skippedHeaderPreview: [],
      detectedDelimiterLabel: 'Foglio Excel (.xlsx)',
      detectedDecimalLabel: 'Nativo Excel',
      totalRawLines: 0,
    };
  }

  // Convert worksheet to 2D array of raw cell values
  const rows2D = XLSX.utils.sheet_to_json<(string | number | null)[]>(worksheet, {
    header: 1,
    defval: null,
    raw: true,
  });

  const skipN = Math.max(0, Math.min(rows2D.length, Math.floor(options.skipLines || 0)));
  const skippedHeaderPreview = rows2D
    .slice(0, skipN)
    .map((r) => (r || []).map((c) => (c === null ? '' : String(c))).join('\t'));

  const activeRows = rows2D
    .slice(skipN)
    .filter((r) => Array.isArray(r) && r.some((cell) => cell !== null && String(cell).trim() !== ''));

  if (activeRows.length === 0) {
    return {
      sheet: createEmptySheet(customOutputName || excelSheetName, 15),
      skippedHeaderPreview,
      detectedDelimiterLabel: 'Foglio Excel (.xlsx)',
      detectedDecimalLabel: 'Nativo Excel',
      totalRawLines: rows2D.length,
    };
  }

  const maxCols = Math.max(1, ...activeRows.map((r) => r.length));

  const parseCellToNum = (cell: string | number | null | undefined): number | null => {
    if (cell === null || cell === undefined) return null;
    if (typeof cell === 'number') return Number.isFinite(cell) ? cell : null;
    const s = String(cell).trim();
    if (!s || s === '--' || s === 'NaN') return null;
    const cleaned =
      options.decimalSeparator === 'comma' ? s.replace(',', '.') : s.replace(',', '.');
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : null;
  };

  const isRowMostlyStrings = (row: (string | number | null)[] | undefined): boolean => {
    if (!row) return false;
    const nonNull = row.filter((c) => c !== null && String(c).trim() !== '');
    if (nonNull.length === 0) return false;
    const numCount = nonNull.filter((c) => parseCellToNum(c) !== null).length;
    return numCount < nonNull.length / 2;
  };

  let cursor = 0;
  let headerLongNames: string[] = [];
  let headerUnits: string[] = [];
  let headerComments: string[] = [];

  if (options.extractHeaderLongName && cursor < activeRows.length) {
    if (isRowMostlyStrings(activeRows[cursor])) {
      headerLongNames = activeRows[cursor].map((c) => (c === null ? '' : String(c).trim()));
      cursor++;
    }
  }

  if (options.extractHeaderUnits && cursor < activeRows.length) {
    if (isRowMostlyStrings(activeRows[cursor])) {
      headerUnits = activeRows[cursor].map((c) =>
        c === null ? '' : String(c).replace(/^[\[\(]|[\]\)]$/g, '').trim()
      );
      cursor++;
    }
  }

  if (options.extractHeaderComments && cursor < activeRows.length) {
    if (isRowMostlyStrings(activeRows[cursor])) {
      headerComments = activeRows[cursor].map((c) => (c === null ? '' : String(c).trim()));
      cursor++;
    }
  }

  const dataRows = activeRows.slice(cursor);
  const columns: DataColumn[] = [];

  for (let colIdx = 0; colIdx < maxCols; colIdx++) {
    const shortName = getColumnShortName(colIdx);
    const rawHeader = headerLongNames[colIdx] || '';
    const extracted = extractNameAndUnitFromHeader(rawHeader);
    const longName = extracted.longName || `Colonna ${shortName}`;
    const unit = headerUnits[colIdx] || extracted.unit || '';
    const comment = headerComments[colIdx] || '';

    const rawCells = dataRows.map((row) => (colIdx < row.length ? row[colIdx] : null));
    const dataType = detectColumnDataType(rawCells);

    const values = rawCells.map((cell, rIdx) => {
      if (dataType === 'datetime' && cell !== null) {
        const dtSec = parseDateTimeTokenToNumeric(String(cell));
        if (dtSec !== null) return dtSec;
      }
      const num = parseCellToNum(cell);
      if (num !== null) return dataType === 'integer' ? Math.round(num) : num;
      return dataType === 'text' && cell !== null && String(cell).trim() !== '' ? rIdx + 1 : null;
    });

    const rawStrings: (string | null)[] | undefined =
      dataType === 'text' || dataType === 'datetime'
        ? rawCells.map((c) => (c !== null && String(c).trim() !== '' ? String(c).trim() : null))
        : undefined;

    let role: DataColumn['role'] = colIdx === 0 ? 'X' : 'Y';
    if (dataType === 'text' && colIdx > 0) {
      role = 'LABEL';
    } else if (/xerr|dx|delta\s*x|err\s*x/i.test(longName)) {
      role = 'X_ERR';
    } else if (/err|error|sigma|dev/i.test(longName) || /err|error|sigma/i.test(unit)) {
      role = 'Y_ERR';
    } else if (/^x\d*$|^asse\s*x/i.test(longName)) {
      role = 'X';
    } else if (/label|etichetta|campione|nome/i.test(longName)) {
      role = 'LABEL';
    }

    columns.push({
      id: `col_xlsx_${Date.now()}_${colIdx}_${Math.random().toString(36).slice(2, 6)}`,
      shortName,
      longName,
      unit,
      comment,
      role,
      dataType,
      values,
      rawStrings,
    });
  }

  const reboundSheet = rebindSheetMultiXAssociations({
    id: `sheet_xlsx_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    name: customOutputName || excelSheetName,
    columns,
  });

  return {
    sheet: reboundSheet,
    skippedHeaderPreview,
    detectedDelimiterLabel: `Foglio Excel "${excelSheetName}"`,
    detectedDecimalLabel: 'Numerico nativo Excel',
    totalRawLines: rows2D.length,
  };
}

/**
 * Combines multiple parsed WorkbookSheets according to MultiFileImportMode:
 * - 'separate-sheets': returns all sheets individually
 * - 'append-rows': concatenates rows vertically into a single sheet
 * - 'append-columns': appends columns horizontally side-by-side into a single sheet
 */
export function combineImportedSheets(
  sheetsList: WorkbookSheet[],
  mode: MultiFileImportMode,
  combinedName = 'MultiFile_Combinato'
): WorkbookSheet[] {
  if (sheetsList.length <= 1 || mode === 'separate-sheets') {
    return sheetsList;
  }

  if (mode === 'append-rows') {
    const base = sheetsList[0];
    const maxColCount = Math.max(...sheetsList.map((s) => s.columns.length));
    const mergedColumns: DataColumn[] = [];

    for (let cIdx = 0; cIdx < maxColCount; cIdx++) {
      const baseCol = base.columns[cIdx];
      const combinedValues: (number | null)[] = [];

      for (const sh of sheetsList) {
        const shRowCount = Math.max(0, ...sh.columns.map((c) => c.values.length));
        const col = sh.columns[cIdx];
        for (let r = 0; r < shRowCount; r++) {
          combinedValues.push(col && r < col.values.length ? col.values[r] : null);
        }
      }

      mergedColumns.push({
        id: `col_merged_row_${Date.now()}_${cIdx}`,
        shortName: getColumnShortName(cIdx),
        longName: baseCol?.longName || `Colonna ${getColumnShortName(cIdx)}`,
        unit: baseCol?.unit || '',
        comment: `Accodato da ${sheetsList.length} file`,
        role: baseCol?.role || (cIdx === 0 ? 'X' : 'Y'),
        values: combinedValues,
      });
    }

    const firstX = mergedColumns.find((c) => c.role === 'X');
    if (firstX) {
      mergedColumns.forEach((c) => {
        if (c.role === 'Y' || c.role === 'Y_ERR') c.linkedXColId = firstX.id;
      });
    }

    return [
      {
        id: `sheet_multi_rows_${Date.now()}`,
        name: combinedName,
        columns: mergedColumns,
      },
    ];
  }

  // 'append-columns': side-by-side columns with each file's X and Y properly linked
  const mergedColumns: DataColumn[] = [];
  let globalColIdx = 0;

  for (const sh of sheetsList) {
    let sheetXColId: string | undefined;
    for (const col of sh.columns) {
      const newColId = `col_merged_col_${Date.now()}_${globalColIdx}`;
      if (col.role === 'X' && !sheetXColId) {
        sheetXColId = newColId;
      }
      mergedColumns.push({
        ...col,
        id: newColId,
        shortName: getColumnShortName(globalColIdx),
        longName: `${col.longName} [${sh.name}]`,
        linkedXColId: col.role !== 'X' ? sheetXColId : undefined,
      });
      globalColIdx++;
    }
  }

  return [
    {
      id: `sheet_multi_cols_${Date.now()}`,
      name: combinedName,
      columns: mergedColumns,
    },
  ];
}

/**
 * Exports a WorkbookSheet to an Excel (.xlsx), CSV, TXT, or DAT file
 */
export function exportWorkbookSheetFile(
  sheet: WorkbookSheet,
  format: 'txt' | 'csv' | 'csv-semi' | 'dat' | 'xlsx',
  decimalChar: '.' | ',' = '.'
) {
  const maxRows = Math.max(0, ...sheet.columns.map((c) => c.values.length));

  if (format === 'xlsx') {
    const aoa: (string | number | null)[][] = [];
    // Row 1: Long Names
    aoa.push(sheet.columns.map((c) => c.longName || c.shortName));
    // Row 2: Units
    if (sheet.columns.some((c) => c.unit)) {
      aoa.push(sheet.columns.map((c) => c.unit || ''));
    }
    // Data Rows
    for (let r = 0; r < maxRows; r++) {
      const row = sheet.columns.map((c) => {
        const v = r < c.values.length ? c.values[r] : null;
        return v !== null && Number.isFinite(v) ? v : null;
      });
      if (row.some((cell) => cell !== null)) {
        aoa.push(row);
      }
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31));
    XLSX.writeFile(wb, `${sheet.name}.xlsx`);
    return;
  }

  const sep =
    format === 'txt' || format === 'dat' ? '\t' : format === 'csv-semi' ? ';' : ',';
  const ext = format === 'csv-semi' ? 'csv' : format;

  const headerLong = sheet.columns
    .map((c) => {
      const text = `${c.longName || c.shortName}`;
      return sep === ',' && text.includes(',') ? `"${text}"` : text;
    })
    .join(sep);

  const lines: string[] = [headerLong];

  if (sheet.columns.some((c) => c.unit.trim() !== '')) {
    const headerUnits = sheet.columns.map((c) => c.unit.trim()).join(sep);
    lines.push(headerUnits);
  }

  for (let r = 0; r < maxRows; r++) {
    const rowTokens = sheet.columns.map((c) => {
      const val = r < c.values.length ? c.values[r] : null;
      if (val === null || !Number.isFinite(val)) return '';
      const str = String(val);
      return decimalChar === ',' ? str.replace('.', ',') : str;
    });
    if (rowTokens.some((t) => t !== '')) {
      lines.push(rowTokens.join(sep));
    }
  }

  const blob = new Blob([lines.join('\n')], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${sheet.name}.${ext}`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Full OriginStudio Project Bundle (.opj.json) Serialization & Deserialization
 */
export interface OriginProjectBundle {
  format: 'OriginStudioLab_Project_v2';
  savedAt: string;
  activeSheetId: string;
  activeGraphId: string;
  dockLayout: DockLayoutMode;
  splitRatio: number;
  sheets: WorkbookSheet[];
  graphs: GraphWindow[];
  mdiWindows: MDIWindowFrame[];
  removedPointsLog: RemovedPointRecord[];
}
