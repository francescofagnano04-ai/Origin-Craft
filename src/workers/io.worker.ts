/// <reference lib="webworker" />
import Papa from 'papaparse';

export interface IOWorkerParseRequest {
  id: string;
  type: 'PARSE_CSV_BINARY';
  rawText: string;
  sheetName: string;
  options: {
    delimiter: 'auto' | '\t' | ';' | ',' | ' ' | 'custom';
    customDelimiter?: string;
    decimalSeparator: 'auto' | '.' | ',';
    skipLines: number;
    skipCommentLines: boolean;
    extractHeaderLongName: boolean;
    extractHeaderUnits: boolean;
  };
}

export interface IOWorkerColumnBinaryPayload {
  shortName: string;
  longName: string;
  unit: string;
  comment: string;
  role: 'X' | 'Y' | 'X_ERR' | 'Y_ERR' | 'LABEL' | 'IGNORE';
  dataType: 'numeric' | 'integer' | 'text' | 'datetime';
  rowCount: number;
  float64Buffer: ArrayBuffer;
  maskUint8Buffer: ArrayBuffer;
  rawStrings?: (string | null)[];
}

export interface IOWorkerParseResponse {
  id: string;
  ok: boolean;
  sheetName: string;
  rowCount: number;
  columns: IOWorkerColumnBinaryPayload[];
  error?: string;
}

function getColumnShortName(index: number): string {
  let name = '';
  let n = index;
  while (n >= 0) {
    name = String.fromCharCode((n % 26) + 65) + name;
    n = Math.floor(n / 26) - 1;
  }
  return name;
}

self.onmessage = (event: MessageEvent<IOWorkerParseRequest>) => {
  const msg = event.data;
  if (!msg || msg.type !== 'PARSE_CSV_BINARY') return;

  try {
    const { rawText, sheetName, options } = msg;
    const allLines = rawText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    const afterSkip = allLines.slice(Math.max(0, options.skipLines || 0));

    const commentComments: string[] = [];
    const activeLines: string[] = [];
    for (const line of afterSkip) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      if (
        options.skipCommentLines &&
        (trimmed.startsWith('#') || trimmed.startsWith('//') || trimmed.startsWith('!'))
      ) {
        commentComments.push(trimmed.replace(/^[#/!]+\s*/, ''));
        continue;
      }
      activeLines.push(line);
    }

    const cleanBody = activeLines.join('\n');
    const papaDelimiter =
      options.delimiter === 'auto'
        ? ''
        : options.delimiter === 'custom'
        ? options.customDelimiter || '\t'
        : options.delimiter;

    const parsed = Papa.parse<string[]>(cleanBody, {
      delimiter: papaDelimiter,
      skipEmptyLines: true,
    });

    const rows = parsed.data;
    if (!rows || rows.length === 0) {
      self.postMessage({
        id: msg.id,
        ok: true,
        sheetName,
        rowCount: 0,
        columns: [],
      } satisfies IOWorkerParseResponse);
      return;
    }

    const colCount = Math.max(1, ...rows.slice(0, 50).map((r) => r.length));

    const parseNumericToken = (tok: string): number => {
      if (!tok) return Number.NaN;
      const t = tok.trim();
      if (!t || t === '--' || t.toLowerCase() === 'nan') return Number.NaN;
      const normalized =
        options.decimalSeparator === ','
          ? t.replace(/\./g, '').replace(',', '.')
          : options.decimalSeparator === 'auto'
          ? t.replace(',', '.')
          : t;
      const v = Number(normalized);
      return Number.isFinite(v) ? v : Number.NaN;
    };

    // Detect header rows
    let dataStartRow = 0;
    const longNames = Array.from({ length: colCount }, (_, i) =>
      i === 0 ? 'Asse X' : `Serie ${getColumnShortName(i)}`
    );
    const units = Array.from({ length: colCount }, () => '');

    const firstRow = rows[0] || [];
    const firstRowNumeric = firstRow.filter((cell) => Number.isFinite(parseNumericToken(cell))).length;

    if (options.extractHeaderLongName && firstRowNumeric < firstRow.length / 2) {
      dataStartRow = 1;
      for (let c = 0; c < colCount; c++) {
        const rawHdr = (firstRow[c] || '').trim();
        if (!rawHdr) continue;
        const m = rawHdr.match(/^(.+?)\s*[\(\[]([^)\]]+)[\)\]]$/);
        if (m && options.extractHeaderUnits) {
          longNames[c] = m[1].trim();
          units[c] = m[2].trim();
        } else {
          longNames[c] = rawHdr;
        }
      }

      if (rows.length > 1 && options.extractHeaderUnits) {
        const secondRow = rows[1] || [];
        const secondRowNumeric = secondRow.filter((cell) =>
          Number.isFinite(parseNumericToken(cell))
        ).length;
        if (secondRowNumeric < secondRow.length / 2) {
          dataStartRow = 2;
          for (let c = 0; c < colCount; c++) {
            const u = (secondRow[c] || '').trim().replace(/^[\(\[]|[\)\]]$/g, '');
            if (u) units[c] = u;
          }
        }
      }
    }

    const rowCount = Math.max(0, rows.length - dataStartRow);
    const f64Columns: Float64Array[] = [];
    const maskColumns: Uint8Array[] = [];

    for (let c = 0; c < colCount; c++) {
      f64Columns.push(new Float64Array(rowCount));
      maskColumns.push(new Uint8Array(rowCount));
    }

    // Stream directly into pre-allocated Float64Array and Uint8Array
    for (let r = 0; r < rowCount; r++) {
      const row = rows[r + dataStartRow];
      for (let c = 0; c < colCount; c++) {
        const tok = c < row.length ? row[c] : '';
        const val = parseNumericToken(tok);
        if (Number.isFinite(val)) {
          f64Columns[c][r] = val;
          maskColumns[c][r] = 0; // POINT_VALID
        } else {
          f64Columns[c][r] = Number.NaN;
          maskColumns[c][r] = 2; // POINT_NAN
        }
      }
    }

    const columns: IOWorkerColumnBinaryPayload[] = [];
    const transferables: Transferable[] = [];

    for (let c = 0; c < colCount; c++) {
      const shortName = getColumnShortName(c);
      const f64Buf = f64Columns[c].buffer as ArrayBuffer;
      const u8Buf = maskColumns[c].buffer as ArrayBuffer;
      transferables.push(f64Buf, u8Buf);

      columns.push({
        shortName,
        longName: longNames[c],
        unit: units[c],
        comment: commentComments[0] || '',
        role: c === 0 ? 'X' : 'Y',
        dataType: 'numeric',
        rowCount,
        float64Buffer: f64Buf,
        maskUint8Buffer: u8Buf,
      });
    }

    const response: IOWorkerParseResponse = {
      id: msg.id,
      ok: true,
      sheetName,
      rowCount,
      columns,
    };

    self.postMessage(response, transferables);
  } catch (err) {
    self.postMessage({
      id: msg.id,
      ok: false,
      sheetName: msg.sheetName,
      rowCount: 0,
      columns: [],
      error: err instanceof Error ? err.message : String(err),
    } satisfies IOWorkerParseResponse);
  }
};
