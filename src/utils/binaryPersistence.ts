import Dexie, { type Table } from 'dexie';
import JSZip from 'jszip';
import {
  ColumnDataType,
  ColumnRole,
  DataColumn,
  DockLayoutMode,
  GraphWindow,
  MDIWindowFrame,
  RemovedPointRecord,
  WorkbookSheet,
} from '../types/origin';
import {
  attachBinaryBuffersToColumn,
  getOrSyncColumnBuffers,
} from './binaryColumnStore';

/**
 * RULE 5: High-Capacity Local Persistence via IndexedDB (Dexie.js) & JSZip (.originclone)
 * NO localStorage is ever used. All numeric columns are saved as raw binary ArrayBuffers.
 */

export interface StoredColumnMeta {
  id: string;
  shortName: string;
  longName: string;
  unit: string;
  comment?: string;
  formula?: string;
  role: ColumnRole;
  dataType?: ColumnDataType;
  linkedXColId?: string;
  linkedErrColId?: string;
  linkedXErrColId?: string;
  rowCount: number;
}

export interface StoredSheetMeta {
  id: string;
  name: string;
  isFitOutput?: boolean;
  sourceFitId?: string;
  rowCount: number;
  columns: StoredColumnMeta[];
}

export interface ProjectManifestRecord {
  id: string; // 'active_workspace'
  savedAt: string;
  activeSheetId: string;
  activeGraphId: string;
  dockLayout: DockLayoutMode;
  splitRatio: number;
  sheetsMeta: StoredSheetMeta[];
  graphs: GraphWindow[];
  mdiWindows: MDIWindowFrame[];
  removedPointsLog: RemovedPointRecord[];
}

export interface ColumnBinaryRecord {
  colId: string;
  sheetId: string;
  rowCount: number;
  float64Buffer: ArrayBuffer;
  maskUint8Buffer: ArrayBuffer;
  rawStrings?: (string | null)[];
}

class OriginStudioIndexedDB extends Dexie {
  manifests!: Table<ProjectManifestRecord, string>;
  columnBuffers!: Table<ColumnBinaryRecord, string>;

  constructor() {
    super('OriginStudioHighPerfDB');
    this.version(1).stores({
      manifests: 'id, savedAt',
      columnBuffers: 'colId, sheetId',
    });
  }
}

export const originDB = new OriginStudioIndexedDB();

/**
 * Saves the entire workspace into IndexedDB as lightweight metadata + raw binary ArrayBuffers
 */
export async function saveWorkspaceToIndexedDB(state: {
  activeSheetId: string;
  activeGraphId: string;
  dockLayout: DockLayoutMode;
  splitRatio: number;
  sheets: WorkbookSheet[];
  graphs: GraphWindow[];
  mdiWindows: MDIWindowFrame[];
  removedPointsLog: RemovedPointRecord[];
}): Promise<void> {
  try {
    const sheetsMeta: StoredSheetMeta[] = [];
    const binaryRecords: ColumnBinaryRecord[] = [];

    for (const sheet of state.sheets) {
      let maxRowCount = 0;
      const colMetas: StoredColumnMeta[] = [];

      for (const col of sheet.columns) {
        const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
        if (length > maxRowCount) maxRowCount = length;

        colMetas.push({
          id: col.id,
          shortName: col.shortName,
          longName: col.longName,
          unit: col.unit,
          comment: col.comment,
          formula: col.formula,
          role: col.role,
          dataType: col.dataType,
          linkedXColId: col.linkedXColId,
          linkedErrColId: col.linkedErrColId,
          linkedXErrColId: col.linkedXErrColId,
          rowCount: length,
        });

        binaryRecords.push({
          colId: col.id,
          sheetId: sheet.id,
          rowCount: length,
          float64Buffer: float64.buffer.slice(0) as ArrayBuffer,
          maskUint8Buffer: maskUint8.buffer.slice(0) as ArrayBuffer,
          rawStrings: col.rawStrings,
        });
      }

      sheetsMeta.push({
        id: sheet.id,
        name: sheet.name,
        isFitOutput: sheet.isFitOutput,
        sourceFitId: sheet.sourceFitId,
        rowCount: maxRowCount,
        columns: colMetas,
      });
    }

    const manifest: ProjectManifestRecord = {
      id: 'active_workspace',
      savedAt: new Date().toISOString(),
      activeSheetId: state.activeSheetId,
      activeGraphId: state.activeGraphId,
      dockLayout: state.dockLayout,
      splitRatio: state.splitRatio,
      sheetsMeta,
      graphs: state.graphs,
      mdiWindows: state.mdiWindows,
      removedPointsLog: state.removedPointsLog,
    };

    await originDB.transaction('rw', originDB.manifests, originDB.columnBuffers, async () => {
      await originDB.manifests.put(manifest);
      await originDB.columnBuffers.clear();
      await originDB.columnBuffers.bulkPut(binaryRecords);
    });
  } catch {
    // Ignore IndexedDB quota/private mode errors silently
  }
}

/**
 * Loads the workspace from IndexedDB reconstructing Float64Array & Uint8Array columns in zero-copy speed
 */
export async function loadWorkspaceFromIndexedDB(): Promise<{
  activeSheetId: string;
  activeGraphId: string;
  dockLayout: DockLayoutMode;
  splitRatio: number;
  sheets: WorkbookSheet[];
  graphs: GraphWindow[];
  mdiWindows: MDIWindowFrame[];
  removedPointsLog: RemovedPointRecord[];
} | null> {
  try {
    const manifest = await originDB.manifests.get('active_workspace');
    if (!manifest || !manifest.sheetsMeta || manifest.sheetsMeta.length === 0) {
      return null;
    }

    const allBuffers = await originDB.columnBuffers.toArray();
    const bufferMap = new Map<string, ColumnBinaryRecord>();
    for (const rec of allBuffers) {
      bufferMap.set(rec.colId, rec);
    }

    const sheets: WorkbookSheet[] = manifest.sheetsMeta.map((sMeta) => {
      const cols: DataColumn[] = sMeta.columns.map((cMeta) => {
        const bin = bufferMap.get(cMeta.id);
        const f64 = bin ? new Float64Array(bin.float64Buffer) : new Float64Array(0);
        const u8 = bin ? new Uint8Array(bin.maskUint8Buffer) : new Uint8Array(0);

        return attachBinaryBuffersToColumn(
          {
            id: cMeta.id,
            shortName: cMeta.shortName,
            longName: cMeta.longName,
            unit: cMeta.unit,
            comment: cMeta.comment,
            formula: cMeta.formula,
            role: cMeta.role,
            dataType: cMeta.dataType,
            linkedXColId: cMeta.linkedXColId,
            linkedErrColId: cMeta.linkedErrColId,
            linkedXErrColId: cMeta.linkedXErrColId,
            rawStrings: bin?.rawStrings,
            values: [],
          },
          f64,
          u8
        );
      });

      return {
        id: sMeta.id,
        name: sMeta.name,
        isFitOutput: sMeta.isFitOutput,
        sourceFitId: sMeta.sourceFitId,
        columns: cols,
      };
    });

    return {
      activeSheetId: manifest.activeSheetId,
      activeGraphId: manifest.activeGraphId,
      dockLayout: manifest.dockLayout,
      splitRatio: manifest.splitRatio,
      sheets,
      graphs: manifest.graphs,
      mdiWindows: manifest.mdiWindows || [],
      removedPointsLog: manifest.removedPointsLog || [],
    };
  } catch {
    return null;
  }
}

/**
 * Exports the entire project as a high-speed binary `.originclone` archive using JSZip
 * (manifest.json + raw binary Float64Array and Uint8Array files)
 */
export async function exportOriginCloneBinaryZip(state: {
  activeSheetId: string;
  activeGraphId: string;
  dockLayout: DockLayoutMode;
  splitRatio: number;
  sheets: WorkbookSheet[];
  graphs: GraphWindow[];
  mdiWindows: MDIWindowFrame[];
  removedPointsLog: RemovedPointRecord[];
}): Promise<void> {
  const zip = new JSZip();
  const sheetsMeta: StoredSheetMeta[] = [];
  const rawStringsMap: Record<string, (string | null)[]> = {};

  const buffersFolder = zip.folder('buffers');

  for (const sheet of state.sheets) {
    let maxRowCount = 0;
    const colMetas: StoredColumnMeta[] = [];

    for (const col of sheet.columns) {
      const { float64, maskUint8, length } = getOrSyncColumnBuffers(col);
      if (length > maxRowCount) maxRowCount = length;

      colMetas.push({
        id: col.id,
        shortName: col.shortName,
        longName: col.longName,
        unit: col.unit,
        comment: col.comment,
        formula: col.formula,
        role: col.role,
        dataType: col.dataType,
        linkedXColId: col.linkedXColId,
        linkedErrColId: col.linkedErrColId,
        linkedXErrColId: col.linkedXErrColId,
        rowCount: length,
      });

      buffersFolder?.file(`${col.id}.f64.bin`, new Uint8Array(float64.buffer, float64.byteOffset, float64.byteLength));
      buffersFolder?.file(`${col.id}.u8.bin`, maskUint8);

      if (col.rawStrings && col.rawStrings.some((s) => s !== null)) {
        rawStringsMap[col.id] = col.rawStrings;
      }
    }

    sheetsMeta.push({
      id: sheet.id,
      name: sheet.name,
      isFitOutput: sheet.isFitOutput,
      sourceFitId: sheet.sourceFitId,
      rowCount: maxRowCount,
      columns: colMetas,
    });
  }

  const manifest = {
    format: 'OriginStudioLab_BinaryArchive_v3',
    savedAt: new Date().toISOString(),
    activeSheetId: state.activeSheetId,
    activeGraphId: state.activeGraphId,
    dockLayout: state.dockLayout,
    splitRatio: state.splitRatio,
    sheetsMeta,
    rawStringsMap,
    graphs: state.graphs,
    mdiWindows: state.mdiWindows,
    removedPointsLog: state.removedPointsLog,
  };

  zip.file('manifest.json', JSON.stringify(manifest, null, 2));

  const blob = await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'Progetto_OriginStudio.originclone';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Loads a `.originclone` binary ZIP archive and reconstructs all Float64Array & Uint8Array buffers
 */
export async function importOriginCloneBinaryZip(file: File): Promise<{
  activeSheetId: string;
  activeGraphId: string;
  dockLayout: DockLayoutMode;
  splitRatio: number;
  sheets: WorkbookSheet[];
  graphs: GraphWindow[];
  mdiWindows: MDIWindowFrame[];
  removedPointsLog: RemovedPointRecord[];
}> {
  const zip = await JSZip.loadAsync(file);
  const manifestFile = zip.file('manifest.json');
  if (!manifestFile) {
    throw new Error('File manifest.json mancante nell\'archivio .originclone');
  }

  const manifestText = await manifestFile.async('string');
  const manifest = JSON.parse(manifestText);
  const sheetsMeta: StoredSheetMeta[] = manifest.sheetsMeta || [];
  const rawStringsMap: Record<string, (string | null)[]> = manifest.rawStringsMap || {};

  const sheets: WorkbookSheet[] = [];

  for (const sMeta of sheetsMeta) {
    const cols: DataColumn[] = [];
    for (const cMeta of sMeta.columns) {
      const f64File = zip.file(`buffers/${cMeta.id}.f64.bin`);
      const u8File = zip.file(`buffers/${cMeta.id}.u8.bin`);

      const f64Buf = f64File ? await f64File.async('arraybuffer') : new ArrayBuffer(0);
      const u8Buf = u8File ? await u8File.async('arraybuffer') : new ArrayBuffer(0);

      const f64 = new Float64Array(f64Buf);
      const u8 = new Uint8Array(u8Buf);

      cols.push(
        attachBinaryBuffersToColumn(
          {
            id: cMeta.id,
            shortName: cMeta.shortName,
            longName: cMeta.longName,
            unit: cMeta.unit,
            comment: cMeta.comment,
            formula: cMeta.formula,
            role: cMeta.role,
            dataType: cMeta.dataType,
            linkedXColId: cMeta.linkedXColId,
            linkedErrColId: cMeta.linkedErrColId,
            linkedXErrColId: cMeta.linkedXErrColId,
            rawStrings: rawStringsMap[cMeta.id],
            values: [],
          },
          f64,
          u8
        )
      );
    }

    sheets.push({
      id: sMeta.id,
      name: sMeta.name,
      isFitOutput: sMeta.isFitOutput,
      sourceFitId: sMeta.sourceFitId,
      columns: cols,
    });
  }

  return {
    activeSheetId: manifest.activeSheetId,
    activeGraphId: manifest.activeGraphId,
    dockLayout: manifest.dockLayout || 'split-horizontal',
    splitRatio: typeof manifest.splitRatio === 'number' ? manifest.splitRatio : 48,
    sheets,
    graphs: manifest.graphs || [],
    mdiWindows: manifest.mdiWindows || [],
    removedPointsLog: manifest.removedPointsLog || [],
  };
}
