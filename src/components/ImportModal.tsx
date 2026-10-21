import React, { useState, useRef, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload,
  FileText,
  ClipboardPaste,
  Sparkles,
  X,
  Check,
  FileSpreadsheet,
  Layers,
  Trash2,
  Sliders,
} from 'lucide-react';
import { WorkbookSheet } from '../types/origin';
import {
  AdvancedParseOptions,
  combineImportedSheets,
  DecimalSeparatorOption,
  DEFAULT_PARSE_OPTIONS,
  DelimiterOption,
  MultiFileImportMode,
  ParsedPreviewResult,
  parseExcelSheetToWorkbook,
  parseTabularTextAdvanced,
  StagedImportFile,
} from '../utils/importExportEngine';
import { parseCsvInWorker } from '../utils/workerClient';
import { ensureSheetBinaryBuffers } from '../utils/binaryColumnStore';

interface ImportModalProps {
  isOpen: boolean;
  onClose: () => void;
  onImportSheets: (sheets: WorkbookSheet[], plotImmediately: boolean) => void;
  initialDroppedFiles?: File[];
  onClearDroppedFiles?: () => void;
  initialClipboardMode?: boolean;
}

const SAMPLE_INSTRUMENT_DATA = `# Tektronix MDO3024 Oscilloscope Acquisition Log
# Date: 2026-10-20 | Sample Rate: 2.5 MS/s | Ch1 & Ch2
# Instrument Calibration Firmware v4.2
Tempo (ms)\tCanale 1 (V)\tCanale 2 Filtrato (V)\tIncertezza (V)
ms\tV\tV\tV
0,00\t0,05\t0,02\t0,04
0,50\t1,18\t1,12\t0,05
1,00\t2,35\t2,24\t0,06
1,50\t3,48\t3,36\t0,07
2,00\t8,90\t4,48\t0,08
2,50\t5,72\t5,60\t0,08
3,00\t6,89\t6,72\t0,09
3,50\t7,95\t7,84\t0,10
4,00\t9,12\t8,96\t0,11
4,50\t10,25\t10,08\t0,11
5,00\t11,40\t11,20\t0,12`;

export const ImportModal: React.FC<ImportModalProps> = ({
  isOpen,
  onClose,
  onImportSheets,
  initialDroppedFiles,
  onClearDroppedFiles,
  initialClipboardMode,
}) => {
  const [stagedFiles, setStagedFiles] = useState<StagedImportFile[]>([]);
  const [selectedFileIndex, setSelectedFileIndex] = useState<number>(0);
  const [rawClipboardText, setRawClipboardText] = useState<string>(SAMPLE_INSTRUMENT_DATA);
  const [sourceMode, setSourceMode] = useState<'files' | 'clipboard'>('files');

  const [parseOptions, setParseOptions] = useState<AdvancedParseOptions>(DEFAULT_PARSE_OPTIONS);
  const [multiFileMode, setMultiFileMode] = useState<MultiFileImportMode>('separate-sheets');
  const [customSheetName, setCustomSheetName] = useState<string>('Dati_Sperimentali');
  const [plotImmediately, setPlotImmediately] = useState<boolean>(true);
  const [isDragOver, setIsDragOver] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Process incoming dropped files from global window drop
  useEffect(() => {
    if (isOpen && initialDroppedFiles && initialDroppedFiles.length > 0) {
      setSourceMode('files');
      ingestFileList(initialDroppedFiles);
      onClearDroppedFiles?.();
    }
  }, [isOpen, initialDroppedFiles]);

  useEffect(() => {
    if (isOpen && initialClipboardMode) {
      setSourceMode('clipboard');
      if (navigator.clipboard && navigator.clipboard.readText) {
        navigator.clipboard
          .readText()
          .then((txt) => {
            if (txt && txt.trim()) {
              setRawClipboardText(txt);
              setCustomSheetName('Incollato_Appunti');
            }
          })
          .catch(() => {
            // User can paste manually in textarea
          });
      }
    }
  }, [isOpen, initialClipboardMode]);

  const ingestFileList = async (files: FileList | File[]) => {
    const fileArray = Array.from(files);
    if (fileArray.length === 0) return;

    const loaded: StagedImportFile[] = [];

    for (const file of fileArray) {
      const ext = file.name.split('.').pop()?.toLowerCase() || '';
      const cleanBaseName = file.name.replace(/\.[^.]+$/, '');

      if (ext === 'xlsx' || ext === 'xls') {
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(buffer, { type: 'array' });
        loaded.push({
          id: `file_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
          fileName: file.name,
          kind: 'excel',
          excelWorkbook: wb,
          excelSheetNames: wb.SheetNames,
          selectedExcelSheet: wb.SheetNames[0] || '',
          importAllExcelSheets: false,
        });
      } else {
        const text = await file.text();
        loaded.push({
          id: `file_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
          fileName: file.name,
          kind: 'text',
          rawText: text,
        });
      }

      if (loaded.length === 1 && stagedFiles.length === 0) {
        setCustomSheetName(cleanBaseName);
      }
    }

    setStagedFiles((prev) => {
      const next = [...prev, ...loaded];
      return next;
    });
    setSelectedFileIndex(0);
    setSourceMode('files');
  };

  const [isWorkerParsing, setIsWorkerParsing] = useState<boolean>(false);

  // Compute live preview on a fast 60-line slice so huge files never block the UI thread
  const activePreview: ParsedPreviewResult = useMemo(() => {
    const sliceTextForFastPreview = (txt: string) => {
      if (txt.length <= 24000) return txt;
      const lines = txt.slice(0, 32000).split(/\r?\n/);
      return lines.slice(0, 60).join('\n');
    };

    if (sourceMode === 'clipboard' || stagedFiles.length === 0) {
      return parseTabularTextAdvanced(
        sliceTextForFastPreview(rawClipboardText),
        customSheetName.trim() || 'Incollato_Appunti',
        parseOptions
      );
    }

    const activeFile = stagedFiles[selectedFileIndex] || stagedFiles[0];
    if (!activeFile) {
      return parseTabularTextAdvanced('', 'Vuoto', parseOptions);
    }

    const baseName = activeFile.fileName.replace(/\.[^.]+$/, '');
    if (activeFile.kind === 'excel' && activeFile.excelWorkbook) {
      const targetSheet =
        activeFile.selectedExcelSheet || activeFile.excelWorkbook.SheetNames[0] || '';
      return parseExcelSheetToWorkbook(
        activeFile.excelWorkbook,
        targetSheet,
        parseOptions,
        stagedFiles.length === 1
          ? customSheetName.trim() || `${baseName}_${targetSheet}`
          : `${baseName}_${targetSheet}`
      );
    }

    return parseTabularTextAdvanced(
      sliceTextForFastPreview(activeFile.rawText || ''),
      stagedFiles.length === 1 ? customSheetName.trim() || baseName : baseName,
      parseOptions
    );
  }, [sourceMode, stagedFiles, selectedFileIndex, rawClipboardText, customSheetName, parseOptions]);

  if (!isOpen) return null;

  const handleConfirmImport = async () => {
    setIsWorkerParsing(true);
    try {
      const mappedDelim: 'auto' | ',' | ';' | '\t' | ' ' | 'custom' =
        parseOptions.delimiter === 'tab'
          ? '\t'
          : parseOptions.delimiter === 'semicolon'
          ? ';'
          : parseOptions.delimiter === 'comma'
          ? ','
          : parseOptions.delimiter === 'space'
          ? ' '
          : parseOptions.delimiter === 'custom'
          ? 'custom'
          : 'auto';

      const mappedDec: 'auto' | '.' | ',' =
        parseOptions.decimalSeparator === 'dot'
          ? '.'
          : parseOptions.decimalSeparator === 'comma'
          ? ','
          : 'auto';

      const workerOpts = {
        delimiter: mappedDelim,
        customDelimiter: parseOptions.customDelimiter,
        decimalSeparator: mappedDec,
        skipLines: parseOptions.skipLines,
        skipCommentLines: parseOptions.skipCommentLines,
        extractHeaderLongName: parseOptions.extractHeaderLongName,
        extractHeaderUnits: parseOptions.extractHeaderUnits,
      };

      if (sourceMode === 'clipboard' || stagedFiles.length === 0) {
        const sheet = await parseCsvInWorker(
          rawClipboardText,
          customSheetName.trim() || 'Dati_Importati',
          workerOpts
        );
        onImportSheets([sheet], plotImmediately);
        onClose();
        return;
      }

      // Parse all staged files using io.worker.ts for text/CSV/TSV
      const parsedSheets: WorkbookSheet[] = [];
      for (const f of stagedFiles) {
        const baseName = f.fileName.replace(/\.[^.]+$/, '');
        if (f.kind === 'excel' && f.excelWorkbook) {
          if (f.importAllExcelSheets && f.excelSheetNames) {
            for (const shName of f.excelSheetNames) {
              const parsed = parseExcelSheetToWorkbook(
                f.excelWorkbook,
                shName,
                parseOptions,
                `${baseName}_${shName}`
              );
              parsedSheets.push(ensureSheetBinaryBuffers(parsed.sheet));
            }
          } else {
            const targetSh = f.selectedExcelSheet || f.excelWorkbook.SheetNames[0] || '';
            const outName =
              stagedFiles.length === 1
                ? customSheetName.trim() || `${baseName}_${targetSh}`
                : `${baseName}_${targetSh}`;
            const parsed = parseExcelSheetToWorkbook(
              f.excelWorkbook,
              targetSh,
              parseOptions,
              outName
            );
            parsedSheets.push(ensureSheetBinaryBuffers(parsed.sheet));
          }
        } else {
          const outName =
            stagedFiles.length === 1 ? customSheetName.trim() || baseName : baseName;
          const sheet = await parseCsvInWorker(f.rawText || '', outName, workerOpts);
          parsedSheets.push(sheet);
        }
      }

      const finalSheets = combineImportedSheets(
        parsedSheets,
        multiFileMode,
        customSheetName.trim() || 'MultiFile_Combinato'
      ).map(ensureSheetBinaryBuffers);

      onImportSheets(finalSheets, plotImmediately);
      setStagedFiles([]);
      onClose();
    } finally {
      setIsWorkerParsing(false);
    }
  };

  const handlePasteFromSystemClipboard = async () => {
    setSourceMode('clipboard');
    if (navigator.clipboard && navigator.clipboard.readText) {
      try {
        const txt = await navigator.clipboard.readText();
        if (txt && txt.trim()) {
          setRawClipboardText(txt);
          setCustomSheetName('Incollato_Appunti');
        }
      } catch {
        // Fallback: user presses Ctrl+V inside the textarea
      }
    }
  };

  const activeStagedFile = stagedFiles[selectedFileIndex] || stagedFiles[0];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/55 p-3">
      <div className="w-full max-w-5xl bg-white border border-slate-300 rounded shadow-2xl flex flex-col max-h-[92vh] overflow-hidden text-xs">
        {/* Modal Header */}
        <div className="flex items-center justify-between px-4 py-2.5 bg-[#dbe4ee] border-b border-slate-300 shrink-0">
          <div className="flex items-center gap-2">
            <FileSpreadsheet className="w-4 h-4 text-blue-800" />
            <h2 className="text-sm font-bold text-slate-900">
              Origin Import Wizard — CSV / TSV / TXT / DAT / Excel (.XLSX) & Smart Clipboard
            </h2>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1 text-slate-500 hover:text-slate-900 rounded"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Source Switcher Bar */}
        <div className="flex flex-wrap items-center justify-between gap-2 px-4 py-2 bg-[#f1f5f9] border-b border-slate-300 shrink-0">
          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={() => setSourceMode('files')}
              className={`flex items-center gap-1.5 px-3 py-1 rounded text-xs font-semibold border transition-colors ${
                sourceMode === 'files'
                  ? 'bg-blue-700 text-white border-blue-700 shadow-2xs'
                  : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
              }`}
            >
              <Upload className="w-3.5 h-3.5" />
              <span>1. Importa File (.CSV, .TXT, .DAT, .XLSX) {stagedFiles.length > 0 && `[${stagedFiles.length}]`}</span>
            </button>

            <button
              type="button"
              onClick={handlePasteFromSystemClipboard}
              className={`flex items-center gap-1.5 px-3 py-1 rounded text-xs font-semibold border transition-colors ${
                sourceMode === 'clipboard'
                  ? 'bg-blue-700 text-white border-blue-700 shadow-2xs'
                  : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
              }`}
            >
              <ClipboardPaste className="w-3.5 h-3.5" />
              <span>2. Copia/Incolla Smart da Appunti (Excel, PDF, Browser)</span>
            </button>
          </div>

          <button
            type="button"
            onClick={() => {
              setSourceMode('clipboard');
              setRawClipboardText(SAMPLE_INSTRUMENT_DATA);
              setParseOptions({ ...DEFAULT_PARSE_OPTIONS, skipLines: 0, skipCommentLines: true });
              setCustomSheetName('Oscilloscopio_Tek');
            }}
            className="flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium text-teal-900 bg-teal-50 border border-teal-300 rounded hover:bg-teal-100"
          >
            <Sparkles className="w-3.5 h-3.5 text-teal-700" />
            <span>Carica Esempio Strumento (con Header & Virgola)</span>
          </button>
        </div>

        {/* Main Wizard Body: 2 Columns on Desktop */}
        <div className="flex-1 grid grid-cols-1 lg:grid-cols-12 min-h-0 overflow-hidden">
          {/* Left Column (5 cols): File Dropzone / Multi-file List / Parser & Header Controls */}
          <div className="lg:col-span-5 border-r border-slate-300 p-3.5 overflow-y-auto space-y-3.5 bg-[#f8fafc]">
            {sourceMode === 'files' ? (
              <>
                {/* Drag & Drop Zone */}
                <div
                  onDragOver={(e) => {
                    e.preventDefault();
                    setIsDragOver(true);
                  }}
                  onDragLeave={() => setIsDragOver(false)}
                  onDrop={(e) => {
                    e.preventDefault();
                    setIsDragOver(false);
                    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                      ingestFileList(e.dataTransfer.files);
                    }
                  }}
                  onClick={() => fileInputRef.current?.click()}
                  className={`p-4 border-2 border-dashed rounded cursor-pointer text-center transition-colors ${
                    isDragOver
                      ? 'border-blue-600 bg-blue-50'
                      : 'border-slate-300 bg-white hover:border-blue-500 hover:bg-blue-50/30'
                  }`}
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    multiple
                    accept=".txt,.csv,.tsv,.dat,.asc,.xlsx,.xls"
                    onChange={(e) => {
                      if (e.target.files) ingestFileList(e.target.files);
                    }}
                    className="hidden"
                  />
                  <Upload className="w-6 h-6 text-blue-700 mx-auto mb-1" />
                  <div className="font-bold text-slate-800 text-xs">
                    Trascina qui uno o più file (.CSV, .TXT, .DAT, .XLSX)
                  </div>
                  <div className="text-[11px] text-slate-500 mt-0.5">
                    oppure clicca per sfogliare dal computer (supporta selezione multi-file)
                  </div>
                </div>

                {/* Staged Files List & Multi-file Mode */}
                {stagedFiles.length > 0 && (
                  <div className="p-2.5 bg-white border border-slate-300 rounded space-y-2">
                    <div className="flex items-center justify-between">
                      <span className="font-bold text-slate-800 text-[11px]">
                        File Selezionati ({stagedFiles.length})
                      </span>
                      <button
                        type="button"
                        onClick={() => setStagedFiles([])}
                        className="text-[10.5px] text-red-600 hover:underline flex items-center gap-0.5"
                      >
                        <Trash2 className="w-3 h-3" />
                        <span>Svuota lista</span>
                      </button>
                    </div>

                    <div className="max-h-28 overflow-y-auto space-y-1">
                      {stagedFiles.map((f, idx) => (
                        <div
                          key={f.id}
                          onClick={() => setSelectedFileIndex(idx)}
                          className={`flex items-center justify-between px-2 py-1 rounded cursor-pointer border text-[11px] ${
                            idx === selectedFileIndex
                              ? 'bg-blue-50 border-blue-500 font-semibold text-blue-950'
                              : 'bg-slate-50 border-slate-200 text-slate-700 hover:bg-slate-100'
                          }`}
                        >
                          <div className="flex items-center gap-1.5 truncate">
                            <FileText className="w-3.5 h-3.5 text-blue-700 shrink-0" />
                            <span className="truncate">{f.fileName}</span>
                            {f.kind === 'excel' && (
                              <span className="px-1 py-0 text-[9.5px] bg-emerald-100 text-emerald-800 rounded font-mono">
                                XLSX
                              </span>
                            )}
                          </div>
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              setStagedFiles((prev) => prev.filter((_, i) => i !== idx));
                              setSelectedFileIndex(0);
                            }}
                            className="text-slate-400 hover:text-red-600 ml-1"
                          >
                            ×
                          </button>
                        </div>
                      ))}
                    </div>

                    {/* Excel Sheet Selector when an .xlsx file is selected */}
                    {activeStagedFile?.kind === 'excel' && activeStagedFile.excelSheetNames && (
                      <div className="p-2 bg-emerald-50 border border-emerald-200 rounded space-y-1.5">
                        <div className="font-semibold text-emerald-950 text-[11px]">
                          Selezione Foglio Excel ({activeStagedFile.fileName})
                        </div>
                        <select
                          value={activeStagedFile.selectedExcelSheet || ''}
                          disabled={activeStagedFile.importAllExcelSheets}
                          onChange={(e) => {
                            const val = e.target.value;
                            setStagedFiles((prev) =>
                              prev.map((item) =>
                                item.id === activeStagedFile.id
                                  ? { ...item, selectedExcelSheet: val }
                                  : item
                              )
                            );
                          }}
                          className="w-full px-2 py-1 text-xs bg-white border border-emerald-300 rounded font-medium"
                        >
                          {activeStagedFile.excelSheetNames.map((shName) => (
                            <option key={shName} value={shName}>
                              Foglio: {shName}
                            </option>
                          ))}
                        </select>

                        {activeStagedFile.excelSheetNames.length > 1 && (
                          <label className="flex items-center gap-1.5 cursor-pointer text-[11px] text-emerald-900">
                            <input
                              type="checkbox"
                              checked={!!activeStagedFile.importAllExcelSheets}
                              onChange={(e) => {
                                const checked = e.target.checked;
                                setStagedFiles((prev) =>
                                  prev.map((item) =>
                                    item.id === activeStagedFile.id
                                      ? { ...item, importAllExcelSheets: checked }
                                      : item
                                  )
                                );
                              }}
                              className="rounded border-emerald-400 text-emerald-700"
                            />
                            <span>
                              Importa tutti i {activeStagedFile.excelSheetNames.length} fogli del file Excel
                            </span>
                          </label>
                        )}
                      </div>
                    )}

                    {/* Multi-File Merge Strategy */}
                    {stagedFiles.length > 1 && (
                      <div className="p-2 bg-blue-50/70 border border-blue-200 rounded space-y-1">
                        <label className="block font-semibold text-blue-950 text-[11px]">
                          Modalità Importazione Multi-File ({stagedFiles.length} file):
                        </label>
                        <select
                          value={multiFileMode}
                          onChange={(e) =>
                            setMultiFileMode(e.target.value as MultiFileImportMode)
                          }
                          className="w-full px-2 py-1 text-xs bg-white border border-blue-300 rounded font-medium text-slate-900"
                        >
                          <option value="separate-sheets">
                            Fogli Separati (un foglio Workbook per ogni file)
                          </option>
                          <option value="append-rows">
                            Accoda Righe nello Stesso Foglio (unione verticale)
                          </option>
                          <option value="append-columns">
                            Affianca Colonne nello Stesso Foglio (confronto multi-curva)
                          </option>
                        </select>
                      </div>
                    )}
                  </div>
                )}
              </>
            ) : (
              /* Smart Clipboard Textarea */
              <div className="space-y-1.5">
                <div className="flex items-center justify-between">
                  <label className="font-bold text-slate-800 text-[11px]">
                    Incolla qui celle da Excel, tabelle PDF o browser (Ctrl+V):
                  </label>
                  <button
                    type="button"
                    onClick={handlePasteFromSystemClipboard}
                    className="px-2 py-0.5 text-[10.5px] font-semibold bg-blue-100 text-blue-900 rounded hover:bg-blue-200"
                  >
                    Leggi Appunti
                  </button>
                </div>
                <textarea
                  rows={6}
                  value={rawClipboardText}
                  onChange={(e) => setRawClipboardText(e.target.value)}
                  placeholder="Incolla qui i dati copiati da Excel, PDF o un file .txt..."
                  className="w-full p-2 font-mono text-[11px] bg-white border border-slate-300 rounded focus:outline-none focus:border-blue-600"
                />
              </div>
            )}

            {/* Advanced CSV/TSV Parser & Decimal Settings */}
            <div className="p-3 bg-white border border-slate-300 rounded space-y-2.5">
              <div className="flex items-center justify-between border-b border-slate-200 pb-1.5">
                <span className="font-bold text-slate-800 flex items-center gap-1.5">
                  <Sliders className="w-3.5 h-3.5 text-blue-700" />
                  <span>Opzioni Parser CSV/TSV & Decimali</span>
                </span>
                <span className="text-[10px] font-mono text-blue-700 bg-blue-50 px-1.5 py-0.5 rounded border border-blue-200">
                  {activePreview.detectedDelimiterLabel}
                </span>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-[10.5px] font-semibold text-slate-700 mb-0.5">
                    Delimitatore Colonne
                  </label>
                  <select
                    value={parseOptions.delimiter}
                    onChange={(e) =>
                      setParseOptions((prev) => ({
                        ...prev,
                        delimiter: e.target.value as DelimiterOption,
                      }))
                    }
                    className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded font-medium"
                  >
                    <option value="auto">Automatico (Smart)</option>
                    <option value="tab">Tabulazione (\t)</option>
                    <option value="semicolon">Punto e virgola (;)</option>
                    <option value="comma">Virgola (,)</option>
                    <option value="space">Spazio / Multi-spazio</option>
                    <option value="custom">Personalizzato...</option>
                  </select>
                </div>

                <div>
                  <label className="block text-[10.5px] font-semibold text-slate-700 mb-0.5">
                    Formato Decimale (. / ,)
                  </label>
                  <div className="grid grid-cols-3 gap-1 bg-slate-200 p-0.5 rounded border border-slate-300">
                    {[
                      { id: 'auto' as DecimalSeparatorOption, label: 'Auto' },
                      { id: 'dot' as DecimalSeparatorOption, label: 'Punto (.)' },
                      { id: 'comma' as DecimalSeparatorOption, label: 'Virgola (,)' },
                    ].map((dec) => (
                      <button
                        key={dec.id}
                        type="button"
                        onClick={() =>
                          setParseOptions((prev) => ({ ...prev, decimalSeparator: dec.id }))
                        }
                        className={`py-0.5 text-[10.5px] font-semibold rounded transition-colors ${
                          parseOptions.decimalSeparator === dec.id
                            ? 'bg-white text-blue-900 shadow-2xs'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        {dec.label}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              {parseOptions.delimiter === 'custom' && (
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[11px] text-slate-700">Carattere separatore:</span>
                  <input
                    type="text"
                    maxLength={4}
                    value={parseOptions.customDelimiter || '|'}
                    onChange={(e) =>
                      setParseOptions((prev) => ({ ...prev, customDelimiter: e.target.value }))
                    }
                    className="w-20 px-2 py-0.5 font-mono text-center border border-slate-300 rounded"
                  />
                </div>
              )}

              {/* Skip Header Lines & Metadata Extraction */}
              <div className="pt-2 border-t border-slate-200 space-y-2">
                <div className="flex items-center justify-between gap-2">
                  <div>
                    <div className="font-semibold text-slate-800 text-[11px]">
                      Skip Righe di Header ($N$ righe)
                    </div>
                    <div className="text-[10px] text-slate-500">
                      Ignora le prime $N$ righe di metadati strumento
                    </div>
                  </div>
                  <input
                    type="number"
                    min={0}
                    max={500}
                    value={parseOptions.skipLines}
                    onChange={(e) =>
                      setParseOptions((prev) => ({
                        ...prev,
                        skipLines: Math.max(0, Number(e.target.value)),
                      }))
                    }
                    className="w-16 px-2 py-1 text-right font-mono font-bold text-xs bg-slate-50 border border-slate-300 rounded"
                  />
                </div>

                <label className="flex items-center gap-2 cursor-pointer text-[11px] text-slate-700">
                  <input
                    type="checkbox"
                    checked={parseOptions.skipCommentLines}
                    onChange={(e) =>
                      setParseOptions((prev) => ({
                        ...prev,
                        skipCommentLines: e.target.checked,
                      }))
                    }
                    className="rounded border-slate-300 text-blue-700"
                  />
                  <span>Ignora automaticamente righe commento (#, %, //, !)</span>
                </label>

                <div className="p-2 bg-slate-50 border border-slate-200 rounded space-y-1.5">
                  <div className="font-semibold text-slate-800 text-[10.5px]">
                    Estrazione Automatica Metadati da Header:
                  </div>
                  <label className="flex items-center gap-2 cursor-pointer text-[11px] text-slate-700">
                    <input
                      type="checkbox"
                      checked={parseOptions.extractHeaderLongName}
                      onChange={(e) =>
                        setParseOptions((prev) => ({
                          ...prev,
                          extractHeaderLongName: e.target.checked,
                        }))
                      }
                      className="rounded border-slate-300 text-blue-700"
                    />
                    <span>Estrai Nome Colonna (Long Name) + Unità tra parentesi</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer text-[11px] text-slate-700">
                    <input
                      type="checkbox"
                      checked={parseOptions.extractHeaderUnits}
                      onChange={(e) =>
                        setParseOptions((prev) => ({
                          ...prev,
                          extractHeaderUnits: e.target.checked,
                        }))
                      }
                      className="rounded border-slate-300 text-blue-700"
                    />
                    <span>Leggi riga dedicata Unità di Misura (2ª riga header)</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer text-[11px] text-slate-700">
                    <input
                      type="checkbox"
                      checked={parseOptions.extractHeaderComments}
                      onChange={(e) =>
                        setParseOptions((prev) => ({
                          ...prev,
                          extractHeaderComments: e.target.checked,
                        }))
                      }
                      className="rounded border-slate-300 text-blue-700"
                    />
                    <span>Leggi riga dedicata Commenti (3ª riga header)</span>
                  </label>
                </div>
              </div>
            </div>
          </div>

          {/* Right Column (7 cols): Live Parsed Table Preview & Skipped Header Inspector */}
          <div className="lg:col-span-7 p-3.5 flex flex-col min-h-0 overflow-hidden bg-white space-y-3">
            {/* Sheet Name & Summary Badges */}
            <div className="flex flex-wrap items-center justify-between gap-2 pb-2 border-b border-slate-200 shrink-0">
              <div className="flex items-center gap-2">
                <label className="text-[11px] font-semibold text-slate-700">Nome Foglio:</label>
                <input
                  type="text"
                  value={customSheetName}
                  onChange={(e) => setCustomSheetName(e.target.value)}
                  className="px-2 py-1 text-xs font-semibold border border-slate-300 rounded bg-slate-50"
                />
              </div>

              <div className="flex items-center gap-2 font-mono text-[11px]">
                <span className="px-2 py-0.5 bg-blue-50 text-blue-900 border border-blue-200 rounded font-semibold">
                  {activePreview.sheet.columns.length} Colonne
                </span>
                <span className="px-2 py-0.5 bg-emerald-50 text-emerald-900 border border-emerald-200 rounded font-semibold">
                  {activePreview.sheet.columns[0]?.values.length || 0} Righe Dati
                </span>
                <span className="px-2 py-0.5 bg-slate-100 text-slate-700 border border-slate-300 rounded">
                  Decimale: {activePreview.detectedDecimalLabel}
                </span>
              </div>
            </div>

            {/* Skipped Header Lines Preview if skipLines > 0 */}
            {activePreview.skippedHeaderPreview.length > 0 && (
              <div className="p-2 bg-amber-50 border border-amber-200 rounded shrink-0">
                <div className="font-semibold text-amber-950 text-[10.5px] mb-1">
                  Righe di Header Strumento Ignorate ({activePreview.skippedHeaderPreview.length}):
                </div>
                <pre className="font-mono text-[10px] text-amber-900 max-h-16 overflow-y-auto whitespace-pre-wrap">
                  {activePreview.skippedHeaderPreview.join('\n')}
                </pre>
              </div>
            )}

            {/* Live Table Preview */}
            <div className="flex-1 border border-slate-300 rounded overflow-auto min-h-0">
              <table className="w-full border-collapse text-left font-mono text-[11px] tabular-nums">
                <thead className="sticky top-0 bg-[#e2e8f0] border-b border-slate-300">
                  <tr>
                    <th className="px-2 py-1 text-center text-slate-500 border-r border-slate-300 w-10">
                      #
                    </th>
                    {activePreview.sheet.columns.map((c) => (
                      <th key={c.id} className="px-2.5 py-1.5 border-r border-slate-300 min-w-[130px]">
                        <div className="flex items-center justify-between">
                          <span className="font-bold text-slate-900">
                            {c.shortName}({c.role})
                          </span>
                          <span className="text-[9.5px] px-1 bg-blue-100 text-blue-900 rounded">
                            {c.unit || 'no-unit'}
                          </span>
                        </div>
                        <div className="text-[10.5px] font-sans font-semibold text-slate-700 truncate mt-0.5">
                          {c.longName}
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {Array.from({
                    length: Math.min(25, activePreview.sheet.columns[0]?.values.length || 0),
                  }).map((_, rIdx) => (
                    <tr key={rIdx} className={rIdx % 2 === 0 ? 'bg-white' : 'bg-slate-50/70'}>
                      <td className="px-2 py-1 text-center text-slate-400 border-r border-slate-200">
                        {rIdx + 1}
                      </td>
                      {activePreview.sheet.columns.map((c) => (
                        <td
                          key={c.id}
                          className="px-2.5 py-1 text-right text-slate-800 border-r border-slate-200"
                        >
                          {c.values[rIdx] !== null && c.values[rIdx] !== undefined
                            ? String(c.values[rIdx])
                            : '--'}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Plot immediately checkbox */}
            <label className="flex items-center gap-2 cursor-pointer pt-1 shrink-0">
              <input
                type="checkbox"
                checked={plotImmediately}
                onChange={(e) => setPlotImmediately(e.target.checked)}
                className="rounded border-slate-300 text-blue-700"
              />
              <span className="text-xs font-medium text-slate-800">
                Traccia automaticamente le colonne Y importate sul grafico attivo
              </span>
            </label>
          </div>
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between px-4 py-2.5 bg-[#eef2f6] border-t border-slate-300 shrink-0">
          <div className="text-[11px] text-slate-600">
            Suggerimento: puoi anche trascinare i file <strong>.csv, .txt, .dat, .xlsx</strong> in qualsiasi punto della finestra principale.
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50"
            >
              Annulla
            </button>
            <button
              type="button"
              onClick={handleConfirmImport}
              className="flex items-center gap-1.5 px-4 py-1.5 text-xs font-bold text-white bg-blue-700 rounded hover:bg-blue-800 shadow-2xs"
            >
              <Check className="w-3.5 h-3.5" />
              <span>
                Importa{' '}
                {sourceMode === 'files' && stagedFiles.length > 1
                  ? `${stagedFiles.length} File (${
                      multiFileMode === 'separate-sheets'
                        ? 'Fogli Separati'
                        : multiFileMode === 'append-rows'
                        ? 'Righe Accodate'
                        : 'Colonne Affiancate'
                    })`
                  : 'Tabella nel Workbook'}
              </span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
