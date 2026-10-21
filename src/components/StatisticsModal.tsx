import React from 'react';
import { X, Download, FileSpreadsheet, Copy } from 'lucide-react';
import { ColumnStatistics, WorkbookSheet } from '../types/origin';
import { formatSciNumber } from '../utils/mathEngine';

interface StatisticsModalProps {
  isOpen: boolean;
  onClose: () => void;
  stats: ColumnStatistics[];
  onCreateStatsSheet: (newSheet: WorkbookSheet) => void;
}

export const StatisticsModal: React.FC<StatisticsModalProps> = ({
  isOpen,
  onClose,
  stats,
  onCreateStatsSheet,
}) => {
  if (!isOpen) return null;

  const handleCreateSheet = () => {
    const colNames = stats.map((s) => `${s.colShortName}: ${s.colLongName}`);
    const rowsMetrics = [
      { metric: 'N (Validi)', vals: stats.map((s) => s.n) },
      { metric: 'Mancanti', vals: stats.map((s) => s.missing) },
      { metric: 'Media', vals: stats.map((s) => Number(s.mean.toFixed(6))) },
      { metric: 'Dev. Standard (SD)', vals: stats.map((s) => Number(s.sd.toFixed(6))) },
      { metric: 'Errore Standard (SEM)', vals: stats.map((s) => Number(s.sem.toFixed(6))) },
      { metric: 'Somma', vals: stats.map((s) => Number(s.sum.toFixed(6))) },
      { metric: 'Minimo', vals: stats.map((s) => s.min) },
      { metric: 'Mediana', vals: stats.map((s) => Number(s.median.toFixed(6))) },
      { metric: 'Massimo', vals: stats.map((s) => s.max) },
      { metric: 'Varianza', vals: stats.map((s) => Number(s.variance.toFixed(6))) },
    ];

    const newSheet: WorkbookSheet = {
      id: `sheet_stats_${Date.now()}`,
      name: 'Stats_Descrittive',
      columns: [
        {
          id: `col_stat_metric_${Date.now()}`,
          shortName: 'A',
          longName: 'Grandezza Statistica',
          unit: '',
          role: 'X',
          values: new Array(rowsMetrics.length).fill(null), // string placeholders
        },
        ...stats.map((s, idx) => ({
          id: `col_stat_${idx}_${Date.now()}`,
          shortName: String.fromCharCode(66 + idx),
          longName: `${s.colShortName} (${s.colLongName})`,
          unit: '',
          role: 'Y' as const,
          values: rowsMetrics.map((r) => r.vals[idx]),
        })),
      ],
    };

    onCreateStatsSheet(newSheet);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
      <div className="w-full max-w-4xl bg-white border border-slate-300 rounded shadow-xl flex flex-col max-h-[85vh] overflow-hidden text-xs">
        {/* Header */}
        <div className="flex items-center justify-between px-4 py-2.5 bg-[#eef2f6] border-b border-slate-300">
          <div className="flex items-center gap-2">
            <span className="font-bold text-slate-800 text-sm">
              Origin Statistics: Statistiche Descrittive su Colonne
            </span>
          </div>
          <button type="button" onClick={onClose} className="p-1 text-slate-500 hover:text-slate-800 rounded">
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Content Table */}
        <div className="flex-1 overflow-auto p-4">
          <table className="w-full border-collapse border border-slate-300 text-left font-mono tabular-nums">
            <thead className="bg-[#e2e8f0] border-b border-slate-300 sticky top-0">
              <tr>
                <th className="px-2.5 py-1.5 border-r border-slate-300 font-semibold text-slate-700">Parametro</th>
                {stats.map((s, idx) => (
                  <th key={idx} className="px-2.5 py-1.5 border-r border-slate-300 text-right font-semibold text-slate-900">
                    {s.colShortName}: {s.colLongName}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 bg-white">
              <tr>
                <td className="px-2.5 py-1.5 border-r border-slate-300 font-sans font-medium text-slate-700">N (Punti Validi)</td>
                {stats.map((s, i) => (
                  <td key={i} className="px-2.5 py-1.5 border-r border-slate-200 text-right font-bold text-blue-900">{s.n}</td>
                ))}
              </tr>
              <tr className="bg-slate-50/50">
                <td className="px-2.5 py-1.5 border-r border-slate-300 font-sans text-slate-700">Mancanti / Nulli</td>
                {stats.map((s, i) => (
                  <td key={i} className="px-2.5 py-1.5 border-r border-slate-200 text-right text-slate-500">{s.missing}</td>
                ))}
              </tr>
              <tr>
                <td className="px-2.5 py-1.5 border-r border-slate-300 font-sans font-medium text-slate-800">Media (Mean)</td>
                {stats.map((s, i) => (
                  <td key={i} className="px-2.5 py-1.5 border-r border-slate-200 text-right font-semibold text-slate-900">{formatSciNumber(s.mean, 5)}</td>
                ))}
              </tr>
              <tr className="bg-slate-50/50">
                <td className="px-2.5 py-1.5 border-r border-slate-300 font-sans text-slate-700">Deviazione Standard (SD)</td>
                {stats.map((s, i) => (
                  <td key={i} className="px-2.5 py-1.5 border-r border-slate-200 text-right">{formatSciNumber(s.sd, 5)}</td>
                ))}
              </tr>
              <tr>
                <td className="px-2.5 py-1.5 border-r border-slate-300 font-sans text-slate-700">Errore Standard (SEM)</td>
                {stats.map((s, i) => (
                  <td key={i} className="px-2.5 py-1.5 border-r border-slate-200 text-right text-slate-600">{formatSciNumber(s.sem, 5)}</td>
                ))}
              </tr>
              <tr className="bg-slate-50/50">
                <td className="px-2.5 py-1.5 border-r border-slate-300 font-sans text-slate-700">Somma Totale</td>
                {stats.map((s, i) => (
                  <td key={i} className="px-2.5 py-1.5 border-r border-slate-200 text-right">{formatSciNumber(s.sum, 5)}</td>
                ))}
              </tr>
              <tr>
                <td className="px-2.5 py-1.5 border-r border-slate-300 font-sans text-slate-700">Minimo</td>
                {stats.map((s, i) => (
                  <td key={i} className="px-2.5 py-1.5 border-r border-slate-200 text-right">{formatSciNumber(s.min, 5)}</td>
                ))}
              </tr>
              <tr className="bg-slate-50/50">
                <td className="px-2.5 py-1.5 border-r border-slate-300 font-sans text-slate-700">Mediana</td>
                {stats.map((s, i) => (
                  <td key={i} className="px-2.5 py-1.5 border-r border-slate-200 text-right">{formatSciNumber(s.median, 5)}</td>
                ))}
              </tr>
              <tr>
                <td className="px-2.5 py-1.5 border-r border-slate-300 font-sans text-slate-700">Massimo</td>
                {stats.map((s, i) => (
                  <td key={i} className="px-2.5 py-1.5 border-r border-slate-200 text-right">{formatSciNumber(s.max, 5)}</td>
                ))}
              </tr>
              <tr className="bg-slate-50/50">
                <td className="px-2.5 py-1.5 border-r border-slate-300 font-sans text-slate-700">Varianza (s²)</td>
                {stats.map((s, i) => (
                  <td key={i} className="px-2.5 py-1.5 border-r border-slate-200 text-right">{formatSciNumber(s.variance, 5)}</td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>

        {/* Footer */}
        <div className="flex items-center justify-end gap-2 px-4 py-2.5 bg-[#eef2f6] border-t border-slate-300">
          <button
            type="button"
            onClick={handleCreateSheet}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-blue-800 bg-blue-50 border border-blue-300 rounded hover:bg-blue-100"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>Crea Foglio Dati con Statistiche</span>
          </button>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50"
          >
            Chiudi
          </button>
        </div>
      </div>
    </div>
  );
};
