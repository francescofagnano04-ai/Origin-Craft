import React, { useState } from 'react';
import { X, Check, Sliders, Layers, Activity, CircleDot, BarChart2 } from 'lucide-react';
import {
  ConnectionType,
  GraphWindow,
  LineStyle,
  MarkerShape,
  PlotSeries,
  PlotType,
  WorkbookSheet,
} from '../types/origin';

interface PlotDetailsModalProps {
  isOpen: boolean;
  onClose: () => void;
  graph: GraphWindow;
  sheets: WorkbookSheet[];
  onUpdateGraph: (updated: GraphWindow) => void;
}

export const PlotDetailsModal: React.FC<PlotDetailsModalProps> = ({
  isOpen,
  onClose,
  graph,
  sheets,
  onUpdateGraph,
}) => {
  const [selectedSeriesId, setSelectedSeriesId] = useState<string>(
    graph.series[0]?.id || ''
  );
  const [activeTab, setActiveTab] = useState<'representation' | 'line' | 'symbol' | 'error_bars' | 'area'>(
    'representation'
  );

  if (!isOpen) return null;

  const currentSeries = graph.series.find((s) => s.id === selectedSeriesId) || graph.series[0];
  const targetSheet = currentSeries ? sheets.find((s) => s.id === currentSeries.sheetId) : undefined;
  const errorColumns = targetSheet ? targetSheet.columns.filter((c) => c.role === 'Y_ERR') : [];

  const handleUpdateCurrentSeries = (patch: Partial<PlotSeries>) => {
    if (!currentSeries) return;
    const updated = graph.series.map((s) => (s.id === currentSeries.id ? { ...s, ...patch } : s));
    onUpdateGraph({ ...graph, series: updated });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
      <div className="w-full max-w-3xl bg-white border border-slate-300 rounded shadow-xl flex flex-col max-h-[85vh] overflow-hidden text-xs">
        {/* Title bar */}
        <div className="flex items-center justify-between px-4 py-2.5 bg-[#eef2f6] border-b border-slate-300">
          <div className="flex items-center gap-2">
            <Sliders className="w-4 h-4 text-blue-700" />
            <span className="font-bold text-slate-800 text-sm">
              Origin Plot Details — {graph.name}
            </span>
          </div>
          <button type="button" onClick={onClose} className="p-1 text-slate-500 hover:text-slate-800 rounded">
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Content: Left tree of layers/series + Right properties tabs */}
        <div className="flex-1 flex min-h-0 overflow-hidden">
          {/* Left Tree Explorer */}
          <div className="w-56 shrink-0 bg-[#f8fafc] border-r border-slate-300 p-2.5 overflow-y-auto space-y-2 select-none">
            <div className="font-semibold text-slate-700 text-[11px] px-1">
              Struttura Grafico (Layer 1)
            </div>

            <div className="space-y-1">
              {graph.series.map((s, idx) => {
                const isSelected = s.id === (currentSeries?.id || '');
                const sh = sheets.find((shItem) => shItem.id === s.sheetId);
                const yc = sh?.columns.find((c) => c.id === s.yColId);

                return (
                  <button
                    key={s.id}
                    type="button"
                    onClick={() => setSelectedSeriesId(s.id)}
                    className={`w-full flex items-center gap-2 px-2 py-1.5 rounded text-left transition-colors ${
                      isSelected
                        ? 'bg-blue-700 text-white font-semibold shadow-2xs'
                        : 'text-slate-700 hover:bg-slate-200/70'
                    }`}
                  >
                    <span
                      className="w-3 h-3 rounded-full shrink-0 border border-white"
                      style={{ backgroundColor: s.color }}
                    />
                    <div className="truncate">
                      <div className="truncate text-xs">
                        {yc ? `${yc.shortName}: ${yc.longName}` : s.label}
                      </div>
                      <div className={`text-[10px] font-mono ${isSelected ? 'text-blue-100' : 'text-slate-400'}`}>
                        {s.plotType}
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Right Properties Tabs & Controls */}
          {currentSeries ? (
            <div className="flex-1 flex flex-col min-h-0 bg-white">
              {/* Secondary Tab Bar */}
              <div className="flex items-center gap-1 px-3 pt-2 border-b border-slate-300 bg-[#edf2f7]">
                <button
                  type="button"
                  onClick={() => setActiveTab('representation')}
                  className={`px-3 py-1.5 text-xs font-medium rounded-t border-t border-x transition-colors ${
                    activeTab === 'representation'
                      ? 'bg-white text-slate-900 border-slate-300 font-semibold shadow-2xs'
                      : 'text-slate-600 border-transparent hover:text-slate-900'
                  }`}
                >
                  Rappresentazione
                </button>

                <button
                  type="button"
                  onClick={() => setActiveTab('line')}
                  className={`px-3 py-1.5 text-xs font-medium rounded-t border-t border-x transition-colors ${
                    activeTab === 'line'
                      ? 'bg-white text-slate-900 border-slate-300 font-semibold shadow-2xs'
                      : 'text-slate-600 border-transparent hover:text-slate-900'
                  }`}
                >
                  Linea (Line)
                </button>

                <button
                  type="button"
                  onClick={() => setActiveTab('symbol')}
                  className={`px-3 py-1.5 text-xs font-medium rounded-t border-t border-x transition-colors ${
                    activeTab === 'symbol'
                      ? 'bg-white text-slate-900 border-slate-300 font-semibold shadow-2xs'
                      : 'text-slate-600 border-transparent hover:text-slate-900'
                  }`}
                >
                  Simbolo (Symbol)
                </button>

                <button
                  type="button"
                  onClick={() => setActiveTab('error_bars')}
                  className={`px-3 py-1.5 text-xs font-medium rounded-t border-t border-x transition-colors ${
                    activeTab === 'error_bars'
                      ? 'bg-white text-slate-900 border-slate-300 font-semibold shadow-2xs'
                      : 'text-slate-600 border-transparent hover:text-slate-900'
                  }`}
                >
                  Barre Errore
                </button>

                <button
                  type="button"
                  onClick={() => setActiveTab('area')}
                  className={`px-3 py-1.5 text-xs font-medium rounded-t border-t border-x transition-colors ${
                    activeTab === 'area'
                      ? 'bg-white text-slate-900 border-slate-300 font-semibold shadow-2xs'
                      : 'text-slate-600 border-transparent hover:text-slate-900'
                  }`}
                >
                  Area & Drop Lines
                </button>
              </div>

              {/* Tab Form Fields */}
              <div className="flex-1 overflow-y-auto p-4 space-y-4">
                {activeTab === 'representation' && (
                  <div className="space-y-3">
                    <div>
                      <label className="block text-xs font-semibold text-slate-800 mb-1.5">
                        Tipo di Tracciamento (Plot Type — 8 Varianti OriginLab)
                      </label>
                      <div className="grid grid-cols-2 gap-2">
                        {[
                          {
                            id: 'line_symbol' as PlotType,
                            title: 'Linea + Simboli (Entrambi)',
                            desc: 'Mostra sia la curva che i marker sperimentali',
                          },
                          {
                            id: 'line' as PlotType,
                            title: 'Solo Linea',
                            desc: 'Linea continua senza simboli',
                          },
                          {
                            id: 'scatter' as PlotType,
                            title: 'Solo Simboli (Scatter)',
                            desc: 'Punti singoli non collegati',
                          },
                          {
                            id: 'spline' as PlotType,
                            title: 'Curva Spline Smussata',
                            desc: 'Interpolazione cubica continua C²',
                          },
                          {
                            id: 'step' as PlotType,
                            title: 'Gradino (Step Plot)',
                            desc: 'Transizioni orizzontali-verticali a gradino',
                          },
                          {
                            id: 'area' as PlotType,
                            title: 'Area Sottesa',
                            desc: 'Riempimento colorato sotto la curva',
                          },
                          {
                            id: 'bar' as PlotType,
                            title: 'Colonne / Istogramma',
                            desc: 'Barre verticali verso la baseline Y=0',
                          },
                          {
                            id: 'stem' as PlotType,
                            title: 'Impulsi / Stem (Lollipop)',
                            desc: 'Steli verticali con marker in testa',
                          },
                        ].map((item) => (
                          <button
                            key={item.id}
                            type="button"
                            onClick={() =>
                              handleUpdateCurrentSeries({
                                plotType: item.id,
                                connection:
                                  item.id === 'spline'
                                    ? 'spline'
                                    : item.id === 'step'
                                    ? 'step'
                                    : item.id === 'scatter' ||
                                      item.id === 'bar' ||
                                      item.id === 'stem'
                                    ? 'none'
                                    : currentSeries.connection === 'none'
                                    ? 'linear'
                                    : currentSeries.connection,
                              })
                            }
                            className={`flex items-center gap-2 p-2 rounded border text-left transition-colors ${
                              currentSeries.plotType === item.id
                                ? 'bg-blue-50 border-blue-500 font-semibold text-blue-900'
                                : 'border-slate-200 hover:bg-slate-50'
                            }`}
                          >
                            <CircleDot className="w-3.5 h-3.5 text-blue-600 shrink-0" />
                            <div>
                              <div className="font-medium text-xs">{item.title}</div>
                              <div className="text-[10px] text-slate-500">{item.desc}</div>
                            </div>
                          </button>
                        ))}
                      </div>
                    </div>

                    <div>
                      <label className="block text-xs font-semibold text-slate-800 mb-1">
                        Connessione Linea
                      </label>
                      <select
                        value={currentSeries.connection}
                        onChange={(e) =>
                          handleUpdateCurrentSeries({
                            connection: e.target.value as ConnectionType,
                          })
                        }
                        className="w-full px-2 py-1.5 text-xs bg-white border border-slate-300 rounded"
                      >
                        <option value="linear">Lineare (Spezzata punto-punto)</option>
                        <option value="spline">Spline Cubica (Curva morbida continua)</option>
                        <option value="step">Step (Gradino Orizzontale-Verticale)</option>
                        <option value="none">Nessuna (Scatter puro)</option>
                      </select>
                    </div>
                  </div>
                )}

                {activeTab === 'line' && (
                  <div className="space-y-3">
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-xs font-semibold text-slate-700 mb-1">
                          Colore Linea
                        </label>
                        <div className="flex items-center gap-2">
                          <input
                            type="color"
                            value={currentSeries.color}
                            onChange={(e) => handleUpdateCurrentSeries({ color: e.target.value })}
                            className="w-8 h-8 rounded border border-slate-300 cursor-pointer"
                          />
                          <span className="font-mono text-xs">{currentSeries.color}</span>
                        </div>
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-slate-700 mb-1">
                          Stile Tratto Linea
                        </label>
                        <select
                          value={currentSeries.lineStyle}
                          onChange={(e) =>
                            handleUpdateCurrentSeries({ lineStyle: e.target.value as LineStyle })
                          }
                          className="w-full px-2 py-1.5 text-xs bg-white border border-slate-300 rounded"
                        >
                          <option value="solid">Continua (Solid)</option>
                          <option value="dashed">Tratteggiata (Dash)</option>
                          <option value="dotted">Punteggiata (Dot)</option>
                          <option value="dash_dot">Tratto-Punto (Dash-Dot)</option>
                        </select>
                      </div>
                    </div>

                    <div>
                      <div className="flex items-center justify-between mb-1">
                        <label className="text-xs font-semibold text-slate-700">
                          Spessore Linea
                        </label>
                        <span className="font-mono font-bold text-xs">{currentSeries.lineWidth} pt</span>
                      </div>
                      <input
                        type="range"
                        min={0.5}
                        max={10}
                        step={0.5}
                        value={currentSeries.lineWidth}
                        onChange={(e) =>
                          handleUpdateCurrentSeries({ lineWidth: Number(e.target.value) })
                        }
                        className="w-full accent-blue-700"
                      />
                    </div>
                  </div>
                )}

                {activeTab === 'symbol' && (
                  <div className="space-y-3">
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-xs font-semibold text-slate-700 mb-1">
                          Forma Simbolo
                        </label>
                        <select
                          value={currentSeries.markerShape}
                          onChange={(e) =>
                            handleUpdateCurrentSeries({
                              markerShape: e.target.value as MarkerShape,
                            })
                          }
                          className="w-full px-2 py-1.5 text-xs bg-white border border-slate-300 rounded"
                        >
                          <option value="circle">Cerchio ●</option>
                          <option value="square">Quadrato ■</option>
                          <option value="triangle_up">Triangolo Su ▲</option>
                          <option value="triangle_down">Triangolo Giù ▼</option>
                          <option value="diamond">Rombo ◆</option>
                          <option value="cross">Croce ✕</option>
                          <option value="star">Stella ★</option>
                          <option value="hexagon">Esagono ⬡</option>
                          <option value="none">Nessuno</option>
                        </select>
                      </div>

                      <div>
                        <div className="flex items-center justify-between mb-1">
                          <label className="text-xs font-semibold text-slate-700">Dimensione Simbolo</label>
                          <span className="font-mono font-bold text-xs">{currentSeries.markerSize} pt</span>
                        </div>
                        <input
                          type="range"
                          min={1}
                          max={20}
                          step={0.5}
                          value={currentSeries.markerSize}
                          onChange={(e) =>
                            handleUpdateCurrentSeries({ markerSize: Number(e.target.value) })
                          }
                          className="w-full accent-blue-700"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-xs font-semibold text-slate-700 mb-1">
                          Colore Riempimento Simbolo
                        </label>
                        <input
                          type="color"
                          value={currentSeries.symbolFill || currentSeries.color}
                          onChange={(e) =>
                            handleUpdateCurrentSeries({ symbolFill: e.target.value })
                          }
                          className="w-full h-8 rounded border border-slate-300 cursor-pointer"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-slate-700 mb-1">
                          Colore Bordo Simbolo
                        </label>
                        <input
                          type="color"
                          value={currentSeries.symbolBorder || '#ffffff'}
                          onChange={(e) =>
                            handleUpdateCurrentSeries({ symbolBorder: e.target.value })
                          }
                          className="w-full h-8 rounded border border-slate-300 cursor-pointer"
                        />
                      </div>
                    </div>
                  </div>
                )}

                {activeTab === 'error_bars' && (
                  <div className="space-y-3">
                    <label className="flex items-center gap-2 cursor-pointer font-semibold text-slate-800">
                      <input
                        type="checkbox"
                        checked={currentSeries.showErrorBars}
                        onChange={(e) =>
                          handleUpdateCurrentSeries({ showErrorBars: e.target.checked })
                        }
                        className="rounded border-slate-300 text-blue-700"
                      />
                      <span>Abilita Barre di Errore (Whiskers Y-Error)</span>
                    </label>

                    {currentSeries.showErrorBars && (
                      <div className="p-3 bg-slate-50 border border-slate-200 rounded space-y-3">
                        <div>
                          <label className="block text-xs text-slate-700 mb-1">
                            Colonna Sorgente Incertezza (±σ)
                          </label>
                          <select
                            value={currentSeries.yErrColId || ''}
                            onChange={(e) =>
                              handleUpdateCurrentSeries({ yErrColId: e.target.value || undefined })
                            }
                            className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                          >
                            <option value="">Nessuna colonna esplicita (Default)</option>
                            {errorColumns.map((c) => (
                              <option key={c.id} value={c.id}>
                                {c.shortName}({c.role}): {c.longName} [{c.unit}]
                              </option>
                            ))}
                          </select>
                        </div>

                        <div className="grid grid-cols-2 gap-3">
                          <div>
                            <label className="block text-xs text-slate-700 mb-1">
                              Ampiezza Alette (Cap Width)
                            </label>
                            <input
                              type="range"
                              min={1}
                              max={12}
                              value={currentSeries.errorBarCap || 4}
                              onChange={(e) =>
                                handleUpdateCurrentSeries({ errorBarCap: Number(e.target.value) })
                              }
                              className="w-full accent-blue-700"
                            />
                          </div>

                          <div>
                            <label className="block text-xs text-slate-700 mb-1">Colore Barre</label>
                            <input
                              type="color"
                              value={currentSeries.errorBarColor || currentSeries.color}
                              onChange={(e) =>
                                handleUpdateCurrentSeries({ errorBarColor: e.target.value })
                              }
                              className="w-full h-7 rounded border border-slate-300 cursor-pointer"
                            />
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {activeTab === 'area' && (
                  <div className="space-y-3">
                    <div className="p-3 bg-slate-50 border border-slate-200 rounded space-y-2.5">
                      <label className="flex items-center gap-2 cursor-pointer font-semibold text-slate-800">
                        <input
                          type="checkbox"
                          checked={currentSeries.dropLines}
                          onChange={(e) =>
                            handleUpdateCurrentSeries({ dropLines: e.target.checked })
                          }
                          className="rounded border-slate-300 text-blue-700"
                        />
                        <span>Linee Cadenti all&apos;Asse X (Drop Lines)</span>
                      </label>
                      <p className="text-[11px] text-slate-500">
                        Traccia linee verticali tratteggiate da ogni punto sperimentale verso l&apos;asse orizzontale.
                      </p>
                    </div>

                    <div className="p-3 bg-slate-50 border border-slate-200 rounded space-y-2.5">
                      <div className="font-semibold text-slate-800">Riempimento Area Sottesa</div>
                      <div className="grid grid-cols-2 gap-3">
                        <div>
                          <label className="block text-xs text-slate-700 mb-1">Colore Area</label>
                          <input
                            type="color"
                            value={currentSeries.fillColor || currentSeries.color}
                            onChange={(e) =>
                              handleUpdateCurrentSeries({ fillColor: e.target.value })
                            }
                            className="w-full h-8 rounded border border-slate-300 cursor-pointer"
                          />
                        </div>
                        <div>
                          <div className="flex items-center justify-between mb-1">
                            <label className="text-xs text-slate-700">Trasparenza</label>
                            <span className="font-mono text-xs">
                              {Math.round((currentSeries.fillOpacity ?? 0.2) * 100)}%
                            </span>
                          </div>
                          <input
                            type="range"
                            min={0.05}
                            max={0.9}
                            step={0.05}
                            value={currentSeries.fillOpacity ?? 0.2}
                            onChange={(e) =>
                              handleUpdateCurrentSeries({ fillOpacity: Number(e.target.value) })
                            }
                            className="w-full accent-blue-700"
                          />
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          ) : (
            <div className="flex-1 flex items-center justify-center text-slate-400">
              Nessuna serie selezionata.
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="flex items-center justify-end gap-2 px-4 py-2 bg-[#eef2f6] border-t border-slate-300">
          <button
            type="button"
            onClick={onClose}
            className="flex items-center gap-1.5 px-4 py-1.5 text-xs font-semibold text-white bg-blue-700 rounded hover:bg-blue-800"
          >
            <Check className="w-3.5 h-3.5" />
            <span>Applica e Chiudi</span>
          </button>
        </div>
      </div>
    </div>
  );
};
