export type ColumnRole = 'X' | 'Y' | 'X_ERR' | 'Y_ERR' | 'LABEL' | 'IGNORE';

export type ColumnDataType = 'numeric' | 'integer' | 'text' | 'datetime';

export type NumericFormatMode = 'auto' | 'decimal' | 'scientific' | 'engineering';

export type NumericPrecisionMode = 'decimals' | 'sigfigs';

export type RowFilterOperator =
  | 'gt'
  | 'gte'
  | 'lt'
  | 'lte'
  | 'eq'
  | 'neq'
  | 'between'
  | 'contains'
  | 'not_nan'
  | 'unmasked'
  | 'custom_expr';

export interface RowFilterConfig {
  enabled: boolean;
  colId: string;
  operator: RowFilterOperator;
  value1: string;
  value2?: string;
  customExpression?: string;
}

export type PlotType =
  | 'line'
  | 'scatter'
  | 'line_symbol'
  | 'spline'
  | 'step'
  | 'area'
  | 'bar'
  | 'stem';

export type LineStyle = 'solid' | 'dashed' | 'dotted' | 'dash_dot';
export type ConnectionType = 'none' | 'linear' | 'spline' | 'step';
export type MarkerShape =
  | 'circle'
  | 'square'
  | 'triangle_up'
  | 'triangle_down'
  | 'diamond'
  | 'cross'
  | 'star'
  | 'hexagon'
  | 'none';

export type FitModelType = 'linear' | 'polynomial' | 'spline' | 'gaussian' | 'exponential';

export type SmoothingFilterMethod =
  | 'savitzky_golay'
  | 'moving_average'
  | 'gaussian'
  | 'median'
  | 'lowpass'
  | 'highpass'
  | 'outlier_zscore';

export type ToolMode = 'pointer' | 'reader' | 'remove_bad' | 'mask_point' | 'zoom' | 'add_text';

export type DockLayoutMode =
  | 'split-horizontal'
  | 'split-vertical'
  | 'mdi-free'
  | 'multi-grid'
  | 'sheet-only'
  | 'graph-only';

export type BadPointRemovalMode = 'cell-only' | 'entire-row' | 'mask-only';

export interface MDIWindowFrame {
  id: string;
  kind: 'sheet' | 'graph';
  targetId: string;
  x: number;
  y: number;
  w: number;
  h: number;
  zIndex: number;
  minimized: boolean;
  maximized: boolean;
}

export interface DataColumn {
  id: string;
  shortName: string;
  longName: string;
  unit: string;
  comment?: string;
  formula?: string;
  role: ColumnRole;
  dataType?: ColumnDataType;
  numericFormat?: NumericFormatMode;
  numericPrecision?: number;
  numericPrecisionMode?: NumericPrecisionMode;
  linkedXColId?: string;
  linkedErrColId?: string;
  linkedXErrColId?: string;
  values: (number | null)[];
  rawStrings?: (string | null)[];
  maskedIndices?: number[];
}

export interface WorkbookSheet {
  id: string;
  name: string;
  isFitOutput?: boolean;
  sourceFitId?: string;
  rowFilter?: RowFilterConfig;
  autoRecalculate?: boolean;
  columns: DataColumn[];
}

export interface AxisConfig {
  label: string;
  autoLabel: boolean;
  scaleType: 'linear' | 'log10';
  min: number | 'auto';
  max: number | 'auto';
  showMajorGrid: boolean;
  showMinorGrid: boolean;
  tickCount: number;
  tickDirection: 'in' | 'out' | 'both';
  labelFormat: 'auto' | 'scientific' | 'decimal';
}

export interface PlotSeries {
  id: string;
  sheetId: string;
  xColId: string;
  yColId: string;
  xErrColId?: string;
  yErrColId?: string;
  labelColId?: string;
  label: string;
  autoLabel: boolean;
  plotType: PlotType;
  color: string;
  lineWidth: number;
  lineStyle: LineStyle;
  connection: ConnectionType;

  // Marker / Symbol properties
  markerShape: MarkerShape;
  markerSize: number;
  symbolFill: string;
  symbolBorder: string;
  symbolBorderWidth: number;

  // Area fill properties
  fillArea: boolean;
  fillColor: string;
  fillOpacity: number;

  // Drop lines to bottom axis
  dropLines: boolean;

  // Error bars
  showErrorBars: boolean;
  errorBarCap: number;
  errorBarColor: string;

  visible: boolean;
}

export interface GraphAnnotation {
  id: string;
  text: string;
  x: number;
  y: number;
  color: string;
  bgColor: string;
  fontSize: number;
  borderColor?: string;
}

export interface FittedCurve {
  id: string;
  name: string;
  sourceSeriesId: string;
  sheetId: string;
  xColId: string;
  yColId: string;
  model: FitModelType;
  polyDegree: number;
  rangeC: number;
  rangeD: number;
  numPointsN: number;
  autoUpdateOnDataChange: boolean;
  outputSheetId?: string;
  equationText: string;
  coefficients: { name: string; value: number; description?: string }[];
  rSquared: number;
  rmse: number;
  sourceDomainMin: number;
  sourceDomainMax: number;
  sourceValidPointsCount: number;
  generatedPoints: { x: number; y: number }[];
  color: string;
  lineWidth: number;
  lineStyle: LineStyle;
  visible: boolean;
}

export interface GraphWindow {
  id: string;
  name: string;
  title: string;
  subtitle: string;
  xAxis: AxisConfig;
  yAxis: AxisConfig;
  showLegend: boolean;
  legendPosition: 'top-right' | 'top-left' | 'bottom-right' | 'bottom-left';
  legendOffset?: { x: number; y: number };
  showFitReportBox: boolean;
  fitBoxOffset?: { x: number; y: number };
  series: PlotSeries[];
  fits: FittedCurve[];
  annotations: GraphAnnotation[];
}

export interface RemovedPointRecord {
  id: string;
  timestamp: string;
  sheetId: string;
  sheetName: string;
  rowIndex: string | number;
  xColId: string;
  yColId: string;
  seriesLabel: string;
  xValue: number;
  yValue: number;
  removalMode: BadPointRemovalMode;
  fullRowSnapshot?: (number | null)[];
}

export interface HoveredDataPoint {
  seriesId: string;
  seriesLabel: string;
  sheetId: string;
  rowIndex: number;
  xColId: string;
  yColId: string;
  x: number;
  y: number;
  yErr?: number;
  svgX: number;
  svgY: number;
  color: string;
  isMasked?: boolean;
}

export interface ColumnStatistics {
  colShortName: string;
  colLongName: string;
  n: number;
  missing: number;
  mean: number;
  sd: number;
  sem: number;
  sum: number;
  min: number;
  median: number;
  max: number;
  variance: number;
}
