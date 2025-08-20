import { Chart, ChartType, ChartData, ChartOptions } from './types';
import { generateUniqueId } from './utils';

/**
 * 圖表實現類別
 */
export class ChartImpl implements Chart {
  name: string;
  type: ChartType;
  data: ChartData[];
  options: ChartOptions;
  position: {
    row: number;
    col: number;
  };

  constructor(
    name: string,
    type: ChartType,
    data: ChartData[],
    options: ChartOptions = {},
    position: { row: number; col: number } = { row: 1, col: 1 }
  ) {
    this.name = name;
    this.type = type;
    this.data = data;
    this.options = {
      title: '',
      xAxisTitle: '',
      yAxisTitle: '',
      width: 400,
      height: 300,
      showLegend: true,
      showDataLabels: false,
      showGridlines: true,
      theme: 'light',
      ...options
    };
    this.position = position;
  }

  /**
   * 添加資料系列
   */
  addSeries(series: ChartData): void {
    this.data.push(series);
  }

  /**
   * 移除資料系列
   */
  removeSeries(seriesName: string): void {
    const index = this.data.findIndex(s => s.series === seriesName);
    if (index !== -1) {
      this.data.splice(index, 1);
    }
  }

  /**
   * 更新圖表選項
   */
  updateOptions(options: Partial<ChartOptions>): void {
    this.options = { ...this.options, ...options };
  }

  /**
   * 移動圖表位置
   */
  moveTo(row: number, col: number): void {
    this.position = { row, col };
  }

  /**
   * 調整圖表大小
   */
  resize(width: number, height: number): void {
    this.options.width = width;
    this.options.height = height;
  }

  /**
   * 取得圖表 XML 表示
   */
  toXml(): { chartXml: string; drawingXml: string; chartId: string } {
    const chartId = generateUniqueId();
    const chartXml = this._buildChartXml(chartId);
    const drawingXml = this._buildDrawingXml(chartId);
    
    return {
      chartXml,
      drawingXml,
      chartId
    };
  }

  /**
   * 驗證圖表資料
   */
  validate(): boolean {
    if (this.data.length === 0) return false;
    
    for (const series of this.data) {
      if (!series.series || !series.categories || !series.values) {
        return false;
      }
    }
    
    return true;
  }

  /**
   * 建立圖表 XML
   */
  private _buildChartXml(chartId: string): string {
    const parts = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      '<c:chart>',
      this._buildChartTitle(),
      this._buildChartType(),
      this._buildChartSeries(),
      this._buildChartAxes(),
      this._buildChartLegend(),
      '</c:chart>',
      '</c:chartSpace>'
    ];
    
    return parts.join('');
  }

  /**
   * 建立圖表標題
   */
  private _buildChartTitle(): string {
    if (!this.options.title) return '';
    
    return [
      '<c:title>',
      '<c:tx>',
      '<c:rich>',
      '<a:bodyPr/>',
      '<a:lstStyle/>',
      '<a:p>',
      '<a:r>',
      '<a:t>' + this.options.title + '</a:t>',
      '</a:r>',
      '</a:p>',
      '</c:rich>',
      '</c:tx>',
      '</c:title>'
    ].join('');
  }

  /**
   * 建立圖表類型
   */
  private _buildChartType(): string {
    const chartTypeMap: Record<ChartType, string> = {
      column: 'bar',
      line: 'line',
      pie: 'pie',
      bar: 'bar',
      area: 'area',
      scatter: 'scatter',
      doughnut: 'doughnut',
      radar: 'radar'
    };
    
    const xmlType = chartTypeMap[this.type] || 'bar';
    
    if (xmlType === 'bar') {
      return [
        '<c:barChart>',
        '<c:barDir val="col"/>',
        '<c:grouping val="clustered"/>',
        this._buildChartSeries(),
        '</c:barChart>'
      ].join('');
    } else if (xmlType === 'line') {
      return [
        '<c:lineChart>',
        '<c:grouping val="standard"/>',
        this._buildChartSeries(),
        '</c:lineChart>'
      ].join('');
    } else if (xmlType === 'pie') {
      return [
        '<c:pieChart>',
        this._buildChartSeries(),
        '</c:pieChart>'
      ].join('');
    }
    
    // 預設為柱狀圖
    return [
      '<c:barChart>',
      '<c:barDir val="col"/>',
      '<c:grouping val="clustered"/>',
      this._buildChartSeries(),
      '</c:barChart>'
    ].join('');
  }

  /**
   * 建立圖表資料系列
   */
  private _buildChartSeries(): string {
    return this.data.map((series, index) => [
      '<c:ser>',
      '<c:idx val="' + index + '"/>',
      '<c:order val="' + index + '"/>',
      '<c:tx>',
      '<c:strRef>',
      '<c:f>' + series.series + '</c:f>',
      '</c:strRef>',
      '</c:tx>',
      '<c:cat>',
      '<c:strRef>',
      '<c:f>' + series.categories + '</c:f>',
      '</c:strRef>',
      '</c:cat>',
      '<c:val>',
      '<c:numRef>',
      '<c:f>' + series.values + '</c:f>',
      '</c:numRef>',
      '</c:val>',
      '</c:ser>'
    ].join('')).join('');
  }

  /**
   * 建立圖表軸線
   */
  private _buildChartAxes(): string {
    if (this.type === 'pie' || this.type === 'doughnut') {
      return '';
    }
    
    return [
      '<c:catAx>',
      '<c:axId val="100"/>',
      '<c:scaling>',
      '<c:orientation val="minMax"/>',
      '</c:scaling>',
      '<c:axPos val="b"/>',
      '<c:crossAx val="200"/>',
      '<c:tickLblPos val="nextTo"/>',
      '<c:crosses val="autoZero"/>',
      '</c:catAx>',
      '<c:valAx>',
      '<c:axId val="200"/>',
      '<c:scaling>',
      '<c:orientation val="minMax"/>',
      '</c:scaling>',
      '<c:axPos val="l"/>',
      '<c:crossAx val="100"/>',
      '<c:tickLblPos val="nextTo"/>',
      '<c:crosses val="autoZero"/>',
      '</c:valAx>'
    ].join('');
  }

  /**
   * 建立圖表圖例
   */
  private _buildChartLegend(): string {
    if (!this.options.showLegend) return '';
    
    return [
      '<c:legend>',
      '<c:legendPos val="r"/>',
      '<c:layout/>',
      '<c:overlay val="0"/>',
      '</c:legend>'
    ].join('');
  }

  /**
   * 建立繪圖 XML
   */
  private _buildDrawingXml(chartId: string): string {
    const parts = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
      '<xdr:twoCellAnchor>',
      '<xdr:from>',
      '<xdr:col>' + (this.position.col - 1) + '</xdr:col>',
      '<xdr:colOff>0</xdr:colOff>',
      '<xdr:row>' + (this.position.row - 1) + '</xdr:row>',
      '<xdr:rowOff>0</xdr:rowOff>',
      '</xdr:from>',
      '<xdr:to>',
      '<xdr:col>' + (this.position.col + Math.floor(this.options.width! / 100)) + '</xdr:col>',
      '<xdr:colOff>0</xdr:colOff>',
      '<xdr:row>' + (this.position.row + Math.floor(this.options.height! / 20)) + '</xdr:row>',
      '<xdr:rowOff>0</xdr:rowOff>',
      '</xdr:to>',
      '<xdr:graphicFrame macro="">',
      '<xdr:nvGraphicFramePr>',
      '<xdr:cNvPr id="' + chartId + '" name="' + this.name + '"/>',
      '<xdr:cNvGraphicFramePr/>',
      '</xdr:nvGraphicFramePr>',
      '<xdr:xfrm>',
      '<a:off x="0" y="0"/>',
      '<a:ext cx="' + this.options.width + '" cy="' + this.options.height + '"/>',
      '</xdr:xfrm>',
      '<a:graphic>',
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">',
      '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' + chartId + '"/>',
      '</a:graphicData>',
      '</a:graphic>',
      '</xdr:graphicFrame>',
      '</xdr:twoCellAnchor>',
      '</xdr:wsDr>'
    ];
    
    return parts.join('');
  }
}

/**
 * 圖表工廠類別
 */
export class ChartFactory {
  /**
   * 建立柱狀圖
   */
  static createColumnChart(
    name: string,
    data: ChartData[],
    options?: ChartOptions,
    position?: { row: number; col: number }
  ): ChartImpl {
    return new ChartImpl(name, 'column', data, options, position);
  }

  /**
   * 建立折線圖
   */
  static createLineChart(
    name: string,
    data: ChartData[],
    options?: ChartOptions,
    position?: { row: number; col: number }
  ): ChartImpl {
    return new ChartImpl(name, 'line', data, options, position);
  }

  /**
   * 建立圓餅圖
   */
  static createPieChart(
    name: string,
    data: ChartData[],
    options?: ChartOptions,
    position?: { row: number; col: number }
  ): ChartImpl {
    return new ChartImpl(name, 'pie', data, options, position);
  }

  /**
   * 建立長條圖
   */
  static createBarChart(
    name: string,
    data: ChartData[],
    options?: ChartOptions,
    position?: { row: number; col: number }
  ): ChartImpl {
    return new ChartImpl(name, 'bar', data, options, position);
  }

  /**
   * 建立面積圖
   */
  static createAreaChart(
    name: string,
    data: ChartData[],
    options?: ChartOptions,
    position?: { row: number; col: number }
  ): ChartImpl {
    return new ChartImpl(name, 'area', data, options, position);
  }

  /**
   * 建立散佈圖
   */
  static createScatterChart(
    name: string,
    data: ChartData[],
    options?: ChartOptions,
    position?: { row: number; col: number }
  ): ChartImpl {
    return new ChartImpl(name, 'scatter', data, options, position);
  }
}
