// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OpenXMLOffice.Global_2013;
using C = DocumentFormat.OpenXml.Drawing.Charts;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Represents the settings for a line chart.
	/// </summary>
	public class LineChart<ApplicationSpecificSetting> : ChartAdvance<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition, new()
	{
		/// <summary>
		/// The settings for the line chart.
		/// </summary>
		protected LineChartSetting<ApplicationSpecificSetting> lineChartSetting;
		internal LineChart(LineChartSetting<ApplicationSpecificSetting> lineChartSetting) : base(lineChartSetting)
		{
			this.lineChartSetting = lineChartSetting;
		}
		/// <summary>
		/// Create Line Chart with provided settings
		/// </summary>
		public LineChart(LineChartSetting<ApplicationSpecificSetting> lineChartSetting, ChartData[][] dataCols, DataRange dataRange = null) : base(lineChartSetting)
		{
			this.lineChartSetting = lineChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}
		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange dataRange)
		{
			C.PlotArea plotArea = new C.PlotArea();
			plotArea.Append(CreateLayout(lineChartSetting.plotAreaOptions != null ? lineChartSetting.plotAreaOptions.manualLayout : null));
			plotArea.Append(CreateLineChart(CreateDataSeries(lineChartSetting.chartDataSetting, dataCols, dataRange)));
			XAxisOptions xAxisOptions = lineChartSetting.chartAxisOptions.xAxisOptions;
			YAxisOptions yAxisOptions = lineChartSetting.chartAxisOptions.yAxisOptions;
			plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
			{
				id = CategoryAxisId,
				crossAxisId = ValueAxisId,
				title = xAxisOptions.chartAxisTitle.title,
				axesLabelPosition = xAxisOptions.chartAxesOptions.axesLabelPosition,
				axesLabelRotationAngle = xAxisOptions.chartAxesOptions.axesLabelAngle,
				axisPosition = xAxisOptions.chartAxesOptions.inReverseOrder ? AxisPosition.TOP : AxisPosition.BOTTOM,
				fontSize = xAxisOptions.chartAxesOptions.fontSize,
				isBold = xAxisOptions.chartAxesOptions.isBold,
				isItalic = xAxisOptions.chartAxesOptions.isItalic,
				isVisible = xAxisOptions.isAxesVisible,
				invertOrder = xAxisOptions.chartAxesOptions.inReverseOrder,
			}));
			plotArea.Append(CreateValueAxis(new ValueAxisSetting()
			{
				id = ValueAxisId,
				crossAxisId = CategoryAxisId,
				title = yAxisOptions.chartAxisTitle.title,
				axesLabelPosition = yAxisOptions.chartAxesOptions.axesLabelPosition,
				axesLabelRotationAngle = yAxisOptions.chartAxesOptions.axesLabelAngle,
				axisPosition = yAxisOptions.chartAxesOptions.inReverseOrder ? AxisPosition.RIGHT : AxisPosition.LEFT,
				fontSize = yAxisOptions.chartAxesOptions.fontSize,
				isBold = yAxisOptions.chartAxesOptions.isBold,
				isItalic = yAxisOptions.chartAxesOptions.isItalic,
				isVisible = yAxisOptions.isAxesVisible,
				invertOrder = yAxisOptions.chartAxesOptions.inReverseOrder,
			}));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}
		internal C.LineChart CreateLineChart(List<ChartDataGrouping> chartDataGroupings)
		{
			C.Grouping grouping;
			switch (lineChartSetting.lineChartType)
			{
				case LineChartTypes.STACKED:
				case LineChartTypes.STACKED_MARKER:
					grouping = new C.Grouping() { Val = C.GroupingValues.Stacked };
					break;
				case LineChartTypes.PERCENT_STACKED:
				case LineChartTypes.PERCENT_STACKED_MARKER:
					grouping = new C.Grouping() { Val = C.GroupingValues.PercentStacked };
					break;
				default:
					grouping = new C.Grouping() { Val = C.GroupingValues.Standard };
					break;
			}
			C.LineChart lineChart = new C.LineChart(grouping, new C.VaryColors { Val = false });
			int seriesIndex = 0;
			chartDataGroupings.ForEach(Series =>
			{
				lineChart.Append(CreateLineChartSeries(seriesIndex, Series));
				seriesIndex++;
			});
			C.DataLabels dataLabels = CreateLineDataLabels(lineChartSetting.lineChartDataLabel);
			if (dataLabels != null)
			{
				lineChart.Append(dataLabels);
			}
			lineChart.Append(new C.AxisId { Val = CategoryAxisId });
			lineChart.Append(new C.AxisId { Val = ValueAxisId });
			return lineChart;
		}
		private SolidFillModel GetBorderColor(int seriesIndex, ChartDataGrouping chartDataGrouping, LineChartLineFormat lineChartLineFormat)
		{
			SolidFillModel solidFillModel = new SolidFillModel();
			string hexColor = lineChartSetting.lineChartSeriesSettings
						.Select(item => item.borderColor)
						.ToList().ElementAtOrDefault(seriesIndex);
			if ((lineChartLineFormat != null && lineChartLineFormat.lineColor != null) || hexColor != null)
			{
				solidFillModel.hexColor = lineChartLineFormat.lineColor ?? hexColor;
				return solidFillModel;
			}
			else
			{
				solidFillModel.schemeColorModel = new SchemeColorModel()
				{
					themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColorCount),
				};
			}
			return solidFillModel;
		}
		private C.LineChartSeries CreateLineChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			MarkerModel marketModel = new MarkerModel()
			{
				markerShapeType = MarkerShapeTypes.NONE,
			};
			if (new[] { LineChartTypes.CLUSTERED_MARKER, LineChartTypes.STACKED_MARKER, LineChartTypes.PERCENT_STACKED_MARKER }.Contains(lineChartSetting.lineChartType))
			{
				marketModel.markerShapeType = MarkerShapeTypes.CIRCLE;
				marketModel.shapeProperties = new ShapePropertiesModel()
				{
					solidFill = new SolidFillModel()
					{
						schemeColorModel = new SchemeColorModel()
						{
							themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColorCount),
						}
					},
					outline = new OutlineModel()
					{
						solidFill = new SolidFillModel()
						{
							schemeColorModel = new SchemeColorModel()
							{
								themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColorCount),
							}
						}
					}
				};
			}
			LineChartSeriesSetting lineChartSeriesSetting = lineChartSetting.lineChartSeriesSettings.ElementAtOrDefault(seriesIndex);
			if (lineChartSeriesSetting != null)
			{
				marketModel.markerShapeType = lineChartSeriesSetting.markerShapeType != MarkerShapeTypes.NONE ? lineChartSeriesSetting.markerShapeType : marketModel.markerShapeType;
			}
			C.DataLabels dataLabels = null;
			if (seriesIndex < lineChartSetting.lineChartSeriesSettings.Count)
			{
				LineChartDataLabel lineChartDataLabel = lineChartSeriesSetting != null ? lineChartSeriesSetting.lineChartDataLabel : null;
				int dataLabelCellsLength = chartDataGrouping.dataLabelCells != null ? chartDataGrouping.dataLabelCells.Length : 0;
				dataLabels = CreateLineDataLabels(lineChartDataLabel ?? new LineChartDataLabel(), dataLabelCellsLength);
			}
			LineChartLineFormat lineChartLineFormat = lineChartSeriesSetting != null ? lineChartSeriesSetting.lineChartLineFormat : null;
			OutlineModel outlineModel = new OutlineModel()
			{
				solidFill = GetBorderColor(seriesIndex, chartDataGrouping, lineChartLineFormat),
			};
			if (lineChartLineFormat != null)
			{
				outlineModel.beginArrowValues = lineChartLineFormat.beginArrowValues ?? DrawingBeginArrowValues.NONE;
				outlineModel.endArrowValues = lineChartLineFormat.endArrowValues ?? DrawingEndArrowValues.NONE;
				if (lineChartLineFormat.width != null)
				{
					outlineModel.width = (int?)ConverterUtils.PointToEmu((int)lineChartLineFormat.width);
				}
				if (lineChartLineFormat.outlineCapTypeValues != null)
				{
					outlineModel.outlineCapTypeValues = lineChartLineFormat.outlineCapTypeValues;
				}
				if (lineChartLineFormat.outlineLineTypeValues != null)
				{
					outlineModel.outlineLineTypeValues = lineChartLineFormat.outlineLineTypeValues;
				}
				if (outlineModel.dashType != null)
				{
					outlineModel.dashType = lineChartLineFormat.dashType;
				}
				if (outlineModel.lineStartWidth != null)
				{
					outlineModel.lineStartWidth = lineChartLineFormat.lineStartWidth;
				}
				if (outlineModel.lineEndWidth != null)
				{
					outlineModel.lineEndWidth = lineChartLineFormat.lineEndWidth;
				}
			}
			ShapePropertiesModel shapePropertiesModel = new ShapePropertiesModel()
			{
				outline = outlineModel,
			};
			C.LineChartSeries series = new C.LineChartSeries(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula, new[] { chartDataGrouping.seriesHeaderCells }));
			series.Append(CreateChartShapeProperties(shapePropertiesModel));
			series.Append(CreateMarker(marketModel));
			if (lineChartSeriesSetting != null)
			{
				lineChartSeriesSetting.trendLines.ForEach(trendLine =>
				{
					TrendLineModel trendLineModel = new TrendLineModel
					{
						secondaryValue = trendLine.secondaryValue,
						trendLineType = trendLine.trendLineType,
						trendLineName = trendLine.trendLineName,
						forecastBackward = trendLine.forecastBackward,
						forecastForward = trendLine.forecastForward,
						setIntercept = trendLine.setIntercept,
						showEquation = trendLine.showEquation,
						showRSquareValue = trendLine.showRSquareValue,
						interceptValue = trendLine.interceptValue
					};
					series.Append(CreateTrendLine(trendLineModel));
				});
			}
			if (dataLabels != null)
			{
				series.Append(dataLabels);
			}
			series.Append(CreateCategoryAxisData(chartDataGrouping.xAxisFormula, chartDataGrouping.xAxisCells));
			series.Append(CreateValueAxisData(chartDataGrouping.yAxisFormula, chartDataGrouping.yAxisCells));
			if (chartDataGrouping.dataLabelCells != null && chartDataGrouping.dataLabelFormula != null)
			{
				series.Append(new C.ExtensionList(new C.Extension(
					CreateDataLabelsRange(chartDataGrouping.dataLabelFormula, chartDataGrouping.dataLabelCells.Skip(1).ToArray())
				)
				{ Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" }));
			}
			return series;
		}
		private C.DataLabels CreateLineDataLabels(LineChartDataLabel lineChartDataLabel, int? dataLabelCounter = 0)
		{
			if (lineChartDataLabel.showValue || lineChartSetting.chartDataSetting.advancedDataLabel.showValueFromColumn || lineChartDataLabel.showCategoryName || lineChartDataLabel.showLegendKey || lineChartDataLabel.showSeriesName)
			{
				C.DataLabels dataLabels = CreateDataLabels(lineChartDataLabel, dataLabelCounter);
				C.DataLabelPosition dataLabelPosition;
				switch (lineChartDataLabel.dataLabelPosition)
				{
					case LineChartDataLabel.DataLabelPositionValues.LEFT:
						dataLabelPosition = new C.DataLabelPosition() { Val = C.DataLabelPositionValues.Left };
						break;
					case LineChartDataLabel.DataLabelPositionValues.RIGHT:
						dataLabelPosition = new C.DataLabelPosition() { Val = C.DataLabelPositionValues.Right };
						break;
					case LineChartDataLabel.DataLabelPositionValues.ABOVE:
						dataLabelPosition = new C.DataLabelPosition() { Val = C.DataLabelPositionValues.Top };
						break;
					case LineChartDataLabel.DataLabelPositionValues.BELOW:
						dataLabelPosition = new C.DataLabelPosition() { Val = C.DataLabelPositionValues.Bottom };
						break;
					default:
						dataLabelPosition = new C.DataLabelPosition() { Val = C.DataLabelPositionValues.Center };
						break;
				}
				dataLabels.InsertAt(dataLabelPosition, 0);
				return dataLabels;
			}
			return null;
		}
	}
}
