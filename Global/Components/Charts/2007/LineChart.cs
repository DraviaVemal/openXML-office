// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
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
	public class LineChart<ApplicationSpecificSetting> : ChartAdvance<ApplicationSpecificSetting>
		where ApplicationSpecificSetting : class, ISizeAndPosition, new()
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
			plotArea.Append(CreateAxis(new AxisSetting<XAxisOptions<CategoryAxis>, CategoryAxis>()
			{
				id = lineChartSetting.isSecondaryAxis ? SecondaryCategoryAxisId : CategoryAxisId,
				crossAxisId = lineChartSetting.isSecondaryAxis ? SecondaryValueAxisId : ValueAxisId,
				axisOptions = lineChartSetting.chartAxisOptions.xAxisOptions,
				axisPosition = lineChartSetting.chartAxisOptions.xAxisOptions.chartAxesOptions.inReverseOrder ? AxisPosition.TOP : AxisPosition.BOTTOM
			}));
			plotArea.Append(CreateAxis(new AxisSetting<YAxisOptions<ValueAxis>, ValueAxis>()
			{
				id = lineChartSetting.isSecondaryAxis ? SecondaryValueAxisId : ValueAxisId,
				crossAxisId = lineChartSetting.isSecondaryAxis ? SecondaryCategoryAxisId : CategoryAxisId,
				axisOptions = lineChartSetting.chartAxisOptions.yAxisOptions,
				axisPosition = lineChartSetting.chartAxisOptions.yAxisOptions.chartAxesOptions.inReverseOrder ? AxisPosition.RIGHT : AxisPosition.LEFT
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
		private ColorOptionModel<SolidOptions> GetBorderColor(int seriesIndex, ChartDataGrouping chartDataGrouping, LineChartLineFormat lineChartLineFormat)
		{
			ColorOptionModel<SolidOptions> solidFillModel = new ColorOptionModel<SolidOptions>();
			string hexColor = lineChartSetting.lineChartSeriesSettings
						.Select(item => item.borderColor)
						.ToList().ElementAtOrDefault(seriesIndex);
			if ((lineChartLineFormat != null && lineChartLineFormat.lineColor != null) || hexColor != null)
			{
				solidFillModel.colorOption.hexColor = lineChartLineFormat.lineColor ?? hexColor;
				return solidFillModel;
			}
			else
			{
				solidFillModel.colorOption.schemeColorModel = new SchemeColorModel()
				{
					themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColorCount),
				};
			}
			return solidFillModel;
		}
		private C.LineChartSeries CreateLineChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{

			LineChartSeriesSetting lineChartSeriesSetting = lineChartSetting.lineChartSeriesSettings.ElementAtOrDefault(seriesIndex);
			C.DataLabels dataLabels = null;
			if (seriesIndex < lineChartSetting.lineChartSeriesSettings.Count)
			{
				LineChartDataLabel lineChartDataLabel = lineChartSeriesSetting != null ? lineChartSeriesSetting.lineChartDataLabel : null;
				int dataLabelCellsLength = chartDataGrouping.dataLabelCells != null ? chartDataGrouping.dataLabelCells.Length : 0;
				dataLabels = CreateLineDataLabels(lineChartDataLabel ?? new LineChartDataLabel(), dataLabelCellsLength);
			}
			LineChartLineFormat lineChartLineFormat = lineChartSeriesSetting != null ? lineChartSeriesSetting.lineChartLineFormat : null;
			OutlineModel<SolidOptions> outlineModel = new OutlineModel<SolidOptions>()
			{
				lineColor = GetBorderColor(seriesIndex, chartDataGrouping, lineChartLineFormat),
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
				if (lineChartLineFormat.transparency != null)
				{
					outlineModel.lineColor.colorOption.transparency = (int)lineChartLineFormat.transparency;
				}
			}
			ShapePropertiesModel<SolidOptions, NoFillOptions> shapePropertiesModel = new ShapePropertiesModel<SolidOptions, NoFillOptions>()
			{
				lineColor = outlineModel,
			};
			C.LineChartSeries series = new C.LineChartSeries(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula, new[] { chartDataGrouping.seriesHeaderCells }));
			series.Append(CreateChartShapeProperties(shapePropertiesModel));
			if (new[] { LineChartTypes.CLUSTERED_MARKER, LineChartTypes.STACKED_MARKER, LineChartTypes.PERCENT_STACKED_MARKER }.Contains(lineChartSetting.lineChartType))
			{
				MarkerModel<SolidOptions, SolidOptions> marketModel = new MarkerModel<SolidOptions, SolidOptions>()
				{
					markerShapeType = MarkerShapeTypes.NONE,
				};
				marketModel.markerShapeType = MarkerShapeTypes.CIRCLE;
				marketModel.shapeProperties = new ShapePropertiesModel<SolidOptions, SolidOptions>()
				{
					fillColor = new ColorOptionModel<SolidOptions>()
					{
						colorOption = new SolidOptions()
						{
							schemeColorModel = new SchemeColorModel()
							{
								themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColorCount),
							}
						}
					},
					lineColor = new OutlineModel<SolidOptions>()
					{
						lineColor = new ColorOptionModel<SolidOptions>()
						{
							colorOption = new SolidOptions()
							{
								schemeColorModel = new SchemeColorModel()
								{
									themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColorCount),
								}
							}
						}
					}
				};
				if (lineChartSeriesSetting != null)
				{
					marketModel.markerShapeType = lineChartSeriesSetting.markerShapeType != MarkerShapeTypes.NONE ? lineChartSeriesSetting.markerShapeType : marketModel.markerShapeType;
				}
				series.Append(CreateMarker(marketModel));
			}
			else
			{
				MarkerModel<NoFillOptions, NoFillOptions> marketModel = new MarkerModel<NoFillOptions, NoFillOptions>()
				{
					markerShapeType = MarkerShapeTypes.NONE,
				};
				if (lineChartSeriesSetting != null)
				{
					marketModel.markerShapeType = lineChartSeriesSetting.markerShapeType != MarkerShapeTypes.NONE ? lineChartSeriesSetting.markerShapeType : marketModel.markerShapeType;
				}
				series.Append(CreateMarker(marketModel));
			}
			if (lineChartSeriesSetting != null)
			{
				lineChartSeriesSetting.trendLines.ForEach(trendLine =>
				{
					if (lineChartSetting.lineChartType != LineChartTypes.CLUSTERED)
					{
						throw new ArgumentException("Treadline is not supported in the given chart type");
					}
					ColorOptionModel<SolidOptions> solidFillModel = new ColorOptionModel<SolidOptions>();
					if (trendLine.hexColor != null)
					{
						solidFillModel.colorOption.hexColor = trendLine.hexColor;
					}
					else
					{
						solidFillModel.colorOption.schemeColorModel = new SchemeColorModel()
						{
							themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColorCount)
						};
					}
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
						interceptValue = trendLine.interceptValue,
						solidFill = solidFillModel,
						drawingPresetLineDashValues = trendLine.lineStye
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
