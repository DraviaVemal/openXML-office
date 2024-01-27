// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a line chart.
	/// </summary>
	public class LineChart : ChartBase
	{

		/// <summary>
		/// The settings for the line chart.
		/// </summary>
		protected LineChartSetting lineChartSetting;

		internal LineChart(LineChartSetting lineChartSetting) : base(lineChartSetting)
		{
			this.lineChartSetting = lineChartSetting;
		}

		/// <summary>
		/// Create Line Chart with provided settings
		/// </summary>
		public LineChart(LineChartSetting lineChartSetting, ChartData[][] dataCols) : base(lineChartSetting)
		{
			this.lineChartSetting = lineChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols));
		}

		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols)
		{
			C.PlotArea plotArea = new();
			plotArea.Append(new C.Layout());
			plotArea.Append(CreateLineChart(CreateDataSeries(dataCols, lineChartSetting.chartDataSetting)));
			plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
			{
				id = CategoryAxisId,
				crossAxisId = ValueAxisId,
				fontSize = lineChartSetting.chartAxesOptions.horizontalFontSize,
				isBold = lineChartSetting.chartAxesOptions.isHorizontalBold,
				isItalic = lineChartSetting.chartAxesOptions.isHorizontalItalic,
				isVisible = lineChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
				invertOrder = lineChartSetting.chartAxesOptions.invertHorizontalAxesOrder,
			}));
			plotArea.Append(CreateValueAxis(new ValueAxisSetting()
			{
				id = ValueAxisId,
				crossAxisId = CategoryAxisId,
				fontSize = lineChartSetting.chartAxesOptions.verticalFontSize,
				isBold = lineChartSetting.chartAxesOptions.isVerticalBold,
				isItalic = lineChartSetting.chartAxesOptions.isVerticalItalic,
				isVisible = lineChartSetting.chartAxesOptions.isVerticalAxesEnabled,
				invertOrder = lineChartSetting.chartAxesOptions.invertVerticalAxesOrder,
			}));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}

		internal C.LineChart CreateLineChart(List<ChartDataGrouping> chartDataGroupings)
		{
			C.LineChart lineChart = new(
							new C.Grouping
							{
								Val = lineChartSetting.lineChartTypes switch
								{
									LineChartTypes.STACKED => C.GroupingValues.Stacked,
									LineChartTypes.STACKED_MARKER => C.GroupingValues.Stacked,
									LineChartTypes.PERCENT_STACKED => C.GroupingValues.PercentStacked,
									LineChartTypes.PERCENT_STACKED_MARKER => C.GroupingValues.PercentStacked,
									// Clusted
									_ => C.GroupingValues.Standard,
								}
							},
							new C.VaryColors { Val = false });
			int seriesIndex = 0;
			chartDataGroupings.ForEach(Series =>
			{
				lineChart.Append(CreateLineChartSeries(seriesIndex, Series));
				seriesIndex++;
			});
			C.DataLabels? dataLabels = CreateLineDataLabels(lineChartSetting.lineChartDataLabel);
			if (dataLabels != null)
			{
				lineChart.Append(dataLabels);
			}
			lineChart.Append(new C.AxisId { Val = CategoryAxisId });
			lineChart.Append(new C.AxisId { Val = ValueAxisId });
			return lineChart;
		}

		private C.LineChartSeries CreateLineChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			MarkerModel marketModel = new()
			{
				markerShapeValues = MarkerModel.MarkerShapeValues.NONE,
			};
			if (new[] { LineChartTypes.CLUSTERED_MARKER, LineChartTypes.STACKED_MARKER, LineChartTypes.PERCENT_STACKED_MARKER }.Contains(lineChartSetting.lineChartTypes))
			{
				marketModel.markerShapeValues = MarkerModel.MarkerShapeValues.CIRCLE;
				marketModel.shapeProperties = new()
				{
					solidFill = new()
					{
						schemeColorModel = new()
						{
							themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColurCount),
						}
					},
					outline = new()
					{
						solidFill = new()
						{
							schemeColorModel = new()
							{
								themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColurCount),
							}
						}
					}
				};
			}
			C.DataLabels? dataLabels = seriesIndex < lineChartSetting.lineChartSeriesSettings.Count ?
				CreateLineDataLabels(lineChartSetting.lineChartSeriesSettings?[seriesIndex]?.lineChartDataLabel ?? new LineChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
			SolidFillModel GetBorderColor()
			{
				SolidFillModel solidFillModel = new();
				string? hexColor = lineChartSetting.lineChartSeriesSettings?
							.Select(item => item?.borderColor)
							.ToList().ElementAtOrDefault(seriesIndex);
				if (hexColor != null)
				{
					solidFillModel.hexColor = hexColor;
					return solidFillModel;
				}
				else
				{
					solidFillModel.schemeColorModel = new()
					{
						themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColurCount),
					};
				}
				return solidFillModel;
			}
			C.LineChartSeries series = new(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula!, new[] { chartDataGrouping.seriesHeaderCells! }));
			ShapePropertiesModel shapePropertiesModel = new()
			{
				outline = new()
				{
					solidFill = GetBorderColor()
				}
			};
			series.Append(CreateChartShapeProperties(shapePropertiesModel));
			series.Append(CreateMarker(marketModel));
			if (dataLabels != null)
			{
				series.Append(dataLabels);
			}
			series.Append(CreateCategoryAxisData(chartDataGrouping.xAxisFormula!, chartDataGrouping.xAxisCells!));
			series.Append(CreateValueAxisData(chartDataGrouping.yAxisFormula!, chartDataGrouping.yAxisCells!));
			if (chartDataGrouping.dataLabelCells != null && chartDataGrouping.dataLabelFormula != null)
			{
				series.Append(new C.ExtensionList(new C.Extension(
					CreateDataLabelsRange(chartDataGrouping.dataLabelFormula, chartDataGrouping.dataLabelCells.Skip(1).ToArray())
				)
				{ Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" }));
			}
			return series;
		}

		private C.DataLabels? CreateLineDataLabels(LineChartDataLabel lineChartDataLabel, int? dataLabelCounter = 0)
		{
			if (lineChartDataLabel.showValue || lineChartDataLabel.showValueFromColumn || lineChartDataLabel.showCategoryName || lineChartDataLabel.showLegendKey || lineChartDataLabel.showSeriesName)
			{
				C.DataLabels dataLabels = CreateDataLabels(lineChartDataLabel, dataLabelCounter);
				dataLabels.InsertAt(new C.DataLabelPosition()
				{
					Val = lineChartDataLabel.dataLabelPosition switch
					{
						LineChartDataLabel.DataLabelPositionValues.LEFT => C.DataLabelPositionValues.Left,
						LineChartDataLabel.DataLabelPositionValues.RIGHT => C.DataLabelPositionValues.Right,
						LineChartDataLabel.DataLabelPositionValues.ABOVE => C.DataLabelPositionValues.Top,
						LineChartDataLabel.DataLabelPositionValues.BELOW => C.DataLabelPositionValues.Bottom,
						//Center
						_ => C.DataLabelPositionValues.Center,
					}
				}, 0);
				return dataLabels;
			}
			return null;
		}


	}
}
