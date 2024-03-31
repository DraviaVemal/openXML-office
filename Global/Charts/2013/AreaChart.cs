// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Aread Chart Core data
	/// </summary>
	public class AreaChart<ApplicationSpecificSetting> : ChartBase<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
	{
		/// <summary>
		/// Area Chart Setting
		/// </summary>
		protected readonly AreaChartSetting<ApplicationSpecificSetting> areaChartSetting;

		internal AreaChart(AreaChartSetting<ApplicationSpecificSetting> areaChartSetting) : base(areaChartSetting)
		{
			this.areaChartSetting = areaChartSetting;
		}

		/// <summary>
		/// Create Area Chart with provided settings
		/// </summary>
		public AreaChart(AreaChartSetting<ApplicationSpecificSetting> areaChartSetting, ChartData[][] dataCols, DataRange? dataRange = null) : base(areaChartSetting)
		{
			this.areaChartSetting = areaChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}

		private C.AreaChartSeries CreateAreaChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			SolidFillModel GetSeriesFillColor()
			{
				SolidFillModel solidFillModel = new();
				string? hexColor = areaChartSetting.areaChartSeriesSettings?
							.Select(item => item?.fillColor)
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
			SolidFillModel GetSeriesBorderColor()
			{
				SolidFillModel solidFillModel = new();
				string? hexColor = areaChartSetting.areaChartSeriesSettings?
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
			ShapePropertiesModel shapePropertiesModel = new()
			{
				solidFill = GetSeriesFillColor(),
				outline = new()
				{
					solidFill = GetSeriesBorderColor()
				}
			};
			C.DataLabels? dataLabels = seriesIndex < areaChartSetting.areaChartSeriesSettings.Count ?
				CreateAreaDataLabels(areaChartSetting.areaChartSeriesSettings?[seriesIndex]?.areaChartDataLabel ?? new AreaChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
			C.AreaChartSeries series = new(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula!, new[] { chartDataGrouping.seriesHeaderCells! }));
			series.Append(CreateChartShapeProperties(shapePropertiesModel));
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

		private C.DataLabels? CreateAreaDataLabels(AreaChartDataLabel areaChartDataLabel, int? dataLabelCounter = 0)
		{
			if (areaChartDataLabel.showValue || areaChartDataLabel.showValueFromColumn || areaChartDataLabel.showCategoryName || areaChartDataLabel.showLegendKey || areaChartDataLabel.showSeriesName)
			{
				C.DataLabels dataLabels = CreateDataLabels(areaChartDataLabel, dataLabelCounter);
				dataLabels.InsertAt(new C.DataLabelPosition()
				{
					Val = areaChartDataLabel.dataLabelPosition switch
					{
						//Show
						_ => C.DataLabelPositionValues.Center,
					}
				}, 0);
				return dataLabels;
			}
			return null;
		}

		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange? dataRange)
		{
			C.PlotArea plotArea = new();
			plotArea.Append(CreateLayout(areaChartSetting.plotAreaOptions?.manualLayout));
			plotArea.Append(CreateAreaChart(CreateDataSeries(areaChartSetting.chartDataSetting, dataCols, dataRange)));
			plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
			{
				id = CategoryAxisId,
				crossAxisId = ValueAxisId,
				fontSize = areaChartSetting.chartAxesOptions.horizontalFontSize,
				isBold = areaChartSetting.chartAxesOptions.isHorizontalBold,
				isItalic = areaChartSetting.chartAxesOptions.isHorizontalItalic,
				isVisible = areaChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
				invertOrder = areaChartSetting.chartAxesOptions.invertHorizontalAxesOrder,
			}));
			plotArea.Append(CreateValueAxis(new ValueAxisSetting()
			{
				id = ValueAxisId,
				crossAxisId = CategoryAxisId,
				fontSize = areaChartSetting.chartAxesOptions.verticalFontSize,
				isBold = areaChartSetting.chartAxesOptions.isVerticalBold,
				isItalic = areaChartSetting.chartAxesOptions.isVerticalItalic,
				isVisible = areaChartSetting.chartAxesOptions.isVerticalAxesEnabled,
				invertOrder = areaChartSetting.chartAxesOptions.invertVerticalAxesOrder,
			}));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}

		internal C.AreaChart CreateAreaChart(List<ChartDataGrouping> chartDataGroupings)
		{
			C.AreaChart areaChart = new(
				new C.Grouping
				{
					Val = areaChartSetting.areaChartTypes switch
					{
						AreaChartTypes.STACKED => C.GroupingValues.Stacked,
						AreaChartTypes.PERCENT_STACKED => C.GroupingValues.PercentStacked,
						// Clusted
						_ => C.GroupingValues.Standard,
					}
				},
				new C.VaryColors { Val = false });
			int seriesIndex = 0;
			chartDataGroupings.ForEach(Series =>
			{
				areaChart.Append(CreateAreaChartSeries(seriesIndex, Series));
				seriesIndex++;
			});
			C.DataLabels? dataLabels = CreateAreaDataLabels(areaChartSetting.areaChartDataLabel);
			if (dataLabels != null)
			{
				areaChart.Append(dataLabels);
			}
			areaChart.Append(new C.AxisId { Val = CategoryAxisId });
			areaChart.Append(new C.AxisId { Val = ValueAxisId });
			return areaChart;
		}


	}
}
