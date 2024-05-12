// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OpenXMLOffice.Global_2013;
using C = DocumentFormat.OpenXml.Drawing.Charts;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Aread Chart Core data
	/// </summary>
	public class AreaChart<ApplicationSpecificSetting> : ChartAdvance<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition, new()
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
		public AreaChart(AreaChartSetting<ApplicationSpecificSetting> areaChartSetting, ChartData[][] dataCols, DataRange dataRange = null) : base(areaChartSetting)
		{
			this.areaChartSetting = areaChartSetting;
			if (areaChartSetting.areaChartType == AreaChartTypes.CLUSTERED_3D ||
			areaChartSetting.areaChartType == AreaChartTypes.STACKED_3D ||
			areaChartSetting.areaChartType == AreaChartTypes.PERCENT_STACKED_3D)
			{
				this.areaChartSetting.is3DChart = true;
				Add3dControl();
			}
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}
		private SolidFillModel GetSeriesBorderColor(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			SolidFillModel solidFillModel = new SolidFillModel();
			string hexColor = areaChartSetting.areaChartSeriesSettings
						.Select(item => item.borderColor)
						.ToList().ElementAtOrDefault(seriesIndex);
			if (hexColor != null)
			{
				solidFillModel.hexColor = hexColor;
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
		private SolidFillModel GetSeriesFillColor(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			SolidFillModel solidFillModel = new SolidFillModel();
			string hexColor = areaChartSetting.areaChartSeriesSettings
						.Select(item => item.fillColor)
						.ToList().ElementAtOrDefault(seriesIndex);
			if (hexColor != null)
			{
				solidFillModel.hexColor = hexColor;
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
		private C.AreaChartSeries CreateAreaChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			ShapePropertiesModel shapePropertiesModel = new ShapePropertiesModel()
			{
				solidFill = GetSeriesFillColor(seriesIndex, chartDataGrouping),
				outline = new OutlineModel()
				{
					solidFill = GetSeriesBorderColor(seriesIndex, chartDataGrouping)
				}
			};
			C.DataLabels dataLabels = null;
			if (seriesIndex < areaChartSetting.areaChartSeriesSettings.Count)
			{
				AreaChartDataLabel areaChartDataLabel = areaChartSetting.areaChartSeriesSettings.ElementAtOrDefault(seriesIndex) != null ? areaChartSetting.areaChartSeriesSettings.ElementAtOrDefault(seriesIndex).areaChartDataLabel : null;
				int dataLabelCellsLength = chartDataGrouping.dataLabelCells != null ? chartDataGrouping.dataLabelCells.Length : 0;
				dataLabels = CreateAreaDataLabels(areaChartDataLabel ?? new AreaChartDataLabel(), dataLabelCellsLength);
			}
			C.AreaChartSeries series = new C.AreaChartSeries(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula, new[] { chartDataGrouping.seriesHeaderCells }));
			series.Append(CreateChartShapeProperties(shapePropertiesModel));
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
		private C.DataLabels CreateAreaDataLabels(AreaChartDataLabel areaChartDataLabel, int dataLabelCounter = 0)
		{
			if (areaChartDataLabel.showValue || areaChartSetting.chartDataSetting.advancedDataLabel.showValueFromColumn || areaChartDataLabel.showCategoryName || areaChartDataLabel.showLegendKey || areaChartDataLabel.showSeriesName)
			{
				C.DataLabels dataLabels = CreateDataLabels(areaChartDataLabel, dataLabelCounter);
				C.DataLabelPositionValues positionValue;
				switch (areaChartDataLabel.dataLabelPosition)
				{
					// Add cases for other dataLabelPosition values as needed
					default:
						positionValue = C.DataLabelPositionValues.Center;
						break;
				}
				C.DataLabelPosition dataLabelPosition = new C.DataLabelPosition { Val = positionValue };
				dataLabels.InsertAt(dataLabelPosition, 0);
				return dataLabels;
			}
			return null;
		}
		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange dataRange)
		{
			C.PlotArea plotArea = new C.PlotArea();
			plotArea.Append(CreateLayout(areaChartSetting.plotAreaOptions != null ? areaChartSetting.plotAreaOptions.manualLayout : null));
			if (areaChartSetting.is3DChart)
			{
				plotArea.Append(CreateAreaChart<C.Area3DChart>(CreateDataSeries(areaChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			else
			{
				plotArea.Append(CreateAreaChart<C.AreaChart>(CreateDataSeries(areaChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
			{
				id = CategoryAxisId,
				crossAxisId = ValueAxisId,
				axisLabelPosition = areaChartSetting.chartAxisOptions.categoryAxisLabelPosition,
				axisLabelRotationAngle = areaChartSetting.chartAxisOptions.categoryAxisLabelAngle,
				axisPosition = areaChartSetting.chartAxisOptions.valuesInReverseOrder ? AxisPosition.TOP : AxisPosition.BOTTOM,
				fontSize = areaChartSetting.chartAxesOptions.horizontalFontSize,
				isBold = areaChartSetting.chartAxesOptions.isHorizontalBold,
				isItalic = areaChartSetting.chartAxesOptions.isHorizontalItalic,
				isVisible = areaChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
				invertOrder = areaChartSetting.chartAxisOptions.categoryInReverseOrder,
			}));
			plotArea.Append(CreateValueAxis(new ValueAxisSetting()
			{
				id = ValueAxisId,
				crossAxisId = CategoryAxisId,
				axisLabelPosition = areaChartSetting.chartAxisOptions.valueAxisLabelPosition,
				axisLabelRotationAngle = areaChartSetting.chartAxisOptions.valueAxisLabelAngle,
				axisPosition = areaChartSetting.chartAxisOptions.categoryInReverseOrder ? AxisPosition.RIGHT : AxisPosition.LEFT,
				fontSize = areaChartSetting.chartAxesOptions.verticalFontSize,
				isBold = areaChartSetting.chartAxesOptions.isVerticalBold,
				isItalic = areaChartSetting.chartAxesOptions.isVerticalItalic,
				isVisible = areaChartSetting.chartAxesOptions.isVerticalAxesEnabled,
				invertOrder = areaChartSetting.chartAxisOptions.valuesInReverseOrder,
			}));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}
		internal ChartType CreateAreaChart<ChartType>(List<ChartDataGrouping> chartDataGroupings) where ChartType : OpenXmlCompositeElement, new()
		{
			ChartType areaChart = new ChartType();
			C.GroupingValues groupingValue;
			switch (areaChartSetting.areaChartType)
			{
				case AreaChartTypes.STACKED:
					groupingValue = C.GroupingValues.Stacked;
					break;
				case AreaChartTypes.PERCENT_STACKED:
					groupingValue = C.GroupingValues.PercentStacked;
					break;
				case AreaChartTypes.CLUSTERED_3D:
					groupingValue = C.GroupingValues.Standard;
					break;
				case AreaChartTypes.STACKED_3D:
					groupingValue = C.GroupingValues.Stacked;
					break;
				case AreaChartTypes.PERCENT_STACKED_3D:
					groupingValue = C.GroupingValues.PercentStacked;
					break;
				default:
					groupingValue = C.GroupingValues.Standard;
					break;
			}
			areaChart.Append(new C.Grouping { Val = groupingValue }, new C.VaryColors { Val = false });
			int seriesIndex = 0;
			chartDataGroupings.ForEach(Series =>
			{
				areaChart.Append(CreateAreaChartSeries(seriesIndex, Series));
				seriesIndex++;
			});
			C.DataLabels dataLabels = CreateAreaDataLabels(areaChartSetting.areaChartDataLabel);
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
