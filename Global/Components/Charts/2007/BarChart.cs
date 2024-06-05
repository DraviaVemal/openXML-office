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
	/// Represents the settings for a bar chart.
	/// </summary>
	public class BarChart<ApplicationSpecificSetting> : ChartAdvance<ApplicationSpecificSetting>
		where ApplicationSpecificSetting : class, ISizeAndPosition, new()
	{
		private const int DefaultGapWidth = 150;
		private const int DefaultOverlap = 100;
		/// <summary>
		/// Bar Chart Setting
		/// </summary>
		protected readonly BarChartSetting<ApplicationSpecificSetting> barChartSetting;
		internal BarChart(BarChartSetting<ApplicationSpecificSetting> barChartSetting) : base(barChartSetting)
		{
			this.barChartSetting = barChartSetting;
		}
		/// <summary>
		/// Create Bar Chart with provided settings
		/// </summary>
		public BarChart(BarChartSetting<ApplicationSpecificSetting> barChartSetting, ChartData[][] dataCols, DataRange dataRange = null) : base(barChartSetting)
		{
			this.barChartSetting = barChartSetting;
			if (barChartSetting.barChartType == BarChartTypes.CLUSTERED_3D ||
			barChartSetting.barChartType == BarChartTypes.STACKED_3D ||
			barChartSetting.barChartType == BarChartTypes.PERCENT_STACKED_3D)
			{
				this.barChartSetting.is3DChart = true;
				Add3dControl();
			}
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}
		private ColorOptionModel<SolidOptions> GetSeriesFillColor(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			ColorOptionModel<SolidOptions> solidFillModel = new ColorOptionModel<SolidOptions>();
			string hexColor = barChartSetting.barChartSeriesSettings
						.Select(item => item.fillColor)
						.ToList().ElementAtOrDefault(seriesIndex);
			if (hexColor != null)
			{
				solidFillModel.colorOption.hexColor = hexColor;
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
		private ColorOptionModel<SolidOptions> GetSeriesBorderColor(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			ColorOptionModel<SolidOptions> solidFillModel = new ColorOptionModel<SolidOptions>();
			string hexColor = barChartSetting.barChartSeriesSettings
						.Select(item => item.borderColor)
						.ToList().ElementAtOrDefault(seriesIndex);
			if (hexColor != null)
			{
				solidFillModel.colorOption.hexColor = hexColor;
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
		private ColorOptionModel<SolidOptions> GetDataPointFill(uint index, int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			ColorOptionModel<SolidOptions> solidFillModel = new ColorOptionModel<SolidOptions>();
			string hexColor = barChartSetting.barChartSeriesSettings.ElementAtOrDefault(seriesIndex) != null ? barChartSetting.barChartSeriesSettings.ElementAtOrDefault(seriesIndex).barChartDataPointSettings
						.Select(item => item.fillColor)
						.ToList().ElementAtOrDefault((int)index) : null;
			if (hexColor != null)
			{
				solidFillModel.colorOption.hexColor = hexColor;
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
		private ColorOptionModel<SolidOptions> GetDataPointBorder(uint index, int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			ColorOptionModel<SolidOptions> solidFillModel = new ColorOptionModel<SolidOptions>();
			string hexColor = barChartSetting.barChartSeriesSettings.ElementAtOrDefault(seriesIndex) != null ? barChartSetting.barChartSeriesSettings.ElementAtOrDefault(seriesIndex).barChartDataPointSettings
						.Select(item => item.borderColor)
						.ToList().ElementAtOrDefault((int)index) : null;
			if (hexColor != null)
			{
				solidFillModel.colorOption.hexColor = hexColor;
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
		private C.BarChartSeries CreateBarChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			C.DataLabels dataLabels = null;
			BarChartSeriesSetting barChartSeriesSetting = barChartSetting.barChartSeriesSettings.ElementAtOrDefault(seriesIndex);
			if (seriesIndex < barChartSetting.barChartSeriesSettings.Count)
			{
				BarChartDataLabel barChartDataLabel = barChartSeriesSetting != null ? barChartSeriesSetting.barChartDataLabel : null;
				int dataLabelCellsLength = chartDataGrouping.dataLabelCells != null ? chartDataGrouping.dataLabelCells.Length : 0;
				dataLabels = CreateBarDataLabels(barChartDataLabel ?? new BarChartDataLabel(), dataLabelCellsLength);
			}
			ShapePropertiesModel<SolidOptions, SolidOptions> shapePropertiesModel = new ShapePropertiesModel<SolidOptions, SolidOptions>()
			{
				fillColor = GetSeriesFillColor(seriesIndex, chartDataGrouping),
				lineColor = new OutlineModel<SolidOptions>()
				{
					lineColor = GetSeriesBorderColor(seriesIndex, chartDataGrouping)
				}
			};
			C.BarChartSeries series = new C.BarChartSeries(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.InvertIfNegative { Val = false },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula, new[] { chartDataGrouping.seriesHeaderCells }));
			series.Append(CreateChartShapeProperties(shapePropertiesModel));
			int dataPointCount = 0;
			if (barChartSeriesSetting != null)
			{
				List<BarChartDataPointSetting> dataPointSettings = barChartSeriesSetting.barChartDataPointSettings;
				if (dataPointSettings != null)
				{
					dataPointCount = dataPointSettings.Count;
				}
			}
			for (uint index = 0; index < dataPointCount; index++)
			{
				if (barChartSetting.barChartSeriesSettings.ElementAtOrDefault(seriesIndex).barChartDataPointSettings != null &&
				index < barChartSetting.barChartSeriesSettings.ElementAtOrDefault(seriesIndex).barChartDataPointSettings.Count &&
				barChartSetting.barChartSeriesSettings.ElementAtOrDefault(seriesIndex).barChartDataPointSettings[(int)index] != null)
				{
					C.DataPoint dataPoint = new C.DataPoint(new C.Index { Val = index }, new C.Bubble3D { Val = false });
					dataPoint.Append(CreateChartShapeProperties(new ShapePropertiesModel<SolidOptions, SolidOptions>()
					{
						fillColor = GetDataPointFill(index, seriesIndex, chartDataGrouping),
						lineColor = new OutlineModel<SolidOptions>()
						{
							lineColor = GetDataPointBorder(index, seriesIndex, chartDataGrouping)
						}
					}));
					series.Append(dataPoint);
				}
			}
			if (barChartSeriesSetting != null)
			{
				barChartSeriesSetting.trendLines.ForEach(trendLine =>
				{
					if (!(barChartSetting.barChartType == BarChartTypes.CLUSTERED || barChartSetting.barChartType == BarChartTypes.CLUSTERED_3D))
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
		private C.DataLabels CreateBarDataLabels(BarChartDataLabel barChartDataLabel, int? dataLabelCounter = 0)
		{
			if (barChartDataLabel.showValue || barChartSetting.chartDataSetting.advancedDataLabel.showValueFromColumn || barChartDataLabel.showCategoryName || barChartDataLabel.showLegendKey || barChartDataLabel.showSeriesName)
			{
				C.DataLabels dataLabels = CreateDataLabels(barChartDataLabel, dataLabelCounter);
				if (!(barChartSetting.barChartType == BarChartTypes.CLUSTERED || barChartSetting.barChartType == BarChartTypes.CLUSTERED_3D) && barChartDataLabel.dataLabelPosition == BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END)
				{
					throw new ArgumentException("'Outside End' Data Label Is only Available with Cluster chart type");
				}
				C.DataLabelPositionValues positionValue;
				switch (barChartDataLabel.dataLabelPosition)
				{
					case BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END:
						positionValue = C.DataLabelPositionValues.OutsideEnd;
						break;
					case BarChartDataLabel.DataLabelPositionValues.INSIDE_END:
						positionValue = C.DataLabelPositionValues.InsideEnd;
						break;
					case BarChartDataLabel.DataLabelPositionValues.INSIDE_BASE:
						positionValue = C.DataLabelPositionValues.InsideBase;
						break;
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
			plotArea.Append(CreateLayout(barChartSetting.plotAreaOptions != null ? barChartSetting.plotAreaOptions.manualLayout : null));
			if (barChartSetting.is3DChart)
			{
				plotArea.Append(CreateBarChart<C.Bar3DChart>(CreateDataSeries(barChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			else
			{
				plotArea.Append(CreateBarChart<C.BarChart>(CreateDataSeries(barChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			plotArea.Append(CreateAxis(new AxisSetting<YAxisOptions<CategoryAxis>, CategoryAxis>()
			{
				id = barChartSetting.isSecondaryAxis ? SecondaryCategoryAxisId : CategoryAxisId,
				crossAxisId = barChartSetting.isSecondaryAxis ? SecondaryValueAxisId : ValueAxisId,
				axisOptions = barChartSetting.chartAxisOptions.yAxisOptions,
				axisPosition = barChartSetting.chartAxisOptions.yAxisOptions.chartAxesOptions.inReverseOrder ? AxisPosition.RIGHT : AxisPosition.LEFT
			}));
			plotArea.Append(CreateAxis(new AxisSetting<XAxisOptions<ValueAxis>, ValueAxis>()
			{
				id = barChartSetting.isSecondaryAxis ? SecondaryValueAxisId : ValueAxisId,
				crossAxisId = barChartSetting.isSecondaryAxis ? SecondaryCategoryAxisId : CategoryAxisId,
				axisOptions = barChartSetting.chartAxisOptions.xAxisOptions,
				axisPosition = barChartSetting.chartAxisOptions.xAxisOptions.chartAxesOptions.inReverseOrder ? AxisPosition.TOP : AxisPosition.BOTTOM
			}));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}
		internal C.ShapeValues GetShapeValue(BarShapeType barShapeType)
		{
			switch (barShapeType)
			{
				case BarShapeType.FULL_PYRAMID:
					return C.ShapeValues.PyramidToMaximum;
				case BarShapeType.PARTIAL_PYRAMID:
					return C.ShapeValues.Pyramid;
				case BarShapeType.FULL_CONE:
					return C.ShapeValues.ConeToMax;
				case BarShapeType.PARTIAL_CONE:
					return C.ShapeValues.Cone;
				case BarShapeType.CYLINDER:
					return C.ShapeValues.Cylinder;
				default:
					return C.ShapeValues.Box;
			}
		}
		internal ChartType CreateBarChart<ChartType>(List<ChartDataGrouping> chartDataGroupings) where ChartType : OpenXmlCompositeElement, new()
		{
			ChartType barChart = new ChartType();
			C.BarGroupingValues groupingValue;
			switch (barChartSetting.barChartType)
			{
				case BarChartTypes.STACKED:
					groupingValue = C.BarGroupingValues.Stacked;
					break;
				case BarChartTypes.PERCENT_STACKED:
					groupingValue = C.BarGroupingValues.PercentStacked;
					break;
				case BarChartTypes.CLUSTERED_3D:
					groupingValue = C.BarGroupingValues.Clustered;
					break;
				case BarChartTypes.STACKED_3D:
					groupingValue = C.BarGroupingValues.Stacked;
					break;
				case BarChartTypes.PERCENT_STACKED_3D:
					groupingValue = C.BarGroupingValues.PercentStacked;
					break;
				default:
					groupingValue = C.BarGroupingValues.Clustered;
					break;
			}
			barChart.Append(new C.BarDirection { Val = C.BarDirectionValues.Bar },
							new C.BarGrouping { Val = groupingValue },
							new C.VaryColors { Val = false });
			int seriesIndex = 0;
			chartDataGroupings.ForEach(Series =>
			{
				barChart.Append(CreateBarChartSeries(seriesIndex, Series));
				seriesIndex++;
			});
			switch (barChartSetting.barChartType)
			{
				case BarChartTypes.CLUSTERED:
					barChart.Append(new C.GapWidth { Val = new UInt16Value((ushort)barChartSetting.barGraphicsSetting.categoryGap) });
					barChart.Append(new C.Overlap { Val = new SByteValue((sbyte)barChartSetting.barGraphicsSetting.seriesGap) });
					break;
				case BarChartTypes.CLUSTERED_3D:
				case BarChartTypes.STACKED_3D:
				case BarChartTypes.PERCENT_STACKED_3D:
					barChart.Append(new C.GapWidth { Val = new UInt16Value((ushort)barChartSetting.barGraphicsSetting.categoryGap) });
					C.ShapeValues shapeValue = GetShapeValue(barChartSetting.barGraphicsSetting.barShapeType);
					barChart.Append(new C.Shape { Val = shapeValue });
					break;
				default:
					barChart.Append(new C.GapWidth { Val = DefaultGapWidth });
					barChart.Append(new C.Overlap { Val = DefaultOverlap });
					break;
			}
			C.DataLabels dataLabels = CreateBarDataLabels(barChartSetting.barChartDataLabel);
			if (dataLabels != null)
			{
				barChart.Append(dataLabels);
			}
			barChart.Append(new C.AxisId { Val = CategoryAxisId });
			barChart.Append(new C.AxisId { Val = ValueAxisId });
			if (barChartSetting.is3DChart)
			{
				barChart.Append(new C.AxisId { Val = 0 });
			}
			return barChart;
		}
	}
}
