// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using OpenXMLOffice.Global_2013;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Represents the settings for a bar chart.
	/// </summary>
	public class BarChart<ApplicationSpecificSetting> : ChartAdvance<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
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
		public BarChart(BarChartSetting<ApplicationSpecificSetting> barChartSetting, ChartData[][] dataCols, DataRange? dataRange = null) : base(barChartSetting)
		{
			this.barChartSetting = barChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}

		private C.BarChartSeries CreateBarChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			SolidFillModel GetSeriesFillColor()
			{
				SolidFillModel solidFillModel = new();
				string? hexColor = barChartSetting.barChartSeriesSettings?
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
				string? hexColor = barChartSetting.barChartSeriesSettings?
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
			C.DataLabels? dataLabels = seriesIndex < barChartSetting.barChartSeriesSettings.Count ?
				CreateBarDataLabels(barChartSetting.barChartSeriesSettings?[seriesIndex]?.barChartDataLabel ?? new BarChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
			ShapePropertiesModel shapePropertiesModel = new()
			{
				solidFill = GetSeriesFillColor(),
				outline = new()
				{
					solidFill = GetSeriesBorderColor()
				}
			};
			C.BarChartSeries series = new(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.InvertIfNegative { Val = true },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula!, new[] { chartDataGrouping.seriesHeaderCells! }));
			series.Append(CreateChartShapeProperties(shapePropertiesModel));
			int dataPointCount = barChartSetting.barChartSeriesSettings?.ElementAtOrDefault(seriesIndex)?.barChartDataPointSettings.Count ?? 0;
			for (uint index = 0; index < dataPointCount; index++)
			{
				if (barChartSetting.barChartSeriesSettings?[seriesIndex]?.barChartDataPointSettings != null &&
				index < barChartSetting.barChartSeriesSettings?[seriesIndex]?.barChartDataPointSettings.Count &&
				barChartSetting.barChartSeriesSettings?[seriesIndex]?.barChartDataPointSettings[(int)index] != null)
				{
					SolidFillModel GetDataPointFill()
					{
						SolidFillModel solidFillModel = new();
						string? hexColor = barChartSetting.barChartSeriesSettings?[seriesIndex]?.barChartDataPointSettings?
									.Select(item => item?.fillColor)
									.ToList().ElementAtOrDefault((int)index);
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
					SolidFillModel GetDataPointBorder()
					{
						SolidFillModel solidFillModel = new();
						string? hexColor = barChartSetting.barChartSeriesSettings?[seriesIndex]?.barChartDataPointSettings?
									.Select(item => item?.borderColor)
									.ToList().ElementAtOrDefault((int)index);
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
					C.DataPoint dataPoint = new(new C.Index { Val = index }, new C.Bubble3D { Val = false });
					dataPoint.Append(CreateChartShapeProperties(new ShapePropertiesModel()
					{
						solidFill = GetDataPointFill(),
						outline = new()
						{
							solidFill = GetDataPointBorder()
						}
					}));
					series.Append(dataPoint);
				}
			}
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

		private C.DataLabels? CreateBarDataLabels(BarChartDataLabel barChartDataLabel, int? dataLabelCounter = 0)
		{
			if (barChartDataLabel.showValue || barChartDataLabel.showValueFromColumn || barChartDataLabel.showCategoryName || barChartDataLabel.showLegendKey || barChartDataLabel.showSeriesName)
			{
				C.DataLabels dataLabels = CreateDataLabels(barChartDataLabel, dataLabelCounter);
				if (barChartSetting.barChartTypes != BarChartTypes.CLUSTERED && barChartDataLabel.dataLabelPosition == BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END)
				{
					throw new ArgumentException("'Outside End' Data Label Is only Available with Cluster chart type");
				}
				dataLabels.InsertAt(new C.DataLabelPosition()
				{
					Val = barChartDataLabel.dataLabelPosition switch
					{
						BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
						BarChartDataLabel.DataLabelPositionValues.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
						BarChartDataLabel.DataLabelPositionValues.INSIDE_BASE => C.DataLabelPositionValues.InsideBase,
						_ => C.DataLabelPositionValues.Center
					}
				}, 0);
				return dataLabels;
			}
			return null;
		}

		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange? dataRange)
		{
			C.PlotArea plotArea = new();
			plotArea.Append(CreateLayout(barChartSetting.plotAreaOptions?.manualLayout));
			if (barChartSetting.barChartTypes == BarChartTypes.CLUSTERED ||
			barChartSetting.barChartTypes == BarChartTypes.STACKED ||
			barChartSetting.barChartTypes == BarChartTypes.PERCENT_STACKED)
			{
				plotArea.Append(CreateBarChart<C.BarChart>(CreateDataSeries(barChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			else
			{
				plotArea.Append(CreateBarChart<C.Bar3DChart>(CreateDataSeries(barChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
			{
				id = CategoryAxisId,
				crossAxisId = ValueAxisId,
				axisPosition = AxisPosition.LEFT,
				fontSize = barChartSetting.chartAxesOptions.verticalFontSize,
				isBold = barChartSetting.chartAxesOptions.isVerticalBold,
				isItalic = barChartSetting.chartAxesOptions.isVerticalItalic,
				isVisible = barChartSetting.chartAxesOptions.isVerticalAxesEnabled,
				invertOrder = barChartSetting.chartAxesOptions.invertVerticalAxesOrder,
			}));
			plotArea.Append(CreateValueAxis(new ValueAxisSetting()
			{
				id = ValueAxisId,
				crossAxisId = CategoryAxisId,
				axisPosition = AxisPosition.BOTTOM,
				fontSize = barChartSetting.chartAxesOptions.horizontalFontSize,
				isBold = barChartSetting.chartAxesOptions.isHorizontalBold,
				isItalic = barChartSetting.chartAxesOptions.isHorizontalItalic,
				isVisible = barChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
				invertOrder = barChartSetting.chartAxesOptions.invertHorizontalAxesOrder,
			}));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}

		internal ChartType CreateBarChart<ChartType>(List<ChartDataGrouping> chartDataGroupings) where ChartType : OpenXmlCompositeElement, new()
		{
			ChartType barChart = new();
			int seriesIndex = 0;
			chartDataGroupings.ForEach(Series =>
			{
				barChart.Append(CreateBarChartSeries(seriesIndex, Series));
				seriesIndex++;
			});
			switch (barChartSetting.barChartTypes)
			{
				case BarChartTypes.CLUSTERED:
					barChart.Append(new C.GapWidth { Val = (UInt16Value)barChartSetting.barGraphicsSetting.categoryGap });
					barChart.Append(new C.Overlap { Val = (SByteValue)barChartSetting.barGraphicsSetting.seriesGap });
					break;
				case BarChartTypes.CLUSTERED_3D:
					barChart.Append(new C.GapWidth { Val = (UInt16Value)barChartSetting.barGraphicsSetting.categoryGap });
					barChart.Append(new C.Shape()
					{
						Val = barChartSetting.barGraphicsSetting.columnShapeType switch
						{
							ColumnShapeType.FULL_PYRAMID => C.ShapeValues.PyramidToMaximum,
							ColumnShapeType.PARTIAL_PYRAMID => C.ShapeValues.Pyramid,
							ColumnShapeType.FULL_CONE => C.ShapeValues.ConeToMax,
							ColumnShapeType.PARTIAL_CONE => C.ShapeValues.Cone,
							ColumnShapeType.CYLINDER => C.ShapeValues.Cylinder,
							_ => C.ShapeValues.Box
						}
					});
					break;
				case BarChartTypes.STACKED_3D:
				case BarChartTypes.PERCENT_STACKED_3D:
					barChart.Append(new C.GapWidth { Val = DefaultGapWidth });
					barChart.Append(new C.Shape()
					{
						Val = barChartSetting.barGraphicsSetting.columnShapeType switch
						{
							ColumnShapeType.FULL_PYRAMID => C.ShapeValues.PyramidToMaximum,
							ColumnShapeType.PARTIAL_PYRAMID => C.ShapeValues.Pyramid,
							ColumnShapeType.FULL_CONE => C.ShapeValues.ConeToMax,
							ColumnShapeType.PARTIAL_CONE => C.ShapeValues.Cone,
							ColumnShapeType.CYLINDER => C.ShapeValues.Cylinder,
							_ => C.ShapeValues.Box
						}
					});
					break;
				default:
					barChart.Append(new C.GapWidth { Val = DefaultGapWidth });
					barChart.Append(new C.Overlap { Val = DefaultOverlap });
					break;

			}
			C.DataLabels? dataLabels = CreateBarDataLabels(barChartSetting.barChartDataLabel);
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
