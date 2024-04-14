// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using DocumentFormat.OpenXml;
using OpenXMLOffice.Global_2013;
using C = DocumentFormat.OpenXml.Drawing.Charts;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Represents the settings for a column chart.
	/// </summary>
	public class ColumnChart<ApplicationSpecificSetting> : ChartAdvance<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
	{
		private const int DefaultGapWidth = 150;
		private const int DefaultOverlap = 100;
		/// <summary>
		/// Column Chart Setting
		/// </summary>
		protected ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting;
		internal ColumnChart(ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting) : base(columnChartSetting)
		{
			this.columnChartSetting = columnChartSetting;
		}
		/// <summary>
		/// Create Column Chart with provided settings
		/// </summary>
		public ColumnChart(ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting, ChartData[][] dataCols, DataRange dataRange = null) : base(columnChartSetting)
		{
			this.columnChartSetting = columnChartSetting;
			if (columnChartSetting.columnChartType == ColumnChartTypes.CLUSTERED_3D ||
			columnChartSetting.columnChartType == ColumnChartTypes.STACKED_3D ||
			columnChartSetting.columnChartType == ColumnChartTypes.PERCENT_STACKED_3D)
			{
				this.columnChartSetting.is3DChart = true;
				Add3Dcontrol();
			}
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}
		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange dataRange)
		{
			C.PlotArea plotArea = new C.PlotArea();
			plotArea.Append(CreateLayout(columnChartSetting.plotAreaOptions?.manualLayout));
			if (columnChartSetting.is3DChart)
			{
				plotArea.Append(CreateColumnChart<C.Bar3DChart>(CreateDataSeries(columnChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			else
			{
				plotArea.Append(CreateColumnChart<C.BarChart>(CreateDataSeries(columnChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
			{
				id = CategoryAxisId,
				crossAxisId = ValueAxisId,
				fontSize = columnChartSetting.chartAxesOptions.horizontalFontSize,
				isBold = columnChartSetting.chartAxesOptions.isVerticalItalic,
				isItalic = columnChartSetting.chartAxesOptions.isVerticalItalic,
				isVisible = columnChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
				invertOrder = columnChartSetting.chartAxesOptions.invertHorizontalAxesOrder,
			}));
			plotArea.Append(CreateValueAxis(new ValueAxisSetting()
			{
				id = ValueAxisId,
				crossAxisId = CategoryAxisId,
				fontSize = columnChartSetting.chartAxesOptions.verticalFontSize,
				isBold = columnChartSetting.chartAxesOptions.isVerticalBold,
				isItalic = columnChartSetting.chartAxesOptions.isVerticalItalic,
				isVisible = columnChartSetting.chartAxesOptions.isVerticalAxesEnabled,
				invertOrder = columnChartSetting.chartAxesOptions.invertVerticalAxesOrder,
			}));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}
		internal ChartType CreateColumnChart<ChartType>(List<ChartDataGrouping> chartDataGroupings) where ChartType : OpenXmlCompositeElement, new()
		{
			ChartType columnChart = new ChartType();
			C.BarGroupingValues barGroupingValue;
			switch (columnChartSetting.columnChartType)
			{
				case ColumnChartTypes.STACKED:
					barGroupingValue = C.BarGroupingValues.Stacked;
					break;
				case ColumnChartTypes.PERCENT_STACKED:
					barGroupingValue = C.BarGroupingValues.PercentStacked;
					break;
				case ColumnChartTypes.CLUSTERED_3D:
					barGroupingValue = C.BarGroupingValues.Clustered;
					break;
				case ColumnChartTypes.STACKED_3D:
					barGroupingValue = C.BarGroupingValues.Stacked;
					break;
				case ColumnChartTypes.PERCENT_STACKED_3D:
					barGroupingValue = C.BarGroupingValues.PercentStacked;
					break;
				default:
					barGroupingValue = C.BarGroupingValues.Clustered;
					break;
			}
			columnChart.Append(new C.BarDirection { Val = C.BarDirectionValues.Column },
								new C.BarGrouping { Val = barGroupingValue },
								new C.VaryColors { Val = false });
			int seriesIndex = 0;
			chartDataGroupings.ForEach(Series =>
			{
				columnChart.Append(CreateColumnChartSeries(seriesIndex, Series));
				seriesIndex++;
			});
			C.ShapeValues shapeValue;
			switch (columnChartSetting.columnChartType)
			{
				case ColumnChartTypes.CLUSTERED:
					columnChart.Append(new C.GapWidth { Val = (UInt16Value)columnChartSetting.columnGraphicsSetting.categoryGap });
					columnChart.Append(new C.Overlap { Val = (SByteValue)columnChartSetting.columnGraphicsSetting.seriesGap });
					break;
				case ColumnChartTypes.CLUSTERED_3D:
					columnChart.Append(new C.GapWidth { Val = (UInt16Value)columnChartSetting.columnGraphicsSetting.categoryGap });
					switch (columnChartSetting.columnGraphicsSetting.columnShapeType)
					{
						case BarShapeType.FULL_PYRAMID:
							shapeValue = C.ShapeValues.PyramidToMaximum;
							break;
						case BarShapeType.PARTIAL_PYRAMID:
							shapeValue = C.ShapeValues.Pyramid;
							break;
						case BarShapeType.FULL_CONE:
							shapeValue = C.ShapeValues.ConeToMax;
							break;
						case BarShapeType.PARTIAL_CONE:
							shapeValue = C.ShapeValues.Cone;
							break;
						case BarShapeType.CYLINDER:
							shapeValue = C.ShapeValues.Cylinder;
							break;
						default:
							shapeValue = C.ShapeValues.Box;
							break;
					}
					columnChart.Append(new C.Shape() { Val = shapeValue });
					break;
				case ColumnChartTypes.STACKED_3D:
				case ColumnChartTypes.PERCENT_STACKED_3D:
					columnChart.Append(new C.GapWidth { Val = DefaultGapWidth });
					switch (columnChartSetting.columnGraphicsSetting.columnShapeType)
					{
						case BarShapeType.FULL_PYRAMID:
							shapeValue = C.ShapeValues.PyramidToMaximum;
							break;
						case BarShapeType.PARTIAL_PYRAMID:
							shapeValue = C.ShapeValues.Pyramid;
							break;
						case BarShapeType.FULL_CONE:
							shapeValue = C.ShapeValues.ConeToMax;
							break;
						case BarShapeType.PARTIAL_CONE:
							shapeValue = C.ShapeValues.Cone;
							break;
						case BarShapeType.CYLINDER:
							shapeValue = C.ShapeValues.Cylinder;
							break;
						default:
							shapeValue = C.ShapeValues.Box;
							break;
					}
					columnChart.Append(new C.Shape { Val = shapeValue });
					break;
				default:
					columnChart.Append(new C.GapWidth { Val = DefaultGapWidth });
					columnChart.Append(new C.Overlap { Val = DefaultOverlap });
					break;
			}
			C.DataLabels dataLabels = CreateColumnDataLabels(columnChartSetting.columnChartDataLabel);
			if (dataLabels != null)
			{
				columnChart.Append(dataLabels);
			}
			columnChart.Append(new C.AxisId { Val = CategoryAxisId });
			columnChart.Append(new C.AxisId { Val = ValueAxisId });
			return columnChart;
		}
		private C.BarChartSeries CreateColumnChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			SolidFillModel GetSeriesFillColor()
			{
				SolidFillModel solidFillModel = new SolidFillModel();
				string hexColor = columnChartSetting.columnChartSeriesSettings?
							.Select(item => item?.fillColor)
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
						themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColurCount),
					};
				}
				return solidFillModel;
			}
			SolidFillModel GetSeriesBorderColor()
			{
				SolidFillModel solidFillModel = new SolidFillModel();
				string hexColor = columnChartSetting.columnChartSeriesSettings?
							.Select(item => item?.borderColor)
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
						themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColurCount),
					};
				}
				return solidFillModel;
			}
			ShapePropertiesModel shapePropertiesModel = new ShapePropertiesModel()
			{
				solidFill = GetSeriesFillColor(),
				outline = new OutlineModel()
				{
					solidFill = GetSeriesBorderColor()
				}
			};
			C.DataLabels dataLabels = seriesIndex < columnChartSetting.columnChartSeriesSettings.Count ?
				CreateColumnDataLabels(columnChartSetting.columnChartSeriesSettings[seriesIndex]?.columnChartDataLabel ?? new ColumnChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
			C.BarChartSeries series = new C.BarChartSeries(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.InvertIfNegative { Val = true },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula, new[] { chartDataGrouping.seriesHeaderCells }));
			series.Append(CreateChartShapeProperties(shapePropertiesModel));
			int dataPointCount = columnChartSetting.columnChartSeriesSettings?.ElementAtOrDefault(seriesIndex)?.columnChartDataPointSettings.Count ?? 0;
			for (uint index = 0; index < dataPointCount; index++)
			{
				if (columnChartSetting.columnChartSeriesSettings?[seriesIndex]?.columnChartDataPointSettings != null &&
				index < columnChartSetting.columnChartSeriesSettings?[seriesIndex]?.columnChartDataPointSettings.Count &&
				columnChartSetting.columnChartSeriesSettings?[seriesIndex]?.columnChartDataPointSettings[(int)index] != null)
				{
					SolidFillModel GetDataPointFill()
					{
						SolidFillModel solidFillModel = new SolidFillModel();
						string hexColor = columnChartSetting.columnChartSeriesSettings?[seriesIndex]?.columnChartDataPointSettings?
									.Select(item => item?.fillColor)
									.ToList().ElementAtOrDefault((int)index);
						if (hexColor != null)
						{
							solidFillModel.hexColor = hexColor;
							return solidFillModel;
						}
						else
						{
							solidFillModel.schemeColorModel = new SchemeColorModel()
							{
								themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColurCount),
							};
						}
						return solidFillModel;
					}
					SolidFillModel GetDataPointBorder()
					{
						SolidFillModel solidFillModel = new SolidFillModel();
						string hexColor = columnChartSetting.columnChartSeriesSettings?[seriesIndex]?.columnChartDataPointSettings?
									.Select(item => item?.borderColor)
									.ToList().ElementAtOrDefault((int)index);
						if (hexColor != null)
						{
							solidFillModel.hexColor = hexColor;
							return solidFillModel;
						}
						else
						{
							solidFillModel.schemeColorModel = new SchemeColorModel()
							{
								themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColurCount),
							};
						}
						return solidFillModel;
					}
					C.DataPoint dataPoint = new C.DataPoint(new C.Index { Val = index }, new C.Bubble3D { Val = false });
					dataPoint.Append(CreateChartShapeProperties(new ShapePropertiesModel()
					{
						solidFill = GetDataPointFill(),
						outline = new OutlineModel()
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
		private C.DataLabels CreateColumnDataLabels(ColumnChartDataLabel columnChartDataLabel, int? dataLabelCounter = 0)
		{
			if (columnChartDataLabel.showValue || columnChartSetting.chartDataSetting.advancedDataLabel.showValueFromColumn || columnChartDataLabel.showCategoryName || columnChartDataLabel.showLegendKey || columnChartDataLabel.showSeriesName)
			{
				C.DataLabels dataLabels = CreateDataLabels(columnChartDataLabel, dataLabelCounter);
				C.DataLabelPosition dataLabelPosition = new C.DataLabelPosition();
				if (columnChartDataLabel.dataLabelPosition == ColumnChartDataLabel.DataLabelPositionValues.OUTSIDE_END)
				{
					dataLabelPosition.Val = C.DataLabelPositionValues.OutsideEnd;
				}
				else if (columnChartDataLabel.dataLabelPosition == ColumnChartDataLabel.DataLabelPositionValues.INSIDE_END)
				{
					dataLabelPosition.Val = C.DataLabelPositionValues.InsideEnd;
				}
				else if (columnChartDataLabel.dataLabelPosition == ColumnChartDataLabel.DataLabelPositionValues.INSIDE_BASE)
				{
					dataLabelPosition.Val = C.DataLabelPositionValues.InsideBase;
				}
				else
				{
					dataLabelPosition.Val = C.DataLabelPositionValues.Center;
				}
				// Insert dataLabelPosition at index 0
				dataLabels.InsertAt(dataLabelPosition, 0);
				return dataLabels;
			}
			return null;
		}
	}
}
