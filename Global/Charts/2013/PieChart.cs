// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the types of pie charts.
	/// </summary>
	public class PieChart<ApplicationSpecificSetting> : ChartBase<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
	{

		/// <summary>
		/// The settings for the pie chart.
		/// </summary>
		protected PieChartSetting<ApplicationSpecificSetting> pieChartSetting;

		internal PieChart(PieChartSetting<ApplicationSpecificSetting> pieChartSetting) : base(pieChartSetting)
		{
			this.pieChartSetting = pieChartSetting;
		}

		/// <summary>
		/// Create Pie Chart with provided settings
		/// </summary>
		public PieChart(PieChartSetting<ApplicationSpecificSetting> pieChartSetting, ChartData[][] dataCols, DataRange? dataRange = null) : base(pieChartSetting)
		{
			this.pieChartSetting = pieChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}

		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange? dataRange)
		{
			C.PlotArea plotArea = new();
			plotArea.Append(CreateLayout(pieChartSetting.plotAreaOptions?.manualLayout));
			plotArea.Append(pieChartSetting.pieChartTypes == PieChartTypes.DOUGHNUT ?
				CreateChart<C.DoughnutChart>(CreateDataSeries(pieChartSetting.chartDataSetting, dataCols, dataRange)) :
				CreateChart<C.PieChart>(CreateDataSeries(pieChartSetting.chartDataSetting, dataCols, dataRange)));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}

		internal T CreateChart<T>(List<ChartDataGrouping> chartDataGroupings) where T : new()
		{
			if (typeof(T) != typeof(C.DoughnutChart) && typeof(T) != typeof(C.PieChart))
			{
				throw new ArgumentException("Invalid type parameter. T must be either C.DoughnutChart or C.PieChart.");
			}
			T chartType = new();
			if (chartType is OpenXmlCompositeElement chart)
			{
				chart.Append(new C.VaryColors { Val = true });
				int seriesIndex = 0;
				chartDataGroupings.ForEach(Series =>
				{
					chart.Append(CreateChartSeries(seriesIndex, Series));
					seriesIndex++;
				});
				C.DataLabels? dataLabels = CreatePieDataLabels(pieChartSetting.pieChartDataLabel);
				if (dataLabels != null)
				{
					chart.Append(dataLabels);
				}
				chart.Append(new C.FirstSliceAngle { Val = (UInt16Value)pieChartSetting.angleOfFirstSlice });
				chart.Append(new C.HoleSize { Val = (ByteValue)pieChartSetting.doughnutHoleSize });
			}
			return chartType;
		}

		private C.PieChartSeries CreateChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			C.DataLabels? dataLabels = seriesIndex < pieChartSetting.pieChartSeriesSettings.Count ?
				CreatePieDataLabels(pieChartSetting.pieChartSeriesSettings?[seriesIndex]?.pieChartDataLabel ?? new PieChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
			C.PieChartSeries series = new(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula!, new[] { chartDataGrouping.seriesHeaderCells! }));
			for (uint index = 0; index < chartDataGrouping.xAxisCells!.Length; index++)
			{
				SolidFillModel GetDataPointFill()
				{
					SolidFillModel solidFillModel = new();
					string? hexColor = pieChartSetting.pieChartSeriesSettings?.ElementAtOrDefault(seriesIndex)?.pieChartDataPointSettings?
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
							themeColorValues = ThemeColorValues.ACCENT_1 + ((int)index % AccentColurCount),
						};
					}
					return solidFillModel;
				}
				SolidFillModel GetDataPointBorder()
				{
					SolidFillModel solidFillModel = new();
					string? hexColor = pieChartSetting.pieChartSeriesSettings?.ElementAtOrDefault(seriesIndex)?.pieChartDataPointSettings?
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
							themeColorValues = ThemeColorValues.ACCENT_1 + ((int)index % AccentColurCount),
						};
					}
					return solidFillModel;
				}
				C.DataPoint dataPoint = new(new C.Index { Val = index }, new C.Bubble3D { Val = false });
				ShapePropertiesModel shapePropertiesModel = new()
				{
					solidFill = GetDataPointFill()
				};
				if (pieChartSetting.pieChartTypes != PieChartTypes.DOUGHNUT)
				{
					shapePropertiesModel.outline = new()
					{
						solidFill = GetDataPointBorder()
					};
				}
				dataPoint.Append(CreateChartShapeProperties(shapePropertiesModel));
				if (dataLabels != null)
				{
					series.Append(dataLabels);
				}
				series.Append(dataPoint);
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

		private C.DataLabels? CreatePieDataLabels(PieChartDataLabel pieChartDataLabel, int? dataLabelCounter = 0)
		{
			if (pieChartDataLabel.showValue || pieChartDataLabel.showValueFromColumn || pieChartDataLabel.showCategoryName || pieChartDataLabel.showLegendKey || pieChartDataLabel.showSeriesName)
			{
				C.DataLabels dataLabels = CreateDataLabels(pieChartDataLabel, dataLabelCounter);
				if (pieChartSetting.pieChartTypes == PieChartTypes.DOUGHNUT &&
					new[] { PieChartDataLabel.DataLabelPositionValues.CENTER, PieChartDataLabel.DataLabelPositionValues.INSIDE_END, PieChartDataLabel.DataLabelPositionValues.OUTSIDE_END, PieChartDataLabel.DataLabelPositionValues.BEST_FIT }.Contains(pieChartDataLabel.dataLabelPosition))
				{
					throw new ArgumentException("DataLabelPosition is not supported for Doughnut Chart.");
				}
				if (pieChartSetting.pieChartTypes != PieChartTypes.DOUGHNUT)
				{
					dataLabels.InsertAt(new C.DataLabelPosition()
					{
						Val = pieChartDataLabel.dataLabelPosition switch
						{
							PieChartDataLabel.DataLabelPositionValues.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
							PieChartDataLabel.DataLabelPositionValues.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
							PieChartDataLabel.DataLabelPositionValues.BEST_FIT => C.DataLabelPositionValues.BestFit,
							//Center
							_ => C.DataLabelPositionValues.Center,
						}
					}, 0);
				}
				return dataLabels;
			}
			return null;
		}


	}
}
