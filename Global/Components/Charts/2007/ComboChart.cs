// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.Collections.Generic;
using System.Linq;
using OpenXMLOffice.Global_2013;
using C = DocumentFormat.OpenXml.Drawing.Charts;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	///
	/// </summary>
	public class ComboChart<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> : ChartAdvance<ApplicationSpecificSetting>
		where ApplicationSpecificSetting : class, ISizeAndPosition, new()
		where XAxisType : class, IAxisTypeOptions, new()
	 	where YAxisType : class, IAxisTypeOptions, new()
	  	where ZAxisType : class, IAxisTypeOptions, new()
	{
		/// <summary>
		///
		/// </summary>
		public ComboChartSetting<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> comboChartSetting { get; private set; }
		/// <summary>
		///
		/// </summary>
		public ComboChart(ComboChartSetting<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> comboChartSetting, ChartData[][] dataCols, DataRange dataRange = null) : base(comboChartSetting)
		{
			this.comboChartSetting = comboChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}
		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange dataRange)
		{
			bool isSecondaryAxisActive = false;
			if (comboChartSetting.ComboChartsSettingList.Count == 0)
			{
				throw new ArgumentException("Combo Chart Series Settings is empty");
			}
			C.PlotArea plotArea = new C.PlotArea();
			plotArea.Append(CreateLayout(comboChartSetting.plotAreaOptions != null ? comboChartSetting.plotAreaOptions.manualLayout : null));
			Dictionary<string, Dictionary<int, object>> groupedCharts = new Dictionary<string, Dictionary<int, object>>();
			int seriesIndex = 0;
			comboChartSetting.ComboChartsSettingList.ForEach(chartSetting =>
			{
				string chartType = chartSetting.GetType().Name;
				if (!groupedCharts.ContainsKey(chartType))
				{
					groupedCharts.Add(chartType, new Dictionary<int, object>());
				}
				Dictionary<int, object> chartList;
				groupedCharts.TryGetValue(chartType, out chartList);
				chartList.Add(seriesIndex, chartSetting);
				++seriesIndex;
			});
			groupedCharts.Keys.ToList().ForEach(chartType =>
			{
				Dictionary<int, object> chartGroup;
				groupedCharts.TryGetValue(chartType, out chartGroup);
				object currentChartSetting = chartGroup.ElementAt(0).Value;
				if (((ChartSetting<ApplicationSpecificSetting>)currentChartSetting).isSecondaryAxis)
				{
					isSecondaryAxisActive = true;
					((ChartSetting<ApplicationSpecificSetting>)currentChartSetting).categoryAxisId = SecondaryCategoryAxisId;
					((ChartSetting<ApplicationSpecificSetting>)currentChartSetting).valueAxisId = SecondaryValueAxisId;
				}
				((ChartSetting<ApplicationSpecificSetting>)currentChartSetting).chartDataSetting = new ChartDataSetting();
				AreaChartSetting<ApplicationSpecificSetting> areaChartSetting = currentChartSetting as AreaChartSetting<ApplicationSpecificSetting>;
				if (areaChartSetting != null)
				{
					AreaChart<ApplicationSpecificSetting> areaChart = new AreaChart<ApplicationSpecificSetting>(areaChartSetting);
					plotArea.Append(areaChart.CreateAreaChart<C.AreaChart>(GetChartPositionData(dataCols, chartGroup.Keys.ToArray(), dataRange)));
				}
				BarChartSetting<ApplicationSpecificSetting> barChartSetting = currentChartSetting as BarChartSetting<ApplicationSpecificSetting>;
				if (barChartSetting != null)
				{
					BarChart<ApplicationSpecificSetting> barChart = new BarChart<ApplicationSpecificSetting>(barChartSetting);
					plotArea.Append(barChart.CreateBarChart<C.BarChart>(GetChartPositionData(dataCols, chartGroup.Keys.ToArray(), dataRange)));
				}
				ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting = currentChartSetting as ColumnChartSetting<ApplicationSpecificSetting>;
				if (columnChartSetting != null)
				{
					ColumnChart<ApplicationSpecificSetting> columnChart = new ColumnChart<ApplicationSpecificSetting>(columnChartSetting);
					plotArea.Append(columnChart.CreateColumnChart<C.BarChart>(GetChartPositionData(dataCols, chartGroup.Keys.ToArray(), dataRange)));
				}
				LineChartSetting<ApplicationSpecificSetting> lineChartSetting = currentChartSetting as LineChartSetting<ApplicationSpecificSetting>;
				if (lineChartSetting != null)
				{
					LineChart<ApplicationSpecificSetting> lineChart = new LineChart<ApplicationSpecificSetting>(lineChartSetting);
					plotArea.Append(lineChart.CreateLineChart(GetChartPositionData(dataCols, chartGroup.Keys.ToArray(), dataRange)));
				}
				PieChartSetting<ApplicationSpecificSetting> pieChartSetting = currentChartSetting as PieChartSetting<ApplicationSpecificSetting>;
				if (pieChartSetting != null)
				{
					PieChart<ApplicationSpecificSetting> pieChart = new PieChart<ApplicationSpecificSetting>(pieChartSetting);
					if (pieChartSetting.pieChartType == PieChartTypes.DOUGHNUT)
					{
						plotArea.Append(pieChart.CreateChart<C.DoughnutChart>(GetChartPositionData(dataCols, chartGroup.Keys.ToArray(), dataRange)));
					}
					else
					{
						plotArea.Append(pieChart.CreateChart<C.PieChart>(GetChartPositionData(dataCols, chartGroup.Keys.ToArray(), dataRange)));
					}
				}
				ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting = currentChartSetting as ScatterChartSetting<ApplicationSpecificSetting>;
				if (scatterChartSetting != null)
				{
					ScatterChart<ApplicationSpecificSetting> scatterChart = new ScatterChart<ApplicationSpecificSetting>(scatterChartSetting);
					if (scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE)
					{
						plotArea.Append(scatterChart.CreateChart<C.BubbleChart>(GetChartPositionData(dataCols, chartGroup.Keys.ToArray(), dataRange)));
					}
					else
					{
						plotArea.Append(scatterChart.CreateChart<C.ScatterChart>(GetChartPositionData(dataCols, chartGroup.Keys.ToArray(), dataRange)));
					}
				}
			});
			plotArea.Append(CreateAxis<C.CategoryAxis, XAxisOptions<XAxisType>, XAxisType>(new AxisSetting<XAxisOptions<XAxisType>, XAxisType>()
			{
				id = CategoryAxisId,
				crossAxisId = ValueAxisId,
				axisOptions = comboChartSetting.chartAxisOptions.xAxisOptions,
				axisPosition = comboChartSetting.chartAxisOptions.xAxisOptions.chartAxesOptions.inReverseOrder ? AxisPosition.TOP : AxisPosition.BOTTOM
			}));
			plotArea.Append(CreateAxis<C.ValueAxis, YAxisOptions<YAxisType>, YAxisType>(new AxisSetting<YAxisOptions<YAxisType>, YAxisType>()
			{
				id = ValueAxisId,
				crossAxisId = CategoryAxisId,
				axisOptions = comboChartSetting.chartAxisOptions.yAxisOptions,
				axisPosition = comboChartSetting.chartAxisOptions.yAxisOptions.chartAxesOptions.inReverseOrder ? AxisPosition.RIGHT : AxisPosition.LEFT
			}));
			if (isSecondaryAxisActive)
			{
				plotArea.Append(CreateAxis<C.CategoryAxis, ZAxisOptions<ZAxisType>, ZAxisType>(new AxisSetting<ZAxisOptions<ZAxisType>, ZAxisType>()
				{
					id = SecondaryCategoryAxisId,
					crossAxisId = SecondaryValueAxisId,
					axisOptions = new ZAxisOptions<ZAxisType>()
					{
						isAxesVisible = false
					},
				}));
				plotArea.Append(CreateAxis<C.ValueAxis, ZAxisOptions<ZAxisType>, ZAxisType>(new AxisSetting<ZAxisOptions<ZAxisType>, ZAxisType>()
				{
					id = SecondaryValueAxisId,
					crossAxisId = SecondaryCategoryAxisId,
					axisOptions = comboChartSetting.chartAxisOptions.zAxisOptions,
					axisPosition = comboChartSetting.secondaryAxisPosition
				}));
			}
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}
		private List<ChartDataGrouping> GetChartPositionData(ChartData[][] dataCols, int[] chartPosition, DataRange dataRange)
		{
			List<ChartDataGrouping> chartDataGroupings = CreateDataSeries(comboChartSetting.chartDataSetting, dataCols, dataRange);
			List<ChartDataGrouping> dataSeries = new List<ChartDataGrouping>();
			chartPosition.ToList().ForEach(index =>
			{
				if (chartDataGroupings.Count > index)
				{
					dataSeries.Add(chartDataGroupings.ElementAt(index));
				}
			});
			return dataSeries;
		}
	}
}
