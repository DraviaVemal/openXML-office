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
	public class ComboChart<ApplicationSpecificSetting> : ChartAdvance<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
	{
		/// <summary>
		///
		/// </summary>
		public ComboChartSetting<ApplicationSpecificSetting> ComboChartSetting { get; private set; }
		/// <summary>
		///
		/// </summary>
		public ComboChart(ComboChartSetting<ApplicationSpecificSetting> comboChartSetting, ChartData[][] dataCols, DataRange dataRange = null) : base(comboChartSetting)
		{
			ComboChartSetting = comboChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}
		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange dataRange)
		{
			bool isSecondaryAxisActive = false;
			if (ComboChartSetting.ComboChartsSettingList.Count == 0)
			{
				throw new ArgumentException("Combo Chart Series Settings is empty");
			}
			C.PlotArea plotArea = new C.PlotArea();
			plotArea.Append(CreateLayout(ComboChartSetting.plotAreaOptions.manualLayout));
			uint chartPosition = 0;
			ComboChartSetting.ComboChartsSettingList.ForEach(chartSetting =>
			{
				if (((ChartSetting<ApplicationSpecificSetting>)chartSetting).isSecondaryAxis)
				{
					isSecondaryAxisActive = true;
					((ChartSetting<ApplicationSpecificSetting>)chartSetting).categoryAxisId = SecondaryCategoryAxisId;
					((ChartSetting<ApplicationSpecificSetting>)chartSetting).valueAxisId = SecondaryValueAxisId;
				}
				((ChartSetting<ApplicationSpecificSetting>)chartSetting).chartDataSetting = new ChartDataSetting();
				AreaChartSetting<ApplicationSpecificSetting> areaChartSetting = chartSetting as AreaChartSetting<ApplicationSpecificSetting>;
				if (areaChartSetting != null)
				{
					AreaChart<ApplicationSpecificSetting> areaChart = new AreaChart<ApplicationSpecificSetting>(areaChartSetting);
					plotArea.Append(areaChart.CreateAreaChart<C.AreaChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
				}
				BarChartSetting<ApplicationSpecificSetting> barChartSetting = chartSetting as BarChartSetting<ApplicationSpecificSetting>;
				if (barChartSetting != null)
				{
					BarChart<ApplicationSpecificSetting> barChart = new BarChart<ApplicationSpecificSetting>(barChartSetting);
					plotArea.Append(barChart.CreateBarChart<C.BarChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
				}
				ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting = chartSetting as ColumnChartSetting<ApplicationSpecificSetting>;
				if (columnChartSetting != null)
				{
					ColumnChart<ApplicationSpecificSetting> columnChart = new ColumnChart<ApplicationSpecificSetting>(columnChartSetting);
					plotArea.Append(columnChart.CreateColumnChart<C.BarChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
				}
				LineChartSetting<ApplicationSpecificSetting> lineChartSetting = chartSetting as LineChartSetting<ApplicationSpecificSetting>;
				if (lineChartSetting != null)
				{
					LineChart<ApplicationSpecificSetting> lineChart = new LineChart<ApplicationSpecificSetting>(lineChartSetting);
					plotArea.Append(lineChart.CreateLineChart(GetChartPositionData(dataCols, chartPosition, dataRange)));
				}
				PieChartSetting<ApplicationSpecificSetting> pieChartSetting = chartSetting as PieChartSetting<ApplicationSpecificSetting>;
				if (pieChartSetting != null)
				{
					PieChart<ApplicationSpecificSetting> pieChart = new PieChart<ApplicationSpecificSetting>(pieChartSetting);
					if (pieChartSetting.pieChartType == PieChartTypes.DOUGHNUT)
					{
						plotArea.Append(pieChart.CreateChart<C.DoughnutChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
					}
					else
					{
						plotArea.Append(pieChart.CreateChart<C.PieChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
					}
				}
				ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting = chartSetting as ScatterChartSetting<ApplicationSpecificSetting>;
				if (scatterChartSetting != null)
				{
					ScatterChart<ApplicationSpecificSetting> scatterChart = new ScatterChart<ApplicationSpecificSetting>(scatterChartSetting);
					if (scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE)
					{
						plotArea.Append(scatterChart.CreateChart<C.BubbleChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
					}
					else
					{
						plotArea.Append(scatterChart.CreateChart<C.ScatterChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
					}
				}
				chartPosition++;
			});
			plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
			{
				id = CategoryAxisId,
				crossAxisId = ValueAxisId,
				fontSize = ComboChartSetting.chartAxesOptions.horizontalFontSize,
				isBold = ComboChartSetting.chartAxesOptions.isHorizontalBold,
				isItalic = ComboChartSetting.chartAxesOptions.isHorizontalItalic,
				isVisible = ComboChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
				invertOrder = ComboChartSetting.chartAxesOptions.invertHorizontalAxesOrder,
			}));
			plotArea.Append(CreateValueAxis(new ValueAxisSetting()
			{
				id = ValueAxisId,
				crossAxisId = CategoryAxisId,
				fontSize = ComboChartSetting.chartAxesOptions.verticalFontSize,
				isBold = ComboChartSetting.chartAxesOptions.isVerticalBold,
				isItalic = ComboChartSetting.chartAxesOptions.isVerticalItalic,
				isVisible = ComboChartSetting.chartAxesOptions.isVerticalAxesEnabled,
				invertOrder = ComboChartSetting.chartAxesOptions.invertVerticalAxesOrder,
			}));
			if (isSecondaryAxisActive)
			{
				plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
				{
					id = SecondaryCategoryAxisId,
					crossAxisId = SecondaryValueAxisId,
					isVisible = false
				}));
				plotArea.Append(CreateValueAxis(new ValueAxisSetting()
				{
					id = SecondaryValueAxisId,
					crossAxisId = SecondaryCategoryAxisId,
					axisPosition = ComboChartSetting.secondaryAxisPosition,
					crosses = C.CrossesValues.Maximum,
					majorTickMark = C.TickMarkValues.Outside
				}));
			}
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}
		private List<ChartDataGrouping> GetChartPositionData(ChartData[][] dataCols, uint chartPosition, DataRange dataRange)
		{
			List<ChartDataGrouping> chartDataGroupings = CreateDataSeries(ComboChartSetting.chartDataSetting, dataCols, dataRange);
			return new List<ChartDataGrouping>() { chartDataGroupings.ElementAt((int)chartPosition) };
		}
	}
}
