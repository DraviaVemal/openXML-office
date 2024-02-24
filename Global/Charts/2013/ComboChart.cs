// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	///
	/// </summary>
	public class ComboChart : ChartBase
	{
		/// <summary>
		///
		/// </summary>
		public ComboChartSetting ComboChartSetting { get; private set; }

		/// <summary>
		///
		/// </summary>
		public ComboChart(ComboChartSetting comboChartSetting, ChartData[][] dataCols) : base(comboChartSetting)
		{
			ComboChartSetting = comboChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols));
		}

		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols)
		{
			bool isSecondaryAxisActive = false;
			if (ComboChartSetting.ComboChartsSettingList.Count == 0)
			{
				throw new ArgumentException("Combo Chart Series Settings is empty");
			}
			C.PlotArea plotArea = new();
			plotArea.Append(CreateLayout(ComboChartSetting.plotAreaOptions?.manualLayout));
			uint chartPosition = 0;
			ComboChartSetting.ComboChartsSettingList.ForEach(chartSetting =>
			{
				if (((ChartSetting)chartSetting).isSecondaryAxis)
				{
					isSecondaryAxisActive = true;
					((ChartSetting)chartSetting).categoryAxisId = SecondaryCategoryAxisId;
					((ChartSetting)chartSetting).valueAxisId = SecondaryValueAxisId;
				}
				((ChartSetting)chartSetting).chartDataSetting = new();
				if (chartSetting is AreaChartSetting areaChartSetting)
				{
					AreaChart areaChart = new(areaChartSetting);
					plotArea.Append(areaChart.CreateAreaChart(GetChartPositionData(dataCols, chartPosition)));
				}
				if (chartSetting is BarChartSetting barChartSetting)
				{
					BarChart barChart = new(barChartSetting);
					plotArea.Append(barChart.CreateBarChart(GetChartPositionData(dataCols, chartPosition)));
				}
				if (chartSetting is ColumnChartSetting columnChartSetting)
				{
					ColumnChart columnChart = new(columnChartSetting);
					plotArea.Append(columnChart.CreateColumnChart(GetChartPositionData(dataCols, chartPosition)));
				}
				if (chartSetting is LineChartSetting lineChartSetting)
				{
					LineChart lineChart = new(lineChartSetting);
					plotArea.Append(lineChart.CreateLineChart(GetChartPositionData(dataCols, chartPosition)));
				}
				if (chartSetting is PieChartSetting pieChartSetting)
				{
					PieChart pieChart = new(pieChartSetting);
					plotArea.Append(pieChartSetting.pieChartTypes == PieChartTypes.DOUGHNUT ?
						pieChart.CreateChart<C.DoughnutChart>(GetChartPositionData(dataCols, chartPosition)) :
						pieChart.CreateChart<C.PieChart>(GetChartPositionData(dataCols, chartPosition)));
				}
				if (chartSetting is ScatterChartSetting scatterChartSetting)
				{
					ScatterChart scatterChart = new(scatterChartSetting);
					plotArea.Append(scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE ?
						scatterChart.CreateChart<C.BubbleChart>(GetChartPositionData(dataCols, chartPosition)) :
						scatterChart.CreateChart<C.ScatterChart>(GetChartPositionData(dataCols, chartPosition)));
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

		private List<ChartDataGrouping> GetChartPositionData(ChartData[][] dataCols, uint chartPosition)
		{
			List<ChartDataGrouping> chartDataGroupings = CreateDataSeries(dataCols, ComboChartSetting.chartDataSetting);
			return new() { chartDataGroupings.ElementAt((int)chartPosition) };
		}
	}
}
