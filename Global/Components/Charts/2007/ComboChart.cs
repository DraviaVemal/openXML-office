// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

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
		public ComboChart(ComboChartSetting<ApplicationSpecificSetting> comboChartSetting, ChartData[][] dataCols, DataRange? dataRange = null) : base(comboChartSetting)
		{
			ComboChartSetting = comboChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}

		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange? dataRange)
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
				if (((ChartSetting<ApplicationSpecificSetting>)chartSetting).isSecondaryAxis)
				{
					isSecondaryAxisActive = true;
					((ChartSetting<ApplicationSpecificSetting>)chartSetting).categoryAxisId = SecondaryCategoryAxisId;
					((ChartSetting<ApplicationSpecificSetting>)chartSetting).valueAxisId = SecondaryValueAxisId;
				}
				((ChartSetting<ApplicationSpecificSetting>)chartSetting).chartDataSetting = new();
				if (chartSetting is AreaChartSetting<ApplicationSpecificSetting> areaChartSetting)
				{
					AreaChart<ApplicationSpecificSetting> areaChart = new(areaChartSetting);
					plotArea.Append(areaChart.CreateAreaChart<C.AreaChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
				}
				if (chartSetting is BarChartSetting<ApplicationSpecificSetting> barChartSetting)
				{
					BarChart<ApplicationSpecificSetting> barChart = new(barChartSetting);
					plotArea.Append(barChart.CreateBarChart<C.BarChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
				}
				if (chartSetting is ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting)
				{
					ColumnChart<ApplicationSpecificSetting> columnChart = new(columnChartSetting);
					plotArea.Append(columnChart.CreateColumnChart<C.BarChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
				}
				if (chartSetting is LineChartSetting<ApplicationSpecificSetting> lineChartSetting)
				{
					LineChart<ApplicationSpecificSetting> lineChart = new(lineChartSetting);
					plotArea.Append(lineChart.CreateLineChart(GetChartPositionData(dataCols, chartPosition, dataRange)));
				}
				if (chartSetting is PieChartSetting<ApplicationSpecificSetting> pieChartSetting)
				{
					PieChart<ApplicationSpecificSetting> pieChart = new(pieChartSetting);
					plotArea.Append(pieChartSetting.pieChartType == PieChartTypes.DOUGHNUT ?
						pieChart.CreateChart<C.DoughnutChart>(GetChartPositionData(dataCols, chartPosition, dataRange)) :
						pieChart.CreateChart<C.PieChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
				}
				if (chartSetting is ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting)
				{
					ScatterChart<ApplicationSpecificSetting> scatterChart = new(scatterChartSetting);
					plotArea.Append(scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE ?
						scatterChart.CreateChart<C.BubbleChart>(GetChartPositionData(dataCols, chartPosition, dataRange)) :
						scatterChart.CreateChart<C.ScatterChart>(GetChartPositionData(dataCols, chartPosition, dataRange)));
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

		private List<ChartDataGrouping> GetChartPositionData(ChartData[][] dataCols, uint chartPosition, DataRange? dataRange)
		{
			List<ChartDataGrouping> chartDataGroupings = CreateDataSeries(ComboChartSetting.chartDataSetting, dataCols, dataRange);
			return new() { chartDataGroupings.ElementAt((int)chartPosition) };
		}
	}
}
