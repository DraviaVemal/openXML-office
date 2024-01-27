// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// 
    /// </summary>
    public class ComboChart : ChartBase
    {
        /// <summary>
        /// 
        /// </summary>
        public ComboChartSetting comboChartSetting { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        public ComboChart(ComboChartSetting comboChartSetting, ChartData[][] dataCols) : base(comboChartSetting)
        {
            this.comboChartSetting = comboChartSetting;
            SetChartPlotArea(CreateChartPlotArea(dataCols));
        }

        private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols)
        {
            if (comboChartSetting.ComboChartsSettingList.Count == 0)
            {
                throw new ArgumentException("Combo Chart Series Settings is empty");
            }
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            uint chartPosition = 0;
            comboChartSetting.ComboChartsSettingList.ForEach(chartSetting =>
            {
                ((ChartSetting)chartSetting).chartDataSetting = new();
                if (chartSetting is AreaChartSetting areaChartSetting)
                {
                    AreaFamilyChart areaChart = new(areaChartSetting);
                    plotArea.Append(areaChart.CreateAreaChart(GetChartPositionData(dataCols, chartPosition)));
                }
                if (chartSetting is BarChartSetting barChartSetting)
                {
                    BarFamilyChart barChart = new(barChartSetting);
                    plotArea.Append(barChart.CreateBarChart(GetChartPositionData(dataCols, chartPosition)));
                }
                if (chartSetting is ColumnChartSetting columnChartSetting)
                {
                    ColumnFamilyChart columnChart = new(columnChartSetting);
                    plotArea.Append(columnChart.CreateColumnChart(GetChartPositionData(dataCols, chartPosition)));
                }
                if (chartSetting is LineChartSetting lineChartSetting)
                {
                    LineFamilyChart lineChart = new(lineChartSetting);
                    plotArea.Append(lineChart.CreateLineChart(GetChartPositionData(dataCols, chartPosition)));
                }
                if (chartSetting is PieChartSetting pieChartSetting)
                {
                    PieFamilyChart pieChart = new(pieChartSetting);
                    plotArea.Append(pieChartSetting.pieChartTypes == PieChartTypes.DOUGHNUT ?
                        pieChart.CreateChart<C.DoughnutChart>(GetChartPositionData(dataCols, chartPosition)) :
                        pieChart.CreateChart<C.PieChart>(GetChartPositionData(dataCols, chartPosition)));
                }
                if (chartSetting is ScatterChartSetting scatterChartSetting)
                {
                    ScatterFamilyChart scatterChart = new(scatterChartSetting);
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
                fontSize = comboChartSetting.chartAxesOptions.horizontalFontSize,
                isBold = comboChartSetting.chartAxesOptions.isHorizontalBold,
                isItalic = comboChartSetting.chartAxesOptions.isHorizontalItalic,
                isVisible = comboChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
                invertOrder = comboChartSetting.chartAxesOptions.invertHorizontalAxesOrder,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                id = ValueAxisId,
                crossAxisId = CategoryAxisId,
                fontSize = comboChartSetting.chartAxesOptions.verticalFontSize,
                isBold = comboChartSetting.chartAxesOptions.isVerticalBold,
                isItalic = comboChartSetting.chartAxesOptions.isVerticalItalic,
                isVisible = comboChartSetting.chartAxesOptions.isVerticalAxesEnabled,
                invertOrder = comboChartSetting.chartAxesOptions.invertVerticalAxesOrder,
            }));
            plotArea.Append(CreateChartShapeProperties());
            return plotArea;
        }

        private List<ChartDataGrouping> GetChartPositionData(ChartData[][] dataCols, uint chartPosition)
        {
            List<ChartDataGrouping> chartDataGroupings = CreateDataSeries(dataCols, comboChartSetting.chartDataSetting);
            return new() { chartDataGroupings.ElementAt((int)chartPosition) };
        }



        /// <summary>
        /// Get Chart Style
        /// </summary>
        /// <returns>
        /// </returns>
        public static CS.ChartStyle GetChartStyle()
        {
            return CreateChartStyles();
        }

        /// <summary>
        /// Get Color Style
        /// </summary>
        /// <returns>
        /// </returns>
        public static CS.ColorStyle GetColorStyle()
        {
            return CreateColorStyles();
        }


    }

}