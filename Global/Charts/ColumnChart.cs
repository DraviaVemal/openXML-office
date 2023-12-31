using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class ColumnChart : ColumnFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, GlobalConstants.ColumnChartTypes columnChartTypes, ColumnChartSetting chartSetting)
        {
            // Apply Properties
            ChartGridLinesOptions = chartSetting.ChartGridLinesOptions;
            C.Chart Chart = CreateChart(chartSetting);
            Chart.PlotArea = columnChartTypes switch
            {
                GlobalConstants.ColumnChartTypes.STACKED => CreateChartPlotArea(DataCols, C.BarGroupingValues.Stacked, chartSetting),
                GlobalConstants.ColumnChartTypes.PERCENT_STACKED => CreateChartPlotArea(DataCols, C.BarGroupingValues.PercentStacked, chartSetting),
                // Clusted
                _ => CreateChartPlotArea(DataCols, C.BarGroupingValues.Clustered, chartSetting),
            };
            GetChartSpace().Append(Chart);
            return GetChartSpace();
        }

        public CS.ChartStyle GetChartStyle()
        {
            return CreateChartStyles();
        }

        public CS.ColorStyle GetColorStyle()
        {
            return CreateColorStyles();
        }

        #endregion Public Methods
    }
}