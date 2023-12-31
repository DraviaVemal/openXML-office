using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class LineChart : LineFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, GlobalConstants.LineChartTypes LineChartType, LineChartSetting chartSetting)
        {
            // Apply Properties
            ChartGridLinesOptions = chartSetting.ChartGridLinesOptions;
            C.Chart Chart = CreateChart(chartSetting);
            Chart.PlotArea = LineChartType switch
            {
                GlobalConstants.LineChartTypes.CLUSTERED_MARKET => CreateChartPlotArea(DataCols, C.GroupingValues.Standard, true),
                GlobalConstants.LineChartTypes.STACKED_MARKET => CreateChartPlotArea(DataCols, C.GroupingValues.Stacked, true),
                GlobalConstants.LineChartTypes.PERCENT_STACKED_MARKET => CreateChartPlotArea(DataCols, C.GroupingValues.PercentStacked, true),
                GlobalConstants.LineChartTypes.STACKED => CreateChartPlotArea(DataCols, C.GroupingValues.Stacked),
                GlobalConstants.LineChartTypes.PERCENT_STACKED => CreateChartPlotArea(DataCols, C.GroupingValues.PercentStacked),
                // Clusted
                _ => CreateChartPlotArea(DataCols, C.GroupingValues.Standard),
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