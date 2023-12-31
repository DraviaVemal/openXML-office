using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class AreaChart : AreaFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, GlobalConstants.AreaChartTypes AreaChartType, AreaChartSetting chartSetting)
        {
            // Apply Properties
            ChartGridLinesOptions = chartSetting.ChartGridLinesOptions;
            C.Chart Chart = CreateChart(chartSetting);
            Chart.PlotArea = AreaChartType switch
            {
                GlobalConstants.AreaChartTypes.STACKED => CreateChartPlotArea(DataCols, C.GroupingValues.Stacked),
                GlobalConstants.AreaChartTypes.PERCENT_STACKED => CreateChartPlotArea(DataCols, C.GroupingValues.PercentStacked),
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