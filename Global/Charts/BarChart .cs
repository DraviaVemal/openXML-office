using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class BarChart : BarFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, GlobalConstants.BarChartTypes barChartType, ChartSetting? chartSetting = null)
        {
            C.Chart Chart = CreateChart();
            Chart.PlotArea = barChartType switch
            {
                GlobalConstants.BarChartTypes.STACKED => CreateChartPlotArea(DataCols, C.BarDirectionValues.Bar, C.BarGroupingValues.Stacked),
                GlobalConstants.BarChartTypes.CENT_STACKED => CreateChartPlotArea(DataCols, C.BarDirectionValues.Bar, C.BarGroupingValues.PercentStacked),
                // Clusted
                _ => CreateChartPlotArea(DataCols, C.BarDirectionValues.Bar, C.BarGroupingValues.Clustered),
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