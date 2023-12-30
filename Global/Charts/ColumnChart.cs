using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class ColumnChart : BarFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, GlobalConstants.ColumnChartTypes columnChartTypes, ColumnChartSetting chartSetting)
        {
            C.Chart Chart = CreateChart(chartSetting);
            Chart.PlotArea = columnChartTypes switch
            {
                GlobalConstants.ColumnChartTypes.STACKED => CreateChartPlotArea(DataCols, C.BarDirectionValues.Column, C.BarGroupingValues.Stacked),
                GlobalConstants.ColumnChartTypes.PERCENT_STACKED => CreateChartPlotArea(DataCols, C.BarDirectionValues.Column, C.BarGroupingValues.PercentStacked),
                // Clusted
                _ => CreateChartPlotArea(DataCols, C.BarDirectionValues.Column, C.BarGroupingValues.Clustered),
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