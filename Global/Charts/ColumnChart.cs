using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class ColumnChart : ColumnFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, ColumnChartSetting ColumnChartSetting)
        {
            // Apply Properties
            ChartGridLinesOptions = ColumnChartSetting.ChartGridLinesOptions;
            C.Chart Chart = CreateChart(ColumnChartSetting);
            Chart.PlotArea = CreateChartPlotArea(DataCols, ColumnChartSetting);
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