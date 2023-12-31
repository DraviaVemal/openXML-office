using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class BarChart : BarFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, BarChartSetting BarChartSetting)
        {
            // Apply Properties
            ChartGridLinesOptions = BarChartSetting.ChartGridLinesOptions;
            // Start Creating Objects
            C.Chart Chart = CreateChart(BarChartSetting);
            Chart.PlotArea = CreateChartPlotArea(DataCols, BarChartSetting);
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