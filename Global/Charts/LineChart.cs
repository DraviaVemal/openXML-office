using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class LineChart : LineFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, LineChartSetting LineChartSetting)
        {
            // Apply Properties
            ChartGridLinesOptions = LineChartSetting.ChartGridLinesOptions;
            C.Chart Chart = CreateChart(LineChartSetting);
            Chart.PlotArea = CreateChartPlotArea(DataCols, LineChartSetting);
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