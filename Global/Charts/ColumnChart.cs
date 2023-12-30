using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class ColumnChart : BarFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, ChartSetting? chartSetting = null)
        {
            C.Chart Chart = CreateChart();
            Chart.PlotArea = CreateChartPlotArea(DataCols);
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