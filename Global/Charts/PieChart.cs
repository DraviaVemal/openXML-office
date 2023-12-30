using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class PieChart : PieFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, GlobalConstants.PieChartTypes PieChartType, ChartSetting? chartSetting = null)
        {
            C.Chart Chart = CreateChart();
            Chart.PlotArea = PieChartType switch
            {
                GlobalConstants.PieChartTypes.DOUGHNUT => CreateDoughnutChartPlotArea(DataCols),
                // Pie
                _ => CreatePieChartPlotArea(DataCols),
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