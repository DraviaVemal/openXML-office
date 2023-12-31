using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class PieChart : PieFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, PieChartSetting PieChartSetting)
        {
            C.Chart Chart = CreateChart(PieChartSetting);
            Chart.PlotArea = PieChartSetting.PieChartTypes switch
            {
                PieChartTypes.DOUGHNUT => CreateDoughnutChartPlotArea(DataCols, PieChartSetting),
                // Pie
                _ => CreatePieChartPlotArea(DataCols, PieChartSetting),
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