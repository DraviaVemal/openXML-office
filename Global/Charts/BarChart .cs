using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class BarChart : BarFamilyChart
    {
        #region Public Methods
        public BarChart(BarChartSetting BarChartSetting, ChartData[][] DataCols) : base(BarChartSetting, DataCols) { }

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