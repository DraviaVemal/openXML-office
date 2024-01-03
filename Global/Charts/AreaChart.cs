using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class AreaChart : AreaFamilyChart
    {
        #region Public Methods
        public AreaChart(AreaChartSetting AreaChartSetting, ChartData[][] DataCols) : base(AreaChartSetting, DataCols) { }

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