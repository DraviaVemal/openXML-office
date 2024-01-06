using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class LineChart : LineFamilyChart
    {
        #region Public Constructors

        public LineChart(LineChartSetting LineChartSetting, ChartData[][] DataCols) : base(LineChartSetting, DataCols)
        {
        }

        #endregion Public Constructors

        #region Public Methods

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