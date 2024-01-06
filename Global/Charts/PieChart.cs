using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class PieChart : PieFamilyChart
    {
        #region Public Constructors

        public PieChart(PieChartSetting PieChartSetting, ChartData[][] DataCols) : base(PieChartSetting, DataCols)
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