using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class ScatterChart : ScatterFamilyChart
    {
        #region Public Constructors

        public ScatterChart(ScatterChartSetting ScatterChartSetting, ChartData[][] DataCols) : base(ScatterChartSetting, DataCols)
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