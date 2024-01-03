using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class ColumnChart : ColumnFamilyChart
    {
        #region Public Methods
        public ColumnChart(ColumnChartSetting ColumnChartSetting, ChartData[][] DataCols) : base(ColumnChartSetting, DataCols) { }

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