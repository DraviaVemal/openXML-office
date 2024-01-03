using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class ScatterChart : ScatterFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, ScatterChartSetting ScatterChartSetting)
        {
            // Apply Properties
            ChartGridLinesOptions = ScatterChartSetting.ChartGridLinesOptions;
            C.Chart Chart = CreateChart(ScatterChartSetting);
            Chart.PlotArea = CreateChartPlotArea(DataCols, ScatterChartSetting);
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