using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class AreaChart : AreaFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, AreaChartSetting AreaChartSetting)
        {
            // Apply Properties
            ChartGridLinesOptions = AreaChartSetting.ChartGridLinesOptions;
            C.Chart Chart = CreateChart(AreaChartSetting);
            Chart.PlotArea = CreateChartPlotArea(DataCols, AreaChartSetting);
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