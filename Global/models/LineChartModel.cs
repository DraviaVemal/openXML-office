namespace OpenXMLOffice.Global
{
    public class LineChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields
        public string? FillColor;
        public string? BorderColor;
        public LineChartDataLabel LineChartDataLabel = new();
        #endregion Public Fields
    }

    public class LineChartDataLabel
    {
        public enum eDataLabelPosition
        {
            NONE,
            LEFT,
            RIGHT,
            CENTER,
            ABOVE,
            BELOW,
            // CALLOUT
        }

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.NONE;
    }
    
    public class LineChartSetting : ChartSetting
    {
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public List<LineChartSeriesSetting>? SeriesSettings;
    }
}