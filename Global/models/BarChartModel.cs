namespace OpenXMLOffice.Global
{
    public class BarChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields
        public string? FillColor;
        public string? BorderColor;
        public BarChartDataLabel BarChartDataLabel = new();
        #endregion Public Fields
    }
    
    public class BarChartDataLabel
    {
        public enum eDataLabelPosition
        {
            NONE,
            CENTER,
            INSIDE_END,
            INSIDE_BASE,
            OUTSIDE_END,
            // CALLOUT
        }

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.NONE;
    }

    public class BarChartSetting : ChartSetting
    {
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public List<AreaChartSeriesSetting>? SeriesSettings;
    }
}