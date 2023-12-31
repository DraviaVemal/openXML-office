namespace OpenXMLOffice.Global
{
    public enum BarChartTypes
    {
        CLUSTERED,
        STACKED,
        PERCENT_STACKED,
        // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D,
    }

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
        public BarChartTypes BarChartTypes = BarChartTypes.CLUSTERED;
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public List<BarChartSeriesSetting> BarChartSeriesSettings = new();
    }
}