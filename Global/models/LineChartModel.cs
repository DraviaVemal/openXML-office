namespace OpenXMLOffice.Global
{
    public enum LineChartTypes
    {
        CLUSTERED,
        STACKED,
        PERCENT_STACKED,
        CLUSTERED_MARKER,
        STACKED_MARKER,
        PERCENT_STACKED_MARKER,
        // CLUSTERED_3D
    }

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
        public LineChartTypes LineChartTypes = LineChartTypes.CLUSTERED;
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public List<LineChartSeriesSetting> LineChartSeriesSettings = new();
    }
}