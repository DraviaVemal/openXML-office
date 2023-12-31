namespace OpenXMLOffice.Global
{
    public enum BarChartTypes
    {
        CLUSTERED,
        STACKED,
        PERCENT_STACKED,
        // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D,
    }

    public class BarChartDataLabel
    {
        #region Public Fields

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.NONE;

        #endregion Public Fields

        #region Public Enums

        public enum eDataLabelPosition
        {
            NONE,
            CENTER,
            INSIDE_END,
            INSIDE_BASE,
            OUTSIDE_END,
            // CALLOUT
        }

        #endregion Public Enums
    }

    public class BarChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        public BarChartDataLabel BarChartDataLabel = new();
        public string? BorderColor;
        public string? FillColor;

        #endregion Public Fields
    }

    public class BarChartSetting : ChartSetting
    {
        #region Public Fields

        public List<BarChartSeriesSetting> BarChartSeriesSettings = new();
        public BarChartTypes BarChartTypes = BarChartTypes.CLUSTERED;
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();

        #endregion Public Fields
    }
}