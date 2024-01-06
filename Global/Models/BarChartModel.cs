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

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.CENTER;
        public bool ShowValue = false;
        public bool ShowLegendKey = false;
        public bool ShowCategoryName = false;
        public bool ShowSeriesName = false;

        #endregion Public Fields

        #region Public Enums

        public enum eDataLabelPosition
        {
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

        #endregion Public Fields
    }
}