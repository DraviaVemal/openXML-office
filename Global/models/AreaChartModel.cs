namespace OpenXMLOffice.Global
{
    public enum AreaChartTypes
    {
        CLUSTERED,
        STACKED,
        PERCENT_STACKED,
        // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D
    }
    public class AreaChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields
        public string? FillColor;
        public string? BorderColor;
        public AreaChartDataLabel AreaChartDataLabel = new();
        #endregion Public Fields
    }

    public class AreaChartDataLabel
    {
        public enum eDataLabelPosition
        {
            NONE,
            SHOW,
            // CALLOUT
        }

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.NONE;
    }

    public class AreaChartSetting : ChartSetting
    {
        public AreaChartTypes AreaChartTypes = AreaChartTypes.CLUSTERED;
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public List<AreaChartSeriesSetting> AreaChartSeriesSettings = new();
    }
}