namespace OpenXMLOffice.Global
{
    public enum AreaChartTypes
    {
        CLUSTERED,
        STACKED,
        PERCENT_STACKED,
        // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D
    }

    public class AreaChartDataLabel
    {
        #region Public Fields

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.NONE;

        #endregion Public Fields

        #region Public Enums

        public enum eDataLabelPosition
        {
            NONE,
            SHOW,
            // CALLOUT
        }

        #endregion Public Enums
    }

    public class AreaChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        public AreaChartDataLabel AreaChartDataLabel = new();
        public string? BorderColor;
        public string? FillColor;

        #endregion Public Fields
    }

    public class AreaChartSetting : ChartSetting
    {
        #region Public Fields

        public List<AreaChartSeriesSetting> AreaChartSeriesSettings = new();
        public AreaChartTypes AreaChartTypes = AreaChartTypes.CLUSTERED;
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();

        #endregion Public Fields
    }
}