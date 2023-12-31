namespace OpenXMLOffice.Global
{
    public enum ColumnChartTypes
    {
        CLUSTERED,
        STACKED,
        PERCENT_STACKED,
        // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D, COLUMN_3D
    }

    public class ColumnChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        public string? FillColor;
        public string? BorderColor;
        public ColumnChartDataLabel ColumnChartDataLabel = new();

        #endregion Public Fields
    }

    public class ColumnChartDataLabel
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

    public class ColumnChartSetting : ChartSetting
    {
        public ColumnChartTypes ColumnChartTypes = ColumnChartTypes.CLUSTERED;
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public List<ColumnChartSeriesSetting> ColumnChartSeriesSettings = new();
    }
}