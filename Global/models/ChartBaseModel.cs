namespace OpenXMLOffice.Global
{

    public class ChartSeriesSetting
    {
        #region Public Fields

        public string? NumberFormat;
        public string? FillColor;
        public string? BorderColor;

        #endregion Public Fields
    }

    public class ChartLegendOptions
    {
        public enum eLegendPosition
        {
            BOTTOM,
            TOP,
            LEFT,
            RIGHT,
            TOP_RIGHT
        }
        public bool IsEnableLegend = true;
        public eLegendPosition legendPosition = eLegendPosition.BOTTOM;
        public bool IsLegendChartOverLap = false;
    }

    public class ChartAxesOptions
    {
        public bool IsHorizontalAxesEnabled = true;
        public bool IsVerticalAxesEnabled = true;
    }

    public class ChartGridLinesOptions
    {
        public bool IsMajorHorizontalEnabled = true;
        public bool IsMinorHorizontalEnabled = false;
        public bool IsMajorVerticalEnabled = false;
        public bool IsMinorVerticalEnabled = false;
    }

    public class ChartAxisOptions
    {
        public string? HorizontalAxisTitle;
        public string? VerticalAxisTitle;
    }

    public class ChartSetting
    {
        internal ChartSetting() { }
        #region Public Fields
        public string? Title;
        public ChartLegendOptions ChartLegendOptions = new();
        public List<ChartSeriesSetting>? SeriesSettings;

        #endregion Public Fields
    }
}