namespace OpenXMLOffice.Global
{
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
        public bool IsMajorCategoryLinesEnabled = false;
        public bool IsMinorCategoryLinesEnabled = false;
        public bool IsMajorValueLinesEnabled = true;
        public bool IsMinorValueLinesEnabled = false;
    }

    public class ChartAxisOptions
    {
        public string? HorizontalAxisTitle;
        public string? VerticalAxisTitle;
    }

    public class ChartSetting
    {
        internal ChartSetting()
        { }

        #region Public Fields

        public string? Title;
        public ChartLegendOptions ChartLegendOptions = new();

        #endregion Public Fields
    }
}