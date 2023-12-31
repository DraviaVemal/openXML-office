namespace OpenXMLOffice.Global
{
    public class ChartAxesOptions
    {
        #region Public Fields

        public bool IsHorizontalAxesEnabled = true;
        public bool IsVerticalAxesEnabled = true;

        #endregion Public Fields
    }

    public class ChartAxisOptions
    {
        #region Public Fields

        public string? HorizontalAxisTitle;
        public string? VerticalAxisTitle;

        #endregion Public Fields
    }

    public class ChartGridLinesOptions
    {
        #region Public Fields

        public bool IsMajorCategoryLinesEnabled = false;
        public bool IsMajorValueLinesEnabled = true;
        public bool IsMinorCategoryLinesEnabled = false;
        public bool IsMinorValueLinesEnabled = false;

        #endregion Public Fields
    }

    public class ChartLegendOptions
    {
        #region Public Fields

        public bool IsEnableLegend = true;

        public bool IsLegendChartOverLap = false;

        public eLegendPosition legendPosition = eLegendPosition.BOTTOM;

        #endregion Public Fields

        #region Public Enums

        public enum eLegendPosition
        {
            BOTTOM,
            TOP,
            LEFT,
            RIGHT,
            TOP_RIGHT
        }

        #endregion Public Enums
    }

    public class ChartSeriesSetting
    {
        #region Public Fields

        public string? NumberFormat;

        #endregion Public Fields

        #region Internal Constructors

        internal ChartSeriesSetting()
        { }

        #endregion Internal Constructors
    }

    public class ChartSetting
    {
        #region Public Fields

        public ChartLegendOptions ChartLegendOptions = new();

        public string? Title;

        #endregion Public Fields

        #region Internal Constructors

        internal ChartSetting()
        { }

        #endregion Internal Constructors
    }
}