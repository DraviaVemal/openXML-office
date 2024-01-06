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

    public class ChartDataSetting
    {
        public uint ChartColumnHeader = 1;
        public uint ChartRowHeader = 1;
        public uint ChartDataRowStart = 1;
        /// <summary>
        /// Set 0 To Use Till End
        /// </summary>
        public uint ChartDataRowEnd = 0;
        public uint ChartDataColumnStart = 1;
        /// <summary>
        /// Set 0 To Use Till End
        /// </summary>
        public uint ChartDataColumnEnd = 0;
        /// <summary>
        /// Key For Data Column Value For Data Label Column
        /// If Data Label Column Are Present Inbetween and Used in the list it will be auto skipped By Data Column
        /// </summary>
        public Dictionary<uint, uint> ValueFromColumn = new();

    }

    public class ChartSetting
    {
        #region Public Fields

        public ChartDataSetting ChartDataSetting = new();

        public ChartLegendOptions ChartLegendOptions = new();

        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public string? Title;

        #endregion Public Fields

        #region Internal Constructors

        internal ChartSetting()
        { }

        #endregion Internal Constructors
    }

    public class ChartDataGrouping
    {
        public string? SeriesHeaderFormula;
        public string? XaxisFormula;
        public string? YaxisFormula;
        public string? DataLabelFormula;
        public ChartData[]? SeriesHeaderCells;
        public ChartData[]? XaxisCells;
        public ChartData[]? YaxisCells;
        public ChartData[]? DataLabelCells;
    }
}