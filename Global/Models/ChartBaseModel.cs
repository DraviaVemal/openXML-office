namespace OpenXMLOffice.Global
{
    public enum AxisPosition
    {
        TOP,
        BOTTOM,
        LEFT,
        RIGHT
    }

    public class CategoryAxisSetting
    {
        #region Internal Fields

        internal AxisPosition AxisPosition = AxisPosition.BOTTOM;
        internal uint CrossAxisId;
        internal uint Id;

        #endregion Internal Fields
    }

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

    public class ChartDataGrouping
    {
        #region Public Fields

        public ChartData[]? DataLabelCells;
        public string? DataLabelFormula;
        public ChartData? SeriesHeaderCells;
        public string? SeriesHeaderFormula;
        public ChartData[]? XaxisCells;
        public string? XaxisFormula;
        public ChartData[]? YaxisCells;
        public string? YaxisFormula;

        #endregion Public Fields
    }

    public class ChartDataSetting
    {
        #region Public Fields

        /// <summary>
        /// Set 0 To Use Till End
        /// </summary>
        public uint ChartDataColumnEnd = 0;

        public uint ChartDataColumnStart = 0;

        /// <summary>
        /// Set 0 To Use Till End
        /// </summary>
        public uint ChartDataRowEnd = 0;

        public uint ChartDataRowStart = 0;

        /// <summary>
        /// Key For Data Column Value For Data Label Column If Data Label Column Are Present
        /// Inbetween and Used in the list it will be auto skipped By Data Column
        /// </summary>
        public Dictionary<uint, uint> ValueFromColumn = new();

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

        public ChartDataSetting ChartDataSetting = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public ChartLegendOptions ChartLegendOptions = new();
        public uint Height = 6858000;
        public string? Title;
        public uint Width = 12192000;
        public uint X = 0;
        public uint Y = 0;

        #endregion Public Fields

        #region Internal Constructors

        internal ChartSetting()
        { }

        #endregion Internal Constructors
    }

    public class ValueAxisSetting
    {
        #region Internal Fields

        internal AxisPosition AxisPosition = AxisPosition.LEFT;
        internal uint CrossAxisId;
        internal uint Id;

        #endregion Internal Fields
    }
}