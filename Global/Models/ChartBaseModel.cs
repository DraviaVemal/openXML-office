/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the position of an axis in a chart.
    /// </summary>
    public enum AxisPosition
    {
        /// <summary>
        /// Top
        /// </summary>
        TOP,
        /// <summary>
        /// Bottom
        /// </summary>
        BOTTOM,
        /// <summary>
        /// Left
        /// </summary>
        LEFT,
        /// <summary>
        /// Right
        /// </summary>
        RIGHT
    }

    /// <summary>
    /// Represents the settings for a category axis in a chart.
    /// </summary>
    public class CategoryAxisSetting
    {
        #region Internal Fields

        internal AxisPosition AxisPosition = AxisPosition.BOTTOM;
        internal uint CrossAxisId;
        internal uint Id;

        #endregion Internal Fields
    }

    /// <summary>
    /// Represents the options for the axes in a chart.
    /// </summary>
    public class ChartAxesOptions
    {
        #region Public Fields
        /// <summary>
        /// Is Horizontal Axes Enabled
        /// </summary>
        public bool IsHorizontalAxesEnabled = true;
        /// <summary>
        /// Is Vertical Axes Enabled
        /// </summary>
        public bool IsVerticalAxesEnabled = true;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the options for a chart axis.
    /// </summary>
    public class ChartAxisOptions
    {
        #region Public Fields
        /// <summary>
        /// Horizontal Axis Title
        /// </summary>
        public string? HorizontalAxisTitle;
        /// <summary>
        /// Vertical Axis Title
        /// </summary>
        public string? VerticalAxisTitle;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the grouping options for chart data.
    /// </summary>
    public class ChartDataGrouping
    {
        #region Public Fields

        /// <summary>
        /// Gets or sets the data label cells.
        /// </summary>
        public ChartData[]? DataLabelCells;

        /// <summary>
        /// Gets or sets the data label formula.
        /// </summary>
        public string? DataLabelFormula;

        /// <summary>
        /// Gets or sets the series header cells.
        /// </summary>
        public ChartData? SeriesHeaderCells;

        /// <summary>
        /// Gets or sets the series header formula.
        /// </summary>
        public string? SeriesHeaderFormula;

        /// <summary>
        /// Gets or sets the series header format.
        /// </summary>
        public string? SeriesHeaderFormat;

        /// <summary>
        /// Gets or sets the X-axis cells.
        /// </summary>
        public ChartData[]? XaxisCells;

        /// <summary>
        /// Gets or sets the X-axis formula.
        /// </summary>
        public string? XaxisFormula;

        /// <summary>
        /// Gets or sets the X-axis format.
        /// </summary>
        public string? XaxisFormat;

        /// <summary>
        /// Gets or sets the Y-axis cells.
        /// </summary>
        public ChartData[]? YaxisCells;

        /// <summary>
        /// Gets or sets the Y-axis formula.
        /// </summary>
        public string? YaxisFormula;

        /// <summary>
        /// Gets or sets the Y-axis format.
        /// </summary>
        public string? YaxisFormat;

        /// <summary>
        /// Gets or sets the Z-axis cells.
        /// </summary>
        public ChartData[]? ZaxisCells;

        /// <summary>
        /// Gets or sets the Z-axis formula.
        /// </summary>
        public string? ZaxisFormula;

        /// <summary>
        /// Gets or sets the Z-axis format.
        /// </summary>
        public string? ZaxisFormat;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the options for chart data labels.
    /// </summary>
    public class ChartDataLabel
    {
        #region Public Fields

        /// <summary>
        /// The separator used for displaying multiple values.
        /// </summary>
        public string Separator = ", ";

        /// <summary>
        /// Determines whether to show the category name in the chart.
        /// </summary>
        public bool ShowCategoryName = false;

        /// <summary>
        /// Determines whether to show the legend key in the chart.
        /// </summary>
        public bool ShowLegendKey = false;

        /// <summary>
        /// Determines whether to show the series name in the chart.
        /// </summary>
        public bool ShowSeriesName = false;

        /// <summary>
        /// Determines whether to show the value from a column in the chart.
        /// </summary>
        public bool ShowValueFromColumn = false;

        /// <summary>
        /// Determines whether to show the value in the chart.
        /// </summary>
        public bool ShowValue = false;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the settings for chart data.
    /// </summary>
    public class ChartDataSetting
    {
        #region Public Fields

        /// <summary>
        /// Set 0 To Use Till End
        /// </summary>
        public uint ChartDataColumnEnd = 0;
        /// <summary>
        /// Chart data Start column 0 based
        /// </summary>
        public uint ChartDataColumnStart = 0;

        /// <summary>
        /// Set 0 To Use Till End
        /// </summary>
        public uint ChartDataRowEnd = 0;
        /// <summary>
        /// Chart data Start Row 0 based
        /// </summary>
        public uint ChartDataRowStart = 0;
        /// <summary>
        /// Is Data is used in 3D chart type
        /// </summary>
        public bool Is3Ddata;
        /// <summary>
        /// Key For Data Column Value For Data Label Column If Data Label Column Are Present
        /// Inbetween and Used in the list it will be auto skipped By Data Column
        /// </summary>
        public Dictionary<uint, uint> ValueFromColumn = new();

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the options for chart grid lines.
    /// </summary>
    public class ChartGridLinesOptions
    {
        #region Public Fields
        /// <summary>
        /// Is Major Category Lines Enabled
        /// </summary>
        public bool IsMajorCategoryLinesEnabled = false;
        /// <summary>
        /// Is Major Value Lines Enabled
        /// </summary>
        public bool IsMajorValueLinesEnabled = true;
        /// <summary>
        /// Is Minor Category Lines Enabled
        /// </summary>
        public bool IsMinorCategoryLinesEnabled = false;
        /// <summary>
        /// Is Minor Value Lines Enabled
        /// </summary>
        public bool IsMinorValueLinesEnabled = false;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the options for chart legend.
    /// </summary>
    public class ChartLegendOptions
    {
        #region Public Fields
        /// <summary>
        /// Is Legend Enabled
        /// </summary>
        public bool IsEnableLegend = true;
        /// <summary>
        /// Is Legend Chart OverLap
        /// </summary>
        public bool IsLegendChartOverLap = false;
        /// <summary>
        /// Legend Position
        /// </summary>
        public LegendPositionValues LegendPosition = LegendPositionValues.BOTTOM;

        #endregion Public Fields

        #region Public Enums
        /// <summary>
        /// Legend Position Values
        /// </summary>
        public enum LegendPositionValues
        {
            /// <summary>
            /// Bottom
            /// </summary>
            BOTTOM,
            /// <summary>
            /// Top
            /// </summary>
            TOP,
            /// <summary>
            /// Left
            /// </summary>
            LEFT,
            /// <summary>
            /// Right
            /// </summary>
            RIGHT,
            /// <summary>
            /// Top Right
            /// </summary>
            TOP_RIGHT
        }

        #endregion Public Enums
    }

    /// <summary>
    /// Represents the settings for a chart series.
    /// </summary>
    public class ChartSeriesSetting
    {
        #region Public Fields

        #endregion Public Fields

        #region Internal Constructors

        internal ChartSeriesSetting()
        { }

        #endregion Internal Constructors
    }

    /// <summary>
    /// Represents the settings for a chart.
    /// </summary>
    public class ChartSetting
    {
        #region Public Fields
        /// <summary>
        /// Chart Data Setting
        /// </summary>
        public ChartDataSetting ChartDataSetting = new();
        /// <summary>
        /// Chart Grid Line Options
        /// </summary>
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        /// <summary>
        /// Chart Legend Options
        /// </summary>
        public ChartLegendOptions ChartLegendOptions = new();
        /// <summary>
        /// Chart Height in EMU
        /// </summary>
        public uint Height = 6858000;
        /// <summary>
        /// Chart Title
        /// </summary>
        public string? Title;
        /// <summary>
        /// Chart Width in EMU
        /// </summary>
        public uint Width = 12192000;
        /// <summary>
        /// Chart X Position in EMU
        /// </summary>
        public uint X = 0;
        /// <summary>
        /// Chart Y Position in EMU
        /// </summary>
        public uint Y = 0;

        #endregion Public Fields

        #region Internal Constructors

        internal ChartSetting()
        { }

        #endregion Internal Constructors
    }

    /// <summary>
    /// Represents the settings for a value axis in a chart.
    /// </summary>
    public class ValueAxisSetting
    {
        #region Internal Fields

        internal AxisPosition AxisPosition = AxisPosition.LEFT;
        internal uint CrossAxisId;
        internal uint Id;

        #endregion Internal Fields
    }
}