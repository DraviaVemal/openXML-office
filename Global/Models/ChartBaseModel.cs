// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using C = DocumentFormat.OpenXml.Drawing.Charts;

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
        internal uint id;
        internal AxisPosition axisPosition = AxisPosition.BOTTOM;
        internal uint crossAxisId;
        /// <summary>
        /// Is Font Bold
        /// </summary>
        internal bool isBold = false;
        /// <summary>
        /// Is Font Italic
        /// </summary>
        internal bool isItalic = false;
        /// <summary>
        ///  Font Size
        /// </summary>
        public float fontSize = 11.97F;
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
        public bool isHorizontalAxesEnabled = true;
        /// <summary>
        /// Is Font Bold
        /// </summary>
        public bool isHorizontalBold = false;
        /// <summary>
        /// Is Font Italic
        /// </summary>
        public bool isHorizontalItalic = false;
        /// <summary>
        ///  Font Size
        /// </summary>
        public float horizontalFontSize = 11.97F;
        /// <summary>
        /// Is Font Bold
        /// </summary>
        public bool isVerticalBold = false;
        /// <summary>
        /// Is Font Italic
        /// </summary>
        public bool isVerticalItalic = false;
        /// <summary>
        ///  Font Size
        /// </summary>
        public float verticalFontSize = 11.97F;
        /// <summary>
        /// Is Vertical Axes Enabled
        /// </summary>
        public bool isVerticalAxesEnabled = true;

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
        public string? horizontalAxisTitle;

        /// <summary>
        /// Vertical Axis Title
        /// </summary>
        public string? verticalAxisTitle;

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
        public ChartData[]? dataLabelCells;

        /// <summary>
        /// Gets or sets the data label formula.
        /// </summary>
        public string? dataLabelFormula;

        /// <summary>
        /// Gets or sets the series header cells.
        /// </summary>
        public ChartData? seriesHeaderCells;

        /// <summary>
        /// Gets or sets the series header format.
        /// </summary>
        public string? seriesHeaderFormat;

        /// <summary>
        /// Gets or sets the series header formula.
        /// </summary>
        public string? seriesHeaderFormula;

        /// <summary>
        /// Gets or sets the X-axis cells.
        /// </summary>
        public ChartData[]? xAxisCells;

        /// <summary>
        /// Gets or sets the X-axis format.
        /// </summary>
        public string? xAxisFormat;

        /// <summary>
        /// Gets or sets the X-axis formula.
        /// </summary>
        public string? xAxisFormula;

        /// <summary>
        /// Gets or sets the Y-axis cells.
        /// </summary>
        public ChartData[]? yAxisCells;

        /// <summary>
        /// Gets or sets the Y-axis format.
        /// </summary>
        public string? yAxisFormat;

        /// <summary>
        /// Gets or sets the Y-axis formula.
        /// </summary>
        public string? yAxisFormula;

        /// <summary>
        /// Gets or sets the Z-axis cells.
        /// </summary>
        public ChartData[]? zAxisCells;

        /// <summary>
        /// Gets or sets the Z-axis format.
        /// </summary>
        public string? zAxisFormat;

        /// <summary>
        /// Gets or sets the Z-axis formula.
        /// </summary>
        public string? zAxisFormula;

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
        public string separator = ", ";

        /// <summary>
        /// Determines whether to show the category name in the chart.
        /// </summary>
        public bool showCategoryName = false;

        /// <summary>
        /// Determines whether to show the legend key in the chart.
        /// </summary>
        public bool showLegendKey = false;

        /// <summary>
        /// Determines whether to show the series name in the chart.
        /// </summary>
        public bool showSeriesName = false;

        /// <summary>
        /// Determines whether to show the value in the chart.
        /// </summary>
        public bool showValue = false;

        /// <summary>
        /// Determines whether to show the value from a column in the chart.
        /// </summary>
        public bool showValueFromColumn = false;
        /// <summary>
        /// Is Font Bold
        /// </summary>
        public bool isBold = false;
        /// <summary>
        /// Is Font Italic
        /// </summary>
        internal bool isItalic = false;
        /// <summary>
        /// Font Size
        /// </summary>
        public float fontSize = 11.97F;

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
        public uint chartDataColumnEnd = 0;

        /// <summary>
        /// Chart data Start column 0 based
        /// </summary>
        public uint chartDataColumnStart = 0;

        /// <summary>
        /// Set 0 To Use Till End
        /// </summary>
        public uint chartDataRowEnd = 0;

        /// <summary>
        /// Chart data Start Row 0 based
        /// </summary>
        public uint chartDataRowStart = 0;

        /// <summary>
        /// Is Data is used in 3D chart type
        /// </summary>
        public bool is3Ddata;

        /// <summary>
        /// Key For Data Column Value For Data Label Column If Data Label Column Are Present
        /// Inbetween and Used in the list it will be auto skipped By Data Column
        /// </summary>
        public Dictionary<uint, uint> valueFromColumn = new();

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
        public bool isMajorCategoryLinesEnabled = false;

        /// <summary>
        /// Is Major Value Lines Enabled
        /// </summary>
        public bool isMajorValueLinesEnabled = true;

        /// <summary>
        /// Is Minor Category Lines Enabled
        /// </summary>
        public bool isMinorCategoryLinesEnabled = false;

        /// <summary>
        /// Is Minor Value Lines Enabled
        /// </summary>
        public bool isMinorValueLinesEnabled = false;

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
        public bool isEnableLegend = true;

        /// <summary>
        /// Is Legend Chart OverLap
        /// </summary>
        public bool isLegendChartOverLap = false;
        /// <summary>
        /// Is Font Bold
        /// </summary>
        public bool isBold = false;
        /// <summary>
        /// Is Font Italic
        /// </summary>
        internal bool isItalic = false;
        /// <summary>
        /// Font Size
        /// </summary>
        public float fontSize = 11.97F;

        /// <summary>
        /// Legend Position
        /// </summary>
        public LegendPositionValues legendPosition = LegendPositionValues.BOTTOM;

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
        #region Internal Constructors

        internal ChartSeriesSetting() { }

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
        public ChartDataSetting chartDataSetting = new();

        /// <summary>
        /// Chart Grid Line Options
        /// </summary>
        public ChartGridLinesOptions chartGridLinesOptions = new();

        /// <summary>
        /// Chart Legend Options
        /// </summary>
        public ChartLegendOptions chartLegendOptions = new();

        /// <summary>
        /// Chart Height in EMU
        /// </summary>
        public uint height = 6858000;

        /// <summary>
        /// Chart Title
        /// </summary>
        public string? title;

        /// <summary>
        /// Chart Width in EMU
        /// </summary>
        public uint width = 12192000;

        /// <summary>
        /// Chart X Position in EMU
        /// </summary>
        public uint x = 0;

        /// <summary>
        /// Chart Y Position in EMU
        /// </summary>
        public uint y = 0;

        #endregion Public Fields

        #region Internal Constructors

        internal ChartSetting() { }

        #endregion Internal Constructors
    }

    /// <summary>
    /// Represents the settings for a value axis in a chart.
    /// </summary>
    public class ValueAxisSetting
    {
        #region Internal Fields
        internal uint id;
        internal AxisPosition axisPosition = AxisPosition.LEFT;
        internal uint crossAxisId;
        /// <summary>
        /// Is Font Bold
        /// </summary>
        public bool isBold = false;
        /// <summary>
        /// Is Font Italic
        /// </summary>
        internal bool isItalic = false;
        /// <summary>
        /// Font Size
        /// </summary>
        public float fontSize = 11.97F;
        #endregion Internal Fields
    }

    /// <summary>
    /// 
    /// </summary>
    public class MarkerModel
    {
        /// <summary>
        /// Market Size
        /// </summary>
        public int size = 5;
        /// <summary>
        /// 
        /// </summary>
        public MarkerShapeValues markerShapeValues = MarkerShapeValues.NONE;
        /// <summary>
        /// 
        /// </summary>
        public ShapePropertiesModel shapeProperties = new();
        /// <summary>
        /// 
        /// </summary>
        public enum MarkerShapeValues
        {
            /// <summary>
            /// 
            /// </summary>
            NONE,
            /// <summary>
            /// 
            /// </summary>
            AUTO,
            /// <summary>
            ///
            /// </summary>
            CIRCLE,
            /// <summary>
            ///
            /// </summary>
            DASH,
            /// <summary>
            ///
            /// </summary>
            DIAMOND,
            /// <summary>
            ///
            /// </summary>
            DOT,
            /// <summary>
            ///
            /// </summary>
            PICTURE,
            /// <summary>
            ///
            /// </summary>
            PLUSE,
            /// <summary>
            ///
            /// </summary>
            SQUARE,
            /// <summary>
            ///
            /// </summary>
            STAR,
            /// <summary>
            ///
            /// </summary>
            TRIANGLE,
            /// <summary>
            ///
            /// </summary>
            X
        }

        internal C.MarkerStyleValues GetMarkerStyleValues(MarkerShapeValues markerShapeValues)
        {
            return markerShapeValues switch
            {
                MarkerShapeValues.AUTO => C.MarkerStyleValues.Auto,
                MarkerShapeValues.CIRCLE => C.MarkerStyleValues.Circle,
                MarkerShapeValues.DASH => C.MarkerStyleValues.Dash,
                MarkerShapeValues.DIAMOND => C.MarkerStyleValues.Diamond,
                MarkerShapeValues.DOT => C.MarkerStyleValues.Dot,
                MarkerShapeValues.PICTURE => C.MarkerStyleValues.Picture,
                MarkerShapeValues.PLUSE => C.MarkerStyleValues.Plus,
                MarkerShapeValues.SQUARE => C.MarkerStyleValues.Square,
                MarkerShapeValues.STAR => C.MarkerStyleValues.Star,
                MarkerShapeValues.TRIANGLE => C.MarkerStyleValues.Triangle,
                MarkerShapeValues.X => C.MarkerStyleValues.X,
                _ => C.MarkerStyleValues.None,
            };
        }
    }
}