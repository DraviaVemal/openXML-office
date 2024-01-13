/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the types of scatter charts.
    /// </summary>
    public enum ScatterChartTypes
    {
        /// <summary>
        /// Scatter Chart
        /// </summary>
        SCATTER,
        /// <summary>
        /// Scatter Chart with Smooth Lines
        /// </summary>
        SCATTER_SMOOTH,
        /// <summary>
        /// Scatter Chart with Smooth Lines and Markers
        /// </summary>
        SCATTER_SMOOTH_MARKER,
        /// <summary>
        /// Scatter Chart with Straight Lines
        /// </summary>
        SCATTER_STRIGHT,
        /// <summary>
        /// Scatter Chart with Straight Lines and Markers
        /// </summary>
        SCATTER_STRIGHT_MARKER,
        /// <summary>
        /// Bubble Chart
        /// </summary>
        BUBBLE,
        // BUBBLE_3D
    }

    /// <summary>
    /// Represents the data label settings for a scatter chart.
    /// </summary>
    public class ScatterChartDataLabel : ChartDataLabel
    {
        #region Public Fields

        /// <summary>
        /// The position of the data labels.
        /// </summary>
        public DataLabelPositionValues DataLabelPosition = DataLabelPositionValues.CENTER;

        /// <summary>
        /// Determines whether to show the bubble size in the data labels.
        /// </summary>
        public bool ShowBubbleSize = false;

        #endregion Public Fields

        #region Public Enums

        /// <summary>
        /// The possible positions for the data labels.
        /// </summary>
        public enum DataLabelPositionValues
        {
            /// <summary>
            /// Left Side
            /// </summary>
            LEFT,
            /// <summary>
            /// Right Side
            /// </summary>
            RIGHT,
            /// <summary>
            /// Center Placement
            /// </summary>
            CENTER,
            /// <summary>
            /// Above content
            /// </summary>
            ABOVE,
            /// <summary>
            /// Below content
            /// </summary>
            BELOW,
            /// <summary>
            /// Data Callout Style
            /// </summary>
            DATA_CALLOUT
        }

        #endregion Public Enums
    }

    /// <summary>
    /// Represents the series settings for a scatter chart.
    /// </summary>
    public class ScatterChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        /// <summary>
        /// The color of the series border.
        /// </summary>
        public string? BorderColor;

        /// <summary>
        /// The color of the series fill.
        /// </summary>
        public string? FillColor;

        /// <summary>
        /// Custom data label settings for the specific data series.
        /// This will override the chart level setting.
        /// </summary>
        public ScatterChartDataLabel ScatterChartDataLabel = new();

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the settings for a scatter chart.
    /// </summary>
    public class ScatterChartSetting : ChartSetting
    {
        #region Public Fields

        /// <summary>
        /// The options for the chart axes.
        /// </summary>
        public ChartAxesOptions ChartAxesOptions = new();

        /// <summary>
        /// The options for the chart axis.
        /// </summary>
        public ChartAxisOptions ChartAxisOptions = new();

        /// <summary>
        /// The data label settings for the scatter chart.
        /// This will get overridden by the series level setting.
        /// </summary>
        public ScatterChartDataLabel ScatterChartDataLabel = new();

        /// <summary>
        /// The list of series settings for the scatter chart.
        /// </summary>
        public List<ScatterChartSeriesSetting?> ScatterChartSeriesSettings = new();

        /// <summary>
        /// The type of scatter chart.
        /// </summary>
        public ScatterChartTypes ScatterChartTypes = ScatterChartTypes.SCATTER;

        #endregion Public Fields
    }
}