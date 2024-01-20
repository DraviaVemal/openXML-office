// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the types of bar charts.
    /// </summary>
    public enum BarChartTypes
    {
        /// <summary>
        /// Clustered Bar Chart
        /// </summary>
        CLUSTERED,

        /// <summary>
        /// Stacked Bar Chart
        /// </summary>
        STACKED,

        /// <summary>
        /// Percent Stacked Bar Chart
        /// </summary>
        PERCENT_STACKED,

        // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D
    }
    /// <summary>
    /// Represents the graphics settings for a bar chart.
    /// </summary>
    public class BarGraphicsSetting
    {
        /// <summary>
        /// The gap width between the bars.
        /// </summary>
        public int CategoryGap = 219;
        /// <summary>
        /// The gap between the series bars.
        /// </summary>
        public int SeriesGap = -27;
    }
    /// <summary>
    /// Represents the data label settings for a bar chart.
    /// </summary>
    public class BarChartDataLabel : ChartDataLabel
    {
        #region Public Fields

        /// <summary>
        /// The position of the data labels.
        /// </summary>
        public DataLabelPositionValues DataLabelPosition = DataLabelPositionValues.CENTER;

        #endregion Public Fields

        #region Public Enums

        /// <summary>
        /// The possible positions for the data labels.
        /// </summary>
        public enum DataLabelPositionValues
        {
            /// <summary>
            /// Center
            /// </summary>
            CENTER,

            /// <summary>
            /// Inside end
            /// </summary>
            INSIDE_END,

            /// <summary>
            /// Inside base
            /// </summary>
            INSIDE_BASE,

            /// <summary>
            /// This option is only for Cluster type chart.
            /// </summary>
            OUTSIDE_END,

            /// <summary>
            /// Data Callout
            /// </summary>
            DATA_CALLOUT
        }

        #endregion Public Enums
    }

    /// <summary>
    /// Represents the series settings for a bar chart.
    /// </summary>
    public class BarChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        /// <summary>
        /// Option to customize specific data series. This will override the chart level setting.
        /// </summary>
        public BarChartDataLabel BarChartDataLabel = new();

        /// <summary>
        /// The color of the border.
        /// </summary>
        public string? BorderColor;

        /// <summary>
        /// The color of the fill.
        /// </summary>
        public string? FillColor;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the settings for a bar chart.
    /// </summary>
    public class BarChartSetting : ChartSetting
    {
        #region Public Fields

        /// <summary>
        /// The data label settings for the entire chart. This will get overridden by series level setting.
        /// </summary>
        public BarChartDataLabel BarChartDataLabel = new();

        /// <summary>
        /// The series settings for the bar chart.
        /// </summary>
        public List<BarChartSeriesSetting?> BarChartSeriesSettings = new();

        /// <summary>
        /// The type of bar chart.
        /// </summary>
        public BarChartTypes BarChartTypes = BarChartTypes.CLUSTERED;

        /// <summary>
        /// The options for the chart axes.
        /// </summary>
        public ChartAxesOptions ChartAxesOptions = new();

        /// <summary>
        /// The options for the chart axis.
        /// </summary>
        public ChartAxisOptions ChartAxisOptions = new();
        /// <summary>
        /// The graphics settings for the bar chart.
        /// </summary>
        public BarGraphicsSetting BarGraphicsSetting = new();

        #endregion Public Fields
    }
}