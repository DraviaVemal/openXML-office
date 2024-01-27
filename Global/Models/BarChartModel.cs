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
        /// Value is used in %.
        /// </summary>
        public int categoryGap = 219;
        /// <summary>
        /// The gap between the series bars.
        /// Value is used in %.
        /// </summary>
        public int seriesGap = -27;
    }
    /// <summary>
    /// Represents the data label settings for a bar chart.
    /// </summary>
    public class BarChartDataLabel : ChartDataLabel
    {        /// <summary>
             /// The position of the data labels.
             /// </summary>
        public DataLabelPositionValues dataLabelPosition = DataLabelPositionValues.CENTER;        /// <summary>
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

            // /// <summary>
            // /// Data Callout
            // /// </summary>
            // DATA_CALLOUT
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public class BarChartDataPointSetting : ChartDataPointSettings
    {

    }

    /// <summary>
    /// Represents the series settings for a bar chart.
    /// </summary>
    public class BarChartSeriesSetting : ChartSeriesSetting
    {        /// <summary>
             /// 
             /// </summary>
        public List<BarChartDataPointSetting?> barChartDataPointSettings = new();

        /// <summary>
        /// Option to customize specific data series. This will override the chart level setting.
        /// </summary>
        public BarChartDataLabel barChartDataLabel = new();

        /// <summary>
        /// The color of the fill.
        /// </summary>
        public string? fillColor;
    }

    /// <summary>
    /// Represents the settings for a bar chart.
    /// </summary>
    public class BarChartSetting : ChartSetting
    {        /// <summary>
             /// The data label settings for the entire chart. This will get overridden by series level setting.
             /// </summary>
        public BarChartDataLabel barChartDataLabel = new();

        /// <summary>
        /// The series settings for the bar chart.
        /// </summary>
        public List<BarChartSeriesSetting?> barChartSeriesSettings = new();

        /// <summary>
        /// The type of bar chart.
        /// </summary>
        public BarChartTypes barChartTypes = BarChartTypes.CLUSTERED;

        /// <summary>
        /// The options for the chart axes.
        /// </summary>
        public ChartAxesOptions chartAxesOptions = new();

        /// <summary>
        /// The options for the chart axis.
        /// </summary>
        public ChartAxisOptions chartAxisOptions = new();
        /// <summary>
        /// The graphics settings for the bar chart.
        /// </summary>
        public BarGraphicsSetting barGraphicsSetting = new();
    }
}