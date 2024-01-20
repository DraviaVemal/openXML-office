// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the types of area charts.
    /// </summary>
    public enum AreaChartTypes
    {
        /// <summary>
        /// Clustered area chart.
        /// </summary>
        CLUSTERED,

        /// <summary>
        /// Stacked area chart.
        /// </summary>
        STACKED,

        /// <summary>
        /// Percent stacked area chart.
        /// </summary>
        PERCENT_STACKED,

        // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D
    }

    /// <summary>
    /// Represents the data label settings for an area chart.
    /// </summary>
    public class AreaChartDataLabel : ChartDataLabel
    {
        #region Public Fields

        /// <summary>
        /// The position of the data labels.
        /// </summary>
        public DataLabelPositionValues dataLabelPosition = DataLabelPositionValues.SHOW;

        #endregion Public Fields

        #region Public Enums

        /// <summary>
        /// The possible values for the data label position.
        /// </summary>
        public enum DataLabelPositionValues
        {
            /// <summary>
            /// Data Label option display type
            /// </summary>
            SHOW,

            /// <summary>
            /// Select Data Callout as Data label style
            /// </summary>
            DATA_CALLOUT
        }

        #endregion Public Enums
    }

    /// <summary>
    /// Represents the series settings for an area chart.
    /// </summary>
    public class AreaChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        /// <summary>
        /// Option to customize specific data series. This will override the chart level setting.
        /// </summary>
        public AreaChartDataLabel areaChartDataLabel = new();

        /// <summary>
        /// The color of the border.
        /// </summary>
        public string? borderColor;

        /// <summary>
        /// The color of the fill.
        /// </summary>
        public string? fillColor;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the settings for an area chart.
    /// </summary>
    public class AreaChartSetting : ChartSetting
    {
        #region Public Fields

        /// <summary>
        /// The data label settings for the entire chart. This will get overridden by series level setting.
        /// </summary>
        public AreaChartDataLabel areaChartDataLabel = new();

        /// <summary>
        /// The series settings for the area chart.
        /// </summary>
        public List<AreaChartSeriesSetting?> areaChartSeriesSettings = new();

        /// <summary>
        /// The type of the area chart.
        /// </summary>
        public AreaChartTypes areaChartTypes = AreaChartTypes.CLUSTERED;

        /// <summary>
        /// The options for the axes of the chart.
        /// </summary>
        public ChartAxesOptions chartAxesOptions = new();

        /// <summary>
        /// The options for the axis of the chart.
        /// </summary>
        public ChartAxisOptions chartAxisOptions = new();

        #endregion Public Fields
    }
}