// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the types of pie charts.
    /// </summary>
    public enum PieChartTypes
    {
        /// <summary>
        /// Pie Chart
        /// </summary>
        PIE,

        // PIE_3D, PIE_PIE, PIE_BAR,
        /// <summary>
        /// Doughnut Chart
        /// </summary>
        DOUGHNUT
    }

    /// <summary>
    /// Represents the data label for a pie chart.
    /// </summary>
    public class PieChartDataLabel : ChartDataLabel
    {
        #region Public Fields

        /// <summary>
        /// The position of the data label.
        /// </summary>
        public DataLabelPositionValues DataLabelPosition = DataLabelPositionValues.CENTER;

        #endregion Public Fields

        #region Public Enums

        /// <summary>
        /// Represents the possible positions of the data label.
        /// </summary>
        public enum DataLabelPositionValues
        {
            /// <summary>
            /// Center
            /// </summary>
            CENTER,

            /// <summary>
            /// Inside End
            /// </summary>
            INSIDE_END,

            /// <summary>
            /// Outside End
            /// </summary>
            OUTSIDE_END,

            /// <summary>
            /// Best Fit
            /// </summary>
            BEST_FIT,

            /// <summary>
            /// Option only for doughnut chart type
            /// </summary>
            SHOW,

            /// <summary>
            /// Data Callout
            /// </summary>
            DATA_CALLOUT
        }

        #endregion Public Enums
    }

    /// <summary>
    /// Represents the series setting for a pie chart.
    /// </summary>
    public class PieChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        /// <summary>
        /// The color of the border.
        /// </summary>
        public string? BorderColor;

        /// <summary>
        /// The color of the fill.
        /// </summary>
        public string? FillColor;

        /// <summary>
        /// Option to customize specific data series, will override chart level setting.
        /// </summary>
        public PieChartDataLabel PieChartDataLabel = new();

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the setting for a pie chart.
    /// </summary>
    public class PieChartSetting : ChartSetting
    {
        #region Public Fields

        /// <summary>
        /// Will get overridden by series level setting.
        /// </summary>
        public PieChartDataLabel PieChartDataLabel = new();

        /// <summary>
        /// The list of series settings for the pie chart.
        /// </summary>
        public List<PieChartSeriesSetting?> PieChartSeriesSettings = new();

        /// <summary>
        /// The type of the pie chart.
        /// </summary>
        public PieChartTypes PieChartTypes = PieChartTypes.PIE;

        #endregion Public Fields
    }
}