/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the types of column charts.
    /// </summary>
    public enum ColumnChartTypes
    {
        /// <summary>
        /// Clustered Column Chart
        /// </summary>
        CLUSTERED,

        /// <summary>
        /// Stacked Column Chart
        /// </summary>
        STACKED,

        /// <summary>
        /// Percent Stacked Column Chart
        /// </summary>
        PERCENT_STACKED,

        // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D, COLUMN_3D
    }

    /// <summary>
    /// Represents the graphics settings for a column chart.
    /// </summary>
    public class ColumnGraphicsSetting
    {
        /// <summary>
        /// The gap width between the Column.
        /// </summary>
        public int CategoryGap = 219;
        /// <summary>
        /// The gap between the series column.
        /// </summary>
        public int SeriesGap = -27;
    }

    /// <summary>
    /// Represents the data label settings for a column chart.
    /// </summary>
    public class ColumnChartDataLabel : ChartDataLabel
    {
        #region Public Fields

        /// <summary>
        /// The position of the data label.
        /// </summary>
        public DataLabelPositionValues DataLabelPosition = DataLabelPositionValues.CENTER;

        #endregion Public Fields

        #region Public Enums

        /// <summary>
        /// The possible positions for the data label.
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
            /// Inside Base
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
    /// Represents the series settings for a column chart.
    /// </summary>
    public class ColumnChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        /// <summary>
        /// Chart Stick Border Color
        /// </summary>
        public string? BorderColor;

        /// <summary>
        /// Option to customize specific data series. Will override chart level setting.
        /// </summary>
        public ColumnChartDataLabel ColumnChartDataLabel = new();

        /// <summary>
        /// Chart Stick Fill Color
        /// </summary>
        public string? FillColor;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the settings for a column chart.
    /// </summary>
    public class ColumnChartSetting : ChartSetting
    {
        #region Public Fields

        /// <summary>
        /// Chart Axes Options
        /// </summary>
        public ChartAxesOptions ChartAxesOptions = new();

        /// <summary>
        /// Chart Axis Options
        /// </summary>
        public ChartAxisOptions ChartAxisOptions = new();

        /// <summary>
        /// Will get overridden by series level setting.
        /// </summary>
        public ColumnChartDataLabel ColumnChartDataLabel = new();

        /// <summary>
        /// Chart Series Settings
        /// </summary>
        public List<ColumnChartSeriesSetting?> ColumnChartSeriesSettings = new();

        /// <summary>
        /// Chart Type. default is CLUSTERED
        /// </summary>
        public ColumnChartTypes ColumnChartTypes = ColumnChartTypes.CLUSTERED;
        /// <summary>
        /// The graphics settings for the column chart.
        /// </summary>
        public ColumnGraphicsSetting ColumnGraphicsSetting = new();

        #endregion Public Fields
    }
}