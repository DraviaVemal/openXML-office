/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    public enum ColumnChartTypes
    {
        CLUSTERED,
        STACKED,
        PERCENT_STACKED,
        // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D, COLUMN_3D
    }

    public class ColumnChartDataLabel : ChartDataLabel
    {
        #region Public Fields

        public DataLabelPositionValues DataLabelPosition = DataLabelPositionValues.CENTER;

        #endregion Public Fields

        #region Public Enums

        public enum DataLabelPositionValues
        {
            CENTER,
            INSIDE_END,
            INSIDE_BASE,

            /// <summary>
            /// This Option is only for Cluster type chart
            /// </summary>
            OUTSIDE_END,

            DATA_CALLOUT
        }

        #endregion Public Enums
    }

    public class ColumnChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields


        /// <summary>
        /// Option To Customise Specific Data Series, Will override Chart Level Setting
        /// </summary>
        public ColumnChartDataLabel ColumnChartDataLabel = new();
        public string? BorderColor;
        public string? FillColor;

        #endregion Public Fields
    }

    public class ColumnChartSetting : ChartSetting
    {
        #region Public Fields

        public ChartAxesOptions ChartAxesOptions = new();

        public ChartAxisOptions ChartAxisOptions = new();

        /// <summary>
        /// Will get override by series level setting
        /// </summary>
        public ColumnChartDataLabel ColumnChartDataLabel = new();

        public List<ColumnChartSeriesSetting?> ColumnChartSeriesSettings = new();
        public ColumnChartTypes ColumnChartTypes = ColumnChartTypes.CLUSTERED;

        #endregion Public Fields
    }
}