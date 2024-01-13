/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    public enum PieChartTypes
    {
        PIE,

        // PIE_3D, PIE_PIE, PIE_BAR,
        DOUGHNUT
    }

    public class PieChartDataLabel : ChartDataLabel
    {
        #region Public Fields

        public DataLabelPositionValues DataLabelPosition = DataLabelPositionValues.CENTER;

        #endregion Public Fields

        #region Public Enums

        public enum DataLabelPositionValues
        {
            CENTER,
            INSIDE_END,
            OUTSIDE_END,
            BEST_FIT,

            /// <summary>
            /// Option only for doughnut chart type
            /// </summary>
            SHOW,

            DATA_CALLOUT
        }

        #endregion Public Enums
    }

    public class PieChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        public string? BorderColor;
        public string? FillColor;

        /// <summary>
        /// Option To Customise Specific Data Series, Will override Chart Level Setting
        /// </summary>
        public PieChartDataLabel PieChartDataLabel = new();

        #endregion Public Fields
    }

    public class PieChartSetting : ChartSetting
    {
        #region Public Fields

        /// <summary>
        /// Will get override by series level setting
        /// </summary>
        public PieChartDataLabel PieChartDataLabel = new();

        public List<PieChartSeriesSetting?> PieChartSeriesSettings = new();
        public PieChartTypes PieChartTypes = PieChartTypes.PIE;

        #endregion Public Fields
    }
}