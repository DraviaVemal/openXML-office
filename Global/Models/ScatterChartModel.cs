/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    public enum ScatterChartTypes
    {
        SCATTER,
        SCATTER_SMOOTH,
        SCATTER_SMOOTH_MARKER,
        SCATTER_STRIGHT,
        SCATTER_STRIGHT_MARKER,
        BUBBLE,
        // BUBBLE_3D
    }

    public class ScatterChartDataLabel : ChartDataLabel
    {
        #region Public Fields

        public DataLabelPositionValues DataLabelPosition = DataLabelPositionValues.CENTER;
        public bool ShowBubbleSize = false;

        #endregion Public Fields

        #region Public Enums

        public enum DataLabelPositionValues
        {
            LEFT,
            RIGHT,
            CENTER,
            ABOVE,
            BELOW,
            DATA_CALLOUT
        }

        #endregion Public Enums
    }

    public class ScatterChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        public string? BorderColor;
        public string? FillColor;

        /// <summary>
        /// Option To Customise Specific Data Series, Will override Chart Level Setting
        /// </summary>
        public ScatterChartDataLabel ScatterChartDataLabel = new();

        #endregion Public Fields
    }

    public class ScatterChartSetting : ChartSetting
    {
        #region Public Fields

        public ChartAxesOptions ChartAxesOptions = new();

        public ChartAxisOptions ChartAxisOptions = new();

        /// <summary>
        /// Will get override by series level setting
        /// </summary>
        public ScatterChartDataLabel ScatterChartDataLabel = new();

        public List<ScatterChartSeriesSetting?> ScatterChartSeriesSettings = new();
        public ScatterChartTypes ScatterChartTypes = ScatterChartTypes.SCATTER;

        #endregion Public Fields
    }
}