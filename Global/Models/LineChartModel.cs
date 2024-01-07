/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    public enum LineChartTypes
    {
        CLUSTERED,
        STACKED,
        PERCENT_STACKED,
        CLUSTERED_MARKER,
        STACKED_MARKER,
        PERCENT_STACKED_MARKER,
        // CLUSTERED_3D
    }

    public class LineChartDataLabel
    {
        #region Public Fields

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.CENTER;
        public bool ShowCategoryName = false;
        public bool ShowLegendKey = false;
        public bool ShowSeriesName = false;
        public bool ShowValue = false;

        #endregion Public Fields

        #region Public Enums

        public enum eDataLabelPosition
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

    public class LineChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        public string? BorderColor;
        public string? FillColor;
        /// <summary>
        /// Option To Customise Specific Data Series, Will override Chart Level Setting
        /// </summary>
        public LineChartDataLabel LineChartDataLabel = new();

        #endregion Public Fields
    }

    public class LineChartSetting : ChartSetting
    {
        #region Public Fields
        /// <summary>
        /// Will get override by series level setting
        /// </summary>
        public LineChartDataLabel LineChartDataLabel = new();
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartAxisOptions ChartAxisOptions = new();
        public List<LineChartSeriesSetting> LineChartSeriesSettings = new();
        public LineChartTypes LineChartTypes = LineChartTypes.CLUSTERED;

        #endregion Public Fields
    }
}