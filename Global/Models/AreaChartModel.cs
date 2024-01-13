/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    public enum AreaChartTypes
    {
        CLUSTERED,
        STACKED,
        PERCENT_STACKED,
        // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D
    }

    public class AreaChartDataLabel : ChartDataLabel
    {
        #region Public Fields

        public DataLabelPositionValues DataLabelPosition = DataLabelPositionValues.SHOW;

        #endregion Public Fields

        #region Public Enums

        public enum DataLabelPositionValues
        {
            SHOW,
            DATA_CALLOUT
        }

        #endregion Public Enums
    }

    public class AreaChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        /// <summary>
        /// Option To Customise Specific Data Series, Will override Chart Level Setting
        /// </summary>
        public AreaChartDataLabel AreaChartDataLabel = new();

        public string? BorderColor;
        public string? FillColor;

        #endregion Public Fields
    }

    public class AreaChartSetting : ChartSetting
    {
        #region Public Fields

        /// <summary>
        /// Will get override by series level setting
        /// </summary>
        public AreaChartDataLabel AreaChartDataLabel = new();

        public List<AreaChartSeriesSetting?> AreaChartSeriesSettings = new();
        public AreaChartTypes AreaChartTypes = AreaChartTypes.CLUSTERED;
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartAxisOptions ChartAxisOptions = new();

        #endregion Public Fields
    }
}