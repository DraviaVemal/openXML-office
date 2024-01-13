/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the settings for a scatter chart.
    /// </summary>
    public class ScatterChart : ScatterFamilyChart
    {
        #region Public Constructors
        /// <summary>
        /// Create Scatter Chart with provided settings
        /// </summary>
        /// <param name="ScatterChartSetting"></param>
        /// <param name="DataCols"></param>
        public ScatterChart(ScatterChartSetting ScatterChartSetting, ChartData[][] DataCols) : base(ScatterChartSetting, DataCols)
        {
        }

        #endregion Public Constructors

        #region Public Methods
        /// <summary>
        /// Get Chart Style
        /// </summary>
        /// <returns></returns>
        public CS.ChartStyle GetChartStyle()
        {
            return CreateChartStyles();
        }
        /// <summary>
        /// Get Color Style
        /// </summary>
        /// <returns></returns>
        public CS.ColorStyle GetColorStyle()
        {
            return CreateColorStyles();
        }

        #endregion Public Methods
    }
}