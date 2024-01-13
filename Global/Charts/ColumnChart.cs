/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the settings for a column chart.
    /// </summary>
    public class ColumnChart : ColumnFamilyChart
    {
        #region Public Constructors
        /// <summary>
        /// Create Column Chart with provided settings
        /// </summary>
        /// <param name="ColumnChartSetting"></param>
        /// <param name="DataCols"></param>
        public ColumnChart(ColumnChartSetting ColumnChartSetting, ChartData[][] DataCols) : base(ColumnChartSetting, DataCols) { }

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