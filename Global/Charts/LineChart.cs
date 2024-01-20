// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the settings for a line chart.
    /// </summary>
    public class LineChart : LineFamilyChart
    {
        #region Public Constructors

        /// <summary>
        /// Create Line Chart with provided settings
        /// </summary>
        /// <param name="LineChartSetting">
        /// </param>
        /// <param name="DataCols">
        /// </param>
        public LineChart(LineChartSetting LineChartSetting, ChartData[][] DataCols) : base(LineChartSetting, DataCols) { }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Get Chart Style
        /// </summary>
        /// <returns>
        /// </returns>
        public CS.ChartStyle GetChartStyle()
        {
            return CreateChartStyles();
        }

        /// <summary>
        /// Get Color Style
        /// </summary>
        /// <returns>
        /// </returns>
        public CS.ColorStyle GetColorStyle()
        {
            return CreateColorStyles();
        }

        #endregion Public Methods
    }
}