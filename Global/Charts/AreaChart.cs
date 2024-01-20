// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents an area chart, which is a type of chart that displays data as a series of points
    /// connected by a line and filled with color.
    /// </summary>
    public class AreaChart : AreaFamilyChart
    {
        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="AreaChart"/> class with the specified area
        /// chart settings and data columns.
        /// </summary>
        /// <param name="AreaChartSetting">
        /// The area chart settings.
        /// </param>
        /// <param name="DataCols">
        /// The data columns.
        /// </param>
        public AreaChart(AreaChartSetting AreaChartSetting, ChartData[][] DataCols) : base(AreaChartSetting, DataCols) { }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Gets the chart style for the area chart.
        /// </summary>
        /// <returns>
        /// The chart style.
        /// </returns>
        public CS.ChartStyle GetChartStyle()
        {
            return CreateChartStyles();
        }

        /// <summary>
        /// Gets the color style for the area chart.
        /// </summary>
        /// <returns>
        /// The color style.
        /// </returns>
        public CS.ColorStyle GetColorStyle()
        {
            return CreateColorStyles();
        }

        #endregion Public Methods
    }
}