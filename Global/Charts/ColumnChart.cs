// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a column chart.
	/// </summary>
	public class ColumnChart : ColumnFamilyChart
	{        /// <summary>
			 /// Create Column Chart with provided settings
			 /// </summary>
			 /// <param name="ColumnChartSetting">
			 /// </param>
			 /// <param name="DataCols">
			 /// </param>
		public ColumnChart(ColumnChartSetting ColumnChartSetting, ChartData[][] DataCols) : base(ColumnChartSetting, DataCols) { }


		/// <summary>
		/// Get Chart Style
		/// </summary>
		/// <returns>
		/// </returns>
		public static CS.ChartStyle GetChartStyle()
		{
			return CreateChartStyles();
		}

		/// <summary>
		/// Get Color Style
		/// </summary>
		/// <returns>
		/// </returns>
		public static CS.ColorStyle GetColorStyle()
		{
			return CreateColorStyles();
		}


	}
}
