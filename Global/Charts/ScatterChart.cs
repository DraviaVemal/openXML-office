// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a scatter chart.
	/// </summary>
	public class ScatterChart : ScatterFamilyChart
	{        /// <summary>
			 /// Create Scatter Chart with provided settings
			 /// </summary>
			 /// <param name="ScatterChartSetting">
			 /// </param>
			 /// <param name="DataCols">
			 /// </param>
		public ScatterChart(ScatterChartSetting ScatterChartSetting, ChartData[][] DataCols) : base(ScatterChartSetting, DataCols)
		{
		}


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
