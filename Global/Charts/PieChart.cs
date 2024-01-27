// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a pie chart.
	/// </summary>
	public class PieChart : PieFamilyChart
	{        /// <summary>
			 /// Create Pie Chart with provided settings
			 /// </summary>
			 /// <param name="PieChartSetting">
			 /// </param>
			 /// <param name="DataCols">
			 /// </param>
		public PieChart(PieChartSetting PieChartSetting, ChartData[][] DataCols) : base(PieChartSetting, DataCols)
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
