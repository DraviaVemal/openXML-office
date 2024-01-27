// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a bar chart.
	/// </summary>
	public class BarChart : BarFamilyChart
	{        /// <summary>
			 /// Create Bar Chart with provided settings
			 /// </summary>
			 /// <param name="BarChartSetting">
			 /// </param>
			 /// <param name="DataCols">
			 /// </param>
		public BarChart(BarChartSetting BarChartSetting, ChartData[][] DataCols) : base(BarChartSetting, DataCols) { }
	}
}
