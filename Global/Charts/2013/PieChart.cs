// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a pie chart.
	/// </summary>
	public class PieChart : PieFamilyChart
	{
		/// <summary>
		/// Create Pie Chart with provided settings
		/// </summary>
		/// <param name="PieChartSetting">
		/// </param>
		/// <param name="DataCols">
		/// </param>
		public PieChart(PieChartSetting PieChartSetting, ChartData[][] DataCols) : base(PieChartSetting, DataCols) { }
	}
}
