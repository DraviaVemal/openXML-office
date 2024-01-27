// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a line chart.
	/// </summary>
	public class LineChart : LineFamilyChart
	{

		/// <summary>
		/// Create Line Chart with provided settings
		/// </summary>
		public LineChart(LineChartSetting LineChartSetting, ChartData[][] DataCols) : base(LineChartSetting, DataCols) { }
	}
}
