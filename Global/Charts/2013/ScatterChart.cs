// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.


namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a scatter chart.
	/// </summary>
	public class ScatterChart : ScatterFamilyChart
	{
		/// <summary>
		/// Create Scatter Chart with provided settings
		/// </summary>
		/// <param name="ScatterChartSetting">
		/// </param>
		/// <param name="DataCols">
		/// </param>
		public ScatterChart(ScatterChartSetting ScatterChartSetting, ChartData[][] DataCols) : base(ScatterChartSetting, DataCols) { }
	}
}
