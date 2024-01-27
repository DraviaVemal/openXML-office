// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a column chart.
	/// </summary>
	public class ColumnChart : ColumnFamilyChart
	{

		/// <summary>
		/// Create Column Chart with provided settings
		/// </summary>
		/// <param name="ColumnChartSetting">
		/// </param>
		/// <param name="DataCols">
		/// </param>
		public ColumnChart(ColumnChartSetting ColumnChartSetting, ChartData[][] DataCols) : base(ColumnChartSetting, DataCols) { }
	}
}
