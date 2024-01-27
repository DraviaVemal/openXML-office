// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents an area chart, which is a type of chart that displays data as a series of points
	/// connected by a line and filled with color.
	/// </summary>
	public class AreaChart : AreaFamilyChart
	{        /// <summary>
			 /// Initializes a new instance of the <see cref="AreaChart"/> class with the specified area
			 /// chart settings and data columns.
			 /// </summary>
			 /// <param name="areaChartSetting">
			 /// The area chart settings.
			 /// </param>
			 /// <param name="dataCols">
			 /// The data columns.
			 /// </param>
		public AreaChart(AreaChartSetting areaChartSetting, ChartData[][] dataCols) : base(areaChartSetting, dataCols) { }
	}
}
