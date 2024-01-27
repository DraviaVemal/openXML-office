// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;

namespace OpenXMLOffice.Global_2016
{
	/// <summary>
	///
	/// </summary>
	public class WaterfallChart : AdvanceCharts
	{
		private WaterfallChartSetting waterfallChartSetting;
		/// <summary>
		///
		/// </summary>
		public WaterfallChart(WaterfallChartSetting waterfallChartSetting, ChartData[][] dataColumns) : base(waterfallChartSetting)
		{
			this.waterfallChartSetting = waterfallChartSetting;
		}

	}
}
