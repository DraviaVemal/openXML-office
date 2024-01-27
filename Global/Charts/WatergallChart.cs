// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Data;

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
		public WaterfallChart(WaterfallChartSetting waterfallChartSetting, DataColumn[][] dataColumns) : base(waterfallChartSetting)
		{
			this.waterfallChartSetting = waterfallChartSetting;
		}

	}
}
