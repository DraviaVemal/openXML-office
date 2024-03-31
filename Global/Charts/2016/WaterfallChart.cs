// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;

namespace OpenXMLOffice.Global_2016
{
	/// <summary>
	///
	/// </summary>
	public class WaterfallChart<ApplicationSpecificSetting> : AdvanceCharts<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
	{
		private readonly WaterfallChartSetting<ApplicationSpecificSetting> waterfallChartSetting;
		/// <summary>
		///
		/// </summary>
		public WaterfallChart(WaterfallChartSetting<ApplicationSpecificSetting> waterfallChartSetting, ChartData[][] dataColumns, DataRange? dataRange = null) : base(waterfallChartSetting)
		{
			this.waterfallChartSetting = waterfallChartSetting;
			CreateWaterfallChart(dataColumns, dataRange);
		}

		private void CreateWaterfallChart(ChartData[][] dataColumns, DataRange? dataRange)
		{
			GetExtendedChartSpace().Append(CreateChartData(CreateDataSeries(waterfallChartSetting.chartDataSetting, dataColumns, dataRange)));
			GetExtendedChartSpace().Append(CreateChart(CreateDataSeries(waterfallChartSetting.chartDataSetting, dataColumns, dataRange)));
		}


	}
}
