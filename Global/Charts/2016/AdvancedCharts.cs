// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OpenXMLOffice.Global_2016
{
	/// <summary>
	///
	/// </summary>
	public class AdvanceCharts : ChartBase
	{
		/// <summary>
		///
		/// </summary>
		/// <param name="chartSetting"></param>
		protected AdvanceCharts(ChartSetting chartSetting) : base(chartSetting) { }

		/// <summary>
		///
		/// </summary>
		/// <returns></returns>
		public CX.ChartSpace GetExtendedChartSpace()
		{
			return new();
		}
	}
}
