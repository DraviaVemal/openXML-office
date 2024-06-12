// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.Collections.Generic;
using System.Linq;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	///
	/// </summary>
	public class ComboChartSetting<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> : ChartSetting<ApplicationSpecificSetting>
		where ApplicationSpecificSetting : class, ISizeAndPosition, new()
		where XAxisType : class, IAxisTypeOptions, new()
	 	where YAxisType : class, IAxisTypeOptions, new()
	  	where ZAxisType : class, IAxisTypeOptions, new()
	{
		/// <summary>
		/// Secondary Axis position
		/// </summary>
		public AxisPosition secondaryAxisPosition = AxisPosition.RIGHT;
		/// <summary>
		/// Add Chart Series Setting Using AddComboChartsSetting Method
		/// </summary>
		public List<object> ComboChartsSettingList = new List<object>();
		/// <summary>
		///
		/// </summary>
		public void AddComboChartsSetting(AreaChartSetting<ApplicationSpecificSetting> areaChartSetting)
		{
			ComboChartsSettingList.Add(areaChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public void AddComboChartsSetting(BarChartSetting<ApplicationSpecificSetting> barChartSetting)
		{
			ComboChartsSettingList.Add(barChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public void AddComboChartsSetting(ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting)
		{
			ComboChartsSettingList.Add(columnChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public void AddComboChartsSetting(LineChartSetting<ApplicationSpecificSetting> lineChartSetting)
		{
			ComboChartsSettingList.Add(lineChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public void AddComboChartsSetting(PieChartSetting<ApplicationSpecificSetting> pieChartSetting)
		{
			ComboChartsSettingList.Add(pieChartSetting);
		}
		/// <summary>
		/// The options for the axis of the chart.
		/// </summary>
		public ChartAxisOptions<XAxisType, YAxisType, ZAxisType> chartAxisOptions = new ChartAxisOptions<XAxisType, YAxisType, ZAxisType>();
	}
}
