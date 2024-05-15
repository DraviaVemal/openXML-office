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
			if (CheckSecondaryAxisAlreadyUsed())
			{
				throw new ArgumentException("Secondary Axis is already used in another series");
			}
			ComboChartsSettingList.Add(areaChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public void AddComboChartsSetting(BarChartSetting<ApplicationSpecificSetting> barChartSetting)
		{
			if (CheckSecondaryAxisAlreadyUsed())
			{
				throw new ArgumentException("Secondary Axis is already used in another series");
			}
			ComboChartsSettingList.Add(barChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public void AddComboChartsSetting(ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting)
		{
			if (CheckSecondaryAxisAlreadyUsed())
			{
				throw new ArgumentException("Secondary Axis is already used in another series");
			}
			ComboChartsSettingList.Add(columnChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public void AddComboChartsSetting(LineChartSetting<ApplicationSpecificSetting> lineChartSetting)
		{
			if (CheckSecondaryAxisAlreadyUsed())
			{
				throw new ArgumentException("Secondary Axis is already used in another series");
			}
			ComboChartsSettingList.Add(lineChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public void AddComboChartsSetting(PieChartSetting<ApplicationSpecificSetting> pieChartSetting)
		{
			if (CheckSecondaryAxisAlreadyUsed())
			{
				throw new ArgumentException("Secondary Axis is already used in another series");
			}
			ComboChartsSettingList.Add(pieChartSetting);
		}
		/// <summary>
		/// The options for the axis of the chart.
		/// </summary>
		public ChartAxisOptions<XAxisType, YAxisType, ZAxisType> chartAxisOptions = new ChartAxisOptions<XAxisType, YAxisType, ZAxisType>();
		private bool CheckSecondaryAxisAlreadyUsed()
		{
			return ComboChartsSettingList.Select(val => ((ChartSetting<ApplicationSpecificSetting>)val).isSecondaryAxis).Count(v => v) > 1;
		}
	}
}
