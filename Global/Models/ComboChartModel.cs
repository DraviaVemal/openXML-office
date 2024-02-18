// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	///
	/// </summary>
	public class ComboChartSetting : ChartSetting
	{
		/// <summary>
		/// Add Chart Series Setting Using AddComboChartsSetting Method
		/// </summary>
		public List<object> ComboChartsSettingList { get; private set; } = new();
		/// <summary>
		///
		/// </summary>
		public void AddComboChartsSetting(AreaChartSetting areaChartSetting)
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
		public void AddComboChartsSetting(BarChartSetting barChartSetting)
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
		public void AddComboChartsSetting(ColumnChartSetting columnChartSetting)
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
		public void AddComboChartsSetting(LineChartSetting lineChartSetting)
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
		public void AddComboChartsSetting(PieChartSetting pieChartSetting)
		{
			if (CheckSecondaryAxisAlreadyUsed())
			{
				throw new ArgumentException("Secondary Axis is already used in another series");
			}
			ComboChartsSettingList.Add(pieChartSetting);
		}

		// /// <summary>
		// ///
		// /// </summary>
		// public void AddComboChartsSetting(ScatterChartSetting scatterChartSetting)
		// {
		// 	if (checkSecondaryAxisAlreadyUsed())
		// 	{
		// 		throw new ArgumentException("Secondary Axis is already used in another series");
		// 	}
		// 	ComboChartsSettingList.Add(scatterChartSetting);
		// }

		/// <summary>
		/// The options for the axes of the chart.
		/// </summary>
		public ChartAxesOptions chartAxesOptions = new();

		/// <summary>
		/// The options for the axis of the chart.
		/// </summary>
		public ChartAxisOptions chartAxisOptions = new();

		private bool CheckSecondaryAxisAlreadyUsed()
		{
			return ComboChartsSettingList.Select(val => ((ChartSetting)val).isSecondaryAxis).Count(v => v) > 1;
		}
	}
}
