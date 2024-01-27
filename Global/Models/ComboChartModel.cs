// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global
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
            ComboChartsSettingList.Add(areaChartSetting);
        }

        /// <summary>
        /// 
        /// </summary>
        public void AddComboChartsSetting(BarChartSetting barChartSetting)
        {
            ComboChartsSettingList.Add(barChartSetting);
        }

        /// <summary>
        /// 
        /// </summary>
        public void AddComboChartsSetting(ColumnChartSetting columnChartSetting)
        {
            ComboChartsSettingList.Add(columnChartSetting);
        }

        /// <summary>
        /// 
        /// </summary>
        public void AddComboChartsSetting(LineChartSetting lineChartSetting)
        {
            ComboChartsSettingList.Add(lineChartSetting);
        }

        /// <summary>
        /// 
        /// </summary>
        public void AddComboChartsSetting(PieChartSetting pieChartSetting)
        {
            ComboChartsSettingList.Add(pieChartSetting);
        }

        // /// <summary>
        // /// 
        // /// </summary>
        // public void AddComboChartsSetting(ScatterChartSetting scatterChartSetting)
        // {
        //     ComboChartsSettingList.Add(scatterChartSetting);
        // }

        /// <summary>
        /// The options for the axes of the chart.
        /// </summary>
        public ChartAxesOptions chartAxesOptions = new();

        /// <summary>
        /// The options for the axis of the chart.
        /// </summary>
        public ChartAxisOptions chartAxisOptions = new();
    }
}