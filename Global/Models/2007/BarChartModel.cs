// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System.Collections.Generic;

namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Represents the types of bar charts.
	/// </summary>
	public enum BarChartTypes
	{
		/// <summary>
		/// Clustered Bar Chart
		/// </summary>
		CLUSTERED,
		/// <summary>
		/// Stacked Bar Chart
		/// </summary>
		STACKED,
		/// <summary>
		/// Percent Stacked Bar Chart
		/// </summary>
		PERCENT_STACKED,
		/// <summary>
		/// Clustered 3D Bar Chart
		/// </summary>
		CLUSTERED_3D,
		/// <summary>
		/// Stacked Bar Chart
		/// </summary>
		STACKED_3D,
		/// <summary>
		/// Percent Stacked Bar Chart
		/// </summary>
		PERCENT_STACKED_3D
	}
	/// <summary>
	/// Represents the types of bar charts.
	/// </summary>
	public enum BarShapeType
	{
		/// <summary>
		///
		/// </summary>
		BOX,
		/// <summary>
		///
		/// </summary>
		FULL_PYRAMID,
		/// <summary>
		///
		/// </summary>
		PARTIAL_PYRAMID,
		/// <summary>
		///
		/// </summary>
		CYLINDER,
		/// <summary>
		///
		/// </summary>
		FULL_CONE,
		/// <summary>
		///
		/// </summary>
		PARTIAL_CONE
	}
	/// <summary>
	/// Represents the graphics settings for a bar chart.
	/// </summary>
	public class BarGraphicsSetting
	{
		/// <summary>
		/// The gap width between the bars.
		/// Value is used in %.
		/// </summary>
		public int categoryGap = 219;
		/// <summary>
		/// 3D Column Shape Options
		/// </summary>
		public BarShapeType barShapeType = BarShapeType.BOX;
		/// <summary>
		/// The gap between the series bars.
		/// Value is used in %.
		/// </summary>
		public int seriesGap = -27;
	}
	/// <summary>
	/// Represents the data label settings for a bar chart.
	/// </summary>
	public class BarChartDataLabel : ChartDataLabel
	{        /// <summary>
			 /// The position of the data labels.
			 /// </summary>
		public DataLabelPositionValues dataLabelPosition = DataLabelPositionValues.CENTER;
		/// <summary>
		/// The possible positions for the data labels.
		/// </summary>
		public enum DataLabelPositionValues
		{
			/// <summary>
			/// Center
			/// </summary>
			CENTER,
			/// <summary>
			/// Inside end
			/// </summary>
			INSIDE_END,
			/// <summary>
			/// Inside base
			/// </summary>
			INSIDE_BASE,
			/// <summary>
			/// This option is only for Cluster type chart.
			/// </summary>
			OUTSIDE_END,
			// /// <summary>
			// /// Data Callout
			// /// </summary>
			// DATA_CALLOUT
		}
	}
	/// <summary>
	///
	/// </summary>
	public class BarChartDataPointSetting : ChartDataPointSettings
	{
	}
	/// <summary>
	/// Represents the series settings for a bar chart.
	/// </summary>
	public class BarChartSeriesSetting : ChartSeriesSetting
	{
		/// <summary>
		///
		/// </summary>
		public List<BarChartDataPointSetting> barChartDataPointSettings = new List<BarChartDataPointSetting>();
		/// <summary>
		/// Option to customize specific data series. This will override the chart level setting.
		/// </summary>
		public BarChartDataLabel barChartDataLabel = new BarChartDataLabel();
		/// <summary>
		/// The color of the fill.
		/// </summary>
		public string fillColor;
	}
	/// <summary>
	/// Represents the settings for a bar chart.
	/// </summary>
	public class BarChartSetting<ApplicationSpecificSetting> : ChartSetting<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
	{
		/// <summary>
		/// The data label settings for the entire chart. This will get overridden by series level setting.
		/// </summary>
		public BarChartDataLabel barChartDataLabel = new BarChartDataLabel();
		/// <summary>
		/// The series settings for the bar chart.
		/// </summary>
		public List<BarChartSeriesSetting> barChartSeriesSettings = new List<BarChartSeriesSetting>();
		/// <summary>
		/// The type of bar chart.
		/// </summary>
		public BarChartTypes barChartType = BarChartTypes.CLUSTERED;
		/// <summary>
		/// The options for the chart axes.
		/// </summary>
		public ChartAxesOptions chartAxesOptions = new ChartAxesOptions();
		/// <summary>
		/// The options for the chart axis.
		/// </summary>
		public ChartAxisOptions chartAxisOptions = new ChartAxisOptions();
		/// <summary>
		/// The graphics settings for the bar chart.
		/// </summary>
		public BarGraphicsSetting barGraphicsSetting = new BarGraphicsSetting();
	}
}
