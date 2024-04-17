// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System.Collections.Generic;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Represents the types of scatter charts.
	/// </summary>
	public enum ScatterChartTypes
	{
		/// <summary>
		/// Scatter Chart
		/// </summary>
		SCATTER,
		/// <summary>
		/// Scatter Chart with Smooth Lines
		/// </summary>
		SCATTER_SMOOTH,
		/// <summary>
		/// Scatter Chart with Smooth Lines and Markers
		/// </summary>
		SCATTER_SMOOTH_MARKER,
		/// <summary>
		/// Scatter Chart with Straight Lines
		/// </summary>
		SCATTER_STRIGHT,
		/// <summary>
		/// Scatter Chart with Straight Lines and Markers
		/// </summary>
		SCATTER_STRIGHT_MARKER,
		/// <summary>
		/// Bubble Chart
		/// </summary>
		BUBBLE,
		// BUBBLE_3D
	}
	/// <summary>
	/// Represents the data label settings for a scatter chart.
	/// </summary>
	public class ScatterChartDataLabel : ChartDataLabel
	{
		/// <summary>
		/// The position of the data labels.
		/// </summary>
		public DataLabelPositionValues dataLabelPosition = DataLabelPositionValues.CENTER;
		/// <summary>
		/// Determines whether to show the bubble size in the data labels.
		/// </summary>
		public bool showBubbleSize = false;
		/// <summary>
		/// The possible positions for the data labels.
		/// </summary>
		public enum DataLabelPositionValues
		{
			/// <summary>
			/// Left Side
			/// </summary>
			LEFT,
			/// <summary>
			/// Right Side
			/// </summary>
			RIGHT,
			/// <summary>
			/// Center Placement
			/// </summary>
			CENTER,
			/// <summary>
			/// Above content
			/// </summary>
			ABOVE,
			/// <summary>
			/// Below content
			/// </summary>
			BELOW,
			// /// <summary>
			// /// Data Callout Style
			// /// </summary>
			// DATA_CALLOUT
		}
	}
	/// <summary>
	/// Represents the series settings for a scatter chart.
	/// </summary>
	public class ScatterChartSeriesSetting : ChartSeriesSetting
	{
		/// <summary>
		/// Custom data label settings for the specific data series. This will override the chart
		/// level setting.
		/// </summary>
		public ScatterChartDataLabel scatterChartDataLabel = new ScatterChartDataLabel();
	}
	/// <summary>
	/// Represents the settings for a scatter chart.
	/// </summary>
	public class ScatterChartSetting<ApplicationSpecificSetting> : ChartSetting<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition, new()
	{
		/// <summary>
		/// The options for the chart axes.
		/// </summary>
		public ChartAxesOptions chartAxesOptions = new ChartAxesOptions();
		/// <summary>
		/// The options for the chart axis.
		/// </summary>
		public ChartAxisOptions chartAxisOptions = new ChartAxisOptions();
		/// <summary>
		/// The data label settings for the scatter chart. This will get overridden by the series
		/// level setting.
		/// </summary>
		public ScatterChartDataLabel scatterChartDataLabel = new ScatterChartDataLabel();
		/// <summary>
		/// The list of series settings for the scatter chart.
		/// </summary>
		public List<ScatterChartSeriesSetting> scatterChartSeriesSettings = new List<ScatterChartSeriesSetting>();
		/// <summary>
		/// The type of scatter chart.
		/// </summary>
		public ScatterChartTypes scatterChartType = ScatterChartTypes.SCATTER;
	}
}
