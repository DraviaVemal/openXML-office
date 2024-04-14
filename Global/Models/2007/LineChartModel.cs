// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System.Collections.Generic;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Represents the types of line charts.
	/// </summary>
	public enum LineChartTypes
	{
		/// <summary>
		/// Clustered Line Chart
		/// </summary>
		CLUSTERED,
		/// <summary>
		/// Stacked Line Chart
		/// </summary>
		STACKED,
		/// <summary>
		/// Percent Stacked Line Chart
		/// </summary>
		PERCENT_STACKED,
		/// <summary>
		/// Clustered Marker Line Chart
		/// </summary>
		CLUSTERED_MARKER,
		/// <summary>
		/// Stacked Marker Line Chart
		/// </summary>
		STACKED_MARKER,
		/// <summary>
		/// Percent Stacked Marker Line Chart
		/// </summary>
		PERCENT_STACKED_MARKER,
		// CLUSTERED_3D
	}
	/// <summary>
	/// Represents the data label settings for a line chart.
	/// </summary>
	public class LineChartDataLabel : ChartDataLabel
	{
		/// <summary>
		/// The position of the data labels.
		/// </summary>
		public DataLabelPositionValues dataLabelPosition = DataLabelPositionValues.CENTER;
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
			/// Center
			/// </summary>
			CENTER,
			/// <summary>
			/// Above
			/// </summary>
			ABOVE,
			/// <summary>
			/// Below
			/// </summary>
			BELOW,
			// /// <summary>
			// /// Data Callout
			// /// </summary>
			// DATA_CALLOUT
		}
	}
	/// <summary>
	///
	/// </summary>
	public class LineChartLineFormat
	{
		/// <summary>
		///
		/// </summary>
		public string lineColor = null;
		/// <summary>
		/// /
		/// </summary>
		public int? transparency = null;
		/// <summary>
		/// /
		/// </summary>
		public int? width = null;
		/// <summary>
		///
		/// </summary>
		public OutlineCapTypeValues? outlineCapTypeValues = OutlineCapTypeValues.FLAT;
		/// <summary>
		///
		/// </summary>
		public OutlineLineTypeValues? outlineLineTypeValues = OutlineLineTypeValues.SINGEL;
		/// <summary>
		///
		/// </summary>
		public DrawingBeginArrowValues? beginArrowValues = DrawingBeginArrowValues.NONE;
		/// <summary>
		///
		/// </summary>
		public DrawingEndArrowValues? endArrowValues = DrawingEndArrowValues.NONE;
		/// <summary>
		///
		/// </summary>
		public DrawingPresetLineDashValues? dashType;
		/// <summary>
		///
		/// </summary>
		public LineWidthValues? lineStartWidth;
		/// <summary>
		///
		/// </summary>
		public LineWidthValues? lineEndWidth;
	}
	/// <summary>
	///
	/// </summary>
	public class LineChartDataPointSetting : ChartDataPointSettings
	{
		// /// <summary>
		// /// Format applied at data point level
		// /// TODO : Data Point Implementation
		// /// </summary>
		// public LineChartLineFormat? lineChartLineFormat = null;
	}
	/// <summary>
	/// Represents the series settings for a line chart.
	/// </summary>
	public class LineChartSeriesSetting : ChartSeriesSetting
	{
		/// <summary>
		/// Format Applied at series level
		/// </summary>
		public LineChartLineFormat lineChartLineFormat = null;
		/// <summary>
		///
		/// </summary>
		public List<LineChartDataPointSetting> lineChartDataPointSettings = new List<LineChartDataPointSetting>();
		/// <summary>
		/// Option to customize specific data series, which will override the chart level setting.
		/// </summary>
		public LineChartDataLabel lineChartDataLabel = new LineChartDataLabel();
	}
	/// <summary>
	/// Represents the settings for a line chart.
	/// </summary>
	public class LineChartSetting<ApplicationSpecificSetting> : ChartSetting<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
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
		/// The data label settings for the line chart, which will get overridden by series level setting.
		/// </summary>
		public LineChartDataLabel lineChartDataLabel = new LineChartDataLabel();
		/// <summary>
		/// The series settings for the line chart.
		/// </summary>
		public List<LineChartSeriesSetting> lineChartSeriesSettings = new List<LineChartSeriesSetting>();
		/// <summary>
		/// The type of the line chart.
		/// </summary>
		public LineChartTypes lineChartType = LineChartTypes.CLUSTERED;
	}
}
