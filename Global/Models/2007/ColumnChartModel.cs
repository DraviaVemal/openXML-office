// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System.Collections.Generic;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Represents the types of column charts.
	/// </summary>
	public enum ColumnChartTypes
	{
		/// <summary>
		/// Clustered Column Chart
		/// </summary>
		CLUSTERED,
		/// <summary>
		/// Stacked Column Chart
		/// </summary>
		STACKED,
		/// <summary>
		/// Percent Stacked Column Chart
		/// </summary>
		PERCENT_STACKED,
		/// <summary>
		///
		/// </summary>
		CLUSTERED_3D,
		/// <summary>
		///
		/// </summary>
		STACKED_3D,
		/// <summary>
		///
		/// </summary>
		PERCENT_STACKED_3D,
		//COLUMN_3D
	}
	/// <summary>
	/// Represents the types of bar charts.
	/// </summary>
	public enum ColumnShapeType
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
	/// Represents the graphics settings for a column chart.
	/// </summary>
	public class ColumnGraphicsSetting
	{
		/// <summary>
		/// The gap width between the Column.
		/// </summary>
		public int categoryGap = 219;
		/// <summary>
		/// The gap between the series column.
		/// </summary>
		public int seriesGap = -27;
		/// <summary>
		/// 3D Column Shape Options
		/// </summary>
		public BarShapeType columnShapeType = BarShapeType.BOX;
	}
	/// <summary>
	/// Represents the data label settings for a column chart.
	/// </summary>
	public class ColumnChartDataLabel : ChartDataLabel
	{
		/// <summary>
		/// The position of the data label.
		/// </summary>
		public DataLabelPositionValues dataLabelPosition = DataLabelPositionValues.CENTER;
		/// <summary>
		/// The possible positions for the data label.
		/// </summary>
		public enum DataLabelPositionValues
		{
			/// <summary>
			/// Center
			/// </summary>
			CENTER,
			/// <summary>
			/// Inside End
			/// </summary>
			INSIDE_END,
			/// <summary>
			/// Inside Base
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
	public class ColumnChartDataPointSetting : ChartDataPointSettings
	{
	}
	/// <summary>
	/// Represents the series settings for a column chart.
	/// </summary>
	public class ColumnChartSeriesSetting : ChartSeriesSetting
	{
		/// <summary>
		///
		/// </summary>
		public List<ColumnChartDataPointSetting> columnChartDataPointSettings = new List<ColumnChartDataPointSetting>();
		/// <summary>
		/// Option to customize specific data series. Will override chart level setting.
		/// </summary>
		public ColumnChartDataLabel columnChartDataLabel = new ColumnChartDataLabel();
		/// <summary>
		/// Chart Stick Fill Color
		/// </summary>
		public string fillColor;
	}
	/// <summary>
	/// Represents the settings for a column chart.
	/// </summary>
	public class ColumnChartSetting<ApplicationSpecificSetting> : ChartSetting<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition, new()
	{
		/// <summary>
		/// Chart Axes Options
		/// </summary>
		public ChartAxesOptions chartAxesOptions = new ChartAxesOptions();
		/// <summary>
		/// Chart Axis Options
		/// </summary>
		public ChartAxisOptions chartAxisOptions = new ChartAxisOptions();
		/// <summary>
		/// Will get overridden by series level setting.
		/// </summary>
		public ColumnChartDataLabel columnChartDataLabel = new ColumnChartDataLabel();
		/// <summary>
		/// Chart Series Settings
		/// </summary>
		public List<ColumnChartSeriesSetting> columnChartSeriesSettings = new List<ColumnChartSeriesSetting>();
		/// <summary>
		/// Chart Type. default is CLUSTERED
		/// </summary>
		public ColumnChartTypes columnChartType = ColumnChartTypes.CLUSTERED;
		/// <summary>
		/// The graphics settings for the column chart.
		/// </summary>
		public ColumnGraphicsSetting columnGraphicsSetting = new ColumnGraphicsSetting();
	}
}
