// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

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

		// CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D, COLUMN_3D
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
	{        /// <summary>
			 ///
			 /// </summary>
		public List<ColumnChartDataPointSetting?> columnChartDataPointSettings = new();
		/// <summary>
		/// Option to customize specific data series. Will override chart level setting.
		/// </summary>
		public ColumnChartDataLabel columnChartDataLabel = new();

		/// <summary>
		/// Chart Stick Fill Color
		/// </summary>
		public string? fillColor;
	}

	/// <summary>
	/// Represents the settings for a column chart.
	/// </summary>
	public class ColumnChartSetting<ApplicationSpecificSetting> : ChartSetting<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
	{

		/// <summary>
		/// Chart Axes Options
		/// </summary>
		public ChartAxesOptions chartAxesOptions = new();

		/// <summary>
		/// Chart Axis Options
		/// </summary>
		public ChartAxisOptions chartAxisOptions = new();

		/// <summary>
		/// Will get overridden by series level setting.
		/// </summary>
		public ColumnChartDataLabel columnChartDataLabel = new();

		/// <summary>
		/// Chart Series Settings
		/// </summary>
		public List<ColumnChartSeriesSetting?> columnChartSeriesSettings = new();

		/// <summary>
		/// Chart Type. default is CLUSTERED
		/// </summary>
		public ColumnChartTypes columnChartType = ColumnChartTypes.CLUSTERED;
		/// <summary>
		/// The graphics settings for the column chart.
		/// </summary>
		public ColumnGraphicsSetting columnGraphicsSetting = new();
	}
}
