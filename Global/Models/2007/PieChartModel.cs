// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Represents the types of pie charts.
	/// </summary>
	public enum PieChartTypes
	{
		/// <summary>
		/// Pie Chart
		/// </summary>
		PIE,
		/// <summary>
		///
		/// </summary>
		PIE_3D,
		// PIE_PIE, PIE_BAR,
		/// <summary>
		/// Doughnut Chart
		/// </summary>
		DOUGHNUT
	}
	/// <summary>
	/// Represents the data label for a pie chart.
	/// </summary>
	public class PieChartDataLabel : ChartDataLabel
	{
		/// <summary>
		/// The position of the data label.
		/// </summary>
		public DataLabelPositionValues dataLabelPosition = DataLabelPositionValues.CENTER;
		/// <summary>
		/// Represents the possible positions of the data label.
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
			/// Outside End
			/// </summary>
			OUTSIDE_END,
			/// <summary>
			/// Best Fit
			/// </summary>
			BEST_FIT,
			/// <summary>
			/// Option only for doughnut chart type
			/// </summary>
			SHOW,
			// /// <summary>
			// /// Data Callout
			// /// </summary>
			// DATA_CALLOUT
		}
	}
	/// <summary>
	///
	/// </summary>
	public class PieChartDataPointSetting : ChartDataPointSettings
	{
	}
	/// <summary>
	/// Represents the series setting for a pie chart.
	/// </summary>
	public class PieChartSeriesSetting : ChartSeriesSetting
	{
		/// <summary>
		/// The color of the fill.
		/// </summary>
		public string fillColor;
		/// <summary>
		/// Border Color is only valid for Doughnut Chart
		/// </summary>
		public override string borderColor { get => base.borderColor; set => base.borderColor = value; }
		/// <summary>
		///
		/// </summary>
		public List<PieChartDataPointSetting> pieChartDataPointSettings = new List<PieChartDataPointSetting>();
		/// <summary>
		/// Option to customize specific data series, will override chart level setting.
		/// </summary>
		public PieChartDataLabel pieChartDataLabel = new PieChartDataLabel();
	}
	/// <summary>
	/// Represents the setting for a pie chart.
	/// </summary>
	public class PieChartSetting<ApplicationSpecificSetting> : ChartSetting<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
	{
		/// <summary>
		/// Will get overridden by series level setting.
		/// </summary>
		public PieChartDataLabel pieChartDataLabel = new PieChartDataLabel();
		/// <summary>
		/// The list of series settings for the pie chart.
		/// </summary>
		public List<PieChartSeriesSetting> pieChartSeriesSettings = new List<PieChartSeriesSetting>();
		/// <summary>
		/// The type of the pie chart.
		/// </summary>
		public PieChartTypes pieChartType = PieChartTypes.PIE;
		/// <summary>
		/// Will be ignored if pieChartTypes is not DOUGHNUT.
		/// Value is assumed in percentage.
		/// </summary>
		public uint doughnutHoleSize = 50;
		/// <summary>
		/// The angle of the first slice of the pie chart in degrees.
		/// </summary>
		public uint angleOfFirstSlice = 0;
		/// <summary>
		/// Pie Explosion Value in percentage
		/// </summary>
		public uint pointExplosion = 0;
	}
}
