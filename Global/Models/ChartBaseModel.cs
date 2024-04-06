// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the position of an axis in a chart.
	/// </summary>
	public enum AxisPosition
	{
		/// <summary>
		/// Top
		/// </summary>
		TOP,

		/// <summary>
		/// Bottom
		/// </summary>
		BOTTOM,

		/// <summary>
		/// Left
		/// </summary>
		LEFT,

		/// <summary>
		/// Right
		/// </summary>
		RIGHT
	}

	/// <summary>
	/// Represents the settings for a category axis in a chart.
	/// </summary>
	public class CategoryAxisSetting
	{
		internal uint id;
		internal AxisPosition axisPosition = AxisPosition.BOTTOM;
		internal uint crossAxisId;
		/// <summary>
		/// Is Font Bold
		/// </summary>
		internal bool isBold = false;
		/// <summary>
		/// Is Font Italic
		/// </summary>
		internal bool isItalic = false;
		/// <summary>
		///  Font Size
		/// </summary>
		internal float fontSize = 11.97F;
		/// <summary>
		///
		/// </summary>
		internal string? fontColor = null;
		/// <summary>
		///
		/// </summary>
		internal UnderLineValues underLineValues = UnderLineValues.NONE;
		/// <summary>
		///
		/// </summary>
		internal StrikeValues strikeValues = StrikeValues.NO_STRIKE;
		/// <summary>
		///
		/// </summary>
		internal bool isVisible = true;
		internal bool invertOrder;
	}

	/// <summary>
	/// Represents the options for the axes in a chart.
	/// </summary>
	public class ChartAxesOptions
	{

		/// <summary>
		/// Invert the axes item order for the Vertical axis
		/// </summary>
		public bool invertVerticalAxesOrder;
		/// <summary>
		/// Invert the axes item order for the Horizontal axis
		/// </summary>
		public bool invertHorizontalAxesOrder;
		/// <summary>
		/// Is Horizontal Axes Enabled
		/// </summary>
		public bool isHorizontalAxesEnabled = true;
		/// <summary>
		/// Is Font Bold
		/// </summary>
		public bool isHorizontalBold = false;
		/// <summary>
		/// Is Font Italic
		/// </summary>
		public bool isHorizontalItalic = false;
		/// <summary>
		///  Font Size
		/// </summary>
		public float horizontalFontSize = 11.97F;
		/// <summary>
		///
		/// </summary>
		public string? horizontalFontColor = null;
		/// <summary>
		///
		/// </summary>
		public UnderLineValues horizontalUnderLineValues = UnderLineValues.NONE;
		/// <summary>
		///
		/// </summary>
		public StrikeValues horizontalStrikeValues = StrikeValues.NO_STRIKE;
		/// <summary>
		/// Is Font Bold
		/// </summary>
		public bool isVerticalBold = false;
		/// <summary>
		/// Is Font Italic
		/// </summary>
		public bool isVerticalItalic = false;
		/// <summary>
		///  Font Size
		/// </summary>
		public float verticalFontSize = 11.97F;
		/// <summary>
		///
		/// </summary>
		public string? verticalFontColor = null;
		/// <summary>
		///
		/// </summary>
		public UnderLineValues verticalUnderLineValues = UnderLineValues.NONE;
		/// <summary>
		///
		/// </summary>
		public StrikeValues verticalStrikeValues = StrikeValues.NO_STRIKE;
		/// <summary>
		/// Is Vertical Axes Enabled
		/// </summary>
		public bool isVerticalAxesEnabled = true;
	}

	/// <summary>
	/// Represents the options for a chart axis.
	/// </summary>
	public class ChartAxisOptions
	{

		/// <summary>
		/// Horizontal Axis Title
		/// </summary>
		public string? horizontalAxisTitle;

		/// <summary>
		/// Vertical Axis Title
		/// </summary>
		public string? verticalAxisTitle;
	}

	/// <summary>
	/// Represents the grouping options for chart data.
	/// </summary>
	public class ChartDataGrouping
	{

		/// <summary>
		/// Gets or sets the data label cells.
		/// </summary>
		public ChartData[]? dataLabelCells;

		/// <summary>
		/// Gets or sets the data label formula.
		/// </summary>
		public string? dataLabelFormula;

		/// <summary>
		/// Gets or sets the series header cells.
		/// </summary>
		public ChartData? seriesHeaderCells;

		/// <summary>
		/// Gets or sets the series header format.
		/// </summary>
		public string? seriesHeaderFormat;

		/// <summary>
		/// Gets or sets the series header formula.
		/// </summary>
		public string? seriesHeaderFormula;

		/// <summary>
		/// Gets or sets the X-axis cells.
		/// </summary>
		public ChartData[]? xAxisCells;

		/// <summary>
		/// Gets or sets the X-axis format.
		/// </summary>
		public string? xAxisFormat;

		/// <summary>
		/// Gets or sets the X-axis formula.
		/// </summary>
		public string? xAxisFormula;

		/// <summary>
		/// Gets or sets the Y-axis cells.
		/// </summary>
		public ChartData[]? yAxisCells;

		/// <summary>
		/// Gets or sets the Y-axis format.
		/// </summary>
		public string? yAxisFormat;

		/// <summary>
		/// Gets or sets the Y-axis formula.
		/// </summary>
		public string? yAxisFormula;

		/// <summary>
		/// Gets or sets the Z-axis cells.
		/// </summary>
		public ChartData[]? zAxisCells;

		/// <summary>
		/// Gets or sets the Z-axis format.
		/// </summary>
		public string? zAxisFormat;

		/// <summary>
		/// Gets or sets the Z-axis formula.
		/// </summary>
		public string? zAxisFormula;
		/// <summary>
		///
		/// </summary>
		public int id;
	}

	/// <summary>
	/// Represents the options for chart data labels.
	/// </summary>
	public class ChartDataLabel
	{

		/// <summary>
		/// The separator used for displaying multiple values.
		/// </summary>
		public string separator = ", ";

		/// <summary>
		/// Determines whether to show the category name in the chart.
		/// </summary>
		public bool showCategoryName = false;

		/// <summary>
		/// Determines whether to show the legend key in the chart.
		/// </summary>
		public bool showLegendKey = false;

		/// <summary>
		/// Determines whether to show the series name in the chart.
		/// </summary>
		public bool showSeriesName = false;

		/// <summary>
		/// Determines whether to show the value in the chart.
		/// </summary>
		public bool showValue = false;

		/// <summary>
		/// Determines whether to show the value from a column in the chart.
		/// </summary>
		public bool showValueFromColumn = false;
		/// <summary>
		/// Is Font Bold
		/// </summary>
		public bool isBold = false;
		/// <summary>
		/// Is Font Italic
		/// </summary>
		internal bool isItalic = false;
		/// <summary>
		/// Font Size
		/// </summary>
		public float fontSize = 11.97F;
		/// <summary>
		///
		/// </summary>
		public string? fontColor = null;
		/// <summary>
		///
		/// </summary>
		public UnderLineValues underLineValues = UnderLineValues.NONE;
		/// <summary>
		///
		/// </summary>
		public StrikeValues strikeValues = StrikeValues.NO_STRIKE;
	}

	/// <summary>
	/// Represents the settings for chart data.
	/// </summary>
	public class ChartDataSetting
	{

		/// <summary>
		/// Set 0 To Use Till End
		/// </summary>
		public uint chartDataColumnEnd = 0;

		/// <summary>
		/// Chart data Start column 0 based
		/// </summary>
		public uint chartDataColumnStart = 0;

		/// <summary>
		/// Set 0 To Use Till End
		/// </summary>
		public uint chartDataRowEnd = 0;

		/// <summary>
		/// Chart data Start Row 0 based
		/// </summary>
		public uint chartDataRowStart = 0;

		/// <summary>
		/// Is Data is used in 3D chart type
		/// </summary>
		internal bool is3Ddata;

		/// <summary>
		/// Key For Data Column Value For Data Label Column If Data Label Column Are Present
		/// Inbetween and Used in the list it will be auto skipped By Data Column
		/// </summary>
		public Dictionary<uint, uint> valueFromColumn = new();
	}

	/// <summary>
	/// Represents the options for chart grid lines.
	/// </summary>
	public class ChartGridLinesOptions
	{

		/// <summary>
		/// Is Major Category Lines Enabled
		/// </summary>
		public bool isMajorCategoryLinesEnabled = false;

		/// <summary>
		/// Is Major Value Lines Enabled
		/// </summary>
		public bool isMajorValueLinesEnabled = true;

		/// <summary>
		/// Is Minor Category Lines Enabled
		/// </summary>
		public bool isMinorCategoryLinesEnabled = false;

		/// <summary>
		/// Is Minor Value Lines Enabled
		/// </summary>
		public bool isMinorValueLinesEnabled = false;
	}

	/// <summary>
	/// Represents the options for chart legend.
	/// </summary>
	public class ChartLegendOptions
	{

		/// <summary>
		/// Is Legend Enabled
		/// </summary>
		public bool isEnableLegend = true;

		/// <summary>
		/// Is Legend Chart OverLap
		/// </summary>
		public bool isLegendChartOverLap = false;
		/// <summary>
		/// Is Font Bold
		/// </summary>
		public bool isBold = false;
		/// <summary>
		/// Is Font Italic
		/// </summary>
		internal bool isItalic = false;
		/// <summary>
		/// Font Size
		/// </summary>
		public float fontSize = 11.97F;
		/// <summary>
		///
		/// </summary>
		public string? fontColor = null;
		/// <summary>
		///
		/// </summary>
		public UnderLineValues underLineValues = UnderLineValues.NONE;
		/// <summary>
		///
		/// </summary>
		public StrikeValues strikeValues = StrikeValues.NO_STRIKE;

		/// <summary>
		/// Legend Position
		/// </summary>
		public LegendPositionValues legendPosition = LegendPositionValues.BOTTOM;

		/// <summary>
		/// Legend Position Values
		/// </summary>
		public enum LegendPositionValues
		{
			/// <summary>
			/// Bottom
			/// </summary>
			BOTTOM,

			/// <summary>
			/// Top
			/// </summary>
			TOP,

			/// <summary>
			/// Left
			/// </summary>
			LEFT,

			/// <summary>
			/// Right
			/// </summary>
			RIGHT,

			/// <summary>
			/// Top Right
			/// </summary>
			TOP_RIGHT
		}
		/// <summary>
		/// Manual Position legend
		/// </summary>
		public LayoutModel? manualLayout;
	}

	/// <summary>
	///
	/// </summary>
	public class ChartDataPointSettings
	{
		/// <summary>
		/// The color of the fill.
		/// </summary>
		public string? fillColor;
		/// <summary>
		///
		/// </summary>
		public string? borderColor;
	}

	/// <summary>
	/// Represents the settings for a chart series.
	/// </summary>
	public class ChartSeriesSetting
	{

		/// <summary>
		/// The color of the border.
		/// </summary>
		public virtual string? borderColor { get; set; }
		internal ChartSeriesSetting() { }
	}

	/// <summary>
	///
	/// </summary>
	public class PlotAreaModel
	{
		/// <summary>
		/// Manual Position Char Ghaph Area
		/// </summary>
		public LayoutModel? manualLayout;
	}

	/// <summary>
	///
	/// </summary>
	public class ChartTitleModel
	{
		/// <summary>
		///
		/// </summary>
		public int fontSize = 11;
		/// <summary>
		///
		/// </summary>
		public string? fontColor;
		/// <summary>
		///
		/// </summary>
		public string? title;
		/// <summary>
		///
		/// </summary>
		public bool isBold;
		/// <summary>
		///
		/// </summary>
		public bool isItalic;
		/// <summary>
		///
		/// </summary>
		internal UnderLineValues underLineValues = UnderLineValues.NONE;

	}

	/// <summary>
	/// 
	/// </summary>
	public class AnchorPosition
	{
		/// <summary>
		/// 
		/// </summary>
		public uint column = 1;
		/// <summary>
		/// 
		/// </summary>
		public uint columnOffset = 0;
		/// <summary>
		/// 
		/// </summary>
		public uint row = 1;
		/// <summary>
		/// 
		/// </summary>
		public uint rowOffset = 0;
	}

	/// <summary>
	/// 
	/// </summary>
	public interface ISizeAndPosition { }

	/// <summary>
	/// 
	/// </summary>
	public class PresentationSetting : ISizeAndPosition
	{

		/// <summary>
		/// Chart Height in EMU
		/// </summary>
		public uint height = 6858000;

		/// <summary>
		/// Chart Width in EMU
		/// </summary>
		public uint width = 12192000;

		/// <summary>
		/// Chart X Position in EMU
		/// </summary>
		public uint x = 0;

		/// <summary>
		/// Chart Y Position in EMU
		/// </summary>
		public uint y = 0;
	}

	/// <summary>
	/// 
	/// </summary>
	public class ExcelSetting : ISizeAndPosition
	{
		/// <summary>
		/// 
		/// </summary>
		public AnchorPosition from = new();

		/// <summary>
		/// 
		/// </summary>
		public AnchorPosition to = new();
	}

	/// <summary>
	/// Represents the settings for a chart.
	/// </summary>
	public class ChartSetting<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
	{
		internal uint? categoryAxisId;
		internal uint? valueAxisId;
		internal bool is3DChart = false;

		/// <summary>
		/// Only useful when used with Combo chart
		/// </summary>
		public bool isSecondaryAxis = false;

		/// <summary>
		///
		/// </summary>
		public PlotAreaModel? plotAreaOptions;

		/// <summary>
		/// Chart Data Setting
		/// </summary>
		public ChartDataSetting chartDataSetting = new();

		/// <summary>
		/// Chart Grid Line Options
		/// </summary>
		public ChartGridLinesOptions chartGridLinesOptions = new();

		/// <summary>
		/// Chart Legend Options
		/// </summary>
		public ChartLegendOptions chartLegendOptions = new();

		/// <summary>
		/// Chart Title
		/// </summary>
		public ChartTitleModel? titleOptions;

		/// <summary>
		/// 
		/// </summary>
		public required ApplicationSpecificSetting applicationSpecificSetting;

		internal ChartSetting() { }
	}

	/// <summary>
	/// Represents the settings for a value axis in a chart.
	/// </summary>
	public class ValueAxisSetting
	{
		internal uint id;
		internal AxisPosition axisPosition = AxisPosition.LEFT;
		internal uint crossAxisId;
		/// <summary>
		/// Is Font Bold
		/// </summary>
		internal bool isBold = false;
		/// <summary>
		/// Is Font Italic
		/// </summary>
		internal bool isItalic = false;
		/// <summary>
		/// Font Size
		/// </summary>
		internal float fontSize = 11.97F;
		/// <summary>
		///
		/// </summary>
		internal string? fontColor = null;
		/// <summary>
		///
		/// </summary>
		internal UnderLineValues underLineValues = UnderLineValues.NONE;
		/// <summary>
		///
		/// </summary>
		internal StrikeValues strikeValues = StrikeValues.NO_STRIKE;
		/// <summary>
		///
		/// </summary>
		internal bool isVisible = true;
		/// <summary>
		/// Invert the axis scale order
		/// </summary>
		internal bool invertOrder;

		internal C.TickMarkValues majorTickMark = C.TickMarkValues.None;
		internal C.TickMarkValues minorTickMark = C.TickMarkValues.None;
		internal C.CrossesValues crosses = C.CrossesValues.AutoZero;
	}

	/// <summary>
	///
	/// </summary>
	public class LayoutModel
	{
		/// <summary>
		/// Considered from left to right
		/// Value is between 0 to 1
		/// </summary>
		public float x = 0;
		/// <summary>
		/// Considered from top to bottom
		/// Value is between 0 to 1
		/// </summary>
		public float y = 0;
		/// <summary>
		/// Value is between 0 to 1
		/// </summary>
		public float width = 1;
		/// <summary>
		/// Value is between 0 to 1
		/// </summary>
		public float height = 1;
	}

	/// <summary>
	///
	/// </summary>
	public class MarkerModel
	{
		/// <summary>
		/// Market Size
		/// </summary>
		public int size = 5;
		/// <summary>
		///
		/// </summary>
		public MarkerShapeValues markerShapeValues = MarkerShapeValues.NONE;
		/// <summary>
		///
		/// </summary>
		public ShapePropertiesModel shapeProperties = new();
		/// <summary>
		///
		/// </summary>
		public enum MarkerShapeValues
		{
			/// <summary>
			///
			/// </summary>
			NONE,
			/// <summary>
			///
			/// </summary>
			AUTO,
			/// <summary>
			///
			/// </summary>
			CIRCLE,
			/// <summary>
			///
			/// </summary>
			DASH,
			/// <summary>
			///
			/// </summary>
			DIAMOND,
			/// <summary>
			///
			/// </summary>
			DOT,
			/// <summary>
			///
			/// </summary>
			PICTURE,
			/// <summary>
			///
			/// </summary>
			PLUSE,
			/// <summary>
			///
			/// </summary>
			SQUARE,
			/// <summary>
			///
			/// </summary>
			STAR,
			/// <summary>
			///
			/// </summary>
			TRIANGLE,
			/// <summary>
			///
			/// </summary>
			X
		}

		internal static C.MarkerStyleValues GetMarkerStyleValues(MarkerShapeValues markerShapeValues)
		{
			return markerShapeValues switch
			{
				MarkerShapeValues.AUTO => C.MarkerStyleValues.Auto,
				MarkerShapeValues.CIRCLE => C.MarkerStyleValues.Circle,
				MarkerShapeValues.DASH => C.MarkerStyleValues.Dash,
				MarkerShapeValues.DIAMOND => C.MarkerStyleValues.Diamond,
				MarkerShapeValues.DOT => C.MarkerStyleValues.Dot,
				MarkerShapeValues.PICTURE => C.MarkerStyleValues.Picture,
				MarkerShapeValues.PLUSE => C.MarkerStyleValues.Plus,
				MarkerShapeValues.SQUARE => C.MarkerStyleValues.Square,
				MarkerShapeValues.STAR => C.MarkerStyleValues.Star,
				MarkerShapeValues.TRIANGLE => C.MarkerStyleValues.Triangle,
				MarkerShapeValues.X => C.MarkerStyleValues.X,
				_ => C.MarkerStyleValues.None,
			};
		}
	}
}
