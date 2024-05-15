// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using C = DocumentFormat.OpenXml.Drawing.Charts;
using OpenXMLOffice.Global_2013;
using System;


namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// 
	/// </summary>
	public enum AxesLabelPosition
	{
		/// <summary>
		/// 
		/// </summary>
		NEXT_TO_AXIS,
		/// <summary>
		/// 
		/// </summary>
		LOW,
		/// <summary>
		/// 
		/// </summary>
		HIGH,
		/// <summary>
		/// 
		/// </summary>
		NONE
	}
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
	///
	/// </summary>
	public enum MarkerShapeTypes
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
		PLUS,
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
	/// <summary>
	/// Text direction in charts
	/// </summary>
	public enum ChartTextDirectionValues
	{
		/// <summary>
		/// 
		/// </summary>
		HORIZONTAL,
		/// <summary>
		/// 
		/// </summary>
		ROTATE_90,
		/// <summary>
		/// 
		/// </summary>
		ROTATE_270,
		/// <summary>
		/// 
		/// </summary>
		STACKED
	}
	/// <summary>
	/// 
	/// </summary>
	public enum ChartVerticalTextAlignmentValues
	{
		/// <summary>
		/// 
		/// </summary>
		RIGHT,
		/// <summary>
		/// 
		/// </summary>
		CENTER,
		/// <summary>
		/// 
		/// </summary>
		LEFT,
		/// <summary>
		/// 
		/// </summary>
		RIGHT_MIDDLE,
		/// <summary>
		/// 
		/// </summary>
		CENTER_MIDDLE,
		/// <summary>
		/// 
		/// </summary>
		LEFT_MIDDLE,
		/// <summary>
		/// 
		/// </summary>
		TOP,
		/// <summary>
		/// 
		/// </summary>
		MIDDLE,
		/// <summary>
		/// 
		/// </summary>
		BOTTOM,
		/// <summary>
		/// 
		/// </summary>
		TOP_CENTER,
		/// <summary>
		/// 
		/// </summary>
		MIDDLE_CENTER,
		/// <summary>
		/// 
		/// </summary>
		BOTTOM_CENTER,
	}
	/// <summary>
	/// 
	/// </summary>
	public class ChartTextOptions : TextOptions
	{
		/// <summary>
		/// TODO
		/// </summary>
		internal ChartTextDirectionValues textDirectionValue = ChartTextDirectionValues.HORIZONTAL;
		internal ChartVerticalTextAlignmentValues chartVerticalTextAlignmentValue;
		/// <summary>
		/// TODO
		/// </summary>
		internal ChartVerticalTextAlignmentValues ChartVerticalTextAlignmentValue
		{
			get
			{
				return chartVerticalTextAlignmentValue;
			}
			set
			{
				switch (textDirectionValue)
				{
					case ChartTextDirectionValues.HORIZONTAL:
						if (value == ChartVerticalTextAlignmentValues.LEFT ||
							value == ChartVerticalTextAlignmentValues.CENTER ||
							value == ChartVerticalTextAlignmentValues.RIGHT ||
							value == ChartVerticalTextAlignmentValues.LEFT_MIDDLE ||
							value == ChartVerticalTextAlignmentValues.CENTER_MIDDLE ||
							value == ChartVerticalTextAlignmentValues.RIGHT_MIDDLE)
						{
							throw new Exception("Selected Text Configuration is not acceptable");
						}
						break;
					default:
						if (value == ChartVerticalTextAlignmentValues.TOP ||
						value == ChartVerticalTextAlignmentValues.MIDDLE ||
						value == ChartVerticalTextAlignmentValues.BOTTOM ||
						value == ChartVerticalTextAlignmentValues.TOP_CENTER ||
						value == ChartVerticalTextAlignmentValues.MIDDLE_CENTER ||
						value == ChartVerticalTextAlignmentValues.BOTTOM_CENTER)
						{
							throw new Exception("Selected Text Configuration is not acceptable");
						}
						break;
				}
				chartVerticalTextAlignmentValue = value;
			}
		}
		private int textAngle = 0;
		/// <summary>
		/// Set Text Angle between -90 to 90 degree
		/// </summary>
		public int TextAngle
		{
			get
			{
				return textAngle;
			}
			set
			{
				if (value > 90)
				{
					textAngle = 90;
				}
				else if (value < -90)
				{
					textAngle = -90;
				}
				else
				{
					textAngle = value;
				}
			}
		}
	}
	/// <summary>
	/// Represents the options for the axes in a chart.
	/// </summary>
	public class ChartAxesLabel : ChartTextOptions
	{
		/// <summary>
		/// Axis Label Position.
		/// </summary>
		public AxesLabelPosition axesLabelPosition = AxesLabelPosition.NEXT_TO_AXIS;
		/// <summary>
		/// Invert the axis order
		/// </summary>
		public bool inReverseOrder = false;
	}
	/// <summary>
	/// Represents the options for the axis title text
	/// </summary>
	public class ChartAxisTitle : ChartTextOptions { }

	/// <summary>
	/// Common Chart Axis Options
	/// </summary>
	public class AxisOptions
	{
		internal C.CrossesValues crosses = C.CrossesValues.AutoZero;
		internal C.TickMarkValues majorTickMark = C.TickMarkValues.None;
		internal C.TickMarkValues minorTickMark = C.TickMarkValues.None;
		/// <summary>
		/// Is Horizontal Axes Enabled
		/// </summary>
		public bool isAxesVisible = true;
		/// <summary>
		/// Option for Axis's Axes label options
		/// </summary>
		public ChartAxesLabel chartAxesOptions = new ChartAxesLabel();
		/// <summary>
		/// Option for Axis title options
		/// </summary>
		public ChartAxisTitle chartAxisTitle = new ChartAxisTitle();

		internal static C.TickLabelPositionValues GetLabelAxesPosition(AxesLabelPosition axesLabelPosition)
		{
			switch (axesLabelPosition)
			{
				case AxesLabelPosition.LOW:
					return C.TickLabelPositionValues.Low;
				case AxesLabelPosition.HIGH:
					return C.TickLabelPositionValues.High;
				case AxesLabelPosition.NONE:
					return C.TickLabelPositionValues.None;
				default:
					return C.TickLabelPositionValues.NextTo;
			}
		}
	}
	/// <summary>
	/// X Axis Specific Options
	/// </summary>
	public class XAxisOptions : AxisOptions { }
	/// <summary>
	/// Y Axis Specific Options
	/// </summary>
	public class YAxisOptions : AxisOptions { }
	/// <summary>
	/// Z Axis Specific Options
	/// </summary>
	public class ZAxisOptions : AxisOptions { }
	/// <summary>
	/// Represents the options for a chart axis.
	/// </summary>
	public class ChartAxisOptions
	{
		/// <summary>
		/// X-Axis and Axes options
		/// </summary>
		public XAxisOptions xAxisOptions = new XAxisOptions();
		/// <summary>
		/// Y-Axis and Axes options
		/// </summary>
		public YAxisOptions yAxisOptions = new YAxisOptions();
		/// <summary>
		/// Z-Axis and Axes options
		/// Totally optional for secondary action options
		/// TODO : Implementation
		/// </summary>
		public ZAxisOptions zAxisOptions = new ZAxisOptions();
	}
	/// <summary>
	/// Represents the grouping options for chart data.
	/// </summary>
	public class ChartDataGrouping
	{
		/// <summary>
		/// Gets or sets the data label cells.
		/// </summary>
		public ChartData[] dataLabelCells;
		/// <summary>
		/// Gets or sets the data label formula.
		/// </summary>
		public string dataLabelFormula;
		/// <summary>
		/// Gets or sets the series header cells.
		/// </summary>
		public ChartData seriesHeaderCells;
		/// <summary>
		/// Gets or sets the series header format.
		/// </summary>
		public string seriesHeaderFormat;
		/// <summary>
		/// Gets or sets the series header formula.
		/// </summary>
		public string seriesHeaderFormula;
		/// <summary>
		/// Gets or sets the X-axis cells.
		/// </summary>
		public ChartData[] xAxisCells;
		/// <summary>
		/// Gets or sets the X-axis format.
		/// </summary>
		public string xAxisFormat;
		/// <summary>
		/// Gets or sets the X-axis formula.
		/// </summary>
		public string xAxisFormula;
		/// <summary>
		/// Gets or sets the Y-axis cells.
		/// </summary>
		public ChartData[] yAxisCells;
		/// <summary>
		/// Gets or sets the Y-axis format.
		/// </summary>
		public string yAxisFormat;
		/// <summary>
		/// Gets or sets the Y-axis formula.
		/// </summary>
		public string yAxisFormula;
		/// <summary>
		/// Gets or sets the Z-axis cells.
		/// </summary>
		public ChartData[] zAxisCells;
		/// <summary>
		/// Gets or sets the Z-axis format.
		/// </summary>
		public string zAxisFormat;
		/// <summary>
		/// Gets or sets the Z-axis formula.
		/// </summary>
		public string zAxisFormula;
		/// <summary>
		///
		/// </summary>
		public int id;
	}
	/// <summary>
	/// Represents the options for chart data labels.
	/// </summary>
	public class ChartDataLabel : TextOptions
	{
		/// <summary>
		/// The separator used for displaying multiple values.
		/// </summary>
		public string separator = ", ";
		/// <summary>
		/// Determines whether to show the category name in the chart.
		/// </summary>
		public bool showCategoryName;
		/// <summary>
		/// Determines whether to show the legend key in the chart.
		/// </summary>
		public bool showLegendKey;
		/// <summary>
		/// Determines whether to show the series name in the chart.
		/// </summary>
		public bool showSeriesName;
		/// <summary>
		/// Determines whether to show the value in the chart.
		/// </summary>
		public bool showValue;
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
		internal bool is3dData;
		/// <summary>
		/// Use 2013 Version Data Label Option
		/// </summary>
		/// <remarks>This Property May get updated in future be in lookout.</remarks>
		public AdvancedDataLabel advancedDataLabel = new AdvancedDataLabel();
	}
	/// <summary>
	/// Represents the options for chart grid lines.
	/// </summary>
	public class ChartGridLinesOptions
	{
		/// <summary>
		/// Is Major Category Lines Enabled
		/// </summary>
		public bool isMajorCategoryLinesEnabled;
		/// <summary>
		/// Is Major Value Lines Enabled
		/// </summary>
		public bool isMajorValueLinesEnabled = true;
		/// <summary>
		/// Is Minor Category Lines Enabled
		/// </summary>
		public bool isMinorCategoryLinesEnabled;
		/// <summary>
		/// Is Minor Value Lines Enabled
		/// </summary>
		public bool isMinorValueLinesEnabled;
	}
	/// <summary>
	/// Represents the options for chart legend.
	/// </summary>
	public class ChartLegendOptions : TextOptions
	{
		/// <summary>
		/// Is Legend Enabled
		/// </summary>
		public bool isEnableLegend = true;
		/// <summary>
		/// Is Legend Chart OverLap
		/// </summary>
		public bool isLegendChartOverLap;
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
		public LayoutModel manualLayout;
	}
	/// <summary>
	///
	/// </summary>
	public class ChartDataPointSettings
	{
		/// <summary>
		/// The color of the fill.
		/// </summary>
		public string fillColor;
		/// <summary>
		///
		/// </summary>
		public string borderColor;
	}
	/// <summary>
	/// Represents the settings for a chart series.
	/// </summary>
	public class ChartSeriesSetting
	{
		/// <summary>
		/// The color of the border.
		/// </summary>
		public virtual string borderColor { get; set; }
		internal ChartSeriesSetting() { }
	}
	/// <summary>
	///
	/// </summary>
	public class PlotAreaModel
	{
		/// <summary>
		/// Manual Position Char Graph Area
		/// </summary>
		public LayoutModel manualLayout;
	}
	/// <summary>
	///
	/// </summary>
	public class ChartTitleModel : ChartTextOptions { }
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
		public AnchorPosition from = new AnchorPosition();
		/// <summary>
		///
		/// </summary>
		public AnchorPosition to = new AnchorPosition();
	}
	/// <summary>
	/// Represents the settings for a chart.
	/// </summary>
	public class ChartSetting<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition, new()
	{
		internal uint? categoryAxisId;
		internal uint? valueAxisId;
		internal bool is3DChart;
		/// <summary>
		///
		/// </summary>
		public HyperlinkProperties hyperlinkProperties;
		/// <summary>
		/// Only useful when used with Combo chart
		/// </summary>
		public bool isSecondaryAxis;
		/// <summary>
		///
		/// </summary>
		public PlotAreaModel plotAreaOptions;
		/// <summary>
		/// Chart Data Setting
		/// </summary>
		public ChartDataSetting chartDataSetting = new ChartDataSetting();
		/// <summary>
		/// Chart Grid Line Options
		/// </summary>
		public ChartGridLinesOptions chartGridLinesOptions = new ChartGridLinesOptions();
		/// <summary>
		/// Chart Legend Options
		/// </summary>
		public ChartLegendOptions chartLegendOptions = new ChartLegendOptions();
		/// <summary>
		/// Chart Title
		/// </summary>
		public ChartTitleModel titleOptions;
		/// <summary>
		///
		/// </summary>
		public ApplicationSpecificSetting applicationSpecificSetting = new ApplicationSpecificSetting();
		internal ChartSetting() { }
	}

	/// <summary>
	/// Represents the Axis settings for chart.
	/// </summary>
	public class AxisSetting<AxisTypeOption> where AxisTypeOption : AxisOptions, new()
	{
		internal uint id;
		internal uint crossAxisId;
		internal AxisPosition axisPosition;
		internal AxisTypeOption axisOptions;
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
	public class TrendLineModel
	{
		/// <summary>
		/// Use for Order value if 'Polynomial' type else Period for 'Moving Average'
		/// Default 2
		/// </summary>
		public int secondaryValue = 2;
		/// <summary>
		/// 
		/// </summary>
		public bool setIntercept = false;
		/// <summary>
		/// 
		/// </summary>
		public float interceptValue = 0.0F;
		/// <summary>
		/// 
		/// </summary>
		public bool showEquation = false;
		/// <summary>
		/// 
		/// </summary>
		public bool showRSquareValue = false;
		/// <summary>
		/// Default 0.0
		/// </summary>
		public float forecastForward = 0.0F;
		/// <summary>
		/// Default 0.0
		/// </summary>
		public float forecastBackward = 0.0F;
		/// <summary>
		/// This is to set custom Trending if null it will assume automatic 
		/// </summary>
		public string trendLineName = null;
		/// <summary>
		/// 
		/// </summary>
		public TrendLineTypes trendLineType = TrendLineTypes.NONE;
		internal static C.TrendlineValues GetTrendlineValues(TrendLineTypes trendLineType)
		{
			switch (trendLineType)
			{
				case TrendLineTypes.EXPONENTIAL:
					return C.TrendlineValues.Exponential;
				case TrendLineTypes.LINEAR:
					return C.TrendlineValues.Linear;
				case TrendLineTypes.LOGARITHMIC:
					return C.TrendlineValues.Logarithmic;
				case TrendLineTypes.POLYNOMIAL:
					return C.TrendlineValues.Polynomial;
				case TrendLineTypes.POWER:
					return C.TrendlineValues.Power;
				case TrendLineTypes.MOVING_AVERAGE:
					return C.TrendlineValues.MovingAverage;
				default:
					return C.TrendlineValues.MovingAverage;
			}
		}
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
		public MarkerShapeTypes markerShapeType = MarkerShapeTypes.NONE;
		/// <summary>
		///
		/// </summary>
		public ShapePropertiesModel shapeProperties = new ShapePropertiesModel();
		internal static C.MarkerStyleValues GetMarkerStyleValues(MarkerShapeTypes markerShapeType)
		{
			switch (markerShapeType)
			{
				case MarkerShapeTypes.AUTO:
					return C.MarkerStyleValues.Auto;
				case MarkerShapeTypes.CIRCLE:
					return C.MarkerStyleValues.Circle;
				case MarkerShapeTypes.DASH:
					return C.MarkerStyleValues.Dash;
				case MarkerShapeTypes.DIAMOND:
					return C.MarkerStyleValues.Diamond;
				case MarkerShapeTypes.DOT:
					return C.MarkerStyleValues.Dot;
				case MarkerShapeTypes.PICTURE:
					return C.MarkerStyleValues.Picture;
				case MarkerShapeTypes.PLUS:
					return C.MarkerStyleValues.Plus;
				case MarkerShapeTypes.SQUARE:
					return C.MarkerStyleValues.Square;
				case MarkerShapeTypes.STAR:
					return C.MarkerStyleValues.Star;
				case MarkerShapeTypes.TRIANGLE:
					return C.MarkerStyleValues.Triangle;
				case MarkerShapeTypes.X:
					return C.MarkerStyleValues.X;
				default:
					return C.MarkerStyleValues.None;
			}
		}
	}
}
