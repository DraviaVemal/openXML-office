// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	///
	/// </summary>
	public enum BorderStyleValues
	{
		/// <summary>
		///
		/// </summary>
		SINGEL,
		/// <summary>
		///
		/// </summary>
		DOUBLE,
		/// <summary>
		///
		/// </summary>
		TRIPLE,
		/// <summary>
		///
		/// </summary>
		THICK_THIN,
		/// <summary>
		///
		/// </summary>
		THIN_THICK,
	}
	/// <summary>
	///
	/// </summary>
	public enum DrawingBeginArrowValues
	{
		/// <summary>
		///
		/// </summary>
		NONE,
		/// <summary>
		///
		/// </summary>
		ARROW,
		/// <summary>
		///
		/// </summary>
		DIAMOND,
		/// <summary>
		///
		/// </summary>
		OVAL,
		/// <summary>
		///
		/// </summary>
		STEALTH,
		/// <summary>
		///
		/// </summary>
		TRIANGLE
	}
	/// <summary>
	///
	/// </summary>
	public enum DrawingEndArrowValues
	{
		/// <summary>
		///
		/// </summary>
		NONE,
		/// <summary>
		///
		/// </summary>
		ARROW,
		/// <summary>
		///
		/// </summary>
		DIAMOND,
		/// <summary>
		///
		/// </summary>
		OVAL,
		/// <summary>
		///
		/// </summary>
		STEALTH,
		/// <summary>
		///
		/// </summary>
		TRIANGLE
	}
	/// <summary>
	///
	/// </summary>
	public enum DrawingPresetLineDashValues
	{
		/// <summary>
		///
		/// </summary>
		DASH,
		/// <summary>
		///
		/// </summary>
		DASH_DOT,
		/// <summary>
		///
		/// </summary>
		DOT,
		/// <summary>
		///
		/// </summary>
		LARGE_DASH,
		/// <summary>
		///
		/// </summary>
		LARGE_DASH_DOT,
		/// <summary>
		///
		/// </summary>
		LARGE_DASH_DOT_DOT,
		/// <summary>
		///
		/// </summary>
		SOLID,
		/// <summary>
		///
		/// </summary>
		SYSTEM_DASH,
		/// <summary>
		///
		/// </summary>
		SYSTEM_DASH_DOT,
		/// <summary>
		///
		/// </summary>
		SYSTEM_DASH_DOT_DOT,
		/// <summary>
		///
		/// </summary>
		SYSTEM_DOT,
	}
	/// <summary>
	///
	/// </summary>
	public enum LineWidthValues
	{
		/// <summary>
		///
		/// </summary>
		SMALL,
		/// <summary>
		///
		/// </summary>
		MEDIUM,
		/// <summary>
		///
		/// </summary>
		LARGE
	}
	/// <summary>
	///
	/// </summary>
	public enum ThemeColorValues
	{
		/// <summary>
		///
		/// </summary>
		ACCENT_1,
		/// <summary>
		///
		/// </summary>
		ACCENT_2,
		/// <summary>
		///
		/// </summary>
		ACCENT_3,
		/// <summary>
		///
		/// </summary>
		ACCENT_4,
		/// <summary>
		///
		/// </summary>
		ACCENT_5,
		/// <summary>
		///
		/// </summary>
		ACCENT_6,
		/// <summary>
		///
		/// </summary>
		DARK_1,
		/// <summary>
		///
		/// </summary>
		DARK_2,
		/// <summary>
		///
		/// </summary>
		BACKGROUND_1,
		/// <summary>
		///
		/// </summary>
		BACKGROUND_2,
		/// <summary>
		///
		/// </summary>
		LIGHT_1,
		/// <summary>
		///
		/// </summary>
		LIGHT_2,
		/// <summary>
		///
		/// </summary>
		TEXT_1,
		/// <summary>
		///
		/// </summary>
		TEXT_2,
		/// <summary>
		///
		/// </summary>
		HYPERLINK,
		/// <summary>
		///
		/// </summary>
		FOLLOW_HYPERLINK,
		/// <summary>
		///
		/// </summary>
		TRANSPARENT
	}
	/// <summary>
	///
	/// </summary>
	public enum OutlineCapTypeValues
	{
		/// <summary>
		///
		/// </summary>
		FLAT,
		/// <summary>
		///
		/// </summary>
		SQUARE,
		/// <summary>
		///
		/// </summary>
		ROUND,
	}
	/// <summary>
	///
	/// </summary>
	public enum OutlineLineTypeValues
	{
		/// <summary>
		///
		/// </summary>
		SINGLE,
		/// <summary>
		///
		/// </summary>
		DOUBLE,
		/// <summary>
		///
		/// </summary>
		TRIPLE,
		/// <summary>
		///
		/// </summary>
		THICK_THIN,
		/// <summary>
		///
		/// </summary>
		THIN_THICK,
	}
	/// <summary>
	///
	/// </summary>
	public enum TextVerticalAlignmentValues
	{
		/// <summary>
		///
		/// </summary>
		EAST_ASIAN_VERTICAL,
		/// <summary>
		///
		/// </summary>
		HORIZONTAL,
		/// <summary>
		///
		/// </summary>
		MONGOLIAN_VERTICAL,
		/// <summary>
		///
		/// </summary>
		VERTICAL,
		/// <summary>
		///
		/// </summary>
		VERTICAL_270,
		/// <summary>
		///
		/// </summary>
		WORD_ART_LEFT_TO_RIGHT,
		/// <summary>
		///
		/// </summary>
		WORD_ART_VERTICAL,
	}
	/// <summary>
	///
	/// </summary>
	public enum TextWrappingValues
	{
		/// <summary>
		///
		/// </summary>
		NONE,
		/// <summary>
		///
		/// </summary>
		SQUARE,
	}
	/// <summary>
	///
	/// </summary>
	public enum TextVerticalOverflowValues
	{
		/// <summary>
		///
		/// </summary>
		CLIP,
		/// <summary>
		///
		/// </summary>
		ELLIPSIS,
		/// <summary>
		///
		/// </summary>
		OVERFLOW,
	}
	/// <summary>
	///
	/// </summary>
	public enum TextAnchoringValues
	{
		/// <summary>
		///
		/// </summary>
		BOTTOM,
		/// <summary>
		///
		/// </summary>
		TOP,
		/// <summary>
		///
		/// </summary>
		CENTER
	}
	/// <summary>
	///
	/// </summary>
	public enum OutlineAlignmentValues
	{
		/// <summary>
		///
		/// </summary>
		CENTER,
		/// <summary>
		///
		/// </summary>
		INSERT,
	}

	/// <summary>
	/// Central Text Options
	/// </summary>
	public class TextOptions
	{
		/// <summary>
		/// 
		/// </summary>
		public string textValue;
		/// <summary>
		/// Is Font Bold
		/// </summary>
		public bool isBold;
		/// <summary>
		/// Is Font Italic
		/// </summary>
		public bool isItalic;
		/// <summary>
		///  Font Size
		/// </summary>
		public float fontSize = 11.97F;
		/// <summary>
		///
		/// </summary>
		public string fontColor;
		/// <summary>
		///
		/// </summary>
		public UnderLineValues underLineValues = UnderLineValues.NONE;
		/// <summary>
		///
		/// </summary>
		public StrikeValues strikeValues = StrikeValues.NO_STRIKE;
		/// <summary>
		///
		/// </summary>
		public string fontFamily = "(Calibri (Body))";
	}
	/// <summary>
	///
	/// </summary>
	public class OutlineModel<LineColorOption> where LineColorOption : class, IColorOptions, new()
	{
		/// <summary>
		///
		/// </summary>
		public int? width;
		/// <summary>
		///
		/// </summary>
		public ColorOptionModel<LineColorOption> lineColor;
		/// <summary>
		///
		/// </summary>
		public OutlineCapTypeValues? outlineCapTypeValues = OutlineCapTypeValues.FLAT;
		/// <summary>
		///
		/// </summary>
		public OutlineLineTypeValues? outlineLineTypeValues = OutlineLineTypeValues.SINGLE;
		/// <summary>
		///
		/// </summary>
		public OutlineAlignmentValues? outlineAlignmentValues = OutlineAlignmentValues.CENTER;
		/// <summary>
		///
		/// </summary>
		public DrawingPresetLineDashValues? dashType;
		/// <summary>
		///
		/// </summary>
		public DrawingBeginArrowValues beginArrowValues = DrawingBeginArrowValues.NONE;
		/// <summary>
		///
		/// </summary>
		public DrawingEndArrowValues endArrowValues = DrawingEndArrowValues.NONE;
		/// <summary>
		///
		/// </summary>
		public LineWidthValues? lineEndWidth;
		/// <summary>
		///
		/// </summary>
		public LineWidthValues? lineStartWidth;
	}
	/// <summary>
	///
	/// </summary>
	public class SchemeColorModel
	{
		/// <summary>
		///
		/// </summary>
		public ThemeColorValues themeColorValues = ThemeColorValues.TRANSPARENT;
		/// <summary>
		///
		/// </summary>
		public int? tint;
		/// <summary>
		///
		/// </summary>
		public int? shade;
		/// <summary>
		///
		/// </summary>
		public int? saturationModulation;
		/// <summary>
		///
		/// </summary>
		public int? saturationOffset;
		/// <summary>
		///
		/// </summary>
		public int? luminanceModulation;
		/// <summary>
		///
		/// </summary>
		public int? luminanceOffset;
	}
	/// <summary>
	///
	/// </summary>
	public class EffectListModel
	{
	}

	/// <summary>
	///	Base interface for color options used with FillColorModel
	/// </summary>
	public interface IColorOptions { }
	/// <summary>
	///	This will update NoFill Color as result
	/// </summary>
	public class NoFillOptions : IColorOptions { }
	/// <summary>
	///	Solid Fill options
	/// </summary>
	public class SolidOptions : IColorOptions
	{
		/// <summary>
		///
		/// </summary>
		public string hexColor;
		/// <summary>
		///
		/// </summary>
		public SchemeColorModel schemeColorModel;
		/// <summary>
		///
		/// </summary>
		public int? transparency;
	}
	/// <summary>
	///
	/// </summary>
	public class GradientOptions : IColorOptions { }
	/// <summary>
	///
	/// </summary>
	public class PictureOrTextureOptions : IColorOptions { }
	/// <summary>
	///
	/// </summary>
	public class PatternOptions : IColorOptions { }
	/// <summary>
	///
	/// </summary>
	public class AutomaticOptions : IColorOptions { }
	/// <summary>
	/// Fill Color Type Options Use Generic to choose the apply style and update the respective options
	/// </summary>
	public class ColorOptionModel<ColorOption> where ColorOption : class, IColorOptions, new()
	{
		/// <summary>
		/// Give you the corresponding color option
		/// </summary>
		public ColorOption colorOption = new ColorOption();
	}
	/// <summary>
	///
	/// </summary>
	public class ShapePropertiesModel<LineColorOption, FillColorOption>
		where LineColorOption : class, IColorOptions, new()
		where FillColorOption : class, IColorOptions, new()
	{
		/// <summary>
		///
		/// </summary>
		public ColorOptionModel<FillColorOption> fillColor = new ColorOptionModel<FillColorOption>();
		/// <summary>
		///
		/// </summary>
		public OutlineModel<LineColorOption> lineColor = new OutlineModel<LineColorOption>();
		/// <summary>
		///
		/// </summary>
		public EffectListModel effectList;
		/// <summary>
		///
		/// </summary>
		public ShapeProperty3D shapeProperty3D;
	}
	/// <summary>
	///
	/// </summary>
	public class ShapeProperty3D { }
	/// <summary>
	///
	/// </summary>
	public enum UnderLineValues
	{
		/// <summary>
		///
		/// </summary>
		NONE,
		/// <summary>
		///
		/// </summary>
		DASH,
		/// <summary>
		///
		/// </summary>
		DASH_HEAVY,
		/// <summary>
		///
		/// </summary>
		DASH_LONG,
		/// <summary>
		///
		/// </summary>
		DASH_LONG_HEAVY,
		/// <summary>
		///
		/// </summary>
		DOT_DASH,
		/// <summary>
		///
		/// </summary>
		DOT_DASH_HEAVY,
		/// <summary>
		///
		/// </summary>
		DOT_DOT_DASH,
		/// <summary>
		///
		/// </summary>
		DOT_DOT_DASH_HEAVY,
		/// <summary>
		///
		/// </summary>
		DOTTED,
		/// <summary>
		///
		/// </summary>
		DOUBLE,
		/// <summary>
		///
		/// </summary>
		HEAVY,
		/// <summary>
		///
		/// </summary>
		HEAVY_DOTTED,
		/// <summary>
		///
		/// </summary>
		SINGLE,
		/// <summary>
		///
		/// </summary>
		WAVY,
		/// <summary>
		///
		/// </summary>
		WAVY_DOUBLE,
		/// <summary>
		///
		/// </summary>
		WAVY_HEAVY,
		/// <summary>
		///
		/// </summary>
		WORDS,
	}
	/// <summary>
	///
	/// </summary>
	public enum StrikeValues
	{
		/// <summary>
		///
		/// </summary>
		NO_STRIKE,
		/// <summary>
		///
		/// </summary>
		SINGLE_STRIKE,
		/// <summary>
		///
		/// </summary>
		DOUBLE_STRIKE,
	}
	/// <summary>
	///
	/// </summary>
	public enum HyperlinkPropertyTypeValues
	{
		/// <summary>
		///
		/// </summary>
		EXISTING_FILE,
		/// <summary>
		///
		/// </summary>
		WEB_URL,
		/// <summary>
		/// 
		/// </summary>
		TARGET_SHEET,
		/// <summary>
		///
		/// </summary>
		TARGET_SLIDE,
		/// <summary>
		///
		/// </summary>
		NEXT_SLIDE,
		/// <summary>
		///
		/// </summary>
		PREVIOUS_SLIDE,
		/// <summary>
		///
		/// </summary>
		FIRST_SLIDE,
		/// <summary>
		///
		/// </summary>
		LAST_SLIDE,
	}
	/// <summary>
	///
	/// </summary>
	public class HyperlinkProperties
	{
		/// <summary>
		/// Hyperlink option type additional address detail can be set using value
		/// </summary>
		public HyperlinkPropertyTypeValues hyperlinkPropertyType = HyperlinkPropertyTypeValues.WEB_URL;
		/// <summary>
		///	Web Url or file path or slide ID use based on "HyperlinkPropertyType"
		/// </summary>
		public string value;
		/// <summary>
		/// Screen Tool Tip
		/// </summary>
		public string toolTip;
		/// <summary>
		/// Internal Use property
		/// </summary>
		public string relationId;
		/// <summary>
		/// Internal Use property
		/// </summary>
		public string action;
	}
	/// <summary>
	///
	/// </summary>
	public class DrawingRunPropertiesModel<TextColorOption> : TextOptions
	where TextColorOption : class, IColorOptions, new()
	{
		/// <summary>
		///
		/// </summary>
		public HyperlinkProperties hyperlinkProperties;
		/// <summary>
		///
		/// </summary>
		public ColorOptionModel<TextColorOption> textColorOption;
	}
	/// <summary>
	///
	/// </summary>
	public class DefaultRunPropertiesModel<TextColorOption> : TextOptions
	where TextColorOption : class, IColorOptions, new()
	{
		/// <summary>
		///
		/// </summary>
		public ColorOptionModel<TextColorOption> textColorOption;
		/// <summary>
		///
		/// </summary>
		public string latinFont;
		/// <summary>
		///
		/// </summary>
		public string eastAsianFont;
		/// <summary>
		///
		/// </summary>
		public string complexScriptFont;
		/// <summary>
		///
		/// </summary>
		public int? kerning;
		/// <summary>
		///
		/// </summary>
		public int? baseline;
	}
	/// <summary>
	///
	/// </summary>
	public class DrawingRunModel<TextColorOption>
	where TextColorOption : class, IColorOptions, new()
	{
		/// <summary>
		///
		/// </summary>
		public DrawingRunPropertiesModel<TextColorOption> drawingRunProperties = new DrawingRunPropertiesModel<TextColorOption>();
		/// <summary>
		///
		/// </summary>
		public string textHighlight;
		/// <summary>
		///
		/// </summary>
		public string text;
	}
	/// <summary>
	///
	/// </summary>
	public class DrawingParagraphModel<TextColorOption>
	where TextColorOption : class, IColorOptions, new()
	{
		/// <summary>
		///
		/// </summary>
		public ParagraphPropertiesModel<TextColorOption> paragraphPropertiesModel;
		/// <summary>
		///
		/// </summary>
		public DrawingRunModel<TextColorOption>[] drawingRuns;
	}
	/// <summary>
	/// /
	/// </summary>
	public class ParagraphPropertiesModel<TextColorOption>
	where TextColorOption : class, IColorOptions, new()
	{
		/// <summary>
		///
		/// </summary>
		public DefaultRunPropertiesModel<TextColorOption> defaultRunProperties;
		/// <summary>
		/// Cell Alignment Option
		/// </summary>
		public HorizontalAlignmentValues? horizontalAlignment;
	}
	/// <summary>
	///
	/// </summary>
	public class DrawingBodyPropertiesModel
	{
		/// <summary>
		///
		/// </summary>
		public int rotation = 0;
		/// <summary>
		///
		/// </summary>
		public int? leftInset;
		/// <summary>
		///
		/// </summary>
		public int? topInset;
		/// <summary>
		///
		/// </summary>
		public int? rightInset;
		/// <summary>
		///
		/// </summary>
		public int? bottomInset;
		/// <summary>
		///
		/// </summary>
		public bool? useParagraphSpacing;
		/// <summary>
		///
		/// </summary>
		public TextVerticalOverflowValues? verticalOverflow;
		/// <summary>
		///
		/// </summary>
		public TextVerticalAlignmentValues? vertical;
		/// <summary>
		///
		/// </summary>
		public TextWrappingValues? wrap;
		/// <summary>
		///
		/// </summary>
		public TextAnchoringValues? anchor;
		/// <summary>
		///
		/// </summary>
		public bool? anchorCenter;
	}
	/// <summary>
	///
	/// </summary>
	public class ChartTextPropertiesModel<TextColorOption>
	where TextColorOption : class, IColorOptions, new()
	{
		/// <summary>
		///
		/// </summary>
		public DrawingBodyPropertiesModel drawingBodyProperties;
		/// <summary>
		///
		/// </summary>
		public DrawingParagraphModel<TextColorOption> drawingParagraph;
	}
}
