// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Common Properties organized in one place to get inherited by child classes
	/// </summary>
	public class CommonProperties
	{
		/// <summary>
		///
		/// </summary>
		public static A.TextAlignmentTypeValues GetTextAlignmentValue(HorizontalAlignmentValues horizontalAlignmentValues)
		{
			switch (horizontalAlignmentValues)
			{
				case HorizontalAlignmentValues.RIGHT:
					return A.TextAlignmentTypeValues.Right;
				case HorizontalAlignmentValues.JUSTIFY:
					return A.TextAlignmentTypeValues.Justified;
				case HorizontalAlignmentValues.CENTER:
					return A.TextAlignmentTypeValues.Center;
				default:
					return A.TextAlignmentTypeValues.Left;
			}
		}
		/// <summary>
		///
		/// </summary>
		public static A.CompoundLineValues GetBorderStyleValue(BorderStyleValues borderStyle)
		{
			switch (borderStyle)
			{
				case BorderStyleValues.DOUBLE:
					return A.CompoundLineValues.Double;
				case BorderStyleValues.TRIPLE:
					return A.CompoundLineValues.Triple;
				case BorderStyleValues.THICK_THIN:
					return A.CompoundLineValues.ThickThin;
				case BorderStyleValues.THIN_THICK:
					return A.CompoundLineValues.ThinThick;
				default:
					return A.CompoundLineValues.Single;
			}
		}
		/// <summary>
		///
		/// </summary>
		public static A.PresetLineDashValues GetDashStyleValue(DrawingPresetLineDashValues dashStyle)
		{
			switch (dashStyle)
			{
				case DrawingPresetLineDashValues.DASH:
					return A.PresetLineDashValues.Dash;
				case DrawingPresetLineDashValues.DASH_DOT:
					return A.PresetLineDashValues.DashDot;
				case DrawingPresetLineDashValues.DOT:
					return A.PresetLineDashValues.Dot;
				case DrawingPresetLineDashValues.LARGE_DASH:
					return A.PresetLineDashValues.LargeDash;
				case DrawingPresetLineDashValues.LARGE_DASH_DOT:
					return A.PresetLineDashValues.LargeDashDot;
				case DrawingPresetLineDashValues.LARGE_DASH_DOT_DOT:
					return A.PresetLineDashValues.LargeDashDotDot;
				case DrawingPresetLineDashValues.SYSTEM_DASH:
					return A.PresetLineDashValues.SystemDash;
				case DrawingPresetLineDashValues.SYSTEM_DASH_DOT:
					return A.PresetLineDashValues.SystemDashDot;
				case DrawingPresetLineDashValues.SYSTEM_DASH_DOT_DOT:
					return A.PresetLineDashValues.SystemDashDotDot;
				case DrawingPresetLineDashValues.SYSTEM_DOT:
					return A.PresetLineDashValues.SystemDot;
				default:
					return A.PresetLineDashValues.Solid;
			}
		}
		/// <summary>
		///
		/// </summary>
		public static A.LineEndValues GetEndArrowValue(DrawingEndArrowValues endArrowValues)
		{
			switch (endArrowValues)
			{
				case DrawingEndArrowValues.ARROW:
					return A.LineEndValues.Arrow;
				case DrawingEndArrowValues.DIAMOND:
					return A.LineEndValues.Diamond;
				case DrawingEndArrowValues.OVAL:
					return A.LineEndValues.Oval;
				case DrawingEndArrowValues.STEALTH:
					return A.LineEndValues.Stealth;
				case DrawingEndArrowValues.TRIANGLE:
					return A.LineEndValues.Triangle;
				default:
					return A.LineEndValues.None;
			}
		}
		/// <summary>
		///
		/// </summary>
		public static A.LineEndLengthValues GetLineEndLengthValue(LineWidthValues lineEndWidth)
		{
			switch (lineEndWidth)
			{
				case LineWidthValues.LARGE:
					return A.LineEndLengthValues.Large;
				case LineWidthValues.MEDIUM:
					return A.LineEndLengthValues.Medium;
				default:
					return A.LineEndLengthValues.Small;
			}
		}
		/// <summary>
		///
		/// </summary>
		public static A.LineEndWidthValues GetLineEndWidthValue(LineWidthValues lineEndWidth)
		{
			switch (lineEndWidth)
			{
				case LineWidthValues.LARGE:
					return A.LineEndWidthValues.Large;
				case LineWidthValues.MEDIUM:
					return A.LineEndWidthValues.Medium;
				default:
					return A.LineEndWidthValues.Small;
			}
		}
		/// <summary>
		///
		/// </summary>
		public static A.LineEndLengthValues GetLineStartLengthValue(LineWidthValues lineStartWidth)
		{
			switch (lineStartWidth)
			{
				case LineWidthValues.LARGE:
					return A.LineEndLengthValues.Large;
				case LineWidthValues.MEDIUM:
					return A.LineEndLengthValues.Medium;
				default:
					return A.LineEndLengthValues.Small;
			}
		}
		/// <summary>
		///
		/// </summary>
		public static A.LineEndValues GetBeginArrowValue(DrawingBeginArrowValues beginArrowValues)
		{
			switch (beginArrowValues)
			{
				case DrawingBeginArrowValues.ARROW:
					return A.LineEndValues.Arrow;
				case DrawingBeginArrowValues.DIAMOND:
					return A.LineEndValues.Diamond;
				case DrawingBeginArrowValues.OVAL:
					return A.LineEndValues.Oval;
				case DrawingBeginArrowValues.STEALTH:
					return A.LineEndValues.Stealth;
				case DrawingBeginArrowValues.TRIANGLE:
					return A.LineEndValues.Triangle;
				default:
					return A.LineEndValues.None;
			}
		}
		/// <summary>
		///
		/// </summary>
		public static A.LineEndWidthValues GetLineStartWidthValue(LineWidthValues lineStartWidth)
		{
			switch (lineStartWidth)
			{
				case LineWidthValues.LARGE:
					return A.LineEndWidthValues.Large;
				case LineWidthValues.MEDIUM:
					return A.LineEndWidthValues.Medium;
				default:
					return A.LineEndWidthValues.Small;
			}
		}
		internal static A.TextAnchoringTypeValues GetAnchorValues(TextAnchoringValues textAnchoring)
		{
			switch (textAnchoring)
			{
				case TextAnchoringValues.BOTTOM:
					return A.TextAnchoringTypeValues.Bottom;
				case TextAnchoringValues.CENTER:
					return A.TextAnchoringTypeValues.Center;
				default:
					return A.TextAnchoringTypeValues.Top;
			}
		}
		internal static A.TextVerticalValues GetTextVerticalAlignmentValues(TextVerticalAlignmentValues textVerticalAlignment)
		{
			switch (textVerticalAlignment)
			{
				case TextVerticalAlignmentValues.EAST_ASIAN_VERTICAL:
					return A.TextVerticalValues.EastAsianVetical;
				case TextVerticalAlignmentValues.HORIZONTAL:
					return A.TextVerticalValues.Horizontal;
				case TextVerticalAlignmentValues.MONGOLIAN_VERTICAL:
					return A.TextVerticalValues.MongolianVertical;
				case TextVerticalAlignmentValues.VERTICAL:
					return A.TextVerticalValues.Vertical;
				case TextVerticalAlignmentValues.VERTICAL_270:
					return A.TextVerticalValues.Vertical270;
				case TextVerticalAlignmentValues.WORD_ART_LEFT_TO_RIGHT:
					return A.TextVerticalValues.WordArtVertical;
				default:
					return A.TextVerticalValues.WordArtVertical;
			}
		}
		internal static A.TextVerticalOverflowValues GetTextVerticalOverflowValues(TextVerticalOverflowValues textVerticalOverflow)
		{
			switch (textVerticalOverflow)
			{
				case TextVerticalOverflowValues.CLIP:
					return A.TextVerticalOverflowValues.Clip;
				case TextVerticalOverflowValues.ELLIPSIS:
					return A.TextVerticalOverflowValues.Ellipsis;
				default:
					return A.TextVerticalOverflowValues.Overflow;
			}
		}
		internal static A.TextWrappingValues GetWarpingValues(TextWrappingValues textWrapping)
		{
			switch (textWrapping)
			{
				case TextWrappingValues.NONE:
					return A.TextWrappingValues.None;
				default:
					return A.TextWrappingValues.Square;
			}
		}
		internal static A.TextStrikeValues GetTextStrikeValues(StrikeValues strikeValues)
		{
			switch (strikeValues)
			{
				case StrikeValues.SINGLE_STRIKE:
					return A.TextStrikeValues.SingleStrike;
				case StrikeValues.DOUBLE_STRIKE:
					return A.TextStrikeValues.DoubleStrike;
				default:
					return A.TextStrikeValues.NoStrike;
			}
		}
		internal static string GetSchemeColorValuesText(ThemeColorValues themeColorValues)
		{
			switch (themeColorValues)
			{
				case ThemeColorValues.ACCENT_1:
					return "accent1";
				case ThemeColorValues.ACCENT_2:
					return "accent2";
				case ThemeColorValues.ACCENT_3:
					return "accent3";
				case ThemeColorValues.ACCENT_4:
					return "accent4";
				case ThemeColorValues.ACCENT_5:
					return "accent5";
				case ThemeColorValues.ACCENT_6:
					return "accent6";
				case ThemeColorValues.DARK_1:
					return "dk1";
				case ThemeColorValues.DARK_2:
					return "dk2";
				case ThemeColorValues.BACKGROUND_1:
					return "bg1";
				case ThemeColorValues.BACKGROUND_2:
					return "bg2";
				case ThemeColorValues.LIGHT_1:
					return "lt1";
				case ThemeColorValues.LIGHT_2:
					return "lt2";
				case ThemeColorValues.TEXT_1:
					return "tx1";
				case ThemeColorValues.TEXT_2:
					return "tx2";
				case ThemeColorValues.HYPERLINK:
					return "hlink";
				case ThemeColorValues.FOLLOW_HYPERLINK:
					return "folHlink";
				default:
					return "phClr";
			}
		}
		internal static A.PenAlignmentValues GetLineAlignmentValues(OutlineAlignmentValues outlineAlignmentValues)
		{
			switch (outlineAlignmentValues)
			{
				case OutlineAlignmentValues.CENTER:
					return A.PenAlignmentValues.Center;
				default:
					return A.PenAlignmentValues.Insert;
			}
		}
		internal static A.LineCapValues GetLineCapValues(OutlineCapTypeValues outlineCapTypeValues)
		{
			switch (outlineCapTypeValues)
			{
				case OutlineCapTypeValues.SQUARE:
					return A.LineCapValues.Square;
				case OutlineCapTypeValues.ROUND:
					return A.LineCapValues.Round;
				default:
					return A.LineCapValues.Flat;
			}
		}
		internal static A.CompoundLineValues GetLineTypeValues(OutlineLineTypeValues outlineLineTypeValues)
		{
			switch (outlineLineTypeValues)
			{
				case OutlineLineTypeValues.DOUBLE:
					return A.CompoundLineValues.Double;
				case OutlineLineTypeValues.TRIPLE:
					return A.CompoundLineValues.Triple;
				case OutlineLineTypeValues.THICK_THIN:
					return A.CompoundLineValues.ThickThin;
				case OutlineLineTypeValues.THIN_THICK:
					return A.CompoundLineValues.ThinThick;
				default:
					return A.CompoundLineValues.Single;
			}
		}
		internal static A.TextUnderlineValues GetTextUnderlineValues(UnderLineValues runPropertiesUnderLineValues)
		{
			switch (runPropertiesUnderLineValues)
			{
				case UnderLineValues.DASH:
					return A.TextUnderlineValues.Dash;
				case UnderLineValues.DASH_HEAVY:
					return A.TextUnderlineValues.DashHeavy;
				case UnderLineValues.DASH_LONG:
					return A.TextUnderlineValues.DashLong;
				case UnderLineValues.DASH_LONG_HEAVY:
					return A.TextUnderlineValues.DashLongHeavy;
				case UnderLineValues.DOT_DASH:
					return A.TextUnderlineValues.DotDash;
				case UnderLineValues.DOT_DASH_HEAVY:
					return A.TextUnderlineValues.DotDashHeavy;
				case UnderLineValues.DOT_DOT_DASH:
					return A.TextUnderlineValues.DotDotDash;
				case UnderLineValues.DOT_DOT_DASH_HEAVY:
					return A.TextUnderlineValues.DotDotDashHeavy;
				case UnderLineValues.DOTTED:
					return A.TextUnderlineValues.Dotted;
				case UnderLineValues.DOUBLE:
					return A.TextUnderlineValues.Double;
				case UnderLineValues.HEAVY:
					return A.TextUnderlineValues.Heavy;
				case UnderLineValues.HEAVY_DOTTED:
					return A.TextUnderlineValues.HeavyDotted;
				case UnderLineValues.SINGLE:
					return A.TextUnderlineValues.Single;
				case UnderLineValues.WAVY:
					return A.TextUnderlineValues.Wavy;
				case UnderLineValues.WAVY_DOUBLE:
					return A.TextUnderlineValues.WavyDouble;
				case UnderLineValues.WAVY_HEAVY:
					return A.TextUnderlineValues.WavyHeavy;
				case UnderLineValues.WORDS:
					return A.TextUnderlineValues.Words;
				default:
					return A.TextUnderlineValues.None;
			}
		}
		/// <summary>
		/// Class is only for inheritance purposes.
		/// </summary>
		protected CommonProperties() { }
		/// <summary>
		/// Create Solid Fill XML Property
		/// </summary>
		protected static OpenXmlElement CreateColorComponent<ColorOption>(ColorOptionModel<ColorOption> colorOptionModel = null)
		where ColorOption : class, IColorOptions, new()
		{
			if (colorOptionModel != null && typeof(ColorOption) == typeof(SolidOptions))
			{
				SolidOptions solidOptions = (colorOptionModel as ColorOptionModel<SolidOptions>).colorOption;
				if (solidOptions.hexColor == null && solidOptions.schemeColorModel == null)
				{
					throw new ArgumentException("Solid Fill Color Error");
				}
				if (solidOptions.hexColor != null)
				{
					A.RgbColorModelHex rgbColorModelHex = new A.RgbColorModelHex() { Val = solidOptions.hexColor };
					if (solidOptions.transparency != null)
					{
						rgbColorModelHex.Append(new A.Alpha() { Val = 100000 - (solidOptions.transparency * 1000) });
					}
					return new A.SolidFill() { RgbColorModelHex = rgbColorModelHex };
				}
				else
				{
					A.SchemeColor schemeColor = new A.SchemeColor()
					{ Val = new A.SchemeColorValues(GetSchemeColorValuesText(solidOptions.schemeColorModel.themeColorValues)) };
					if (solidOptions.transparency != null)
					{
						schemeColor.Append(new A.Alpha() { Val = 100000 - (solidOptions.transparency * 1000) });
					}
					if (solidOptions.schemeColorModel.tint != null)
					{
						schemeColor.Append(new A.Tint() { Val = solidOptions.schemeColorModel.tint });
					}
					if (solidOptions.schemeColorModel.shade != null)
					{
						schemeColor.Append(new A.Shade() { Val = solidOptions.schemeColorModel.shade });
					}
					if (solidOptions.schemeColorModel.saturationModulation != null)
					{
						schemeColor.Append(new A.SaturationModulation() { Val = solidOptions.schemeColorModel.saturationModulation });
					}
					if (solidOptions.schemeColorModel.saturationOffset != null)
					{
						schemeColor.Append(new A.SaturationOffset() { Val = solidOptions.schemeColorModel.saturationOffset });
					}
					if (solidOptions.schemeColorModel.luminanceModulation != null)
					{
						schemeColor.Append(new A.LuminanceModulation() { Val = solidOptions.schemeColorModel.luminanceModulation });
					}
					if (solidOptions.schemeColorModel.luminanceOffset != null)
					{
						schemeColor.Append(new A.LuminanceOffset() { Val = solidOptions.schemeColorModel.luminanceOffset });
					}
					return new A.SolidFill(schemeColor);
				}
			}
			return new A.NoFill();
		}
		/// <summary>
		/// Create Shape Properties With Default Settings
		/// </summary>
		/// <returns></returns>
		protected C.ShapeProperties CreateChartShapeProperties()
		{
			return CreateChartShapeProperties(new ShapePropertiesModel<NoFillOptions, NoFillOptions>());
		}
		/// <summary>
		/// Create Shape Properties
		/// </summary>
		/// <returns></returns>
		protected static C.ShapeProperties CreateChartShapeProperties<LineColorOption, FillColorOption>(ShapePropertiesModel<LineColorOption, FillColorOption> shapePropertiesModel)
		where LineColorOption : class, IColorOptions, new()
		where FillColorOption : class, IColorOptions, new()
		{
			C.ShapeProperties shapeProperties = new C.ShapeProperties();
			shapeProperties.Append(CreateColorComponent(shapePropertiesModel.fillColor));
			shapeProperties.Append(CreateOutline(shapePropertiesModel.lineColor));
			if (shapePropertiesModel.effectList != null)
			{
				shapeProperties.Append(CreateEffectList(shapePropertiesModel.effectList));
			}
			else
			{
				shapeProperties.Append(new A.EffectList());
			}
			if (shapePropertiesModel.shapeProperty3D != null)
			{
				shapeProperties.Append(new A.Shape3DType());
			}
			return shapeProperties;
		}
		/// <summary>
		/// Create Effect List
		/// </summary>
		protected static A.EffectList CreateEffectList(EffectListModel effectListModel)
		{
			return new A.EffectList();
		}
		/// <summary>
		/// Create Outline
		/// </summary>
		protected static A.Outline CreateOutline<LineColorOption>(OutlineModel<LineColorOption> outlineModel)
		where LineColorOption : class, IColorOptions, new()
		{
			A.Outline outline = new A.Outline();
			if (outlineModel.width != null)
			{
				outline.Width = outlineModel.width;
			}
			if (outlineModel.outlineCapTypeValues != null)
			{
				outline.CapType = GetLineCapValues((OutlineCapTypeValues)outlineModel.outlineCapTypeValues);
			}
			if (outlineModel.outlineLineTypeValues != null)
			{
				outline.CompoundLineType = GetLineTypeValues((OutlineLineTypeValues)outlineModel.outlineLineTypeValues);
			}
			if (outlineModel.outlineAlignmentValues != null)
			{
				outline.Alignment = GetLineAlignmentValues((OutlineAlignmentValues)outlineModel.outlineAlignmentValues);
			}
			outline.Append(CreateColorComponent(outlineModel.lineColor));
			if (outlineModel.lineColor != null)
			{
				outline.Append(new A.Round());
			}
			if (outlineModel.dashType != null)
			{
				outline.Append(new A.PresetDash { Val = GetDashStyleValue((DrawingPresetLineDashValues)outlineModel.dashType) });
			}
			A.HeadEnd headEnd = new A.HeadEnd() { Type = GetBeginArrowValue(outlineModel.beginArrowValues) };
			if (outlineModel.lineStartWidth != null)
			{
				headEnd.Width = GetLineStartWidthValue((LineWidthValues)outlineModel.lineStartWidth);
				headEnd.Length = GetLineStartLengthValue((LineWidthValues)outlineModel.lineStartWidth);
			}
			outline.Append(headEnd);
			A.TailEnd tailEnd = new A.TailEnd() { Type = GetEndArrowValue(outlineModel.endArrowValues) };
			if (outlineModel.lineEndWidth != null)
			{
				tailEnd.Width = GetLineEndWidthValue((LineWidthValues)outlineModel.lineEndWidth);
				tailEnd.Length = GetLineEndLengthValue((LineWidthValues)outlineModel.lineEndWidth);
			}
			outline.Append(tailEnd);
			return outline;
		}
		/// <summary>
		/// Create Default Run Properties
		/// </summary>
		protected static A.DefaultRunProperties CreateDefaultRunProperties<TextColorOption>()
		where TextColorOption : class, IColorOptions, new()
		{
			return CreateDefaultRunProperties(new DefaultRunPropertiesModel<TextColorOption>());
		}
		/// <summary>
		///     Create Default Run Properties
		/// </summary>
		protected static A.DefaultRunProperties CreateDefaultRunProperties<TextColorOption>(DefaultRunPropertiesModel<TextColorOption> defaultRunPropertiesModel)
		where TextColorOption : class, IColorOptions, new()
		{
			A.DefaultRunProperties defaultRunProperties = new A.DefaultRunProperties
			{
				FontSize = (int)defaultRunPropertiesModel.fontSize,
				Bold = defaultRunPropertiesModel.isBold,
				Italic = defaultRunPropertiesModel.isItalic,
				Underline = GetTextUnderlineValues(defaultRunPropertiesModel.underLineValues),
				Strike = GetTextStrikeValues(defaultRunPropertiesModel.strikeValues)
			};
			if (defaultRunPropertiesModel.textColorOption != null)
			{
				defaultRunProperties.Append(CreateColorComponent(defaultRunPropertiesModel.textColorOption));
			}
			if (defaultRunPropertiesModel.latinFont != null)
			{
				defaultRunProperties.Append(new A.LatinFont { Typeface = defaultRunPropertiesModel.latinFont });
			}
			if (defaultRunPropertiesModel.eastAsianFont != null)
			{
				defaultRunProperties.Append(new A.EastAsianFont { Typeface = defaultRunPropertiesModel.eastAsianFont });
			}
			if (defaultRunPropertiesModel.complexScriptFont != null)
			{
				defaultRunProperties.Append(new A.ComplexScriptFont { Typeface = defaultRunPropertiesModel.complexScriptFont });
			}
			if (defaultRunPropertiesModel.kerning != null)
			{
				defaultRunProperties.Kerning = defaultRunPropertiesModel.kerning;
			}
			if (defaultRunPropertiesModel.baseline != null)
			{
				defaultRunProperties.Baseline = defaultRunPropertiesModel.baseline;
			}
			return defaultRunProperties;
		}
		/// <summary>
		///
		/// </summary>
		protected A.Paragraph CreateDrawingParagraph<TextColorOption>(DrawingParagraphModel<TextColorOption> drawingParagraphModel)
		where TextColorOption : class, IColorOptions, new()
		{
			A.Paragraph paragraph = new A.Paragraph();
			if (drawingParagraphModel.paragraphPropertiesModel != null)
			{
				paragraph.Append(CreateDrawingParagraphProperties(drawingParagraphModel.paragraphPropertiesModel));
			}
			if (drawingParagraphModel.drawingRuns != null && drawingParagraphModel.drawingRuns.Length > 0)
			{
				paragraph.Append(CreateDrawingRun(drawingParagraphModel.drawingRuns));
			}
			else
			{
				if (drawingParagraphModel.paragraphPropertiesModel != null)
				{
					paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-IN" });
				}
			}
			return paragraph;
		}
		/// <summary>
		///
		/// </summary>
		private static A.ParagraphProperties CreateDrawingParagraphProperties<TextColorOption>(ParagraphPropertiesModel<TextColorOption> paragraphPropertiesModel)
		where TextColorOption : class, IColorOptions, new()
		{
			A.ParagraphProperties paragraphProperties = new A.ParagraphProperties();
			if (paragraphPropertiesModel.defaultRunProperties != null)
			{
				paragraphProperties.Append(CreateDefaultRunProperties(paragraphPropertiesModel.defaultRunProperties));
			}
			if (paragraphPropertiesModel.bulletsAndNumbering != null)
			{
				paragraphProperties.LeftMargin = 342900;
				paragraphProperties.Indent = -342900;
				switch (paragraphPropertiesModel.bulletsAndNumbering)
				{
					case BulletsAndNumberingValues.FILLED_ROUND:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "Arial",
							Panose = "020B0604020202020204",
							PitchFamily = 34,
							CharacterSet = 0
						});
						paragraphProperties.Append(new A.CharacterBullet() { Char = "•" });
						break;
					case BulletsAndNumberingValues.HOLLOW_ROUND:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "Arial",
							Panose = "020B0604020202020204",
							PitchFamily = 34,
							CharacterSet = 0
						});
						paragraphProperties.Append(new A.CharacterBullet() { Char = "o" });
						break;
					case BulletsAndNumberingValues.FILLED_SQUARE:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "Wingdings",
							Panose = "05000000000000000000",
							PitchFamily = 2,
							CharacterSet = 2
						});
						paragraphProperties.Append(new A.CharacterBullet()
						{
							Char = "§"
						});
						break;
					case BulletsAndNumberingValues.HOLLOW_SQUARE:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "Wingdings",
							Panose = "05000000000000000000",
							PitchFamily = 2,
							CharacterSet = 2
						});
						paragraphProperties.Append(new A.CharacterBullet()
						{
							Char = "q"
						});
						break;
					case BulletsAndNumberingValues.STAR_BULLET:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "Wingdings",
							Panose = "05000000000000000000",
							PitchFamily = 2,
							CharacterSet = 2
						});
						paragraphProperties.Append(new A.CharacterBullet()
						{
							Char = "v"
						});
						break;
					case BulletsAndNumberingValues.ARROW_BULLET:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "Wingdings",
							Panose = "05000000000000000000",
							PitchFamily = 2,
							CharacterSet = 2
						});
						paragraphProperties.Append(new A.CharacterBullet()
						{
							Char = "Ø"
						});
						break;
					case BulletsAndNumberingValues.CHECK_BULLET:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "Wingdings",
							Panose = "05000000000000000000",
							PitchFamily = 2,
							CharacterSet = 2
						});
						paragraphProperties.Append(new A.CharacterBullet()
						{
							Char = "ü"
						});
						break;
					case BulletsAndNumberingValues.NUMERIC_DOT:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "+mj-lt"
						});
						paragraphProperties.Append(new A.AutoNumberedBullet()
						{
							Type = A.TextAutoNumberSchemeValues.ArabicPeriod
						});
						break;
					case BulletsAndNumberingValues.NUMERIC_BRACKET:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "+mj-lt"
						});
						paragraphProperties.Append(new A.AutoNumberedBullet()
						{
							Type = A.TextAutoNumberSchemeValues.ArabicParenR
						});
						break;
					case BulletsAndNumberingValues.ROMAN_CAPS:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "+mj-lt"
						});
						paragraphProperties.Append(new A.AutoNumberedBullet()
						{
							Type = A.TextAutoNumberSchemeValues.RomanUpperCharacterPeriod
						});
						break;
					case BulletsAndNumberingValues.ROMAN_SMALL:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "+mj-lt"
						});
						paragraphProperties.Append(new A.AutoNumberedBullet()
						{
							Type = A.TextAutoNumberSchemeValues.RomanLowerCharacterPeriod
						});
						break;
					case BulletsAndNumberingValues.ALPHABET_CAPS:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "+mj-lt"
						});
						paragraphProperties.Append(new A.AutoNumberedBullet()
						{
							Type = A.TextAutoNumberSchemeValues.AlphaUpperCharacterPeriod
						});
						break;
					case BulletsAndNumberingValues.ALPHABET_SMALL:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "+mj-lt"
						});
						paragraphProperties.Append(new A.AutoNumberedBullet()
						{
							Type = A.TextAutoNumberSchemeValues.AlphaLowerCharacterPeriod
						});
						break;
					case BulletsAndNumberingValues.ALPHABET_SMALL_BRACKET:
						paragraphProperties.Append(new A.BulletFont()
						{
							Typeface = "+mj-lt"
						});
						paragraphProperties.Append(new A.AutoNumberedBullet()
						{
							Type = A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenR
						});
						break;
					default: // NONE
						break;
				}
			}
			if (paragraphPropertiesModel.horizontalAlignment != null)
			{
				paragraphProperties.Alignment = GetTextAlignmentValue((HorizontalAlignmentValues)paragraphPropertiesModel.horizontalAlignment);
			}
			return paragraphProperties;
		}
		/// <summary>
		///
		/// </summary>
		protected static A.ListStyle CreateDrawingListStyle()
		{
			return new A.ListStyle();
		}
		/// <summary>
		///     Create Chart Text Properties
		/// </summary>
		protected C.TextProperties CreateChartTextProperties<TextColorOption>(ChartTextPropertiesModel<TextColorOption> chartTextPropertiesModel)
		where TextColorOption : class, IColorOptions, new()
		{
			C.TextProperties textProperties = new C.TextProperties();
			if (chartTextPropertiesModel.drawingBodyProperties != null)
			{
				textProperties.Append(CreateDrawingBodyProperties(chartTextPropertiesModel.drawingBodyProperties));
			}
			textProperties.Append(CreateDrawingListStyle());
			if (chartTextPropertiesModel.drawingParagraph != null)
			{
				textProperties.Append(CreateDrawingParagraph(chartTextPropertiesModel.drawingParagraph));
			}
			return textProperties;
		}
		/// <summary>
		///
		/// </summary>
		protected C.RichText CreateChartRichText<TextColorOption>(ChartTextPropertiesModel<TextColorOption> chartTextPropertiesModel)
		where TextColorOption : class, IColorOptions, new()
		{
			C.RichText richText = new C.RichText();
			if (chartTextPropertiesModel.drawingBodyProperties != null)
			{
				richText.Append(CreateDrawingBodyProperties(chartTextPropertiesModel.drawingBodyProperties));
			}
			richText.Append(CreateDrawingListStyle());
			if (chartTextPropertiesModel.drawingParagraph != null)
			{
				richText.Append(CreateDrawingParagraph(chartTextPropertiesModel.drawingParagraph));
			}
			return richText;
		}
		/// <summary>
		///
		/// </summary>
		protected static A.Run[] CreateDrawingRun<TextColorOption>(DrawingRunModel<TextColorOption>[] drawingRunModels)
		where TextColorOption : class, IColorOptions, new()
		{
			List<A.Run> runs = new List<A.Run>();
			foreach (DrawingRunModel<TextColorOption> drawingRunModel in drawingRunModels)
			{
				A.Run run = new A.Run(CreateDrawingRunProperties(drawingRunModel.drawingRunProperties));
				if (drawingRunModel.text != null)
				{
					run.Append(new A.Text(drawingRunModel.text));
				}
				if (drawingRunModel.textHighlight != null)
				{
					run.Append(new A.Highlight(new A.RgbColorModelHex { Val = drawingRunModel.textHighlight }));
				}
				runs.Add(run);
			}
			return runs.ToArray();
		}
		/// <summary>
		///
		/// </summary>
		protected static A.RunProperties CreateDrawingRunProperties<TextColorOption>(DrawingRunPropertiesModel<TextColorOption> drawingRunPropertiesModel)
		where TextColorOption : class, IColorOptions, new()
		{
			A.RunProperties runProperties = new A.RunProperties
			{
				FontSize = ConverterUtils.FontSizeToFontSize(drawingRunPropertiesModel.fontSize),
				Bold = drawingRunPropertiesModel.isBold,
				Italic = drawingRunPropertiesModel.isItalic,
				Dirty = false,
				Underline = GetTextUnderlineValues(drawingRunPropertiesModel.underLineValues)
			};
			if (drawingRunPropertiesModel.hyperlinkProperties != null)
			{
				runProperties.Append(CreateHyperLink(drawingRunPropertiesModel.hyperlinkProperties));
			}
			if (drawingRunPropertiesModel.textColorOption != null)
			{
				runProperties.Append(CreateColorComponent(drawingRunPropertiesModel.textColorOption));
			}
			if (drawingRunPropertiesModel.fontFamily != null)
			{
				runProperties.Append(new A.LatinFont { Typeface = drawingRunPropertiesModel.fontFamily });
			}
			if (drawingRunPropertiesModel.fontFamily != null)
			{
				runProperties.Append(new A.EastAsianFont { Typeface = drawingRunPropertiesModel.fontFamily });
			}
			if (drawingRunPropertiesModel.fontFamily != null)
			{
				runProperties.Append(new A.ComplexScriptFont { Typeface = drawingRunPropertiesModel.fontFamily });
			}
			return runProperties;
		}
		/// <summary>
		///
		/// </summary>
		protected static A.HyperlinkOnClick CreateHyperLink(HyperlinkProperties hyperlinkProperties)
		{
			A.HyperlinkOnClick hyperlinkOnClick = new A.HyperlinkOnClick()
			{
				Id = hyperlinkProperties.relationId ?? ""
			};
			if (hyperlinkProperties.action != null)
			{
				hyperlinkOnClick.Action = hyperlinkProperties.action;
			}
			if (hyperlinkProperties.toolTip != null)
			{
				hyperlinkOnClick.Tooltip = hyperlinkProperties.toolTip;
			}
			return hyperlinkOnClick;
		}

		/// <summary>
		///    Create Drawing Body Properties
		/// </summary>
		/// <param name="drawingBodyPropertiesModel"></param>
		/// <returns></returns>
		private static A.BodyProperties CreateDrawingBodyProperties(DrawingBodyPropertiesModel drawingBodyPropertiesModel)
		{
			A.BodyProperties bodyProperties = new A.BodyProperties(new A.ShapeAutoFit())
			{
				Rotation = ConverterUtils.ConvertAngleDegreeToRotation(drawingBodyPropertiesModel.rotation)
			};
			if (drawingBodyPropertiesModel.leftInset != null)
			{
				bodyProperties.LeftInset = drawingBodyPropertiesModel.leftInset;
			}
			if (drawingBodyPropertiesModel.topInset != null)
			{
				bodyProperties.TopInset = drawingBodyPropertiesModel.topInset;
			}
			if (drawingBodyPropertiesModel.rightInset != null)
			{
				bodyProperties.RightInset = drawingBodyPropertiesModel.rightInset;
			}
			if (drawingBodyPropertiesModel.bottomInset != null)
			{
				bodyProperties.BottomInset = drawingBodyPropertiesModel.bottomInset;
			}
			if (drawingBodyPropertiesModel.useParagraphSpacing != null)
			{
				bodyProperties.UseParagraphSpacing = drawingBodyPropertiesModel.useParagraphSpacing;
			}
			if (drawingBodyPropertiesModel.verticalOverflow != null)
			{
				bodyProperties.VerticalOverflow = GetTextVerticalOverflowValues((TextVerticalOverflowValues)drawingBodyPropertiesModel.verticalOverflow);
			}
			if (drawingBodyPropertiesModel.vertical != null)
			{
				bodyProperties.Vertical = GetTextVerticalAlignmentValues((TextVerticalAlignmentValues)drawingBodyPropertiesModel.vertical);
			}
			if (drawingBodyPropertiesModel.wrap != null)
			{
				bodyProperties.Wrap = GetWarpingValues((TextWrappingValues)drawingBodyPropertiesModel.wrap);
			}
			if (drawingBodyPropertiesModel.anchor != null)
			{
				bodyProperties.Anchor = GetAnchorValues((TextAnchoringValues)drawingBodyPropertiesModel.anchor);
			}
			return bodyProperties;
		}
	}
}
