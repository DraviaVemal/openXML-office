// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using LiteDB;
using OpenXMLOffice.Global_2007;
using X = DocumentFormat.OpenXml.Spreadsheet;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// This class serves as a versatile tool for working with Excel spreadsheets, styles component
	/// </summary>
	public class StylesService
	{
		private readonly LiteDatabase liteDatabase = new LiteDatabase(Path.ChangeExtension(Path.GetTempFileName(), "db"));
		private readonly ILiteCollection<BorderStyle> borderStyleCollection;
		private readonly ILiteCollection<CellXfs> cellXfsCollection;
		private readonly ILiteCollection<FillStyle> fillStyleCollection;
		private readonly ILiteCollection<FontStyle> fontStyleCollection;
		private readonly ILiteCollection<NumberFormats> numberFormatCollection;
		internal StylesService()
		{
			numberFormatCollection = liteDatabase.GetCollection<NumberFormats>("NumberFormats");
			fontStyleCollection = liteDatabase.GetCollection<FontStyle>("FontStyle");
			fillStyleCollection = liteDatabase.GetCollection<FillStyle>("FillStyle");
			borderStyleCollection = liteDatabase.GetCollection<BorderStyle>("BorderStyle");
			cellXfsCollection = liteDatabase.GetCollection<CellXfs>("CellXfs");
		}
		private void InitializeDefault()
		{
			if (fontStyleCollection.Count() < 1)
			{
				fontStyleCollection.Insert(new FontStyle()
				{
					Id = (uint)fontStyleCollection.Count()
				});
			}
			if (fillStyleCollection.Count() < 1)
			{
				fillStyleCollection.Insert(new FillStyle()
				{
					Id = (uint)fillStyleCollection.Count(),
					PatternType = PatternTypeValues.NONE,
				});
			}
			if (fillStyleCollection.Count() < 2)
			{
				fillStyleCollection.Insert(new FillStyle()
				{
					Id = (uint)fillStyleCollection.Count(),
					PatternType = PatternTypeValues.GRAY125,
				});

			}
			if (borderStyleCollection.Count() < 1)
			{
				borderStyleCollection.Insert(new BorderStyle()
				{
					Id = (uint)borderStyleCollection.Count()
				});
			}
			if (cellXfsCollection.Count() < 1 && cellXfsCollection.FindOne(item =>
				item.HorizontalAlignment == HorizontalAlignmentValues.NONE &&
				item.VerticalAlignment == VerticalAlignmentValues.NONE) == null)
			{
				cellXfsCollection.Insert(new CellXfs()
				{
					Id = (uint)cellXfsCollection.Count()
				});
			}
		}
		/// <summary>
		/// Return Style details for the provided style ID
		/// </summary>
		public CellStyleSetting GetStyleForId(uint styleId)
		{
			CellXfs cellXfs = cellXfsCollection.FindOne(item => item.Id == styleId);
			FontStyle fontStyle = fontStyleCollection.FindOne(item => item.Id == cellXfs.FontId);
			BorderStyle borderStyle = borderStyleCollection.FindOne(item => item.Id == cellXfs.BorderId);
			FillStyle fillStyle = fillStyleCollection.FindOne(item => item.Id == cellXfs.FillId);
			NumberFormats numberFormats = numberFormatCollection.FindOne(item =>
				item.Id == cellXfs.NumberFormatId);
			CellStyleSetting cellStyleSetting = new CellStyleSetting()
			{
				isWrapText = cellXfs.IsWrapText,
				fontFamily = fontStyle.Name,
				fontSize = fontStyle.Size,
				isItalic = fontStyle.IsItalic,
				isBold = fontStyle.IsBold,
				isUnderline = fontStyle.IsUnderline,
				isDoubleUnderline = fontStyle.IsDoubleUnderline,
				textColor = fontStyle.Color,
				borderLeft = borderStyle.Left,
				borderTop = borderStyle.Top,
				borderRight = borderStyle.Right,
				borderBottom = borderStyle.Bottom,
				backgroundColor = fillStyle.BackgroundColor.Value,
				foregroundColor = fillStyle.ForegroundColor.Value,
				numberFormat = numberFormats.FormatCode,
			};
			return cellStyleSetting;
		}
		/// <summary>
		/// Get the Cell Style Id based on user specified CellStyleSetting
		/// </summary>
		public uint GetCellStyleId(CellStyleSetting CellStyleSetting)
		{
			uint FontId = GetFontId(CellStyleSetting);
			uint BorderId = GetBorderId(CellStyleSetting);
			uint FillId = GetFillId(CellStyleSetting);
			uint NumberFormatId = GetNumberFormat(CellStyleSetting);
			bool IsNumberFormat = NumberFormatId > 0;
			bool IsFill = FillId > 0;
			bool IsFont = FontId > 0;
			bool IsBorder = BorderId > 0;
			bool IsAlignment = CellStyleSetting.horizontalAlignment != HorizontalAlignmentValues.NONE ||
				CellStyleSetting.verticalAlignment != VerticalAlignmentValues.NONE;
			CellXfs CellXfs = cellXfsCollection.FindOne(item =>
				item.FontId == FontId &&
				item.BorderId == BorderId &&
				item.FillId == FillId &&
				item.NumberFormatId == NumberFormatId &&
				item.ApplyFill == IsFill &&
				item.ApplyFont == IsFont &&
				item.ApplyBorder == IsBorder &&
				item.ApplyAlignment == IsAlignment &&
				item.ApplyNumberFormat == IsNumberFormat &&
				item.IsWrapText == CellStyleSetting.isWrapText);
			if (CellXfs != null)
			{
				return CellXfs.Id;
			}
			else
			{
				BsonValue Result = cellXfsCollection.Insert(new CellXfs()
				{
					Id = (uint)cellXfsCollection.Count(),
					FontId = FontId,
					BorderId = BorderId,
					FillId = FillId,
					NumberFormatId = NumberFormatId,
					ApplyFill = IsFill,
					ApplyFont = IsFont,
					ApplyBorder = IsBorder,
					ApplyAlignment = IsAlignment,
					ApplyNumberFormat = IsNumberFormat,
					IsWrapText = CellStyleSetting.isWrapText
				});
				return (uint)Result.AsInt64;
			}
		}
		/// <summary>
		/// Load the style from the Existing Sheet
		/// </summary>
		internal void LoadStyleFromSheet(X.Stylesheet Stylesheet)
		{
			SetFonts(Stylesheet.Fonts);
			SetFills(Stylesheet.Fills);
			SetBorders(Stylesheet.Borders);
			SetCellFormats(Stylesheet.CellFormats);
			SetNumberFormats(Stylesheet.NumberingFormats);
			InitializeDefault();
		}
		/// <summary>
		/// Save the style properties to the xlsx file
		/// </summary>
		internal void SaveStyleProps(X.Stylesheet Stylesheet)
		{
			Stylesheet.Fonts = GetFonts();
			Stylesheet.Fills = GetFills();
			Stylesheet.Borders = GetBorders();
			if (Stylesheet.CellStyleFormats == null)
			{
				Stylesheet.CellStyleFormats = new X.CellStyleFormats(
				new X.CellFormat() { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 })
				{ Count = 1 };//cellStyleXfs
			}
			Stylesheet.CellFormats = GetCellFormats();//cellXfs
			if (Stylesheet.CellStyles == null)
			{
				Stylesheet.CellStyles = new X.CellStyles(
					new X.CellStyle() { Name = "Normal", FormatId = 0, BuiltinId = 0 })
				{ Count = 1 };//cellStyles
			}
			if (Stylesheet.DifferentialFormats == null)
			{
				Stylesheet.DifferentialFormats = new X.DifferentialFormats() { Count = 0 };//dxfs
			}
			Stylesheet.NumberingFormats = GetNumberFormats();//numFmts
		}
		private uint GetBorderId(CellStyleSetting CellStyleSetting)
		{
			BorderStyle BorderStyle = borderStyleCollection.FindOne(item =>
				item.Left == CellStyleSetting.borderLeft &&
				item.Right == CellStyleSetting.borderRight &&
				item.Top == CellStyleSetting.borderTop &&
				item.Bottom == CellStyleSetting.borderBottom);
			if (BorderStyle != null)
			{
				return BorderStyle.Id;
			}
			else
			{
				BsonValue Result = borderStyleCollection.Insert(new BorderStyle()
				{
					Id = (uint)borderStyleCollection.Count(),
					Left = CellStyleSetting.borderLeft,
					Right = CellStyleSetting.borderRight,
					Top = CellStyleSetting.borderTop,
					Bottom = CellStyleSetting.borderBottom
				});
				return (uint)Result.AsInt64;
			}
		}
		private X.BorderStyleValues GetBorderStyle(StyleValues Style)
		{
			switch (Style)
			{
				case StyleValues.THIN:
					return X.BorderStyleValues.Thin;
				case StyleValues.THICK:
					return X.BorderStyleValues.Thick;
				case StyleValues.DOTTED:
					return X.BorderStyleValues.Dotted;
				case StyleValues.DOUBLE:
					return X.BorderStyleValues.Double;
				case StyleValues.DASHED:
					return X.BorderStyleValues.Dashed;
				case StyleValues.DASH_DOT:
					return X.BorderStyleValues.DashDot;
				case StyleValues.DASH_DOT_DOT:
					return X.BorderStyleValues.DashDotDot;
				case StyleValues.MEDIUM:
					return X.BorderStyleValues.Medium;
				case StyleValues.MEDIUM_DASHED:
					return X.BorderStyleValues.MediumDashed;
				case StyleValues.MEDIUM_DASH_DOT:
					return X.BorderStyleValues.MediumDashDot;
				case StyleValues.MEDIUM_DASH_DOT_DOT:
					return X.BorderStyleValues.MediumDashDotDot;
				case StyleValues.SLANT_DASH_DOT:
					return X.BorderStyleValues.SlantDashDot;
				case StyleValues.HAIR:
					return X.BorderStyleValues.Hair;
				default:
					return X.BorderStyleValues.None;
			}
		}
		private StyleValues GetInvertedBorderStyle(string Style)
		{
			switch (Style)
			{
				case "thin":
					return StyleValues.THIN;
				case "thick":
					return StyleValues.THICK;
				case "dotted":
					return StyleValues.DOTTED;
				case "double":
					return StyleValues.DOUBLE;
				case "dashed":
					return StyleValues.DASHED;
				case "dashDot":
					return StyleValues.DASH_DOT;
				case "dashDotDot":
					return StyleValues.DASH_DOT_DOT;
				case "medium":
					return StyleValues.MEDIUM;
				case "mediumDashed":
					return StyleValues.MEDIUM_DASHED;
				case "mediumDashDot":
					return StyleValues.MEDIUM_DASH_DOT;
				case "mediumDashDotDot":
					return StyleValues.MEDIUM_DASH_DOT_DOT;
				case "slantDashDot":
					return StyleValues.SLANT_DASH_DOT;
				case "hair":
					return StyleValues.HAIR;
				default:
					return StyleValues.NONE;
			}
		}
		private X.Borders GetBorders()
		{
			return new X.Borders(borderStyleCollection.FindAll().ToList().Select(item =>
			{
				X.Border Border = new X.Border()
				{
					LeftBorder = new X.LeftBorder(),
					RightBorder = new X.RightBorder(),
					BottomBorder = new X.BottomBorder(),
					TopBorder = new X.TopBorder(),
					DiagonalBorder = new X.DiagonalBorder(),
				};
				if (item.Left.Style != StyleValues.NONE)
				{
					Border.LeftBorder.Style = GetBorderStyle(item.Left.Style);
					SetBorderColor(item.Left, Border.LeftBorder);
				}
				if (item.Right.Style != StyleValues.NONE)
				{
					Border.RightBorder.Style = GetBorderStyle(item.Right.Style);
					SetBorderColor(item.Right, Border.RightBorder);
				}
				if (item.Top.Style != StyleValues.NONE)
				{
					Border.TopBorder.Style = GetBorderStyle(item.Top.Style);
					SetBorderColor(item.Top, Border.TopBorder);
				}
				if (item.Bottom.Style != StyleValues.NONE)
				{
					Border.BottomBorder.Style = GetBorderStyle(item.Bottom.Style);
					SetBorderColor(item.Bottom, Border.BottomBorder);
				}
				return Border;
			}))
			{ Count = (uint)borderStyleCollection.Count() };
		}

		private static void SetBorderColor<T>(BorderSetting borderSetting, T Border) where T : X.BorderPropertiesType
		{
			if (borderSetting.BorderColor.ColorSettingTypeValues == ColorSettingTypeValues.RGB)
			{
				Border.AppendChild(new X.Color() { Rgb = borderSetting.BorderColor.Value });
			}
			else if (borderSetting.BorderColor.ColorSettingTypeValues == ColorSettingTypeValues.THEME)
			{
				Border.AppendChild(new X.Color() { Theme = new UInt32Value((uint)int.Parse(borderSetting.BorderColor.Value)) });
			}
			else if (borderSetting.BorderColor.ColorSettingTypeValues == ColorSettingTypeValues.INDEXED)
			{
				Border.AppendChild(new X.Color() { Indexed = new UInt32Value((uint)int.Parse(borderSetting.BorderColor.Value)) });
			}
		}

		private X.CellFormats GetCellFormats()
		{
			return new X.CellFormats(
						cellXfsCollection.FindAll().ToList().Select(item =>
						{
							X.CellFormat CellFormat = new X.CellFormat()
							{
								NumberFormatId = item.NumberFormatId,
								FontId = item.FontId,
								FillId = item.FillId,
								BorderId = item.BorderId,
								FormatId = item.FormatId,
							};
							if (item.ApplyAlignment)
							{
								CellFormat.ApplyAlignment = item.ApplyAlignment;
							}
							if (item.ApplyBorder)
							{
								CellFormat.ApplyBorder = item.ApplyBorder;
							}
							if (item.ApplyNumberFormat)
							{
								CellFormat.ApplyNumberFormat = item.ApplyNumberFormat;
							}
							if (item.ApplyFill)
							{
								CellFormat.ApplyFill = item.ApplyFill;
							}
							if (item.ApplyFont)
							{
								CellFormat.ApplyFont = item.ApplyFont;
							}
							if (item.VerticalAlignment != VerticalAlignmentValues.NONE ||
								item.HorizontalAlignment != HorizontalAlignmentValues.NONE ||
								item.IsWrapText)
							{
								CellFormat.Alignment = new X.Alignment();
								if (item.VerticalAlignment != VerticalAlignmentValues.NONE)
								{
									if (item.VerticalAlignment == VerticalAlignmentValues.TOP)
									{
										CellFormat.Alignment.Vertical = X.VerticalAlignmentValues.Top;
									}
									else if (item.VerticalAlignment == VerticalAlignmentValues.MIDDLE)
									{
										CellFormat.Alignment.Vertical = X.VerticalAlignmentValues.Center;
									}
									else
									{
										CellFormat.Alignment.Vertical = X.VerticalAlignmentValues.Bottom;
									}
								}
								if (item.HorizontalAlignment != HorizontalAlignmentValues.NONE)
								{
									if (item.HorizontalAlignment == HorizontalAlignmentValues.LEFT)
									{
										CellFormat.Alignment.Horizontal = X.HorizontalAlignmentValues.Left;
									}
									else if (item.HorizontalAlignment == HorizontalAlignmentValues.CENTER)
									{
										CellFormat.Alignment.Horizontal = X.HorizontalAlignmentValues.Center;
									}
									else
									{
										CellFormat.Alignment.Horizontal = X.HorizontalAlignmentValues.Right;
									}
								}
								if (item.IsWrapText)
								{
									CellFormat.Alignment.WrapText = true;
								}
							}
							return CellFormat;
						}))
			{ Count = (uint)cellXfsCollection.Count() };
		}
		private uint GetFillId(CellStyleSetting CellStyleSetting)
		{
			FillStyle FillStyle = fillStyleCollection.FindOne(item =>
				item.BackgroundColor.ColorSettingTypeValues == ColorSettingTypeValues.RGB &&
				item.BackgroundColor.Value == CellStyleSetting.backgroundColor &&
				item.ForegroundColor.ColorSettingTypeValues == ColorSettingTypeValues.RGB &&
				item.ForegroundColor.Value == CellStyleSetting.foregroundColor);
			if (FillStyle != null)
			{
				return FillStyle.Id;
			}
			else
			{
				BsonValue Result = fillStyleCollection.Insert(new FillStyle()
				{
					Id = (uint)fillStyleCollection.Count(),
					BackgroundColor = new ColorSetting()
					{
						ColorSettingTypeValues = ColorSettingTypeValues.RGB,
						Value = CellStyleSetting.backgroundColor
					},
					ForegroundColor = new ColorSetting()
					{
						ColorSettingTypeValues = ColorSettingTypeValues.RGB,
						Value = CellStyleSetting.foregroundColor
					}
				});
				return (uint)Result.AsInt64;
			}
		}
		private X.Fills GetFills()
		{
			return new X.Fills(fillStyleCollection.FindAll().ToList().Select(item =>
			{
				X.PatternValues patternType;
				switch (item.PatternType)
				{
					case PatternTypeValues.SOLID:
						patternType = X.PatternValues.Solid;
						break;
					case PatternTypeValues.GRAY125:
						patternType = X.PatternValues.Gray125;
						break;
					default:
						patternType = X.PatternValues.None;
						break;
				}
				X.PatternFill patternFill = new X.PatternFill()
				{
					PatternType = patternType
				};
				X.Fill fill = new X.Fill()
				{
					PatternFill = patternFill
				};
				if (item.BackgroundColor != null)
				{
					switch (item.BackgroundColor.ColorSettingTypeValues)
					{
						case ColorSettingTypeValues.INDEXED:
							fill.PatternFill.BackgroundColor = new X.BackgroundColor() { Indexed = new UInt32Value((uint)int.Parse(item.BackgroundColor.Value)) };
							break;
						case ColorSettingTypeValues.THEME:
							fill.PatternFill.BackgroundColor = new X.BackgroundColor() { Theme = new UInt32Value((uint)int.Parse(item.BackgroundColor.Value)) };
							break;
						default:
							fill.PatternFill.BackgroundColor = new X.BackgroundColor() { Rgb = item.BackgroundColor.Value };
							break;
					}
				}
				if (item.ForegroundColor != null)
				{
					switch (item.ForegroundColor.ColorSettingTypeValues)
					{
						case ColorSettingTypeValues.INDEXED:
							fill.PatternFill.ForegroundColor = new X.ForegroundColor() { Indexed = new UInt32Value((uint)int.Parse(item.ForegroundColor.Value)) };
							break;
						case ColorSettingTypeValues.THEME:
							fill.PatternFill.ForegroundColor = new X.ForegroundColor() { Theme = new UInt32Value((uint)int.Parse(item.ForegroundColor.Value)) };
							break;
						default:
							fill.PatternFill.ForegroundColor = new X.ForegroundColor() { Rgb = item.ForegroundColor.Value };
							break;
					}
				}
				return fill;
			}))
			{ Count = (uint)fillStyleCollection.Count() };
		}
		private uint GetFontId(CellStyleSetting CellStyleSetting)
		{
			FontStyle FontStyle = fontStyleCollection.FindOne(item =>
				item.IsBold == CellStyleSetting.isBold &&
				item.IsItalic == CellStyleSetting.isItalic &&
				item.IsUnderline == CellStyleSetting.isUnderline &&
				item.IsDoubleUnderline == CellStyleSetting.isDoubleUnderline &&
				item.Size == CellStyleSetting.fontSize &&
				item.Color.ColorSettingTypeValues == CellStyleSetting.textColor.ColorSettingTypeValues &&
				item.Color.Value == CellStyleSetting.textColor.Value &&
				item.Name == CellStyleSetting.fontFamily);
			if (FontStyle != null)
			{
				return FontStyle.Id;
			}
			else
			{
				BsonValue Result = fontStyleCollection.Insert(new FontStyle()
				{
					Id = (uint)fontStyleCollection.Count(),
					IsBold = CellStyleSetting.isBold,
					IsItalic = CellStyleSetting.isItalic,
					IsUnderline = CellStyleSetting.isUnderline,
					IsDoubleUnderline = CellStyleSetting.isDoubleUnderline,
					Size = CellStyleSetting.fontSize,
					Color = new ColorSetting()
					{
						ColorSettingTypeValues = ColorSettingTypeValues.RGB,
						Value = CellStyleSetting.textColor.Value,
					},
					Name = CellStyleSetting.fontFamily
				});
				return (uint)Result.AsInt64;
			}
		}
		private X.Fonts GetFonts()
		{
			return new X.Fonts(fontStyleCollection.FindAll().ToList().Select(item =>
			{
				X.FontSchemeValues fontScheme;
				switch (item.FontScheme)
				{
					case SchemeValues.MINOR:
						fontScheme = X.FontSchemeValues.Minor;
						break;
					case SchemeValues.MAJOR:
						fontScheme = X.FontSchemeValues.Major;
						break;
					default:
						fontScheme = X.FontSchemeValues.None;
						break;
				}
				X.FontScheme fontSchemeElement = new X.FontScheme()
				{
					Val = fontScheme
				};
				X.Font Font = new X.Font()
				{
					FontSize = new X.FontSize() { Val = item.Size },
					FontName = new X.FontName() { Val = item.Name },
					FontFamilyNumbering = new X.FontFamilyNumbering() { Val = item.Family },
					FontScheme = fontSchemeElement
				};
				if (item.Color.Value != null)
				{
					if (item.Color.ColorSettingTypeValues == ColorSettingTypeValues.RGB)
					{
						Font.Color = new X.Color() { Rgb = item.Color.Value };
					}
					else
					{
						Font.Color = new X.Color() { Theme = new UInt32Value((uint)int.Parse(item.Color.Value)) };
					}
				}
				if (item.IsBold)
				{
					Font.Bold = new X.Bold();
				}
				if (item.IsItalic)
				{
					Font.Italic = new X.Italic();
				}
				if (item.IsUnderline || item.IsDoubleUnderline)
				{
					Font.Underline = item.IsDoubleUnderline ? new X.Underline() { Val = X.UnderlineValues.Double } : new X.Underline();
				}
				return Font;
			}))
			{ Count = (uint)fontStyleCollection.Count() };
		}
		private uint GetNumberFormat(CellStyleSetting CellStyleSetting)
		{
			NumberFormats NumberFormats = numberFormatCollection.FindOne(item =>
				item.FormatCode == CellStyleSetting.numberFormat);
			if (NumberFormats != null)
			{
				return NumberFormats.Id;
			}
			else
			{
				uint numberFormatId = (uint)numberFormatCollection.Count();
				if (numberFormatId != 0)
				{
					numberFormatId = ((uint)numberFormatCollection.Max().AsInt32) + 1;
				}
				BsonValue Result = numberFormatCollection.Insert(new NumberFormats()
				{
					Id = numberFormatId,
					FormatCode = CellStyleSetting.numberFormat
				});
				return (uint)Result.AsInt64;
			}
		}
		private X.NumberingFormats GetNumberFormats()
		{
			return new X.NumberingFormats(numberFormatCollection.FindAll().ToList().Select(item =>
			{
				X.NumberingFormat NumberingFormat = new X.NumberingFormat()
				{
					NumberFormatId = item.Id,
					FormatCode = item.FormatCode
				};
				return NumberingFormat;
			}))
			{ Count = (uint)numberFormatCollection.Count() };
		}
		private void SetNumberFormats(X.NumberingFormats numberingFormats)
		{
			numberingFormats.Descendants<X.NumberingFormat>().ToList()
			.ForEach(numberingFormat =>
			{
				NumberFormats numFormat = new NumberFormats()
				{
					Id = numberingFormat.NumberFormatId
				};
				if (numberingFormat.FormatCode != null)
				{
					numFormat.FormatCode = numberingFormat.FormatCode;
				}
				numberFormatCollection.Insert(numFormat);
			});
		}
		private void SetCellFormats(X.CellFormats cellFormats)
		{
			cellFormats.Descendants<X.CellFormat>().ToList()
			.ForEach(cellFormat =>
			{
				CellXfs cellXfs = new CellXfs()
				{
					Id = (uint)cellXfsCollection.Count()
				};
				if (cellFormat.FormatId != null)
				{
					cellXfs.FormatId = cellFormat.FormatId;
				}
				if (cellFormat.NumberFormatId != null)
				{
					cellXfs.NumberFormatId = cellFormat.NumberFormatId;
				}
				if (cellFormat.FontId != null)
				{
					cellXfs.FontId = cellFormat.FontId;
				}
				if (cellFormat.FillId != null)
				{
					cellXfs.FillId = cellFormat.FillId;
				}
				if (cellFormat.BorderId != null)
				{
					cellXfs.BorderId = cellFormat.BorderId;
				}
				if (cellFormat.ApplyAlignment != null)
				{
					cellXfs.ApplyAlignment = cellFormat.ApplyAlignment;
				}
				if (cellFormat.ApplyBorder != null)
				{
					cellXfs.ApplyBorder = cellFormat.ApplyBorder;
				}
				if (cellFormat.ApplyFill != null)
				{
					cellXfs.ApplyFill = cellFormat.ApplyFill;
				}
				if (cellFormat.ApplyFont != null)
				{
					cellXfs.ApplyFont = cellFormat.ApplyFont;
				}
				if (cellFormat.ApplyNumberFormat != null)
				{
					cellXfs.ApplyNumberFormat = cellFormat.ApplyNumberFormat;
				}
				if (cellFormat.ApplyProtection != null)
				{
					cellXfs.ApplyProtection = cellFormat.ApplyProtection;
				}
				if (cellFormat.Alignment != null)
				{
					if (cellFormat.Alignment.Horizontal != null)
					{
						switch (cellFormat.Alignment.Horizontal.InnerText)
						{
							case "left":
								cellXfs.HorizontalAlignment = HorizontalAlignmentValues.LEFT;
								break;
							case "center":
								cellXfs.HorizontalAlignment = HorizontalAlignmentValues.CENTER;
								break;
							case "right":
								cellXfs.HorizontalAlignment = HorizontalAlignmentValues.RIGHT;
								break;
							case "justify":
								cellXfs.HorizontalAlignment = HorizontalAlignmentValues.JUSTIFY;
								break;
							default:
								cellXfs.HorizontalAlignment = HorizontalAlignmentValues.NONE;
								break;
						}
					}
					if (cellFormat.Alignment.Vertical != null)
					{
						switch (cellFormat.Alignment.Vertical.InnerText)
						{
							case "top":
								cellXfs.VerticalAlignment = VerticalAlignmentValues.TOP;
								break;
							case "center":
								cellXfs.VerticalAlignment = VerticalAlignmentValues.MIDDLE;
								break;
							case "bottom":
								cellXfs.VerticalAlignment = VerticalAlignmentValues.BOTTOM;
								break;
							default:
								cellXfs.VerticalAlignment = VerticalAlignmentValues.NONE;
								break;
						}
					}
					if (cellFormat.Alignment.WrapText != null)
					{
						cellXfs.IsWrapText = cellFormat.Alignment.WrapText;
					}
				}
				cellXfsCollection.Insert(cellXfs);
			});
		}
		private void SetBorders(X.Borders borders)
		{
			borders.Descendants<X.Border>().ToList()
			.ForEach(border =>
			{
				BorderStyle borderStyle = new BorderStyle()
				{
					Id = (uint)borderStyleCollection.Count()
				};
				if (border.LeftBorder.Style != null)
				{
					borderStyle.Left = new BorderSetting
					{
						Style = GetInvertedBorderStyle(border.LeftBorder.Style.InnerText),
						BorderColor = GetColorSetting(border.LeftBorder.Color)
					};
				}
				if (border.TopBorder.Style != null)
				{
					borderStyle.Top = new BorderSetting
					{
						Style = GetInvertedBorderStyle(border.TopBorder.Style.InnerText),
						BorderColor = GetColorSetting(border.TopBorder.Color)
					};
				}
				if (border.RightBorder.Style != null)
				{
					borderStyle.Right = new BorderSetting
					{
						Style = GetInvertedBorderStyle(border.RightBorder.Style.InnerText),
						BorderColor = GetColorSetting(border.RightBorder.Color)
					};
				}
				if (border.BottomBorder.Style != null)
				{
					borderStyle.Bottom = new BorderSetting
					{
						Style = GetInvertedBorderStyle(border.BottomBorder.Style.InnerText),
						BorderColor = GetColorSetting(border.BottomBorder.Color)
					};
				}
				borderStyleCollection.Insert(borderStyle);
			});
		}

		private ColorSetting GetColorSetting(X.Color color)
		{
			if (color.Indexed != null)
			{
				return new ColorSetting()
				{
					ColorSettingTypeValues = ColorSettingTypeValues.INDEXED,
					Value = color.Indexed
				};
			}
			if (color.Theme != null)
			{
				return new ColorSetting()
				{
					ColorSettingTypeValues = ColorSettingTypeValues.THEME,
					Value = color.Theme
				};
			}
			if (color.Rgb != null)
			{
				return new ColorSetting()
				{
					ColorSettingTypeValues = ColorSettingTypeValues.RGB,
					Value = color.Rgb
				};
			}
			throw new InvalidOperationException("Found Not Supported Color Type in source File");
		}

		private void SetFills(X.Fills fills)
		{
			fills.Descendants<X.Fill>().ToList()
			.ForEach(fill =>
			{
				FillStyle fillStyle = new FillStyle()
				{
					Id = (uint)fillStyleCollection.Count()
				};
				if (fill.PatternFill.ForegroundColor != null)
				{
					if (fill.PatternFill.ForegroundColor.Indexed != null)
					{
						fillStyle.ForegroundColor = new ColorSetting()
						{
							ColorSettingTypeValues = ColorSettingTypeValues.INDEXED,
							Value = fill.PatternFill.ForegroundColor.Indexed
						};
					}
					else if (fill.PatternFill.ForegroundColor.Theme != null)
					{
						fillStyle.ForegroundColor = new ColorSetting()
						{
							ColorSettingTypeValues = ColorSettingTypeValues.THEME,
							Value = fill.PatternFill.ForegroundColor.Theme
						};
					}
					else if (fill.PatternFill.ForegroundColor.Rgb != null)
					{
						fillStyle.ForegroundColor = new ColorSetting()
						{
							ColorSettingTypeValues = ColorSettingTypeValues.RGB,
							Value = fill.PatternFill.ForegroundColor.Rgb
						};
					}
				}
				if (fill.PatternFill.BackgroundColor != null)
				{
					if (fill.PatternFill.BackgroundColor.Indexed != null)
					{
						fillStyle.BackgroundColor = new ColorSetting()
						{
							ColorSettingTypeValues = ColorSettingTypeValues.INDEXED,
							Value = fill.PatternFill.BackgroundColor.Indexed
						};
					}
					else if (fill.PatternFill.BackgroundColor.Theme != null)
					{
						fillStyle.BackgroundColor = new ColorSetting()
						{
							ColorSettingTypeValues = ColorSettingTypeValues.THEME,
							Value = fill.PatternFill.BackgroundColor.Theme
						};
					}
					else if (fill.PatternFill.BackgroundColor.Rgb != null)
					{
						fillStyle.BackgroundColor = new ColorSetting()
						{
							ColorSettingTypeValues = ColorSettingTypeValues.RGB,
							Value = fill.PatternFill.BackgroundColor.Rgb
						};
					}
				}
				switch (fill.PatternFill.PatternType.InnerText)
				{
					case "gray125":
						fillStyle.PatternType = PatternTypeValues.GRAY125;
						break;
					case "solid":
						fillStyle.PatternType = PatternTypeValues.SOLID;
						break;
					default:
						fillStyle.PatternType = PatternTypeValues.NONE;
						break;
				}
				fillStyleCollection.Insert(fillStyle);
			});
		}
		private void SetFonts(X.Fonts fonts)
		{
			fonts.Descendants<X.Font>().ToList()
			.ForEach(font =>
			{
				FontStyle fontStyle = new FontStyle()
				{
					Id = (uint)fontStyleCollection.Count(),
					IsBold = font.Bold != null,
					IsItalic = font.Italic != null,
					IsUnderline = font.Underline != null,
					IsDoubleUnderline = font.Underline != null && font.Underline.Val != null && font.Underline.Val == X.UnderlineValues.Double,
				};
				if (font.FontSize.Val != null)
				{
					fontStyle.Size = (uint)font.FontSize.Val;
				}
				if (font.Color != null)
				{
					fontStyle.Color = font.Color.Rgb != null ?
					new ColorSetting()
					{
						ColorSettingTypeValues = ColorSettingTypeValues.RGB,
						Value = font.Color.Rgb
					} :
					new ColorSetting() { Value = font.Color.Theme ?? "1" };
				}
				if (font.FontName.Val != null)
				{
					fontStyle.Name = font.FontName.Val;
				}
				if (font.FontFamilyNumbering.Val != null)
				{
					fontStyle.Family = font.FontFamilyNumbering.Val;
				}
				if (font.FontScheme != null && font.FontScheme.Val != null && font.FontScheme.Val.InnerText != null)
				{
					SchemeValues fontScheme;
					switch (font.FontScheme.Val.InnerText)
					{
						case "minor":
							fontScheme = SchemeValues.MINOR;
							break;
						case "major":
							fontScheme = SchemeValues.MAJOR;
							break;
						default:
							fontScheme = SchemeValues.NONE;
							break;
					}
					fontStyle.FontScheme = fontScheme;
				}
				fontStyleCollection.Insert(fontStyle);
			});
		}
	}
}
