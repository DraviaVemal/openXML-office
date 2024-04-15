// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
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
			InitializeDefault();
		}
		private void InitializeDefault()
		{
			fontStyleCollection.Insert(new FontStyle()
			{
				Id = (uint)fontStyleCollection.Count()
			});
			fillStyleCollection.Insert(new FillStyle()
			{
				Id = (uint)fillStyleCollection.Count(),
				PatternType = PatternTypeValues.NONE,
			});
			fillStyleCollection.Insert(new FillStyle()
			{
				Id = (uint)fillStyleCollection.Count(),
				PatternType = PatternTypeValues.GRAY125,
			});
			borderStyleCollection.Insert(new BorderStyle()
			{
				Id = (uint)borderStyleCollection.Count()
			});
			cellXfsCollection.Insert(new CellXfs()
			{
				Id = (uint)cellXfsCollection.Count()
			});
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
				isWrapText = cellXfs.IsWrapetext,
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
				backgroundColor = fillStyle.BackgroundColor,
				foregroundColor = fillStyle.ForegroundColor,
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
				item.IsWrapetext == CellStyleSetting.isWrapText);
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
					IsWrapetext = CellStyleSetting.isWrapText
				});
				return (uint)Result.AsInt64;
			}
		}
		/// <summary>
		/// Load the style from the Exisiting Sheet
		/// </summary>
		internal void LoadStyleFromSheet(X.Stylesheet Stylesheet)
		{
			SetFonts(Stylesheet.Fonts);
			SetFills(Stylesheet.Fills);
			SetBorders(Stylesheet.Borders);
			SetCellFormats(Stylesheet.CellFormats);
			SetNumberFormats(Stylesheet.NumberingFormats);
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
		private static X.BorderStyleValues GetBorderStyle(StyleValues Style)
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
				if (item.Left.style != StyleValues.NONE)
				{
					Border.LeftBorder.Style = GetBorderStyle(item.Left.style);
					Border.LeftBorder.AppendChild(new X.Color() { Rgb = item.Left.color });
				}
				if (item.Right.style != StyleValues.NONE)
				{
					Border.RightBorder.Style = GetBorderStyle(item.Right.style);
					Border.RightBorder.AppendChild(new X.Color() { Rgb = item.Left.color });
				}
				if (item.Top.style != StyleValues.NONE)
				{
					Border.TopBorder.Style = GetBorderStyle(item.Top.style);
					Border.TopBorder.AppendChild(new X.Color() { Rgb = item.Left.color });
				}
				if (item.Bottom.style != StyleValues.NONE)
				{
					Border.BottomBorder.Style = GetBorderStyle(item.Bottom.style);
					Border.BottomBorder.AppendChild(new X.Color() { Rgb = item.Left.color });
				}
				return Border;
			}))
			{ Count = (uint)borderStyleCollection.Count() };
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
								FormatId = 0,
								ApplyAlignment = item.ApplyAlignment,
								ApplyBorder = item.ApplyBorder,
								ApplyNumberFormat = item.ApplyNumberFormat,
								ApplyFill = item.ApplyFill,
								ApplyFont = item.ApplyFont,
							};
							if (item.VerticalAlignment != VerticalAlignmentValues.NONE ||
								item.HorizontalAlignment != HorizontalAlignmentValues.NONE ||
								item.IsWrapetext)
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
								if (item.IsWrapetext)
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
				item.BackgroundColor == CellStyleSetting.backgroundColor &&
				item.ForegroundColor == CellStyleSetting.foregroundColor);
			if (FillStyle != null)
			{
				return FillStyle.Id;
			}
			else
			{
				BsonValue Result = fillStyleCollection.Insert(new FillStyle()
				{
					Id = (uint)fillStyleCollection.Count(),
					BackgroundColor = CellStyleSetting.backgroundColor,
					ForegroundColor = CellStyleSetting.foregroundColor
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
					fill.PatternFill.BackgroundColor = new X.BackgroundColor() { Rgb = item.BackgroundColor };
				}
				if (item.ForegroundColor != null)
				{
					fill.PatternFill.ForegroundColor = new X.ForegroundColor() { Rgb = item.ForegroundColor };
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
				item.Color.FontColorType == CellStyleSetting.textColor.FontColorType &&
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
					Color = new FontColor()
					{
						FontColorType = FontColorTypeValues.RGB,
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
					if (item.Color.FontColorType == FontColorTypeValues.RGB)
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
				if (cellFormat.NumberFormatId != null ||
				cellFormat.FontId != null ||
				cellFormat.FillId != null ||
				cellFormat.BorderId != null)
				{
					CellXfs cellXfs = new CellXfs()
					{
						Id = (uint)cellXfsCollection.Count()
					};
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
					cellXfsCollection.Insert(cellXfs);
				}
			});
		}
		private void SetBorders(X.Borders borders)
		{
			borders.Descendants<X.Border>().ToList()
			.ForEach(border =>
			{
			});
		}
		private void SetFills(X.Fills fills)
		{
			fills.Descendants<X.Fill>().ToList()
			.ForEach(fill =>
			{
				if ((fill.PatternFill.BackgroundColor != null && fill.PatternFill.BackgroundColor.Rgb != null) ||
				(fill.PatternFill.BackgroundColor != null && fill.PatternFill.BackgroundColor.Rgb != null))
				{
					FillStyle fillStyle = new FillStyle()
					{
						Id = (uint)fillStyleCollection.Count()
					};
					if (fill.PatternFill.BackgroundColor != null && fill.PatternFill.ForegroundColor.Rgb != null)
					{
						fillStyle.ForegroundColor = fill.PatternFill.ForegroundColor.Rgb;
					}
					if (fill.PatternFill.BackgroundColor != null && fill.PatternFill.BackgroundColor.Rgb != null)
					{
						fillStyle.BackgroundColor = fill.PatternFill.BackgroundColor.Rgb;
					}
					fillStyleCollection.Insert(fillStyle);
				}
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
					new FontColor()
					{
						FontColorType = FontColorTypeValues.RGB,
						Value = font.Color.Rgb
					} :
					new FontColor() { Value = font.Color.Theme ?? "1" };
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
