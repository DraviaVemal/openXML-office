/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using LiteDB;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLOffice.Excel
{
    /// <summary>
    /// This class serves as a versatile tool for working with Excel spreadsheets, styles component
    /// </summary>
    public class Styles
    {
        #region Private Fields

        private static readonly LiteDatabase LiteDatabase = new(Path.ChangeExtension(Path.GetTempFileName(), "db"));
        private static Styles? instance = null;
        private readonly ILiteCollection<BorderStyle> BorderStyleCollection;
        private readonly ILiteCollection<CellXfs> CellXfsCollection;
        private readonly ILiteCollection<FillStyle> FillStyleCollection;
        private readonly ILiteCollection<FontStyle> FontStyleCollection;
        private readonly ILiteCollection<NumberFormats> NumberFormatCollection;

        #endregion Private Fields

        #region Private Constructors

        private Styles()
        {
            NumberFormatCollection = LiteDatabase.GetCollection<NumberFormats>("NumberFormats");
            FontStyleCollection = LiteDatabase.GetCollection<FontStyle>("FontStyle");
            FillStyleCollection = LiteDatabase.GetCollection<FillStyle>("FillStyle");
            BorderStyleCollection = LiteDatabase.GetCollection<BorderStyle>("BorderStyle");
            CellXfsCollection = LiteDatabase.GetCollection<CellXfs>("CellXfs");
        }

        #endregion Private Constructors

        #region Public Properties

        /// <summary>
        /// Get the Cell Style Id based on user specified CellStyleSetting
        /// </summary>
        public static Styles Instance
        {
            get
            {
                instance ??= new Styles();
                return instance;
            }
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Get the Cell Style Id based on user specified CellStyleSetting
        /// </summary>
        /// <param name="CellStyleSetting">
        /// </param>
        /// <returns>
        /// </returns>
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
            bool IsAlignment = CellStyleSetting.HorizontalAlignment != HorizontalAlignmentValues.NONE ||
                CellStyleSetting.VerticalAlignment != VerticalAlignmentValues.NONE;
            CellXfs? CellXfs = CellXfsCollection.FindOne(item =>
                item.FontId == FontId &&
                item.BorderId == BorderId &&
                item.FillId == FillId &&
                item.NumberFormatId == NumberFormatId &&
                item.ApplyFill == IsFill &&
                item.ApplyFont == IsFont &&
                item.ApplyBorder == IsBorder &&
                item.ApplyAlignment == IsAlignment &&
                item.ApplyNumberFormat == IsNumberFormat &&
                item.IsWrapetext == CellStyleSetting.IsWrapText);
            if (CellXfs != null)
            {
                return CellXfs.Id;
            }
            else
            {
                BsonValue Result = CellXfsCollection.Insert(new CellXfs()
                {
                    Id = (uint)CellXfsCollection.Count(),
                    FontId = FontId,
                    BorderId = BorderId,
                    FillId = FillId,
                    NumberFormatId = NumberFormatId,
                    ApplyFill = IsFill,
                    ApplyFont = IsFont,
                    ApplyBorder = IsBorder,
                    ApplyAlignment = IsAlignment,
                    ApplyNumberFormat = IsNumberFormat,
                    IsWrapetext = CellStyleSetting.IsWrapText
                });
                return (uint)Result.AsInt64;
            }
        }

        #endregion Public Methods

        #region Internal Methods

        /// <summary>
        /// Load the style from the Exisiting Sheet
        /// TODO: Load Exisiting Style from the Excel Sheet For Update
        /// </summary>
        /// <exception cref="NotImplementedException">
        /// </exception>
        internal void LoadStyleFromSheet(X.Stylesheet Stylesheet)
        {
        }

        /// <summary>
        /// Save the style properties to the xlsx file
        /// </summary>
        /// <exception cref="NotImplementedException">
        /// </exception>
        internal void SaveStyleProps(X.Stylesheet Stylesheet)
        {
            Stylesheet.Fonts = GetFonts();
            Stylesheet.Fills = GetFills();
            Stylesheet.Borders = GetBorders();
            Stylesheet.CellStyleFormats ??= new(
                new X.CellFormat() { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 })
            { Count = 1 };//cellStyleXfs
            Stylesheet.CellFormats = GetCellFormats();//cellXfs
            Stylesheet.CellStyles ??= new(
                new X.CellStyle() { Name = "Normal", FormatId = 0, BuiltinId = 0 })
            { Count = 1 };//cellStyles
            Stylesheet.DifferentialFormats ??= new() { Count = 0 };//dxfs
            Stylesheet.NumberingFormats = GetNumberFormats();//numFmts
        }

        private X.NumberingFormats GetNumberFormats()
        {
            return new(NumberFormatCollection.FindAll().ToList().Select(item =>
            {
                X.NumberingFormat NumberingFormat = new()
                {
                    NumberFormatId = item.Id,
                    FormatCode = item.FormatCode
                };
                return NumberingFormat;
            }))
            { Count = (uint)NumberFormatCollection.Count() };
        }

        private X.CellFormats GetCellFormats()
        {
            return new(
                            CellXfsCollection.FindAll().ToList().Select(item =>
                            {
                                X.CellFormat CellFormat = new()
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
                                    CellFormat.Alignment = new();
                                    if (item.VerticalAlignment != VerticalAlignmentValues.NONE)
                                    {
                                        CellFormat.Alignment.Vertical = item.VerticalAlignment switch
                                        {
                                            VerticalAlignmentValues.TOP => X.VerticalAlignmentValues.Top,
                                            VerticalAlignmentValues.MIDDLE => X.VerticalAlignmentValues.Center,
                                            _ => X.VerticalAlignmentValues.Bottom
                                        };
                                    }
                                    if (item.HorizontalAlignment != HorizontalAlignmentValues.NONE)
                                    {
                                        CellFormat.Alignment.Horizontal = item.HorizontalAlignment switch
                                        {
                                            HorizontalAlignmentValues.LEFT => X.HorizontalAlignmentValues.Left,
                                            HorizontalAlignmentValues.CENTER => X.HorizontalAlignmentValues.Center,
                                            _ => X.HorizontalAlignmentValues.Right
                                        };
                                    }
                                    if (item.IsWrapetext)
                                    {
                                        CellFormat.Alignment.WrapText = true;
                                    }
                                }
                                return CellFormat;
                            }))
            { Count = (uint)CellXfsCollection.Count() };
        }

        private X.Borders GetBorders()
        {
            return new(BorderStyleCollection.FindAll().ToList().Select(item =>
            {
                static X.BorderStyleValues GetBorderStyle(BorderSetting.StyleValues Style)
                {
                    return Style switch
                    {
                        BorderSetting.StyleValues.THIN => X.BorderStyleValues.Thin,
                        BorderSetting.StyleValues.THICK => X.BorderStyleValues.Thick,
                        BorderSetting.StyleValues.DOTTED => X.BorderStyleValues.Dotted,
                        BorderSetting.StyleValues.DOUBLE => X.BorderStyleValues.Double,
                        BorderSetting.StyleValues.DASHED => X.BorderStyleValues.Dashed,
                        BorderSetting.StyleValues.DASH_DOT => X.BorderStyleValues.DashDot,
                        BorderSetting.StyleValues.DASH_DOT_DOT => X.BorderStyleValues.DashDotDot,
                        BorderSetting.StyleValues.MEDIUM => X.BorderStyleValues.Medium,
                        BorderSetting.StyleValues.MEDIUM_DASHED => X.BorderStyleValues.MediumDashed,
                        BorderSetting.StyleValues.MEDIUM_DASH_DOT => X.BorderStyleValues.MediumDashDot,
                        BorderSetting.StyleValues.MEDIUM_DASH_DOT_DOT => X.BorderStyleValues.MediumDashDotDot,
                        BorderSetting.StyleValues.SLANT_DASH_DOT => X.BorderStyleValues.SlantDashDot,
                        BorderSetting.StyleValues.HAIR => X.BorderStyleValues.Hair,
                        _ => X.BorderStyleValues.None
                    };
                }
                X.Border Border = new()
                {
                    LeftBorder = new(),
                    RightBorder = new(),
                    BottomBorder = new(),
                    TopBorder = new(),
                };
                if (item.Left.Style != BorderSetting.StyleValues.NONE)
                {
                    Border.LeftBorder.Style = GetBorderStyle(item.Left.Style);
                    Border.LeftBorder.AppendChild(new X.Color() { Rgb = item.Left.Color });
                }
                if (item.Right.Style != BorderSetting.StyleValues.NONE)
                {
                    Border.RightBorder.Style = GetBorderStyle(item.Right.Style);
                    Border.RightBorder.AppendChild(new X.Color() { Rgb = item.Left.Color });
                }
                if (item.Top.Style != BorderSetting.StyleValues.NONE)
                {
                    Border.TopBorder.Style = GetBorderStyle(item.Top.Style);
                    Border.TopBorder.AppendChild(new X.Color() { Rgb = item.Left.Color });
                }
                if (item.Bottom.Style != BorderSetting.StyleValues.NONE)
                {
                    Border.BottomBorder.Style = GetBorderStyle(item.Bottom.Style);
                    Border.BottomBorder.AppendChild(new X.Color() { Rgb = item.Left.Color });
                }
                return Border;
            }))
            { Count = (uint)BorderStyleCollection.Count() };
        }

        private X.Fills GetFills()
        {
            return new(FillStyleCollection.FindAll().ToList().Select(item =>
            {
                X.Fill Fill = new()
                {
                    PatternFill = new()
                    {
                        PatternType = item.PatternType switch
                        {
                            FillStyle.PatternTypeValues.SOLID => X.PatternValues.Solid,
                            _ => X.PatternValues.None,
                        }
                    }
                };
                if (item.BackgroundColor != null)
                {
                    Fill.PatternFill.BackgroundColor = new() { Rgb = item.BackgroundColor };
                }
                if (item.ForegroundColor != null)
                {
                    Fill.PatternFill.ForegroundColor = new() { Rgb = item.ForegroundColor };
                }
                return Fill;
            }))
            { Count = (uint)FillStyleCollection.Count() };
        }

        private X.Fonts GetFonts()
        {
            return new(FontStyleCollection.FindAll().ToList().Select(item =>
            {
                X.Font Font = new()
                {
                    FontSize = new() { Val = item.Size },
                    FontName = new() { Val = item.Name },
                    FontFamilyNumbering = new() { Val = item.Family },
                    FontScheme = new()
                    {
                        Val = item.FontScheme switch
                        {
                            FontStyle.SchemeValues.MINOR => X.FontSchemeValues.Minor,
                            FontStyle.SchemeValues.MAJOR => X.FontSchemeValues.Major,
                            _ => X.FontSchemeValues.None
                        }
                    }
                };
                if (item.Color != null)
                {
                    Font.Color = new() { Rgb = item.Color };
                }
                return Font;
            }))
            { Count = (uint)FontStyleCollection.Count() };
        }

        #endregion Internal Methods

        #region Private Methods

        private uint GetBorderId(CellStyleSetting CellStyleSetting)
        {
            BorderStyle? BorderStyle = BorderStyleCollection.FindOne(item =>
                item.Left == CellStyleSetting.BorderLeft &&
                item.Right == CellStyleSetting.BorderRight &&
                item.Top == CellStyleSetting.BorderTop &&
                item.Bottom == CellStyleSetting.BorderBottom);
            if (BorderStyle != null)
            {
                return BorderStyle.Id;
            }
            else
            {
                BsonValue Result = BorderStyleCollection.Insert(new BorderStyle()
                {
                    Id = (uint)BorderStyleCollection.Count(),
                    Left = CellStyleSetting.BorderLeft,
                    Right = CellStyleSetting.BorderRight,
                    Top = CellStyleSetting.BorderTop,
                    Bottom = CellStyleSetting.BorderBottom
                });
                return (uint)Result.AsInt64;
            }
        }

        private uint GetFillId(CellStyleSetting CellStyleSetting)
        {
            FillStyle? FillStyle = FillStyleCollection.FindOne(item =>
                item.BackgroundColor == CellStyleSetting.BackgroundColor &&
                item.ForegroundColor == CellStyleSetting.ForegroundColor);
            if (FillStyle != null)
            {
                return FillStyle.Id;
            }
            else
            {
                BsonValue Result = FillStyleCollection.Insert(new FillStyle()
                {
                    Id = (uint)FillStyleCollection.Count(),
                    BackgroundColor = CellStyleSetting.BackgroundColor,
                    ForegroundColor = CellStyleSetting.ForegroundColor
                });
                return (uint)Result.AsInt64;
            }
        }

        private uint GetFontId(CellStyleSetting CellStyleSetting)
        {
            FontStyle? FontStyle = FontStyleCollection.FindOne(item =>
                item.IsBold == CellStyleSetting.IsBold &&
                item.IsItalic == CellStyleSetting.IsItalic &&
                item.IsUnderline == CellStyleSetting.IsUnderline &&
                item.IsDoubleUnderline == CellStyleSetting.IsDoubleUnderline &&
                item.Size == CellStyleSetting.FontSize &&
                item.Color == CellStyleSetting.TextColor &&
                item.Name == CellStyleSetting.FontFamily);
            if (FontStyle != null)
            {
                return FontStyle.Id;
            }
            else
            {
                BsonValue Result = FontStyleCollection.Insert(new FontStyle()
                {
                    Id = (uint)FontStyleCollection.Count(),
                    IsBold = CellStyleSetting.IsBold,
                    IsItalic = CellStyleSetting.IsItalic,
                    IsUnderline = CellStyleSetting.IsUnderline,
                    IsDoubleUnderline = CellStyleSetting.IsDoubleUnderline,
                    Size = CellStyleSetting.FontSize,
                    Color = CellStyleSetting.TextColor,
                    Name = CellStyleSetting.FontFamily
                });
                return (uint)Result.AsInt64;
            }
        }

        private uint GetNumberFormat(CellStyleSetting CellStyleSetting)
        {
            NumberFormats? NumberFormats = NumberFormatCollection.FindOne(item =>
                item.FormatCode == CellStyleSetting.NumberFormat);
            if (NumberFormats != null)
            {
                return NumberFormats.Id;
            }
            else
            {
                BsonValue Result = NumberFormatCollection.Insert(new NumberFormats()
                {
                    Id = (uint)NumberFormatCollection.Count(),
                    FormatCode = CellStyleSetting.NumberFormat
                });
                return (uint)Result.AsInt64;
            }
        }

        #endregion Private Methods
    }
}