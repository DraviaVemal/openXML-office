// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using LiteDB;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLOffice.Excel {
    /// <summary>
    /// This class serves as a versatile tool for working with Excel spreadsheets, styles component
    /// </summary>
    public class Styles {
        #region Private Fields

        private static readonly LiteDatabase liteDatabase = new(Path.ChangeExtension(Path.GetTempFileName(),"db"));
        private static Styles? instance = null;
        private readonly ILiteCollection<BorderStyle> borderStyleCollection;
        private readonly ILiteCollection<CellXfs> cellXfsCollection;
        private readonly ILiteCollection<FillStyle> fillStyleCollection;
        private readonly ILiteCollection<FontStyle> fontStyleCollection;
        private readonly ILiteCollection<NumberFormats> numberFormatCollection;

        #endregion Private Fields

        #region Private Constructors

        private Styles() {
            numberFormatCollection = liteDatabase.GetCollection<NumberFormats>("NumberFormats");
            fontStyleCollection = liteDatabase.GetCollection<FontStyle>("FontStyle");
            fillStyleCollection = liteDatabase.GetCollection<FillStyle>("FillStyle");
            borderStyleCollection = liteDatabase.GetCollection<BorderStyle>("BorderStyle");
            cellXfsCollection = liteDatabase.GetCollection<CellXfs>("CellXfs");
        }

        #endregion Private Constructors

        #region Public Properties

        /// <summary>
        /// Get the Cell Style Id based on user specified CellStyleSetting
        /// </summary>
        public static Styles Instance {
            get {
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
        public uint GetCellStyleId(CellStyleSetting CellStyleSetting) {
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
            CellXfs? CellXfs = cellXfsCollection.FindOne(item =>
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
            if(CellXfs != null) {
                return CellXfs.Id;
            } else {
                BsonValue Result = cellXfsCollection.Insert(new CellXfs() {
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

        #endregion Public Methods

        #region Internal Methods

        /// <summary>
        /// Load the style from the Exisiting Sheet
        /// TODO: Load Exisiting Style from the Excel Sheet For Update
        /// </summary>
        /// <exception cref="NotImplementedException">
        /// </exception>
        internal void LoadStyleFromSheet(X.Stylesheet Stylesheet) {
        }

        /// <summary>
        /// Save the style properties to the xlsx file
        /// </summary>
        /// <exception cref="NotImplementedException">
        /// </exception>
        internal void SaveStyleProps(X.Stylesheet Stylesheet) {
            Stylesheet.Fonts = GetFonts();
            Stylesheet.Fills = GetFills();
            Stylesheet.Borders = GetBorders();
            Stylesheet.CellStyleFormats ??= new(
                new X.CellFormat() { NumberFormatId = 0,FontId = 0,FillId = 0,BorderId = 0 }) { Count = 1 };//cellStyleXfs
            Stylesheet.CellFormats = GetCellFormats();//cellXfs
            Stylesheet.CellStyles ??= new(
                new X.CellStyle() { Name = "Normal",FormatId = 0,BuiltinId = 0 }) { Count = 1 };//cellStyles
            Stylesheet.DifferentialFormats ??= new() { Count = 0 };//dxfs
            Stylesheet.NumberingFormats = GetNumberFormats();//numFmts
        }

        #endregion Internal Methods

        #region Private Methods

        private uint GetBorderId(CellStyleSetting CellStyleSetting) {
            BorderStyle? BorderStyle = borderStyleCollection.FindOne(item =>
                item.Left == CellStyleSetting.borderLeft &&
                item.Right == CellStyleSetting.borderRight &&
                item.Top == CellStyleSetting.borderTop &&
                item.Bottom == CellStyleSetting.borderBottom);
            if(BorderStyle != null) {
                return BorderStyle.Id;
            } else {
                BsonValue Result = borderStyleCollection.Insert(new BorderStyle() {
                    Id = (uint)borderStyleCollection.Count(),
                    Left = CellStyleSetting.borderLeft,
                    Right = CellStyleSetting.borderRight,
                    Top = CellStyleSetting.borderTop,
                    Bottom = CellStyleSetting.borderBottom
                });
                return (uint)Result.AsInt64;
            }
        }

        private X.Borders GetBorders() {
            return new(borderStyleCollection.FindAll().ToList().Select(item => {
                static X.BorderStyleValues GetBorderStyle(BorderSetting.StyleValues Style) {
                    return Style switch {
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
                X.Border Border = new() {
                    LeftBorder = new(),
                    RightBorder = new(),
                    BottomBorder = new(),
                    TopBorder = new(),
                };
                if(item.Left.style != BorderSetting.StyleValues.NONE) {
                    Border.LeftBorder.Style = GetBorderStyle(item.Left.style);
                    Border.LeftBorder.AppendChild(new X.Color() { Rgb = item.Left.color });
                }
                if(item.Right.style != BorderSetting.StyleValues.NONE) {
                    Border.RightBorder.Style = GetBorderStyle(item.Right.style);
                    Border.RightBorder.AppendChild(new X.Color() { Rgb = item.Left.color });
                }
                if(item.Top.style != BorderSetting.StyleValues.NONE) {
                    Border.TopBorder.Style = GetBorderStyle(item.Top.style);
                    Border.TopBorder.AppendChild(new X.Color() { Rgb = item.Left.color });
                }
                if(item.Bottom.style != BorderSetting.StyleValues.NONE) {
                    Border.BottomBorder.Style = GetBorderStyle(item.Bottom.style);
                    Border.BottomBorder.AppendChild(new X.Color() { Rgb = item.Left.color });
                }
                return Border;
            })) { Count = (uint)borderStyleCollection.Count() };
        }

        private X.CellFormats GetCellFormats() {
            return new(
                        cellXfsCollection.FindAll().ToList().Select(item => {
                            X.CellFormat CellFormat = new() {
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
                            if(item.VerticalAlignment != VerticalAlignmentValues.NONE ||
                                item.HorizontalAlignment != HorizontalAlignmentValues.NONE ||
                                item.IsWrapetext) {
                                CellFormat.Alignment = new();
                                if(item.VerticalAlignment != VerticalAlignmentValues.NONE) {
                                    CellFormat.Alignment.Vertical = item.VerticalAlignment switch {
                                        VerticalAlignmentValues.TOP => X.VerticalAlignmentValues.Top,
                                        VerticalAlignmentValues.MIDDLE => X.VerticalAlignmentValues.Center,
                                        _ => X.VerticalAlignmentValues.Bottom
                                    };
                                }
                                if(item.HorizontalAlignment != HorizontalAlignmentValues.NONE) {
                                    CellFormat.Alignment.Horizontal = item.HorizontalAlignment switch {
                                        HorizontalAlignmentValues.LEFT => X.HorizontalAlignmentValues.Left,
                                        HorizontalAlignmentValues.CENTER => X.HorizontalAlignmentValues.Center,
                                        _ => X.HorizontalAlignmentValues.Right
                                    };
                                }
                                if(item.IsWrapetext) {
                                    CellFormat.Alignment.WrapText = true;
                                }
                            }
                            return CellFormat;
                        })) { Count = (uint)cellXfsCollection.Count() };
        }

        private uint GetFillId(CellStyleSetting CellStyleSetting) {
            FillStyle? FillStyle = fillStyleCollection.FindOne(item =>
                item.BackgroundColor == CellStyleSetting.backgroundColor &&
                item.ForegroundColor == CellStyleSetting.foregroundColor);
            if(FillStyle != null) {
                return FillStyle.Id;
            } else {
                BsonValue Result = fillStyleCollection.Insert(new FillStyle() {
                    Id = (uint)fillStyleCollection.Count(),
                    BackgroundColor = CellStyleSetting.backgroundColor,
                    ForegroundColor = CellStyleSetting.foregroundColor
                });
                return (uint)Result.AsInt64;
            }
        }

        private X.Fills GetFills() {
            return new(fillStyleCollection.FindAll().ToList().Select(item => {
                X.Fill Fill = new() {
                    PatternFill = new() {
                        PatternType = item.PatternType switch {
                            FillStyle.PatternTypeValues.SOLID => X.PatternValues.Solid,
                            _ => X.PatternValues.None,
                        }
                    }
                };
                if(item.BackgroundColor != null) {
                    Fill.PatternFill.BackgroundColor = new() { Rgb = item.BackgroundColor };
                }
                if(item.ForegroundColor != null) {
                    Fill.PatternFill.ForegroundColor = new() { Rgb = item.ForegroundColor };
                }
                return Fill;
            })) { Count = (uint)fillStyleCollection.Count() };
        }

        private uint GetFontId(CellStyleSetting CellStyleSetting) {
            FontStyle? FontStyle = fontStyleCollection.FindOne(item =>
                item.IsBold == CellStyleSetting.isBold &&
                item.IsItalic == CellStyleSetting.isItalic &&
                item.IsUnderline == CellStyleSetting.isUnderline &&
                item.IsDoubleUnderline == CellStyleSetting.isDoubleUnderline &&
                item.Size == CellStyleSetting.fontSize &&
                item.Color == CellStyleSetting.textColor &&
                item.Name == CellStyleSetting.fontFamily);
            if(FontStyle != null) {
                return FontStyle.Id;
            } else {
                BsonValue Result = fontStyleCollection.Insert(new FontStyle() {
                    Id = (uint)fontStyleCollection.Count(),
                    IsBold = CellStyleSetting.isBold,
                    IsItalic = CellStyleSetting.isItalic,
                    IsUnderline = CellStyleSetting.isUnderline,
                    IsDoubleUnderline = CellStyleSetting.isDoubleUnderline,
                    Size = CellStyleSetting.fontSize,
                    Color = CellStyleSetting.textColor,
                    Name = CellStyleSetting.fontFamily
                });
                return (uint)Result.AsInt64;
            }
        }

        private X.Fonts GetFonts() {
            return new(fontStyleCollection.FindAll().ToList().Select(item => {
                X.Font Font = new() {
                    FontSize = new() { Val = item.Size },
                    FontName = new() { Val = item.Name },
                    FontFamilyNumbering = new() { Val = item.Family },
                    FontScheme = new() {
                        Val = item.FontScheme switch {
                            FontStyle.SchemeValues.MINOR => X.FontSchemeValues.Minor,
                            FontStyle.SchemeValues.MAJOR => X.FontSchemeValues.Major,
                            _ => X.FontSchemeValues.None
                        }
                    }
                };
                if(item.Color != null) {
                    Font.Color = new() { Rgb = item.Color };
                }
                return Font;
            })) { Count = (uint)fontStyleCollection.Count() };
        }

        private uint GetNumberFormat(CellStyleSetting CellStyleSetting) {
            NumberFormats? NumberFormats = numberFormatCollection.FindOne(item =>
                item.FormatCode == CellStyleSetting.numberFormat);
            if(NumberFormats != null) {
                return NumberFormats.Id;
            } else {
                BsonValue Result = numberFormatCollection.Insert(new NumberFormats() {
                    Id = (uint)numberFormatCollection.Count(),
                    FormatCode = CellStyleSetting.numberFormat
                });
                return (uint)Result.AsInt64;
            }
        }

        private X.NumberingFormats GetNumberFormats() {
            return new(numberFormatCollection.FindAll().ToList().Select(item => {
                X.NumberingFormat NumberingFormat = new() {
                    NumberFormatId = item.Id,
                    FormatCode = item.FormatCode
                };
                return NumberingFormat;
            })) { Count = (uint)numberFormatCollection.Count() };
        }

        #endregion Private Methods
    }
}