/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using DocumentFormat.OpenXml.Spreadsheet;
using LiteDB;

namespace OpenXMLOffice.Excel
{
    /// <summary>
    /// This class serves as a versatile tool for working with Excel spreadsheets, styles component
    /// </summary>
    public class Styles
    {
        #region Private Fields

        private static readonly LiteDatabase LiteDatabase = new(Path.ChangeExtension(Path.GetTempFileName(), "db"));
        private readonly ILiteCollection<BorderStyle> BorderStyleCollection;
        private readonly ILiteCollection<CellXfs> CellXfsCollection;
        private readonly ILiteCollection<FillStyle> FillStyleCollection;
        private readonly ILiteCollection<FontStyle> FontStyleCollection;
        private readonly ILiteCollection<NumberFormats> NumberFormatCollection;
        private readonly Stylesheet Stylesheet;

        #endregion Private Fields

        #region Internal Constructors

        internal Styles(Stylesheet Stylesheet)
        {
            this.Stylesheet = Stylesheet;
            NumberFormatCollection = LiteDatabase.GetCollection<NumberFormats>("NumberFormats");
            FontStyleCollection = LiteDatabase.GetCollection<FontStyle>("FontStyle");
            FillStyleCollection = LiteDatabase.GetCollection<FillStyle>("FillStyle");
            BorderStyleCollection = LiteDatabase.GetCollection<BorderStyle>("BorderStyle");
            CellXfsCollection = LiteDatabase.GetCollection<CellXfs>("CellXfs");
            Initialise();
        }

        #endregion Internal Constructors

        #region Public Methods

        /// <summary>
        /// Get the Cell Style Id based on user specified CellStyleSetting
        /// </summary>
        /// <param name="CellStyleSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public int GetCellStyleId(CellStyleSetting CellStyleSetting)
        {
            int FontId = GetFontId(CellStyleSetting);
            int BorderId = GetBorderId(CellStyleSetting);
            int FillId = GetFillId(CellStyleSetting);
            int NumberFormatId = GetNumberFormat(CellStyleSetting);
            bool IsNumberFormat = NumberFormatId > 0;
            bool IsFill = FillId > 0;
            bool IsFont = FontId > 0;
            bool IsBorder = BorderId > 0;
            bool IsAlignment = CellStyleSetting.HorizontalAlignment != HorizontalAlignmentValues.NONE ||
                CellStyleSetting.VerticalAlignment != VerticalAlignmentValues.NONE;
            CellXfs? CellXfs = CellXfsCollection.Query().Where(item =>
                item.FontId == FontId &&
                item.BorderId == BorderId &&
                item.FillId == FillId &&
                item.NumberFormatId == NumberFormatId &&
                item.ApplyFill == IsFill &&
                item.ApplyFont == IsFont &&
                item.ApplyBorder == IsBorder &&
                item.ApplyAlignment == IsAlignment)
             .FirstOrDefault();
            if (CellXfs != null)
            {
                return CellXfs.Id;
            }
            else
            {
                BsonValue Result = CellXfsCollection.Insert(new CellXfs()
                {
                    FontId = FontId,
                    BorderId = BorderId,
                    FillId = FillId,
                    NumberFormatId = NumberFormatId,
                    ApplyFill = IsFill,
                    ApplyFont = IsFont,
                    ApplyBorder = IsBorder,
                    ApplyAlignment = IsAlignment
                });
                return Result.AsInt32;
            }
        }

        #endregion Public Methods

        #region Internal Methods

        internal void SaveStyleProps()
        {
            throw new NotImplementedException();
        }

        #endregion Internal Methods

        #region Private Methods

        private int GetBorderId(CellStyleSetting CellStyleSetting)
        {
            BorderStyle? BorderStyle = BorderStyleCollection.Query().Where(item =>
                item.Left == CellStyleSetting.Left &&
                item.Right == CellStyleSetting.Right &&
                item.Top == CellStyleSetting.Top &&
                item.Bottom == CellStyleSetting.Bottom)
            .FirstOrDefault();
            if (BorderStyle != null)
            {
                return BorderStyle.Id;
            }
            else
            {
                BsonValue Result = BorderStyleCollection.Insert(new BorderStyle()
                {
                    Left = CellStyleSetting.Left,
                    Right = CellStyleSetting.Right,
                    Top = CellStyleSetting.Top,
                    Bottom = CellStyleSetting.Bottom
                });
                return Result.AsInt32;
            }
        }

        private int GetFillId(CellStyleSetting CellStyleSetting)
        {
            FillStyle? FillStyle = FillStyleCollection.Query().Where(item =>
                item.BackgroundColor == CellStyleSetting.BackgroundColor &&
                item.ForegroundColor == CellStyleSetting.ForegroundColor)
                .FirstOrDefault();
            if (FillStyle != null)
            {
                return FillStyle.Id;
            }
            else
            {
                BsonValue Result = FillStyleCollection.Insert(new FillStyle()
                {
                    BackgroundColor = CellStyleSetting.BackgroundColor,
                    ForegroundColor = CellStyleSetting.ForegroundColor
                });
                return Result.AsInt32;
            }
        }

        private int GetFontId(CellStyleSetting CellStyleSetting)
        {
            FontStyle? FontStyle = FontStyleCollection.Query().Where(item =>
                item.IsBold == CellStyleSetting.IsBold &&
                item.IsItalic == CellStyleSetting.IsItalic &&
                item.IsUnderline == CellStyleSetting.IsUnderline &&
                item.IsDoubleUnderline == CellStyleSetting.IsDoubleUnderline &&
                item.Size == CellStyleSetting.FontSize &&
                item.Color == CellStyleSetting.TextColor &&
                item.Name == CellStyleSetting.FontFamily)
            .FirstOrDefault();
            if (FontStyle != null)
            {
                return FontStyle.Id;
            }
            else
            {
                BsonValue Result = FontStyleCollection.Insert(new FontStyle()
                {
                    IsBold = CellStyleSetting.IsBold,
                    IsItalic = CellStyleSetting.IsItalic,
                    IsUnderline = CellStyleSetting.IsUnderline,
                    IsDoubleUnderline = CellStyleSetting.IsDoubleUnderline,
                    Size = CellStyleSetting.FontSize,
                    Color = CellStyleSetting.TextColor,
                    Name = CellStyleSetting.FontFamily
                });
                return Result.AsInt32;
            }
        }

        private int GetNumberFormat(CellStyleSetting CellStyleSetting)
        {
            NumberFormats? NumberFormats = NumberFormatCollection.Query().Where(item =>
                item.FormatCode == CellStyleSetting.NumberFormat)
                .FirstOrDefault();
            if (NumberFormats != null)
            {
                return NumberFormats.Id;
            }
            else
            {
                BsonValue Result = NumberFormatCollection.Insert(new NumberFormats()
                {
                    FormatCode = CellStyleSetting.NumberFormat
                });
                return Result.AsInt32;
            }
        }

        private void Initialise()
        {
            Stylesheet.Fonts ??= new(
                new Font(
                    new FontSize() { Val = 11 },
                    new Color() { Theme = 1 },
                    new FontName() { Val = "Calibri" },
                    new FontFamily() { Val = 2 },
                    new FontScheme() { Val = FontSchemeValues.Minor }
                ))
            { Count = 1 };
            Stylesheet.Fills ??= new(
                new Fill(
                    new PatternFill() { PatternType = PatternValues.None }
                ),
                new Fill(
                    new PatternFill() { PatternType = PatternValues.DarkGray }
                ))
            { Count = 2 };
            Stylesheet.Borders ??= new(
                new Border(
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()
                )
            )
            { Count = 1 };
            Stylesheet.CellStyleFormats ??= new(
                new CellFormat() { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 })
            { Count = 1 };//cellStyleXfs
            Stylesheet.CellFormats ??= new(
                new CellFormat() { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 })
            { Count = 1 };//cellXfs
            Stylesheet.CellStyles ??= new(
                new CellStyle() { Name = "Normal", FormatId = 0, BuiltinId = 0 })
            { Count = 1 };//cellStyles
            Stylesheet.DifferentialFormats ??= new() { Count = 0 };//dxfs
        }

        #endregion Private Methods
    }
}