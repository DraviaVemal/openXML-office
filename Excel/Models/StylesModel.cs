/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Excel
{
    /// <summary>
    /// Represents the base class for a border in a style.
    /// </summary>
    public class BorderSetting
    {
        #region Public Fields
        /// <summary>
        /// Gets or sets the color of the border.
        /// </summary>
        public string Color = "64";
        /// <summary>
        /// Gets or sets the style of the border.
        /// </summary>
        public StyleValues Style = StyleValues.NONE;

        #endregion Public Fields

        #region Public Enums
        /// <summary>
        /// Border style values
        /// </summary>
        public enum StyleValues
        {
            /// <summary>
            /// None Border option
            /// </summary>
            NONE,
            /// <summary>
            /// Thin Border option
            /// </summary>
            THIN,
            /// <summary>
            /// Medium Border option
            /// </summary>
            THICK
        }

        #endregion Public Enums
    }
    /// <summary>
    /// Horizontal alignment values
    /// </summary>
    public enum HorizontalAlignmentValues
    {
        /// <summary>
        /// Unused
        /// </summary>
        NONE,
        /// <summary>
        /// Left alignment
        /// </summary>
        LEFT,
        /// <summary>
        /// Center alignment
        /// </summary>
        CENTER,
        /// <summary>
        /// Right alignment
        /// </summary>
        RIGHT
    }
    /// <summary>
    /// Vertical alignment values
    /// </summary>
    public enum VerticalAlignmentValues
    {
        /// <summary>
        /// Unused
        /// </summary>
        NONE,
        /// <summary>
        /// Top alignment
        /// </summary>
        TOP,
        /// <summary>
        /// Middle alignment
        /// </summary>
        MIDDLE,
        /// <summary>
        /// Bottom alignment
        /// </summary>
        BOTTOM
    }
    /// <summary>
    /// Represents the border style of a cell in a worksheet.
    /// </summary>
    public class BorderStyle
    {
        #region Public Fields
        /// <summary>
        /// Bottom border style
        /// </summary>
        public BorderSetting Bottom = new();
        /// <summary>
        /// Left border style
        /// </summary>
        public BorderSetting Left = new();
        /// <summary>
        /// Right border style
        /// </summary>
        public BorderSetting Right = new();
        /// <summary>
        /// Top border style
        /// </summary>
        public BorderSetting Top = new();

        #endregion Public Fields

        #region Public Properties
        /// <summary>
        /// Gets or sets the ID of the border style.
        /// </summary>
        public int Id;

        #endregion Public Properties
    }
    /// <summary>
    /// Represents the style of a cell in a worksheet.
    /// </summary>
    public class CellStyleSetting
    {
        #region Public Fields
        /// <summary>
        /// Gets or sets the background color of the cell.
        /// </summary>
        public string? BackgroundColor;
        /// <summary>
        /// Gets or sets the font family of the cell. default is Calibri
        /// </summary>
        public string FontFamily = "Calibri";
        /// <summary>
        /// Gets or sets the font size of the cell. default is 11
        /// </summary>
        public int FontSize = 11;
        /// <summary>
        /// Is Cell Bold. default is false
        /// </summary>
        public bool IsBold = false;
        /// <summary>
        /// Is Cell Underline. default is false
        /// </summary>
        public bool IsUnderline = false;
        /// <summary>
        /// Is Cell Double Underline. default is false
        /// </summary>
        public bool IsDoubleUnderline = false;
        /// <summary>
        /// Is Cell Italic. default is false
        /// </summary>
        public bool IsItalic = false;
        /// <summary>
        /// Is Wrap Text. default is false
        /// </summary>
        public bool IsWrapText = false;
        /// <summary>
        /// Gets or sets the number format of the cell. default is General
        /// </summary>
        public string NumberFormat = "General";
        /// <summary>
        /// Gets or sets the text color of the cell. default is 000000
        /// </summary>
        public string TextColor = "000000";
        /// <summary>
        /// Horizontal alignment of the cell. default is left
        /// </summary>
        public HorizontalAlignmentValues HorizontalAlignment = HorizontalAlignmentValues.NONE;
        /// <summary>
        /// Vertical alignment of the cell. default is bottom
        /// </summary>
        public VerticalAlignmentValues VerticalAlignment = VerticalAlignmentValues.NONE;
        /// <summary>
        /// Bottom border style
        /// </summary>
        public BorderSetting Bottom = new();
        /// <summary>
        /// Left border style
        /// </summary>
        public BorderSetting Left = new();
        /// <summary>
        /// Right border style
        /// </summary>
        public BorderSetting Right = new();
        /// <summary>
        /// Top border style
        /// </summary>
        public BorderSetting Top = new();
        /// <summary>
        /// Get or Set Foreground Color
        /// </summary>
        public string? ForegroundColor;
        #endregion Public Fields
    }
    /// <summary>
    /// Represents the fill style of a cell in a worksheet.
    /// </summary>
    public class FillStyle
    {
        #region Public Properties
        /// <summary>
        /// Fill style ID
        /// </summary>
        public int Id;
        /// <summary>
        /// Gets or sets the background color of the cell.
        /// </summary>
        public string? BackgroundColor;
        /// <summary>
        /// Gets or sets the foreground color of the cell.
        /// </summary>
        public string? ForegroundColor;
        #endregion Public Properties
    }
    /// <summary>
    /// Represents the font style of a cell in a worksheet.
    /// </summary>
    public class FontStyle
    {
        /// <summary>
        /// Font Scheme values
        /// </summary>
        public enum SchemeValues
        {
            /// <summary>
            /// None Scheme
            /// </summary>
            NONE,
            /// <summary>
            /// Minor Scheme
            /// </summary>
            MINOR,
            /// <summary>
            /// Major Scheme
            /// </summary>
            MAJOR
        }
        #region Public Fields
        /// <summary>
        /// Gets or sets the size of the font. default is 11
        /// </summary>
        public int Size = 11;
        /// <summary>
        /// Gets or sets the color of the font. default is accent1
        /// </summary>
        public string Color = "accent1";
        /// <summary>
        /// Font name default is Calibri
        /// </summary>
        public string Name = "Calibri";
        /// <summary>
        /// Gets or sets the font family of the font.
        /// </summary>
        public string Family = "2";
        /// <summary>
        /// Is Cell Underline. default is false
        /// </summary>
        public bool IsUnderline = false;
        /// <summary>
        /// Is Cell Italic. default is false
        /// </summary>
        public bool IsItalic = false;
        /// <summary>
        /// Is Cell Bold
        /// </summary>
        public bool IsBold = false;
        /// <summary>
        /// Is Cell Double Underline. default is false
        /// </summary>
        public bool IsDoubleUnderline = false;
        /// <summary>
        /// Configure Font Scheme
        /// </summary>
        public SchemeValues FontScheme = SchemeValues.NONE;
        #endregion Public Fields

        #region Public Properties
        /// <summary>
        /// Font style ID
        /// </summary>
        public int Id;

        #endregion Public Properties
    }

    /// <summary>
    /// Represents the number format of a cell in a worksheet.
    /// </summary>
    public class NumberFormats
    {
        #region Public Properties
        /// <summary>
        /// Number format ID
        /// </summary>
        public int Id;
        /// <summary>
        /// Number format code
        /// </summary>
        public string? FormatCode;
        #endregion Public Properties
    }
    /// <summary>
    /// Represents the cell style of a cell in a worksheet.
    /// </summary>
    public class CellXfs
    {
        #region Public Properties
        /// <summary>
        /// CellXfs ID
        /// </summary>
        public int Id;
        /// <summary>
        /// Number Format Id from collection
        /// </summary>
        public int NumberFormatId;
        /// <summary>
        /// Font Id from collection
        /// </summary>
        public int FontId;
        /// <summary>
        /// Fill Id from collection
        /// </summary>
        public int FillId;
        /// <summary>
        /// Border Id from collection
        /// </summary>
        public int BorderId;
        /// <summary>
        /// Apply Number Format
        /// </summary>
        public bool ApplyNumberFormat = false;
        /// <summary>
        /// Apply Alignment
        /// </summary>
        public bool ApplyAlignment = false;
        /// <summary>
        /// Apply Font style
        /// </summary>
        public bool ApplyFont = false;
        /// <summary>
        /// Apply Fill style
        /// </summary>
        public bool ApplyFill = false;
        /// <summary>
        /// Apply Border style
        /// </summary>
        public bool ApplyBorder = false;
        /// <summary>
        /// Horizontal alignment of the cell. default is left
        /// </summary>
        public HorizontalAlignmentValues HorizontalAlignment = HorizontalAlignmentValues.NONE;
        /// <summary>
        /// Vertical alignment of the cell. default is bottom
        /// </summary>
        public VerticalAlignmentValues VerticalAlignment = VerticalAlignmentValues.NONE;
        #endregion Public Properties
    }
}