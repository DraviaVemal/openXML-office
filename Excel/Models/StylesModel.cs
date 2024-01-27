// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global;

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
        public string color = "64";

        /// <summary>
        /// Gets or sets the style of the border.
        /// </summary>
        public StyleValues style = StyleValues.NONE;

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
            THICK,

            /// <summary>
            /// Dotted Border option
            /// </summary>
            DOTTED,

            /// <summary>
            /// Double Border option
            /// </summary>
            DOUBLE,

            /// <summary>
            /// Dashed Border option
            /// </summary>
            DASHED,

            /// <summary>
            /// Dash Dot Border option
            /// </summary>
            DASH_DOT,

            /// <summary>
            /// Dash Dot Dot Border option
            /// </summary>
            DASH_DOT_DOT,

            /// <summary>
            /// Medium Border option
            /// </summary>
            MEDIUM,

            /// <summary>
            /// Medium Dashed Border option
            /// </summary>
            MEDIUM_DASHED,

            /// <summary>
            /// Medium Dash Dot Border option
            /// </summary>
            MEDIUM_DASH_DOT,

            /// <summary>
            /// Medium Dash Dot Dot Border option
            /// </summary>
            MEDIUM_DASH_DOT_DOT,

            /// <summary>
            /// Slant Dash Dot Border option
            /// </summary>
            SLANT_DASH_DOT,

            /// <summary>
            /// Hair Border option
            /// </summary>
            HAIR
        }

        #endregion Public Enums
    }

    /// <summary>
    /// Represents the border style of a cell in a worksheet.
    /// </summary>
    public class BorderStyle
    {
        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="BorderStyle"/> class.
        /// </summary>
        public BorderStyle()
        {
            Bottom = new();
            Left = new();
            Right = new();
            Top = new();
        }

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Bottom border style
        /// </summary>
        public BorderSetting Bottom { get; set; }

        /// <summary>
        /// Gets or sets the ID of the border style.
        /// </summary>
        public uint Id { get; set; }

        /// <summary>
        /// Left border style
        /// </summary>
        public BorderSetting Left { get; set; }

        /// <summary>
        /// Right border style
        /// </summary>
        public BorderSetting Right { get; set; }

        /// <summary>
        /// Top border style
        /// </summary>
        public BorderSetting Top { get; set; }

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
        public string? backgroundColor;

        /// <summary>
        /// Bottom border style
        /// </summary>
        public BorderSetting borderBottom = new();

        /// <summary>
        /// Left border style
        /// </summary>
        public BorderSetting borderLeft = new();

        /// <summary>
        /// Right border style
        /// </summary>
        public BorderSetting borderRight = new();

        /// <summary>
        /// Top border style
        /// </summary>
        public BorderSetting borderTop = new();

        /// <summary>
        /// Gets or sets the font family of the cell. default is Calibri
        /// </summary>
        public string fontFamily = "Calibri";

        /// <summary>
        /// Gets or sets the font size of the cell. default is 11
        /// </summary>
        public uint fontSize = 11;

        /// <summary>
        /// Get or Set Foreground Color
        /// </summary>
        public string? foregroundColor;

        /// <summary>
        /// Horizontal alignment of the cell. default is left
        /// </summary>
        public HorizontalAlignmentValues horizontalAlignment = HorizontalAlignmentValues.NONE;

        /// <summary>
        /// Is Cell Bold. default is false
        /// </summary>
        public bool isBold = false;

        /// <summary>
        /// Is Cell Double Underline. default is false
        /// </summary>
        public bool isDoubleUnderline = false;

        /// <summary>
        /// Is Cell Italic. default is false
        /// </summary>
        public bool isItalic = false;

        /// <summary>
        /// Is Cell Underline. default is false
        /// </summary>
        public bool isUnderline = false;

        /// <summary>
        /// Is Wrap Text. default is false
        /// </summary>
        public bool isWrapText = false;

        /// <summary>
        /// Gets or sets the number format of the cell. default is General
        /// </summary>
        public string numberFormat = "General";

        /// <summary>
        /// Gets or sets the text color of the cell. default is 000000
        /// </summary>
        public string textColor = "000000";

        /// <summary>
        /// Vertical alignment of the cell. default is bottom
        /// </summary>
        public VerticalAlignmentValues verticalAlignment = VerticalAlignmentValues.NONE;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the cell style of a cell in a worksheet.
    /// </summary>
    public class CellXfs
    {
        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CellXfs"/> class.
        /// </summary>
        public CellXfs()
        {
            HorizontalAlignment = HorizontalAlignmentValues.NONE;
            VerticalAlignment = VerticalAlignmentValues.NONE;
        }

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Apply Alignment
        /// </summary>
        public bool ApplyAlignment { get; set; }

        /// <summary>
        /// Apply Border style
        /// </summary>
        public bool ApplyBorder { get; set; }

        /// <summary>
        /// Apply Fill style
        /// </summary>
        public bool ApplyFill { get; set; }

        /// <summary>
        /// Apply Font style
        /// </summary>
        public bool ApplyFont { get; set; }

        /// <summary>
        /// Apply Number Format
        /// </summary>
        public bool ApplyNumberFormat { get; set; }

        /// <summary>
        /// Border Id from collection
        /// </summary>
        public uint BorderId { get; set; }

        /// <summary>
        /// Fill Id from collection
        /// </summary>
        public uint FillId { get; set; }

        /// <summary>
        /// Font Id from collection
        /// </summary>
        public uint FontId { get; set; }

        /// <summary>
        /// Horizontal alignment of the cell. default is left
        /// </summary>
        public HorizontalAlignmentValues HorizontalAlignment { get; set; }

        /// <summary>
        /// CellXfs ID
        /// </summary>
        public uint Id { get; set; }

        /// <summary>
        /// Is Wrap Text. default is false
        /// </summary>
        public bool IsWrapetext { get; internal set; }

        /// <summary>
        /// Number Format Id from collection
        /// </summary>
        public uint NumberFormatId { get; set; }

        /// <summary>
        /// Vertical alignment of the cell. default is bottom
        /// </summary>
        public VerticalAlignmentValues VerticalAlignment { get; set; }

        #endregion Public Properties
    }

    /// <summary>
    /// Represents the fill style of a cell in a worksheet.
    /// </summary>
    public class FillStyle
    {
        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="FillStyle"/> class.
        /// </summary>
        public FillStyle()
        {
            PatternType = PatternTypeValues.NONE;
        }

        #endregion Public Constructors

        #region Public Enums

        /// <summary>
        /// Color Pattern Type
        /// TODO: Add more pattern types
        /// </summary>
        public enum PatternTypeValues
        {
            /// <summary>
            /// None Pattern
            /// </summary>
            NONE,

            /// <summary>
            /// Solid Pattern Type
            /// </summary>
            SOLID
        }

        #endregion Public Enums

        #region Public Properties

        /// <summary>
        /// Gets or sets the background color of the cell.
        /// </summary>
        public string? BackgroundColor { get; set; }

        /// <summary>
        /// Gets or sets the foreground color of the cell.
        /// </summary>
        public string? ForegroundColor { get; set; }

        /// <summary>
        /// Fill style ID
        /// </summary>
        public uint Id { get; set; }

        /// <summary>
        /// Pattern Type
        /// </summary>
        public PatternTypeValues PatternType { get; set; }

        #endregion Public Properties
    }

    /// <summary>
    /// Represents the font style of a cell in a worksheet.
    /// </summary>
    public class FontStyle
    {
        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="FontStyle"/> class.
        /// </summary>
        public FontStyle()
        {
            Color = "accent1";
            Family = 2;
            Size = 11;
            Name = "Calibri";
            FontScheme = SchemeValues.NONE;
        }

        #endregion Public Constructors

        #region Public Enums

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

        #endregion Public Enums

        #region Public Properties

        /// <summary>
        /// Gets or sets the color of the font. default is accent1
        /// </summary>
        public string Color { get; set; }

        /// <summary>
        /// Gets or sets the font family of the font.
        /// </summary>
        public int Family { get; set; }

        /// <summary>
        /// Configure Font Scheme
        /// </summary>
        public SchemeValues FontScheme { get; set; }

        /// <summary>
        /// Font style ID
        /// </summary>
        public uint Id { get; set; }

        /// <summary>
        /// Is Cell Bold
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// Is Cell Double Underline. default is false
        /// </summary>
        public bool IsDoubleUnderline { get; set; }

        /// <summary>
        /// Is Cell Italic. default is false
        /// </summary>
        public bool IsItalic { get; set; }

        /// <summary>
        /// Is Cell Underline. default is false
        /// </summary>
        public bool IsUnderline { get; set; }

        /// <summary>
        /// Font name default is Calibri
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the size of the font. default is 11
        /// </summary>
        public uint Size { get; set; }

        #endregion Public Properties
    }

    /// <summary>
    /// Represents the number format of a cell in a worksheet.
    /// </summary>
    public class NumberFormats
    {
        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="NumberFormats"/> class.
        /// </summary>
        public NumberFormats()
        {
            FormatCode = "General";
        }

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Number format code
        /// </summary>
        public string FormatCode { get; set; }

        /// <summary>
        /// Number format ID
        /// </summary>
        public uint Id { get; set; }

        #endregion Public Properties
    }
}