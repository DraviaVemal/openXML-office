// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using A = DocumentFormat.OpenXml.Drawing;
using OpenXMLOffice.Global;

namespace OpenXMLOffice.Presentation
{
    /// <summary>
    /// 
    /// </summary>
    public class TableBorderSetting
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
        public bool showBorder = false;
        /// <summary>
        /// 
        /// </summary>
        public string borderColor = "000000";
        /// <summary>
        /// 
        /// </summary>
        public float width = 1.27F;
        /// <summary>
        /// 
        /// </summary>
        public BorderStyleValues borderStyle = BorderStyleValues.SINGEL;
        /// <summary>
        /// 
        /// </summary>
        public DrawingPresetLineDashValues dashStyle = DrawingPresetLineDashValues.SOLID;

        internal static A.CompoundLineValues GetBorderStyleValue(BorderStyleValues borderStyle)
        {
            return borderStyle switch
            {
                BorderStyleValues.DOUBLE => A.CompoundLineValues.Double,
                BorderStyleValues.TRIPLE => A.CompoundLineValues.Triple,
                BorderStyleValues.THICK_THIN => A.CompoundLineValues.ThickThin,
                BorderStyleValues.THIN_THICK => A.CompoundLineValues.ThinThick,
                _ => A.CompoundLineValues.Single,
            };
        }

        internal static A.PresetLineDashValues GetDashStyleValue(DrawingPresetLineDashValues dashStyle)
        {
            return dashStyle switch
            {
                DrawingPresetLineDashValues.DASH => A.PresetLineDashValues.Dash,
                DrawingPresetLineDashValues.DASH_DOT => A.PresetLineDashValues.DashDot,
                DrawingPresetLineDashValues.DOT => A.PresetLineDashValues.Dot,
                DrawingPresetLineDashValues.LARGE_DASH => A.PresetLineDashValues.LargeDash,
                DrawingPresetLineDashValues.LARGE_DASH_DOT => A.PresetLineDashValues.LargeDashDot,
                DrawingPresetLineDashValues.LARGE_DASH_DOT_DOT => A.PresetLineDashValues.LargeDashDotDot,
                DrawingPresetLineDashValues.SYSTEM_DASH => A.PresetLineDashValues.SystemDash,
                DrawingPresetLineDashValues.SYSTEM_DASH_DOT => A.PresetLineDashValues.SystemDashDot,
                DrawingPresetLineDashValues.SYSTEM_DASH_DOT_DOT => A.PresetLineDashValues.SystemDashDotDot,
                DrawingPresetLineDashValues.SYSTEM_DOT => A.PresetLineDashValues.SystemDot,
                _ => A.PresetLineDashValues.Solid,
            };
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public class TableBorderSettings
    {
        /// <summary>
        /// 
        /// </summary>
        public TableBorderSetting leftBorder = new();
        /// <summary>
        /// 
        /// </summary>
        public TableBorderSetting topBorder = new();
        /// <summary>
        /// 
        /// </summary>
        public TableBorderSetting rightBorder = new();
        /// <summary>
        /// 
        /// </summary>
        public TableBorderSetting bottomBorder = new();
        /// <summary>
        /// 
        /// </summary>
        public TableBorderSetting topLeftToBottomRightBorder = new();
        /// <summary>
        /// 
        /// </summary>
        public TableBorderSetting bottomLeftToTopRightBorder = new();
    }

    /// <summary>
    /// Presentation Table Cell Class for setting the cell properties.
    /// </summary>
    public class TableCell
    {
        #region Public Fields

        /// <summary>
        /// Cell Alignment Option
        /// </summary>
        public HorizontalAlignmentValues? horizontalAlignment;

        /// <summary>
        /// 
        /// </summary>
        public VerticalAlignmentValues? verticalAlignment;

        /// <summary>
        /// 
        /// </summary>
        public TableBorderSettings borderSettings = new();

        /// <summary>
        /// Cell Background Color
        /// </summary>
        public string? cellBackground;

        /// <summary>
        /// Cell Font Family
        /// Default: Calibri (Body)
        /// </summary>
        public string fontFamily = "Calibri (Body)";

        /// <summary>
        /// Cell Font Size
        /// </summary>
        public int fontSize = 16;

        /// <summary>
        /// Is Bold text
        /// </summary>
        public bool isBold = false;

        /// <summary>
        /// Is Italic text
        /// </summary>
        public bool isItalic = false;

        /// <summary>
        /// Is Underline text
        /// </summary>
        public bool isUnderline = false;

        /// <summary>
        /// Text Background Color aka Highlight Color
        /// </summary>
        public string? textBackground;

        /// <summary>
        /// Text Color
        /// </summary>
        public string textColor = "000000";

        /// <summary>
        /// Cell Value
        /// </summary>
        public string? value;

        #endregion Public Fields
    }

    /// <summary>
    /// Table Row Customisation Properties
    /// </summary>
    public class TableRow
    {
        #region Public Fields

        /// <summary>
        /// Row Height
        /// </summary>
        public int height = 370840;

        /// <summary>
        /// Row Background Color.Will get overriden by TableCell.CellBackground
        /// </summary>
        public string? rowBackground;

        /// <summary>
        /// Table Cell List
        /// </summary>
        public List<TableCell> tableCells = new();

        /// <summary>
        /// Default Text Color for the row. Will get overriden by TableCell.TextColor
        /// </summary>
        public string textColor = "000000";

        #endregion Public Fields
    }

    /// <summary>
    /// Table Customisation Properties
    /// </summary>
    public class TableSetting
    {
        #region Public Fields

        /// <summary>
        /// Overall Table Height
        /// </summary>
        public uint height = 741680;

        /// <summary>
        /// Table Name. Default: Table 1
        /// </summary>
        public string name = "Table 1";

        /// <summary>
        /// Table Column Width List.Works based on WidthType Setting
        /// </summary>
        public List<float> tableColumnWidth = new();

        /// <summary>
        /// Overall Table Width
        /// </summary>
        public uint width = 8128000;

        /// <summary>
        /// AUTO - Ignore User Width value and space the colum equally EMU - (English Metric Units)
        /// Direct PPT standard Sizing 1 Inch * 914400 EMU's PIXEL - Based on Target DPI the pixel
        /// is converted to EMU and used when running PERCENTAGE - 0-100 Width percentage split for
        /// each column RATIO - 0-10 Width ratio of each column
        /// </summary>
        public WidthOptionValues widthType = WidthOptionValues.AUTO;

        /// <summary>
        /// Table X Position in the slide in EMUs (English Metric Units).
        /// </summary>
        public uint x = 0;

        /// <summary>
        /// Table Y Position in the slide in EMUs (English Metric Units).
        /// </summary>
        public uint y = 0;

        #endregion Public Fields

        #region Public Enums

        /// <summary>
        /// Width Option Values
        /// </summary>
        public enum WidthOptionValues
        {
            /// <summary>
            /// AUTO - Ignore User Width value and space the colum equally
            /// </summary>
            AUTO,

            /// <summary>
            /// EMU - (English Metric Units) Direct PPT standard Sizing 1 Inch * 914400 EMU's
            /// </summary>
            EMU,

            /// <summary>
            /// PIXEL - Based on Target DPI the pixel is converted to EMU and used when running
            /// </summary>
            PIXEL,

            /// <summary>
            /// PERCENTAGE - 0-100 Width percentage split for each column
            /// </summary>
            PERCENTAGE,

            /// <summary>
            /// RATIO - 0-10 Width ratio of each column
            /// </summary>
            RATIO
        }

        #endregion Public Enums
    }
}