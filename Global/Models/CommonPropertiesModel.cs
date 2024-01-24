// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenXMLOffice.Global
{
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
    /// 
    /// </summary>
    public class OutlineModel
    {
        /// <summary>
        /// 
        /// </summary>
        public int? width = null;
        /// <summary>
        /// 
        /// </summary>
        public OutlineCapTypeValues? outlineCapTypeValues = OutlineCapTypeValues.FLAT;
        /// <summary>
        /// 
        /// </summary>
        public OutlineLineTypeValues? outlineLineTypeValues = OutlineLineTypeValues.SINGEL;
        /// <summary>
        /// 
        /// </summary>
        public OutlineAlignmentValues? outlineAlignmentValues = OutlineAlignmentValues.CENTER;
        /// <summary>
        /// 
        /// </summary>
        public SolidFillModel? solidFill = null;

        internal A.PenAlignmentValues GetLineAlignmentValues(OutlineAlignmentValues outlineAlignmentValues)
        {
            return outlineAlignmentValues switch
            {
                OutlineAlignmentValues.CENTER => A.PenAlignmentValues.Center,
                _ => A.PenAlignmentValues.Insert
            };
        }

        internal A.LineCapValues GetLineCapValues(OutlineCapTypeValues outlineCapTypeValues)
        {
            return outlineCapTypeValues switch
            {
                OutlineCapTypeValues.SQUARE => A.LineCapValues.Square,
                OutlineCapTypeValues.ROUND => A.LineCapValues.Round,
                _ => A.LineCapValues.Flat
            };
        }

        internal A.CompoundLineValues GetLineTypeValues(OutlineLineTypeValues outlineLineTypeValues)
        {
            return outlineLineTypeValues switch
            {
                OutlineLineTypeValues.DOUBLE => A.CompoundLineValues.Double,
                OutlineLineTypeValues.TRIPLE => A.CompoundLineValues.Triple,
                OutlineLineTypeValues.THICK_THIN => A.CompoundLineValues.ThickThin,
                OutlineLineTypeValues.THIN_THICK => A.CompoundLineValues.ThinThick,
                _ => A.CompoundLineValues.Single
            };
        }
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
    /// 
    /// </summary>
    public class SolidFillModel
    {
        /// <summary>
        /// 
        /// </summary>
        public string? hexColor = null;
        /// <summary>
        /// 
        /// </summary>
        public SchemeColorModel? schemeColorModel = null;
        /// <summary>
        /// 
        /// </summary>
        internal string GetSchemeColorValuesText(ThemeColorValues themeColorValues)
        {
            return themeColorValues switch
            {
                ThemeColorValues.ACCENT_1 => "accent1",
                ThemeColorValues.ACCENT_2 => "accent2",
                ThemeColorValues.ACCENT_3 => "accent3",
                ThemeColorValues.ACCENT_4 => "accent4",
                ThemeColorValues.ACCENT_5 => "accent5",
                ThemeColorValues.ACCENT_6 => "accent6",
                ThemeColorValues.DARK_1 => "dk1",
                ThemeColorValues.DARK_2 => "dk2",
                ThemeColorValues.BACKGROUND_1 => "bg1",
                ThemeColorValues.BACKGROUND_2 => "bg2",
                ThemeColorValues.LIGHT_1 => "lt1",
                ThemeColorValues.LIGHT_2 => "lt2",
                ThemeColorValues.TEXT_1 => "tx1",
                ThemeColorValues.TEXT_2 => "tx2",
                ThemeColorValues.HYPERLINK => "hlink",
                ThemeColorValues.FOLLOW_HYPERLINK => "folHlink",
                _ => "phClr"
            };
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public class ShapePropertiesModel
    {
        /// <summary>
        /// 
        /// </summary>
        public SolidFillModel? solidFill = null;
        /// <summary>
        /// 
        /// </summary>
        public OutlineModel outline = new();
        /// <summary>
        /// 
        /// </summary>
        public EffectListModel? effectList = null;

    }
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
    public class DefaultRunPropertiesModel
    {

        /// <summary>
        /// 
        /// </summary>
        public SolidFillModel? solidFill = null;

        /// <summary>
        /// 
        /// </summary>
        public UnderLineValues? underline = null;

        /// <summary>
        /// 
        /// </summary>
        public string? latinFont;

        /// <summary>
        /// 
        /// </summary>
        public string? eastAsianFont;

        /// <summary>
        /// 
        /// </summary>
        public string? complexScriptFont;

        /// <summary>
        /// 
        /// </summary>
        public int? fontSize;

        /// <summary>
        /// 
        /// </summary>
        public bool? bold;

        /// <summary>
        /// 
        /// </summary>
        public bool? italic;

        /// <summary>
        /// 
        /// </summary>
        public StrikeValues? strike;

        /// <summary>
        /// 
        /// </summary>
        public int? kerning;

        /// <summary>
        /// 
        /// </summary>
        public int? baseline;

        /// <summary>
        /// 
        /// </summary>
        public A.TextStrikeValues GetTextStrikeValues(StrikeValues strikeValues)
        {
            return strikeValues switch
            {
                StrikeValues.SINGLE_STRIKE => A.TextStrikeValues.SingleStrike,
                StrikeValues.DOUBLE_STRIKE => A.TextStrikeValues.DoubleStrike,
                _ => A.TextStrikeValues.NoStrike
            };
        }

        /// <summary>
        /// 
        /// </summary>
        public A.TextUnderlineValues GetTextUnderlineValues(UnderLineValues runPropertiesUnderLineValues)
        {
            return runPropertiesUnderLineValues switch
            {
                UnderLineValues.DASH => A.TextUnderlineValues.Dash,
                UnderLineValues.DASH_HEAVY => A.TextUnderlineValues.DashHeavy,
                UnderLineValues.DASH_LONG => A.TextUnderlineValues.DashLong,
                UnderLineValues.DASH_LONG_HEAVY => A.TextUnderlineValues.DashLongHeavy,
                UnderLineValues.DOT_DASH => A.TextUnderlineValues.DotDash,
                UnderLineValues.DOT_DASH_HEAVY => A.TextUnderlineValues.DotDashHeavy,
                UnderLineValues.DOT_DOT_DASH => A.TextUnderlineValues.DotDotDash,
                UnderLineValues.DOT_DOT_DASH_HEAVY => A.TextUnderlineValues.DotDotDashHeavy,
                UnderLineValues.DOTTED => A.TextUnderlineValues.Dotted,
                UnderLineValues.DOUBLE => A.TextUnderlineValues.Double,
                _ => A.TextUnderlineValues.None
            };
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public class DrawingParagraphModel
    {
        /// <summary>
        /// 
        /// </summary>
        public ParagraphPropertiesModel? paragraphPropertiesModel = null;
    }

    /// <summary>
    /// /
    /// </summary>
    public class ParagraphPropertiesModel
    {
        /// <summary>
        /// 
        /// </summary>
        public DefaultRunPropertiesModel? defaultRunProperties = null;
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
        public int? leftInset = null;
        /// <summary>
        /// 
        /// </summary>
        public int? topInset = null;
        /// <summary>
        /// 
        /// </summary>
        public int? rightInset = null;
        /// <summary>
        /// 
        /// </summary>
        public int? bottomInset = null;
        /// <summary>
        /// 
        /// </summary>
        public bool? useParagraphSpacing = null;
        /// <summary>
        /// 
        /// </summary>
        public TextVerticalOverflowValues? verticalOverflow = null;
        /// <summary>
        /// 
        /// </summary>
        public TextVerticalAlignmentValues? vertical = null;
        /// <summary>
        /// 
        /// </summary>
        public TextWrappingValues? wrap = null;
        /// <summary>
        /// 
        /// </summary>
        public TextAnchoringValues? anchor = null;
        /// <summary>
        /// 
        /// </summary>
        public bool? anchorCenter = null;

        internal A.TextAnchoringTypeValues GetAnchorValues(TextAnchoringValues textAnchoring)
        {
            return textAnchoring switch
            {
                TextAnchoringValues.BOTTOM => A.TextAnchoringTypeValues.Bottom,
                TextAnchoringValues.CENTER => A.TextAnchoringTypeValues.Center,
                _ => A.TextAnchoringTypeValues.Top
            };
        }

        internal A.TextVerticalValues GetTextVerticalAlignmentValues(TextVerticalAlignmentValues textVerticalAlignment)
        {
            return textVerticalAlignment switch
            {
                TextVerticalAlignmentValues.EAST_ASIAN_VERTICAL => A.TextVerticalValues.EastAsianVetical,
                TextVerticalAlignmentValues.HORIZONTAL => A.TextVerticalValues.Horizontal,
                TextVerticalAlignmentValues.MONGOLIAN_VERTICAL => A.TextVerticalValues.MongolianVertical,
                TextVerticalAlignmentValues.VERTICAL => A.TextVerticalValues.Vertical,
                TextVerticalAlignmentValues.VERTICAL_270 => A.TextVerticalValues.Vertical270,
                TextVerticalAlignmentValues.WORD_ART_LEFT_TO_RIGHT => A.TextVerticalValues.WordArtVertical,
                _ => A.TextVerticalValues.WordArtVertical
            };
        }

        internal A.TextVerticalOverflowValues GetTextVerticalOverflowValues(TextVerticalOverflowValues textVerticalOverflow)
        {
            return textVerticalOverflow switch
            {
                TextVerticalOverflowValues.CLIP => A.TextVerticalOverflowValues.Clip,
                TextVerticalOverflowValues.ELLIPSIS => A.TextVerticalOverflowValues.Ellipsis,
                _ => A.TextVerticalOverflowValues.Overflow
            };
        }

        internal A.TextWrappingValues GetWrapingValues(TextWrappingValues textWrapping)
        {
            return textWrapping switch
            {
                TextWrappingValues.NONE => A.TextWrappingValues.None,
                _ => A.TextWrappingValues.Square
            };
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public class ChartTextPropertiesModel
    {
        /// <summary>
        /// 
        /// </summary>
        public DrawingBodyPropertiesModel? bodyProperties = null;
    }
}