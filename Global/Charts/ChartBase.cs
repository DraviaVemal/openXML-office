using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global;

public class ChartBase
{
    #region Protected Methods

    private C.ChartSpace OpenXMLChartSpace;

    protected ChartBase()
    {
        OpenXMLChartSpace = CreateChartSpace();
    }

    protected C.ChartSpace GetChartSpace()
    {
        return OpenXMLChartSpace;
    }

    private C.ChartSpace CreateChartSpace()
    {
        C.ChartSpace ChartSpace = new();
        ChartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        ChartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        ChartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        ChartSpace.RoundedCorners = new C.RoundedCorners()
        {
            Val = false
        };
        ChartSpace.Date1904 = new C.Date1904()
        {
            Val = false
        };
        ChartSpace.EditingLanguage = new C.EditingLanguage()
        {
            Val = "en-US"
        };
        return ChartSpace;
    }

    protected CS.ChartStyle CreateChartStyles()
    {
        ChartStyle ChartStyle = new();
        return ChartStyle.CreateChartStyles();
    }

    protected CS.ColorStyle CreateColorStyles()
    {
        ChartColor ChartColor = new();
        return ChartColor.CreateColorStyles();
    }

    #endregion Protected Methods

    #region Private Methods


    protected C.Chart CreateChart()
    {
        C.Chart Chart = new()
        {
            Title = CreateTitle(),
            Legend = CreateChartLegend(),
            PlotVisibleOnly = new C.PlotVisibleOnly()
            {
                Val = true
            },
            AutoTitleDeleted = new C.AutoTitleDeleted()
            {
                Val = false
            },
            DisplayBlanksAs = new C.DisplayBlanksAs()
            {
                Val = C.DisplayBlanksAsValues.Gap
            }
        };
        return Chart;
    }

    private C.Legend CreateChartLegend()
    {
        C.Legend legend = new();
        legend.Append(new C.LegendPosition() { Val = C.LegendPositionValues.Bottom });
        legend.Append(new C.Overlay() { Val = false });
        C.ShapeProperties ShapeProperties = new();
        ShapeProperties.Append(new A.NoFill());
        A.Outline ln = new();
        ln.Append(new A.NoFill());
        ShapeProperties.Append(ln);
        ShapeProperties.Append(new A.EffectList());
        legend.Append(ShapeProperties);
        C.TextProperties TextProperties = new();
        TextProperties.Append(new A.BodyProperties()
        {
            Rotation = 0,
            UseParagraphSpacing = true,
            VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
            Vertical = A.TextVerticalValues.Horizontal,
            Wrap = A.TextWrappingValues.Square,
            Anchor = A.TextAnchoringTypeValues.Center,
            AnchorCenter = true
        });
        TextProperties.Append(new A.ListStyle());
        A.Paragraph paragraph = new();
        A.ParagraphProperties paragraphProperties = new();
        A.DefaultRunProperties defaultRunProperties = new()
        {
            FontSize = 1197,
            Bold = false,
            Italic = false,
            Underline = A.TextUnderlineValues.None,
            Strike = A.TextStrikeValues.NoStrike,
            Kerning = 1200,
            Baseline = 0
        };
        defaultRunProperties.Append(new A.SolidFill(new A.SchemeColor(
        new A.LuminanceModulation() { Val = 65000 },
        new A.LuminanceOffset() { Val = 35000 })
        { Val = A.SchemeColorValues.Text1 }));
        defaultRunProperties.Append(new A.LatinFont() { Typeface = "+mn-lt" });
        defaultRunProperties.Append(new A.EastAsianFont() { Typeface = "+mn-ea" });
        defaultRunProperties.Append(new A.ComplexScriptFont() { Typeface = "+mn-cs" });
        paragraphProperties.Append(defaultRunProperties);
        paragraph.Append(paragraphProperties);
        paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-US" });
        TextProperties.Append(paragraph);
        legend.Append(TextProperties);
        return legend;
    }

    private C.Title CreateTitle()
    {
        C.Title title = new();
        title.Append(new C.Overlay() { Val = false });
        C.ShapeProperties ShapeProperties = new();
        ShapeProperties.Append(new A.NoFill());
        A.Outline ln = new();
        ln.Append(new A.NoFill());
        ShapeProperties.Append(ln);
        ShapeProperties.Append(new A.EffectList());
        title.Append(ShapeProperties);
        C.TextProperties TextProperties = new();
        TextProperties.Append(new A.BodyProperties()
        {
            Rotation = 0,
            UseParagraphSpacing = true,
            VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
            Vertical = A.TextVerticalValues.Horizontal,
            Wrap = A.TextWrappingValues.Square,
            Anchor = A.TextAnchoringTypeValues.Center,
            AnchorCenter = true
        });
        TextProperties.Append(new A.ListStyle());
        A.Paragraph paragraph = new();
        A.ParagraphProperties paragraphProperties = new();
        A.DefaultRunProperties defaultRunProperties = new()
        {
            FontSize = 1862,
            Bold = false,
            Italic = false,
            Underline = A.TextUnderlineValues.None,
            Strike = A.TextStrikeValues.NoStrike,
            Kerning = 1200,
            Spacing = 0,
            Baseline = 0
        };
        defaultRunProperties.Append(new A.SolidFill(new A.SchemeColor(
        new A.LuminanceModulation() { Val = 65000 },
        new A.LuminanceOffset() { Val = 35000 })
        {
            Val = A.SchemeColorValues.Text1
        }));
        defaultRunProperties.Append(new A.LatinFont() { Typeface = "+mn-lt" });
        defaultRunProperties.Append(new A.EastAsianFont() { Typeface = "+mn-ea" });
        defaultRunProperties.Append(new A.ComplexScriptFont() { Typeface = "+mn-cs" });
        paragraphProperties.Append(defaultRunProperties);
        paragraph.Append(paragraphProperties);
        paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-US" });
        TextProperties.Append(paragraph);
        title.Append(TextProperties);
        return title;
    }

    #endregion Private Methods

    protected C.NumberingCache AddNumberCacheValue(ChartData[] Cells, ChartSeriesSetting? ChartSeriesSetting)
    {
        try
        {
            C.NumberingCache NumberingCache = new()
            {
                FormatCode = new C.FormatCode(ChartSeriesSetting?.NumberFormat ?? "General"),
                PointCount = new C.PointCount()
                {
                    Val = (UInt32Value)(uint)Cells.Length
                },
            };
            int count = 0;
            foreach (ChartData Cell in Cells)
            {
                C.StringPoint StringPoint = new()
                {
                    Index = (UInt32Value)(uint)count
                };
                StringPoint.AppendChild(new C.NumericValue(Cell.Value ?? ""));
                NumberingCache.AppendChild(StringPoint);
                ++count;
            }
            return NumberingCache;
        }
        catch
        {
            throw new Exception("Chart. Numeric Ref Error");
        }
    }

    protected C.StringCache AddStringCacheValue(ChartData[] Cells)
    {
        try
        {
            C.StringCache StringCache = new()
            {
                PointCount = new C.PointCount()
                {
                    Val = (UInt32Value)(uint)Cells.Length
                },
            };
            int count = 0;
            foreach (ChartData Cell in Cells)
            {
                C.StringPoint StringPoint = new()
                {
                    Index = (UInt32Value)(uint)count
                };
                StringPoint.AppendChild(new C.NumericValue(Cell.Value ?? ""));
                StringCache.AppendChild(StringPoint);
                ++count;
            }
            return StringCache;
        }
        catch
        {
            throw new Exception("Chart. String Ref Error");
        }
    }
}