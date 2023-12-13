using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global;
public class ChartBase
{
    protected C.ChartSpace CreateChartSpace()
    {
        C.ChartSpace ChartSpace = new();
        ChartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        ChartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        ChartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        C.Chart Chart = CreateChart();
        ChartSpace.Append(Chart);
        return ChartSpace;
    }

    private C.Chart CreateChart()
    {
        C.Chart Chart = new()
        {
            Title = CreateTitle(),
            PlotArea = CreateChartPlotArea(),
            Legend = CreateChartLegend(),
        };
        C.Layout Layout = CreateChartLayout();
        Chart.AppendChild(Layout);
        return Chart;
    }

    private C.Title CreateTitle()
    {
        C.Title title = new();
        title.Append(new C.Overlay() { Val = false });
        C.ShapeProperties spPr = new();
        spPr.Append(new A.NoFill());
        A.Outline ln = new();
        ln.Append(new A.NoFill());
        spPr.Append(ln);
        spPr.Append(new A.EffectList());
        title.Append(spPr);
        C.TextProperties txPr = new();
        txPr.Append(new A.BodyProperties()
        {
            Rotation = 0,
            UseParagraphSpacing = true,
            VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
            Vertical = A.TextVerticalValues.Horizontal,
            Wrap = A.TextWrappingValues.Square,
            Anchor = A.TextAnchoringTypeValues.Center,
            AnchorCenter = true
        });
        txPr.Append(new A.ListStyle());
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
        txPr.Append(paragraph);
        title.Append(txPr);
        return title;
    }

    private C.PlotArea CreateChartPlotArea()
    {
        C.PlotArea plotArea = new();
        plotArea.Append(new C.Layout());
        C.BarChart barChart = new(
            new C.BarDirection() { Val = C.BarDirectionValues.Column },
            new C.BarGrouping() { Val = C.BarGroupingValues.Clustered },
            new C.VaryColors() { Val = false });
        barChart.Append(CreateBarChartSeries(0, "Sheet1!$B$1", "Sheet1!$A$2:$A$5", "Sheet1!$B$2:$B$5", "accent1"));
        barChart.Append(CreateBarChartSeries(1, "Sheet1!$C$1", "Sheet1!$A$2:$A$5", "Sheet1!$C$2:$C$5", "accent2"));
        C.DataLabels dLbls = new(
            new C.ShowLegendKey() { Val = false },
            new C.ShowValue() { Val = false },
            new C.ShowCategoryName() { Val = false },
            new C.ShowSeriesName() { Val = false },
            new C.ShowPercent() { Val = false },
            new C.ShowBubbleSize() { Val = false });
        barChart.Append(dLbls);
        barChart.Append(new C.GapWidth() { Val = 219 });
        barChart.Append(new C.Overlap() { Val = -27 });
        barChart.Append(new C.AxisId() { Val = 1362418656 });
        barChart.Append(new C.AxisId() { Val = 1358349936 });
        plotArea.Append(barChart);
        plotArea.Append(CreateCategoryAxis(1362418656, "Sheet1!$A$2:$A$5"));
        plotArea.Append(CreateValueAxis(1358349936));
        C.ShapeProperties spPr = new();
        spPr.Append(new A.NoFill());
        spPr.Append(new A.Outline(new A.NoFill()));
        spPr.Append(new A.EffectList());
        plotArea.Append(spPr);
        return plotArea;
    }
    private C.BarChartSeries CreateBarChartSeries(int seriesIndex, string seriesTextFormula, string categoryFormula, string valueFormula, string accent)
    {
        C.BarChartSeries series = new(
            new C.Index() { Val = new UInt32Value((uint)seriesIndex) },
            new C.Order() { Val = new UInt32Value((uint)seriesIndex) },
            new C.SeriesText(new C.StringReference(new C.Formula(seriesTextFormula))),
            new C.InvertIfNegative() { Val = false });
        C.ShapeProperties spPr = new();
        spPr.Append(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }));
        spPr.Append(new A.Outline(new A.NoFill()));
        spPr.Append(new A.EffectList());
        series.Append(spPr);
        series.Append(new C.CategoryAxisData(new C.StringReference(new C.Formula(categoryFormula))));
        series.Append(new C.Values(new C.NumberReference(new C.Formula(valueFormula))));
        return series;
    }
    private C.CategoryAxis CreateCategoryAxis(UInt32Value axisId, string formula)
    {
        C.CategoryAxis catAx = new(
            new C.AxisId() { Val = axisId },
            new C.Scaling(new C.Orientation() { Val = C.OrientationValues.MinMax }),
            new C.Delete() { Val = false },
            new C.AxisPosition() { Val = C.AxisPositionValues.Bottom },
            new C.MajorTickMark() { Val = C.TickMarkValues.None },
            new C.MinorTickMark() { Val = C.TickMarkValues.None },
            new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis() { Val = axisId },
            new C.Crosses() { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled() { Val = true },
            new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset() { Val = 100 },
            new C.NoMultiLevelLabels() { Val = false });
        C.ShapeProperties spPr = new();
        spPr.Append(new A.NoFill());
        spPr.Append(new A.Outline(new A.NoFill()));
        spPr.Append(new A.EffectList());
        catAx.Append(spPr);
        return catAx;
    }

    private C.ValueAxis CreateValueAxis(UInt32Value axisId)
    {
        C.ValueAxis valAx = new(
            new C.AxisId() { Val = axisId },
            new C.Scaling(new C.Orientation() { Val = C.OrientationValues.MinMax }),
            new C.Delete() { Val = false },
            new C.AxisPosition() { Val = C.AxisPositionValues.Left },
            new C.MajorGridlines(),
            new C.NumberingFormat() { FormatCode = "General", SourceLinked = true },
            new C.MajorTickMark() { Val = C.TickMarkValues.None },
            new C.MinorTickMark() { Val = C.TickMarkValues.None },
            new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis() { Val = axisId },
            new C.Crosses() { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween() { Val = C.CrossBetweenValues.Between });
        C.ShapeProperties spPr = new();
        spPr.Append(new A.NoFill());
        spPr.Append(new A.Outline(new A.NoFill()));
        spPr.Append(new A.EffectList());
        valAx.Append(spPr);
        return valAx;
    }
    private C.Legend CreateChartLegend()
    {
        C.Legend legend = new();
        legend.Append(new C.LegendPosition() { Val = C.LegendPositionValues.Bottom });
        legend.Append(new C.Overlay() { Val = false });
        C.ShapeProperties spPr = new();
        spPr.Append(new A.NoFill());
        A.Outline ln = new();
        ln.Append(new A.NoFill());
        spPr.Append(ln);
        spPr.Append(new A.EffectList());
        legend.Append(spPr);
        C.TextProperties txPr = new();
        txPr.Append(new A.BodyProperties()
        {
            Rotation = 0,
            UseParagraphSpacing = true,
            VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
            Vertical = A.TextVerticalValues.Horizontal,
            Wrap = A.TextWrappingValues.Square,
            Anchor = A.TextAnchoringTypeValues.Center,
            AnchorCenter = true
        });
        txPr.Append(new A.ListStyle());
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
        new A.LuminanceOffset() { Val = 35000 }
        )
        {
            Val = A.SchemeColorValues.Text1
        }));
        defaultRunProperties.Append(new A.LatinFont() { Typeface = "+mn-lt" });
        defaultRunProperties.Append(new A.EastAsianFont() { Typeface = "+mn-ea" });
        defaultRunProperties.Append(new A.ComplexScriptFont() { Typeface = "+mn-cs" });
        paragraphProperties.Append(defaultRunProperties);
        paragraph.Append(paragraphProperties);
        paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-US" });
        txPr.Append(paragraph);
        legend.Append(txPr);
        return legend;
    }

    private C.Layout CreateChartLayout()
    {
        return new();
    }

    protected CS.ChartStyle CreateChartStyles()
    {
        CS.ChartStyle ChartStyle = new()
        {
            AxisTitle = CreateAxisTitle(),
            CategoryAxis = CreateCategoryAxis(),
            ChartArea = CreateChartArea(),
            DataLabel = CreateDataLabel(),
            DataLabelCallout = CreateDataLabelCallout(),
            DataPoint = CreateDataPoint(),
            DataPoint3D = CreateDataPoint3D(),
            DataPointLine = CreateDataPointLine(),
            DataPointMarker = CreateDataPointMarker(),
            DataTableStyle = CreateDataTableStyle(),
            DownBar = CreateDownBar(),
            DropLine = CreateDropLine(),
            ErrorBar = CreateErrorBar(),
            Floor = CreateFloor(),
            GridlineMajor = CreateGridlineMajor(),
            GridlineMinor = CreateGridlineMinor(),
            HiLoLine = CreateHiLoLine(),
            LeaderLine = CreateLeaderLine(),
            LegendStyle = CreateLegendStyle(),
            PlotArea = CreatePlotArea(),
            PlotArea3D = CreatePlotArea3D(),
            SeriesAxis = CreateSeriesAxis(),
            SeriesLine = CreateSeriesLine(),
            TitleStyle = CreateTitleStyle(),
            TrendlineStyle = CreateTrendlineStyle(),
            TrendlineLabel = CreateTrendlineLabel(),
            UpBar = CreateUpBar(),
            ValueAxis = CreateValueAxis(),
            Wall = CreateWall()
        };
        ChartStyle.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        return ChartStyle;
    }
    private CS.AxisTitle CreateAxisTitle()
    {
        CS.AxisTitle axisTitle = new();
        axisTitle.Append(new CS.LineReference() { Index = (UInt32Value)0 });
        axisTitle.Append(new CS.FillReference() { Index = (UInt32Value)0 });
        axisTitle.Append(new CS.EffectReference() { Index = (UInt32Value)0 });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        A.SchemeColor schemeClr = new() { Val = A.SchemeColorValues.Text1 };
        schemeClr.Append(new A.LuminanceModulation() { Val = 65000 });
        schemeClr.Append(new A.LuminanceOffset() { Val = 35000 });
        fontRef.Append(schemeClr);
        axisTitle.Append(fontRef);
        CS.TextCharacterPropertiesType defRPr = new() { FontSize = 1330, Kerning = 1200 };
        axisTitle.Append(defRPr);
        return axisTitle;
    }

    private CS.CategoryAxis CreateCategoryAxis()
    {
        CS.CategoryAxis categoryAxis = new();
        categoryAxis.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        categoryAxis.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        categoryAxis.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        A.SchemeColor schemeClrFont = new() { Val = A.SchemeColorValues.Text1 };
        schemeClrFont.Append(new A.LuminanceModulation() { Val = 65000 });
        schemeClrFont.Append(new A.LuminanceOffset() { Val = 35000 });
        fontRef.Append(schemeClrFont);
        categoryAxis.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
        A.SolidFill solidFill = new();
        A.SchemeColor schemeClrLn = new() { Val = A.SchemeColorValues.Text1 };
        schemeClrLn.Append(new A.LuminanceModulation() { Val = 15000 });
        schemeClrLn.Append(new A.LuminanceOffset() { Val = 85000 });
        solidFill.Append(schemeClrLn);
        ln.Append(solidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        categoryAxis.Append(spPr);
        CS.TextCharacterPropertiesType defRPr = new() { FontSize = 1197, Kerning = 1200 };
        categoryAxis.Append(defRPr);
        return categoryAxis;
    }

    private CS.ChartArea CreateChartArea()
    {
        CS.ChartArea chartArea = new();
        chartArea.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        chartArea.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        chartArea.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        chartArea.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.SolidFill solidFill = new();
        solidFill.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Background1 });
        spPr.Append(solidFill);
        A.Outline ln = new() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
        A.SolidFill lnSolidFill = new();
        A.SchemeColor lnSchemeClr = new() { Val = A.SchemeColorValues.Text1 };
        lnSchemeClr.Append(new A.LuminanceModulation() { Val = 15000 });
        lnSchemeClr.Append(new A.LuminanceOffset() { Val = 85000 });
        lnSolidFill.Append(lnSchemeClr);
        ln.Append(lnSolidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        chartArea.Append(spPr);
        CS.TextCharacterPropertiesType defRPr = new() { FontSize = 1330, Kerning = 1200 };
        chartArea.Append(defRPr);
        return chartArea;
    }

    private CS.DataLabel CreateDataLabel()
    {
        CS.DataLabel dataLabel = new();
        dataLabel.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        dataLabel.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        dataLabel.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        A.SchemeColor schemeClr = new() { Val = A.SchemeColorValues.Text1 };
        schemeClr.Append(new A.LuminanceModulation() { Val = 75000 });
        schemeClr.Append(new A.LuminanceOffset() { Val = 25000 });
        fontRef.Append(schemeClr);
        dataLabel.Append(fontRef);
        CS.TextCharacterPropertiesType defRPr = new() { FontSize = 1197, Kerning = 1200 };
        dataLabel.Append(defRPr);
        return dataLabel;
    }

    private CS.DataLabelCallout CreateDataLabelCallout()
    {
        CS.DataLabelCallout dataLabelCallout = new();
        dataLabelCallout.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        dataLabelCallout.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        dataLabelCallout.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        A.SchemeColor schemeClr = new() { Val = A.SchemeColorValues.Dark1 };
        schemeClr.Append(new A.LuminanceModulation() { Val = 65000 });
        schemeClr.Append(new A.LuminanceOffset() { Val = 35000 });
        fontRef.Append(schemeClr);
        dataLabelCallout.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.SolidFill solidFill = new(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 });
        spPr.Append(solidFill);
        A.Outline ln = new();
        A.SolidFill lnSolidFill = new();
        A.SchemeColor lnSchemeClr = new() { Val = A.SchemeColorValues.Dark1 };
        lnSchemeClr.Append(new A.LuminanceModulation() { Val = 25000 });
        lnSchemeClr.Append(new A.LuminanceOffset() { Val = 75000 });
        lnSolidFill.Append(lnSchemeClr);
        ln.Append(lnSolidFill);
        spPr.Append(ln);
        dataLabelCallout.Append(spPr);
        CS.TextCharacterPropertiesType defRPr = new() { FontSize = 1197, Kerning = 1200 };
        dataLabelCallout.Append(defRPr);
        CS.TextBodyProperties bodyPr = new()
        {
            Rotation = 0,
            UseParagraphSpacing = true,
            VerticalOverflow = A.TextVerticalOverflowValues.Clip,
            HorizontalOverflow = A.TextHorizontalOverflowValues.Clip,
            Vertical = A.TextVerticalValues.Horizontal,
            Wrap = A.TextWrappingValues.Square,
            LeftInset = 36576,
            TopInset = 18288,
            RightInset = 36576,
            BottomInset = 18288,
            Anchor = A.TextAnchoringTypeValues.Center,
            AnchorCenter = true
        };
        bodyPr.Append(new A.ShapeAutoFit());
        dataLabelCallout.Append(bodyPr);
        return dataLabelCallout;
    }

    private CS.DataPoint CreateDataPoint()
    {
        CS.DataPoint dataPoint = new();
        dataPoint.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        CS.FillReference fillRef = new() { Index = (UInt32Value)1U };
        fillRef.Append(new CS.StyleColor() { Val = "auto" });
        dataPoint.Append(fillRef);
        dataPoint.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        dataPoint.Append(fontRef);
        return dataPoint;
    }

    private CS.DataPoint3D CreateDataPoint3D()
    {
        CS.DataPoint3D dataPoint3D = new();
        dataPoint3D.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        CS.FillReference fillRef = new() { Index = (UInt32Value)1U };
        fillRef.Append(new CS.StyleColor() { Val = "auto" });
        dataPoint3D.Append(fillRef);
        dataPoint3D.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        dataPoint3D.Append(fontRef);
        return dataPoint3D;
    }

    private CS.DataPointLine CreateDataPointLine()
    {
        CS.DataPointLine dataPointLine = new();
        CS.LineReference lnRef = new() { Index = (UInt32Value)0U };
        lnRef.Append(new CS.StyleColor() { Val = "auto" });
        dataPointLine.Append(lnRef);
        dataPointLine.Append(new CS.FillReference() { Index = (UInt32Value)1U });
        dataPointLine.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        dataPointLine.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 28575, CapType = A.LineCapValues.Round };
        A.SolidFill solidFill = new(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor });
        ln.Append(solidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        dataPointLine.Append(spPr);
        return dataPointLine;
    }

    private CS.DataPointMarker CreateDataPointMarker()
    {
        CS.DataPointMarker dataPointMarker = new();
        CS.LineReference lnRef = new() { Index = (UInt32Value)0U };
        lnRef.Append(new CS.StyleColor() { Val = "auto" });
        dataPointMarker.Append(lnRef);
        CS.FillReference fillRef = new() { Index = (UInt32Value)1U };
        fillRef.Append(new CS.StyleColor() { Val = "auto" });
        dataPointMarker.Append(fillRef);
        dataPointMarker.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        dataPointMarker.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 9525 };
        A.SolidFill solidFill = new(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor });
        ln.Append(solidFill);
        spPr.Append(ln);
        dataPointMarker.Append(spPr);
        return dataPointMarker;
    }

    private CS.DataTableStyle CreateDataTableStyle()
    {
        CS.DataTableStyle dataTableStyle = new();
        dataTableStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        dataTableStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        dataTableStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        A.SchemeColor schemeClrFont = new() { Val = A.SchemeColorValues.Text1 };
        schemeClrFont.Append(new A.LuminanceModulation() { Val = 65000 });
        schemeClrFont.Append(new A.LuminanceOffset() { Val = 35000 });
        fontRef.Append(schemeClrFont);
        dataTableStyle.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
        A.SolidFill lnSolidFill = new();
        A.SchemeColor lnSchemeClr = new() { Val = A.SchemeColorValues.Text1 };
        lnSchemeClr.Append(new A.LuminanceModulation() { Val = 15000 });
        lnSchemeClr.Append(new A.LuminanceOffset() { Val = 85000 });
        lnSolidFill.Append(lnSchemeClr);
        ln.Append(lnSolidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        dataTableStyle.Append(spPr);
        CS.TextCharacterPropertiesType defRPr = new() { FontSize = 1197, Kerning = 1200 };
        dataTableStyle.Append(defRPr);
        return dataTableStyle;
    }

    private CS.DownBar CreateDownBar()
    {
        CS.DownBar downBar = new();

        downBar.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        downBar.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        downBar.Append(new CS.EffectReference() { Index = (UInt32Value)0U });

        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 });
        downBar.Append(fontRef);

        CS.ShapeProperties spPr = new();
        A.SolidFill solidFill = new(new A.SchemeColor(
            new A.LuminanceModulation() { Val = 65000 },
            new A.LuminanceOffset() { Val = 35000 })
        {
            Val = A.SchemeColorValues.Dark1
        });
        spPr.Append(solidFill);

        A.Outline ln = new() { Width = 9525 };
        A.SolidFill lnSolidFill = new(new A.SchemeColor(
            new A.LuminanceModulation() { Val = 65000 },
            new A.LuminanceOffset() { Val = 35000 })
        {
            Val = A.SchemeColorValues.Text1
        });
        ln.Append(lnSolidFill);
        spPr.Append(ln);

        downBar.Append(spPr);

        return downBar;
    }

    private CS.DropLine CreateDropLine()
    {
        CS.DropLine dropLine = new();
        dropLine.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        dropLine.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        dropLine.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        dropLine.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
        A.SolidFill lnSolidFill = new(new A.SchemeColor(
            new A.LuminanceModulation() { Val = 35000 },
            new A.LuminanceOffset() { Val = 65000 })
        {
            Val = A.SchemeColorValues.Text1
        });
        ln.Append(lnSolidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        dropLine.Append(spPr);
        return dropLine;
    }

    private CS.ErrorBar CreateErrorBar()
    {
        CS.ErrorBar errorBar = new();
        errorBar.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        errorBar.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        errorBar.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        errorBar.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
        A.SolidFill lnSolidFill = new(new A.SchemeColor(
            new A.LuminanceModulation() { Val = 65000 },
            new A.LuminanceOffset() { Val = 35000 })
        {
            Val = A.SchemeColorValues.Text1
        });
        ln.Append(lnSolidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        errorBar.Append(spPr);
        return errorBar;
    }

    private CS.Floor CreateFloor()
    {
        CS.Floor floor = new();
        floor.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        floor.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        floor.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        floor.Append(fontRef);
        CS.ShapeProperties spPr = new();
        spPr.Append(new A.NoFill());
        A.Outline ln = new();
        ln.Append(new A.NoFill());
        spPr.Append(ln);
        floor.Append(spPr);
        return floor;
    }

    private CS.GridlineMajor CreateGridlineMajor()
    {
        CS.GridlineMajor gridlineMajor = new();
        gridlineMajor.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        gridlineMajor.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        gridlineMajor.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        gridlineMajor.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
        A.SolidFill lnSolidFill = new(new A.SchemeColor(
            new A.LuminanceModulation() { Val = 15000 },
            new A.LuminanceOffset() { Val = 85000 })
        { Val = A.SchemeColorValues.Text1 });
        ln.Append(lnSolidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        gridlineMajor.Append(spPr);
        return gridlineMajor;
    }

    private CS.GridlineMinor CreateGridlineMinor()
    {
        CS.GridlineMinor gridlineMinor = new();
        gridlineMinor.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        gridlineMinor.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        gridlineMinor.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        gridlineMinor.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
        A.SolidFill lnSolidFill = new(new A.SchemeColor(
            new A.LuminanceModulation() { Val = 5000 },
            new A.LuminanceOffset() { Val = 95000 })
        { Val = A.SchemeColorValues.Text1 });
        ln.Append(lnSolidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        gridlineMinor.Append(spPr);
        return gridlineMinor;
    }

    private CS.HiLoLine CreateHiLoLine()
    {
        CS.HiLoLine hiLoLine = new();
        hiLoLine.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        hiLoLine.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        hiLoLine.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        hiLoLine.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
        A.SolidFill lnSolidFill = new(
            new A.SchemeColor(new A.LuminanceModulation() { Val = 75000 },
            new A.LuminanceOffset() { Val = 25000 })
            { Val = A.SchemeColorValues.Text1 });
        ln.Append(lnSolidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        hiLoLine.Append(spPr);
        return hiLoLine;
    }

    private CS.LeaderLine CreateLeaderLine()
    {
        CS.LeaderLine leaderLine = new();
        leaderLine.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        leaderLine.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        leaderLine.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        leaderLine.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
        A.SolidFill lnSolidFill = new(new A.SchemeColor(new A.LuminanceModulation() { Val = 35000 },
        new A.LuminanceOffset() { Val = 65000 })
        { Val = A.SchemeColorValues.Text1 });
        ln.Append(lnSolidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        leaderLine.Append(spPr);
        return leaderLine;
    }

    private CS.LegendStyle CreateLegendStyle()
    {
        CS.LegendStyle legendStyle = new();
        legendStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        legendStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        legendStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        A.SchemeColor schemeClr = new() { Val = A.SchemeColorValues.Text1 };
        schemeClr.Append(new A.LuminanceModulation() { Val = 65000 });
        schemeClr.Append(new A.LuminanceOffset() { Val = 35000 });
        fontRef.Append(schemeClr);
        legendStyle.Append(fontRef);
        CS.TextCharacterPropertiesType defRPr = new() { FontSize = 1197, Kerning = 1200 };
        legendStyle.Append(defRPr);
        return legendStyle;
    }

    private CS.PlotArea CreatePlotArea()
    {
        CS.PlotArea plotAreaStyle = new();
        plotAreaStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        plotAreaStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        plotAreaStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });

        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        plotAreaStyle.Append(fontRef);

        return plotAreaStyle;
    }

    private CS.PlotArea3D CreatePlotArea3D()
    {
        CS.PlotArea3D plotArea3DStyle = new();
        plotArea3DStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        plotArea3DStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        plotArea3DStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        plotArea3DStyle.Append(fontRef);

        return plotArea3DStyle;
    }

    private CS.SeriesAxis CreateSeriesAxis()
    {
        CS.SeriesAxis seriesAxisStyle = new();

        seriesAxisStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        seriesAxisStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        seriesAxisStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });

        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        A.SchemeColor schemeClr = new() { Val = A.SchemeColorValues.Text1 };
        schemeClr.Append(new A.LuminanceModulation() { Val = 65000 });
        schemeClr.Append(new A.LuminanceOffset() { Val = 35000 });
        fontRef.Append(schemeClr);
        seriesAxisStyle.Append(fontRef);

        CS.TextCharacterPropertiesType defRPr = new() { FontSize = 1197, Kerning = 1200 };
        seriesAxisStyle.Append(defRPr);

        return seriesAxisStyle;
    }

    private CS.SeriesLine CreateSeriesLine()
    {
        CS.SeriesLine seriesLineStyle = new();
        seriesLineStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        seriesLineStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        seriesLineStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        seriesLineStyle.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
        A.SolidFill lnSolidFill = new(new A.SchemeColor(new A.LuminanceModulation() { Val = 35000 },
        new A.LuminanceOffset() { Val = 65000 })
        {
            Val = A.SchemeColorValues.Text1
        });
        ln.Append(lnSolidFill);
        ln.Append(new A.Round());
        spPr.Append(ln);
        seriesLineStyle.Append(spPr);
        return seriesLineStyle;
    }

    private CS.TitleStyle CreateTitleStyle()
    {
        CS.TitleStyle titleStyle = new();
        titleStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        titleStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        titleStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        A.SchemeColor schemeClr = new() { Val = A.SchemeColorValues.Text1 };
        schemeClr.Append(new A.LuminanceModulation() { Val = 65000 });
        schemeClr.Append(new A.LuminanceOffset() { Val = 35000 });
        fontRef.Append(schemeClr);
        titleStyle.Append(fontRef);
        CS.TextCharacterPropertiesType defRPr = new()
        {
            FontSize = 1862,
            Bold = false,
            Kerning = 1200,
            Spacing = 0,
            Baseline = 0
        };
        titleStyle.Append(defRPr);
        return titleStyle;
    }

    private CS.TrendlineStyle CreateTrendlineStyle()
    {
        CS.TrendlineStyle trendlineStyle = new();
        CS.LineReference lnRef = new() { Index = (UInt32Value)0U };
        lnRef.Append(new CS.StyleColor() { Val = "auto" });
        trendlineStyle.Append(lnRef);
        trendlineStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        trendlineStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        trendlineStyle.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.Outline ln = new() { Width = 19050, CapType = A.LineCapValues.Round };
        A.SolidFill lnSolidFill = new(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor });
        ln.Append(lnSolidFill);
        ln.Append(new A.PresetDash() { Val = A.PresetLineDashValues.SystemDot });
        spPr.Append(ln);
        trendlineStyle.Append(spPr);
        return trendlineStyle;
    }

    private CS.TrendlineLabel CreateTrendlineLabel()
    {
        CS.TrendlineLabel trendlineLabelStyle = new();
        trendlineLabelStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        trendlineLabelStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        trendlineLabelStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        A.SchemeColor schemeClr = new() { Val = A.SchemeColorValues.Text1 };
        schemeClr.Append(new A.LuminanceModulation() { Val = 65000 });
        schemeClr.Append(new A.LuminanceOffset() { Val = 35000 });
        fontRef.Append(schemeClr);
        trendlineLabelStyle.Append(fontRef);
        CS.TextCharacterPropertiesType defRPr = new()
        {
            FontSize = 1197,
            Kerning = 1200
        };
        trendlineLabelStyle.Append(defRPr);
        return trendlineLabelStyle;
    }

    private CS.UpBar CreateUpBar()
    {
        CS.UpBar upBarStyle = new();
        upBarStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        upBarStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        upBarStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 });
        upBarStyle.Append(fontRef);
        CS.ShapeProperties spPr = new();
        A.SolidFill solidFill = new(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 });
        spPr.Append(solidFill);

        A.Outline ln = new() { Width = 9525 };
        A.SolidFill lnSolidFill = new(new A.SchemeColor(
            new A.LuminanceModulation() { Val = 15000 },
            new A.LuminanceOffset() { Val = 85000 })
        { Val = A.SchemeColorValues.Text1 });
        ln.Append(lnSolidFill);
        spPr.Append(ln);
        upBarStyle.Append(spPr);
        return upBarStyle;
    }

    private CS.ValueAxis CreateValueAxis()
    {
        CS.ValueAxis valueAxisStyle = new();
        valueAxisStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        valueAxisStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        valueAxisStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        A.SchemeColor schemeClr = new() { Val = A.SchemeColorValues.Text1 };
        schemeClr.Append(new A.LuminanceModulation() { Val = 65000 });
        schemeClr.Append(new A.LuminanceOffset() { Val = 35000 });
        fontRef.Append(schemeClr);
        valueAxisStyle.Append(fontRef);
        CS.TextCharacterPropertiesType defRPr = new()
        {
            FontSize = 1197,
            Kerning = 1200
        };
        valueAxisStyle.Append(defRPr);
        return valueAxisStyle;
    }

    private CS.Wall CreateWall()
    {
        CS.Wall wallStyle = new();
        wallStyle.Append(new CS.LineReference() { Index = (UInt32Value)0U });
        wallStyle.Append(new CS.FillReference() { Index = (UInt32Value)0U });
        wallStyle.Append(new CS.EffectReference() { Index = (UInt32Value)0U });
        CS.FontReference fontRef = new() { Index = A.FontCollectionIndexValues.Minor };
        fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Text1 });
        wallStyle.Append(fontRef);
        CS.ShapeProperties spPr = new();
        spPr.Append(new A.NoFill());
        A.Outline ln = new();
        ln.Append(new A.NoFill());
        spPr.Append(ln);
        wallStyle.Append(spPr);
        return wallStyle;
    }

    protected CS.ColorStyle CreateColorStyles()
    {
        CS.ColorStyle colorStyle = new() { Method = "cycle", Id = 10 };
        colorStyle.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent1
        });
        colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent2
        }); colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent3
        }); colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent4
        }); colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent5
        }); colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent6
        });
        colorStyle.Append(new CS.ColorStyleVariation());
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 60000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 80000
        }, new A.LuminanceOffset()
        {
            Val = 20000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 80000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 60000
        }, new A.LuminanceOffset()
        {
            Val = 40000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 50000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 70000
        }, new A.LuminanceOffset()
        {
            Val = 30000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 70000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 50000
        }, new A.LuminanceOffset()
        {
            Val = 50000
        }));
        return colorStyle;
    }
}
