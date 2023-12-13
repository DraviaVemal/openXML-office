using OpenXMLOffice.Global;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using OpenXMLOffice.Excel;

namespace OpenXMLOffice.Presentation;
public class Chart
{
    private readonly ChartPart OpenXMLChartPart;
    private readonly Slide CurrentSlide;
    public int X = 0;
    public int Y = 0;
    public int Height = 100;
    public int Width = 100;
    public Chart(Slide Slide)
    {
        CurrentSlide = Slide;
        OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
        OpenXMLChartPart.AddNewPart<ChartStylePart>(GetNextChartRelationId());
        OpenXMLChartPart.AddNewPart<ChartColorStylePart>(GetNextChartRelationId());
        OpenXMLChartPart.AddNewPart<EmbeddedPackagePart>(EmbeddedPackagePartType.Xlsx.ContentType, GetNextChartRelationId());
    }

    private ChartPart GetChartPart()
    {
        return OpenXMLChartPart;
    }
    internal string GetNextChartRelationId()
    {
        return string.Format("rId{0}", GetChartPart().Parts.Count() + 1);
    }
    public Chart(Slide Slide, ChartPart ChartPart)
    {
        OpenXMLChartPart = ChartPart;
        CurrentSlide = Slide;
    }

    public P.GraphicFrame CreateChart(GlobalConstants.ChartTypes ChartTypes)
    {
        ChartSpace ChartSpace = new();
        switch (ChartTypes)
        {
            default:
                ChartBase ChartBase = new();
                ChartSpace = ChartBase.CreateChartSpace();
                break;
        }
        GetChartPart().ChartSpace = ChartSpace;
        string? relationshipId = CurrentSlide.GetSlidePart().GetIdOfPart(GetChartPart());
        P.NonVisualGraphicFrameProperties NonVisualProperties = new()
        {
            NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = (UInt32Value)2U, Name = "Chart" },
            NonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties(),
            ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties()
        };
        P.GraphicFrame GraphicFrame = new()
        {
            NonVisualGraphicFrameProperties = NonVisualProperties,
            Transform = new P.Transform(
                new A.Offset
                {
                    X = X,
                    Y = Y
                },
                new A.Extents
                {
                    Cx = Width,
                    Cy = Height
                }),
            Graphic = new A.Graphic(
                new A.GraphicData(
                    new ChartReference { Id = relationshipId }
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
        };
        GetChartPart().ChartSpace.Save();
        Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
        Spreadsheet spreadsheet = new(stream, SpreadsheetDocumentType.Workbook);
        Worksheet Worksheet = spreadsheet.AddSheet();
        Worksheet.SetRow(1, 1, new DataCell[5]{
            new(){
                CellValue = "test1",
                DataType = CellDataType.STRING
            },
             new(){
                CellValue = "test2",
                DataType = CellDataType.STRING
            },
             new(){
                CellValue = "test3",
                DataType = CellDataType.STRING
            },
             new(){
                CellValue = "test4",
                DataType = CellDataType.STRING
            },
             new(){
                CellValue = "test5",
                DataType = CellDataType.STRING
            }
        }, new RowProperties()
        {
            height = 20
        });
        spreadsheet.Save();
        GetChartPart().ChartSpace.Save();
        return GraphicFrame;
    }

    public void Save()
    {
        CurrentSlide.GetSlidePart().Slide.Save();
    }

}