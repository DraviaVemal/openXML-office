using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Excel;
using OpenXMLOffice.Global;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    public class Chart
    {
        #region Public Fields

        public int Height = 100;
        public int Width = 100;
        public int X = 0;
        public int Y = 0;

        #endregion Public Fields

        #region Private Fields

        private readonly Slide CurrentSlide;
        private readonly ChartPart OpenXMLChartPart;

        #endregion Private Fields

        #region Public Constructors

        public Chart(Slide Slide)
        {
            CurrentSlide = Slide;
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            InitialiseChartParts();
        }

        public Chart(Slide Slide, ChartPart ChartPart)
        {
            OpenXMLChartPart = ChartPart;
            CurrentSlide = Slide;
        }

        #endregion Public Constructors

        #region Public Methods

        public P.GraphicFrame CreateChart(GlobalConstants.ChartTypes ChartTypes, DataCell[][] DataRows, ChartSetting? chartSetting = null)
        {
            switch (ChartTypes)
            {
                case GlobalConstants.ChartTypes.BAR:
                    BarChart BarChart = new();
                    GetChartPart().ChartSpace = BarChart.GetChartSpace();
                    GetChartStylePart().ChartStyle = BarChart.GetChartStyle();
                    GetChartColorStylePart().ColorStyle = BarChart.GetColorStyle();
                    break;
            }
            // Load Data To Embeded Sheet
            Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
            Spreadsheet spreadsheet = new(stream, SpreadsheetDocumentType.Workbook);
            Worksheet Worksheet = spreadsheet.AddSheet();
            int RowIndex = 1;
            foreach (DataCell[] DataCells in DataRows)
            {
                Worksheet.SetRow(RowIndex, 1, DataCells);
                ++RowIndex;
            }
            spreadsheet.Save();
            // Load Chart Part To Graphics Frame For Export
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
                        new C.ChartReference { Id = relationshipId }
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
            };
            // Save All Changes
            GetChartPart().ChartSpace.Save();
            GetChartStylePart().ChartStyle.Save();
            GetChartColorStylePart().ColorStyle.Save();
            return GraphicFrame;
        }

        public void Save()
        {
            CurrentSlide.GetSlidePart().Slide.Save();
        }

        #endregion Public Methods

        #region Internal Methods

        internal string GetNextChartRelationId()
        {
            return string.Format("rId{0}", GetChartPart().Parts.Count() + 1);
        }

        #endregion Internal Methods

        #region Private Methods

        private ChartColorStylePart GetChartColorStylePart()
        {
            return OpenXMLChartPart.ChartColorStyleParts.FirstOrDefault()!;
        }

        private ChartPart GetChartPart()
        {
            return OpenXMLChartPart;
        }

        private ChartStylePart GetChartStylePart()
        {
            return OpenXMLChartPart.ChartStyleParts.FirstOrDefault()!;
        }

        private void InitialiseChartParts()
        {
            GetChartPart().AddNewPart<EmbeddedPackagePart>(EmbeddedPackagePartType.Xlsx.ContentType, GetNextChartRelationId());
            GetChartPart().AddNewPart<ChartColorStylePart>(GetNextChartRelationId());
            GetChartPart().AddNewPart<ChartStylePart>(GetNextChartRelationId());
        }

        #endregion Private Methods
    }
}