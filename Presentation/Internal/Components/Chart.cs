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
        #region Private Fields

        private readonly Slide CurrentSlide;
        private readonly ChartPart OpenXMLChartPart;
        private P.GraphicFrame? GraphicFrame;
        private int Height = 6858000;
        private int Width = 12192000;
        private int X = 0;
        private int Y = 0;

        #endregion Private Fields

        #region Public Constructors

        public Chart(Slide Slide, DataCell[][] DataRows, AreaChartSetting AreaChartSetting)
        {
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, AreaChartSetting);
        }

        public Chart(Slide Slide, DataCell[][] DataRows, BarChartSetting BarChartSetting)
        {
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, BarChartSetting);
        }

        public Chart(Slide Slide, DataCell[][] DataRows, ColumnChartSetting ColumnChartSetting)
        {
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, ColumnChartSetting);
        }

        public Chart(Slide Slide, DataCell[][] DataRows, LineChartSetting LineChartSetting)
        {
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, LineChartSetting);
        }

        public Chart(Slide Slide, DataCell[][] DataRows, PieChartSetting PieChartSetting)
        {
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, PieChartSetting);
        }

        public Chart(Slide Slide, DataCell[][] DataRows, ScatterChartSetting ScatterChartSetting)
        {
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, ScatterChartSetting);
        }

        #endregion Public Constructors

        #region Public Methods

        public P.GraphicFrame GetChartGraphicFrame()
        {
            // Load Chart Part To Graphics Frame For Export
            string? relationshipId = CurrentSlide.GetSlidePart().GetIdOfPart(GetChartPart());
            P.NonVisualGraphicFrameProperties NonVisualProperties = new()
            {
                NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = (UInt32Value)2U, Name = "Chart" },
                NonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties(),
                ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties()
            };
            GraphicFrame = new()
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

        public Spreadsheet GetChartWorkBook()
        {
            Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
            return new(stream);
        }

        /// <summary>
        /// </summary>
        /// <returns>
        /// X,Y
        /// </returns>
        public (int, int) GetPosition()
        {
            return (X, Y);
        }

        /// <summary>
        /// </summary>
        /// <returns>
        /// Width,Height
        /// </returns>
        public (int, int) GetSize()
        {
            return (Width, Height);
        }

        public void Save()
        {
            CurrentSlide.GetSlidePart().Slide.Save();
        }

        public void UpdatePosition(int X, int Y)
        {
            this.X = X;
            this.Y = Y;
            if (GraphicFrame != null)
            {
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = X, Y = Y },
                    Extents = new A.Extents { Cx = Width, Cy = Height }
                };
            }
        }

        public void UpdateSize(int Width, int Height)
        {
            this.Width = Width;
            this.Height = Height;
            if (GraphicFrame != null)
            {
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = X, Y = Y },
                    Extents = new A.Extents { Cx = Width, Cy = Height }
                };
            }
        }

        #endregion Public Methods

        #region Internal Methods

        internal string GetNextChartRelationId()
        {
            return string.Format("rId{0}", GetChartPart().Parts.Count() + 1);
        }

        #endregion Internal Methods

        #region Private Methods

        private P.GraphicFrame CreateChart(DataCell[][] DataRows, AreaChartSetting AreaChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
                col.Select(cell => new ChartData { Value = cell?.CellValue }).ToArray()).ToArray();
            AreaChart AreaChart = new(AreaChartSetting, ChartData);
            GetChartPart().ChartSpace = AreaChart.GetChartSpace();
            GetChartStylePart().ChartStyle = AreaChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = AreaChart.GetColorStyle();
            return GetChartGraphicFrame();
        }

        private P.GraphicFrame CreateChart(DataCell[][] DataRows, BarChartSetting BarChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
               col.Select(cell => new ChartData { Value = cell?.CellValue }).ToArray()).ToArray();
            BarChart BarChart = new(BarChartSetting, ChartData);
            GetChartPart().ChartSpace = BarChart.GetChartSpace();
            GetChartStylePart().ChartStyle = BarChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = BarChart.GetColorStyle();
            return GetChartGraphicFrame();
        }

        private P.GraphicFrame CreateChart(DataCell[][] DataRows, ColumnChartSetting ColumnChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
                col.Select(cell => new ChartData { Value = cell?.CellValue }).ToArray()).ToArray();
            ColumnChart ColumnChart = new(ColumnChartSetting, ChartData);
            GetChartPart().ChartSpace = ColumnChart.GetChartSpace();
            GetChartStylePart().ChartStyle = ColumnChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = ColumnChart.GetColorStyle();
            return GetChartGraphicFrame();
        }

        private P.GraphicFrame CreateChart(DataCell[][] DataRows, LineChartSetting LineChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
                col.Select(cell => new ChartData { Value = cell?.CellValue }).ToArray()).ToArray();
            LineChart LineChart = new(LineChartSetting, ChartData);
            GetChartPart().ChartSpace = LineChart.GetChartSpace();
            GetChartStylePart().ChartStyle = LineChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = LineChart.GetColorStyle();
            return GetChartGraphicFrame();
        }

        private P.GraphicFrame CreateChart(DataCell[][] DataRows, PieChartSetting PieChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
                col.Select(cell => new ChartData { Value = cell?.CellValue }).ToArray()).ToArray();
            PieChart PieChart = new(PieChartSetting, ChartData);
            GetChartPart().ChartSpace = PieChart.GetChartSpace();
            GetChartStylePart().ChartStyle = PieChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = PieChart.GetColorStyle();
            return GetChartGraphicFrame();
        }

        private P.GraphicFrame CreateChart(DataCell[][] DataRows, ScatterChartSetting ScatterChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
                col.Select(cell => new ChartData { Value = cell?.CellValue }).ToArray()).ToArray();
            ScatterChart ScatterChart = new(ScatterChartSetting, ChartData);
            GetChartPart().ChartSpace = ScatterChart.GetChartSpace();
            GetChartStylePart().ChartStyle = ScatterChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = ScatterChart.GetColorStyle();
            return GetChartGraphicFrame();
        }

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

        private void LoadDataToExcel(DataCell[][] DataRows)
        {
            // Load Data To Embeded Sheet
            Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
            Spreadsheet spreadsheet = new(stream);
            Worksheet Worksheet = spreadsheet.AddSheet();
            int RowIndex = 1;
            foreach (DataCell[] DataCells in DataRows)
            {
                Worksheet.SetRow(RowIndex, 1, DataCells);
                ++RowIndex;
            }
            spreadsheet.Save();
        }

        #endregion Private Methods
    }
}