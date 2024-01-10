/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

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

        private readonly ChartSetting ChartSetting;
        private readonly Slide CurrentSlide;
        private readonly ChartPart OpenXMLChartPart;
        private P.GraphicFrame? GraphicFrame;

        #endregion Private Fields

        #region Public Constructors

        public Chart(Slide Slide, DataCell[][] DataRows, AreaChartSetting AreaChartSetting)
        {
            ChartSetting = AreaChartSetting;
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, AreaChartSetting);
        }

        public Chart(Slide Slide, DataCell[][] DataRows, BarChartSetting BarChartSetting)
        {
            ChartSetting = BarChartSetting;
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, BarChartSetting);
        }

        public Chart(Slide Slide, DataCell[][] DataRows, ColumnChartSetting ColumnChartSetting)
        {
            ChartSetting = ColumnChartSetting;
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, ColumnChartSetting);
        }

        public Chart(Slide Slide, DataCell[][] DataRows, LineChartSetting LineChartSetting)
        {
            ChartSetting = LineChartSetting;
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, LineChartSetting);
        }

        public Chart(Slide Slide, DataCell[][] DataRows, PieChartSetting PieChartSetting)
        {
            ChartSetting = PieChartSetting;
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, PieChartSetting);
        }

        public Chart(Slide Slide, DataCell[][] DataRows, ScatterChartSetting ScatterChartSetting)
        {
            ChartSetting = ScatterChartSetting;
            OpenXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            CurrentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, ScatterChartSetting);
        }

        #endregion Public Constructors

        #region Public Methods

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
        public (uint, uint) GetPosition()
        {
            return (ChartSetting.X, ChartSetting.Y);
        }

        /// <summary>
        /// </summary>
        /// <returns>
        /// Width,Height
        /// </returns>
        public (uint, uint) GetSize()
        {
            return (ChartSetting.Width, ChartSetting.Height);
        }

        public void Save()
        {
            CurrentSlide.GetSlidePart().Slide.Save();
        }

        public void UpdatePosition(uint X, uint Y)
        {
            ChartSetting.X = X;
            ChartSetting.Y = Y;
            if (GraphicFrame != null)
            {
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = ChartSetting.X, Y = ChartSetting.Y },
                    Extents = new A.Extents { Cx = ChartSetting.Width, Cy = ChartSetting.Height }
                };
            }
        }

        public void UpdateSize(uint Width, uint Height)
        {
            ChartSetting.Width = Width;
            ChartSetting.Height = Height;
            if (GraphicFrame != null)
            {
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = ChartSetting.X, Y = ChartSetting.Y },
                    Extents = new A.Extents { Cx = ChartSetting.Width, Cy = ChartSetting.Height }
                };
            }
        }

        #endregion Public Methods

        #region Internal Methods

        internal P.GraphicFrame GetChartGraphicFrame()
        {
            return GraphicFrame!;
        }

        internal string GetNextChartRelationId()
        {
            return string.Format("rId{0}", GetChartPart().Parts.Count() + 1);
        }

        #endregion Internal Methods

        #region Private Methods

        private void CreateChart(DataCell[][] DataRows, AreaChartSetting AreaChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
                col.Select(Cell => new ChartData
                {
                    Value = Cell?.CellValue,
                    DataType = Cell?.DataType switch
                    {
                        CellDataType.NUMBER => DataType.NUMBER,
                        CellDataType.DATE => DataType.DATE,
                        _ => DataType.STRING
                    }
                }).ToArray()).ToArray();
            AreaChart AreaChart = new(AreaChartSetting, ChartData);
            GetChartPart().ChartSpace = AreaChart.GetChartSpace();
            GetChartStylePart().ChartStyle = AreaChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = AreaChart.GetColorStyle();
            CreateChartGraphicFrame();
        }

        private void CreateChart(DataCell[][] DataRows, BarChartSetting BarChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
               col.Select(Cell => new ChartData
               {
                   Value = Cell?.CellValue,
                   DataType = Cell?.DataType switch
                   {
                       CellDataType.NUMBER => DataType.NUMBER,
                       CellDataType.DATE => DataType.DATE,
                       _ => DataType.STRING
                   }
               }).ToArray()).ToArray();
            BarChart BarChart = new(BarChartSetting, ChartData);
            GetChartPart().ChartSpace = BarChart.GetChartSpace();
            GetChartStylePart().ChartStyle = BarChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = BarChart.GetColorStyle();
            CreateChartGraphicFrame();
        }

        private void CreateChart(DataCell[][] DataRows, ColumnChartSetting ColumnChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
                col.Select(Cell => new ChartData
                {
                    Value = Cell?.CellValue,
                    DataType = Cell?.DataType switch
                    {
                        CellDataType.NUMBER => DataType.NUMBER,
                        CellDataType.DATE => DataType.DATE,
                        _ => DataType.STRING
                    }
                }).ToArray()).ToArray();
            ColumnChart ColumnChart = new(ColumnChartSetting, ChartData);
            GetChartPart().ChartSpace = ColumnChart.GetChartSpace();
            GetChartStylePart().ChartStyle = ColumnChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = ColumnChart.GetColorStyle();
            CreateChartGraphicFrame();
        }

        private void CreateChart(DataCell[][] DataRows, LineChartSetting LineChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
                col.Select(Cell => new ChartData
                {
                    Value = Cell?.CellValue,
                    DataType = Cell?.DataType switch
                    {
                        CellDataType.NUMBER => DataType.NUMBER,
                        CellDataType.DATE => DataType.DATE,
                        _ => DataType.STRING
                    }
                }).ToArray()).ToArray();
            LineChart LineChart = new(LineChartSetting, ChartData);
            GetChartPart().ChartSpace = LineChart.GetChartSpace();
            GetChartStylePart().ChartStyle = LineChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = LineChart.GetColorStyle();
            CreateChartGraphicFrame();
        }

        private void CreateChart(DataCell[][] DataRows, PieChartSetting PieChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
                col.Select(Cell => new ChartData
                {
                    Value = Cell?.CellValue,
                    DataType = Cell?.DataType switch
                    {
                        CellDataType.NUMBER => DataType.NUMBER,
                        CellDataType.DATE => DataType.DATE,
                        _ => DataType.STRING
                    }
                }).ToArray()).ToArray();
            PieChart PieChart = new(PieChartSetting, ChartData);
            GetChartPart().ChartSpace = PieChart.GetChartSpace();
            GetChartStylePart().ChartStyle = PieChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = PieChart.GetColorStyle();
            CreateChartGraphicFrame();
        }

        private void CreateChart(DataCell[][] DataRows, ScatterChartSetting ScatterChartSetting)
        {
            LoadDataToExcel(DataRows);
            // Prepare Excel Data for PPT Cache
            ChartData[][] ChartData = CommonTools.TransposeArray(DataRows).Select(col =>
                col.Select(Cell => new ChartData
                {
                    Value = Cell?.CellValue,
                    DataType = Cell?.DataType switch
                    {
                        CellDataType.NUMBER => DataType.NUMBER,
                        CellDataType.DATE => DataType.DATE,
                        _ => DataType.STRING
                    }
                }).ToArray()).ToArray();
            ScatterChart ScatterChart = new(ScatterChartSetting, ChartData);
            GetChartPart().ChartSpace = ScatterChart.GetChartSpace();
            GetChartStylePart().ChartStyle = ScatterChart.GetChartStyle();
            GetChartColorStylePart().ColorStyle = ScatterChart.GetColorStyle();
            CreateChartGraphicFrame();
        }

        private void CreateChartGraphicFrame()
        {
            // Load Chart Part To Graphics Frame For Export
            string? relationshipId = CurrentSlide.GetSlidePart().GetIdOfPart(GetChartPart());
            P.NonVisualGraphicFrameProperties NonVisualProperties = new()
            {
                NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = (uint)CurrentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), Name = "Chart" },
                NonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties(),
                ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties()
            };
            GraphicFrame = new()
            {
                NonVisualGraphicFrameProperties = NonVisualProperties,
                Transform = new P.Transform(
                   new A.Offset
                   {
                       X = ChartSetting.X,
                       Y = ChartSetting.Y
                   },
                   new A.Extents
                   {
                       Cx = ChartSetting.Width,
                       Cy = ChartSetting.Height
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
                Worksheet.SetRow(RowIndex, 1, DataCells, new RowProperties());
                ++RowIndex;
            }
            spreadsheet.Save();
        }

        #endregion Private Methods
    }
}