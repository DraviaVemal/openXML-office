// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Excel;
using OpenXMLOffice.Global;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    /// <summary>
    /// Chart Class Exported out of PPT importing from Global
    /// </summary>
    public class Chart
    {
        #region Private Fields

        private readonly ChartSetting chartSetting;
        private readonly Slide currentSlide;
        private readonly ChartPart openXMLChartPart;
        private P.GraphicFrame? graphicFrame;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Create Area Chart with provided settings
        /// </summary>
        /// <param name="Slide">
        /// </param>
        /// <param name="DataRows">
        /// </param>
        /// <param name="AreaChartSetting">
        /// </param>
        public Chart(Slide Slide, DataCell[][] DataRows, AreaChartSetting AreaChartSetting)
        {
            chartSetting = AreaChartSetting;
            openXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            currentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, AreaChartSetting);
        }

        /// <summary>
        /// Create Bar Chart with provided settings
        /// </summary>
        /// <param name="Slide">
        /// </param>
        /// <param name="DataRows">
        /// </param>
        /// <param name="BarChartSetting">
        /// </param>
        public Chart(Slide Slide, DataCell[][] DataRows, BarChartSetting BarChartSetting)
        {
            chartSetting = BarChartSetting;
            openXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            currentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, BarChartSetting);
        }

        /// <summary>
        /// Create Column Chart with provided settings
        /// </summary>
        /// <param name="Slide">
        /// </param>
        /// <param name="DataRows">
        /// </param>
        /// <param name="ColumnChartSetting">
        /// </param>
        public Chart(Slide Slide, DataCell[][] DataRows, ColumnChartSetting ColumnChartSetting)
        {
            chartSetting = ColumnChartSetting;
            openXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            currentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, ColumnChartSetting);
        }

        /// <summary>
        /// Create Line Chart with provided settings
        /// </summary>
        /// <param name="Slide">
        /// </param>
        /// <param name="DataRows">
        /// </param>
        /// <param name="LineChartSetting">
        /// </param>
        public Chart(Slide Slide, DataCell[][] DataRows, LineChartSetting LineChartSetting)
        {
            chartSetting = LineChartSetting;
            openXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            currentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, LineChartSetting);
        }

        /// <summary>
        /// Create Pie Chart with provided settings
        /// </summary>
        /// <param name="Slide">
        /// </param>
        /// <param name="DataRows">
        /// </param>
        /// <param name="PieChartSetting">
        /// </param>
        public Chart(Slide Slide, DataCell[][] DataRows, PieChartSetting PieChartSetting)
        {
            chartSetting = PieChartSetting;
            openXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            currentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, PieChartSetting);
        }

        /// <summary>
        /// Create Scatter Chart with provided settings
        /// </summary>
        /// <param name="Slide">
        /// </param>
        /// <param name="DataRows">
        /// </param>
        /// <param name="ScatterChartSetting">
        /// </param>
        public Chart(Slide Slide, DataCell[][] DataRows, ScatterChartSetting ScatterChartSetting)
        {
            chartSetting = ScatterChartSetting;
            openXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
            currentSlide = Slide;
            InitialiseChartParts();
            CreateChart(DataRows, ScatterChartSetting);
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Get Worksheet control for the chart embedded object
        /// </summary>
        /// <returns>
        /// </returns>
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
            return (chartSetting.x, chartSetting.y);
        }

        /// <summary>
        /// </summary>
        /// <returns>
        /// Width,Height
        /// </returns>
        public (uint, uint) GetSize()
        {
            return (chartSetting.width, chartSetting.height);
        }

        /// <summary>
        /// Save Chart Part
        /// </summary>
        public void Save()
        {
            currentSlide.GetSlidePart().Slide.Save();
        }

        /// <summary>
        /// Update Chart Position
        /// </summary>
        /// <param name="X">
        /// </param>
        /// <param name="Y">
        /// </param>
        public void UpdatePosition(uint X, uint Y)
        {
            chartSetting.x = X;
            chartSetting.y = Y;
            if (graphicFrame != null)
            {
                graphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = chartSetting.x, Y = chartSetting.y },
                    Extents = new A.Extents { Cx = chartSetting.width, Cy = chartSetting.height }
                };
            }
        }

        /// <summary>
        /// Update Chart Size
        /// </summary>
        /// <param name="Width">
        /// </param>
        /// <param name="Height">
        /// </param>
        public void UpdateSize(uint Width, uint Height)
        {
            chartSetting.width = Width;
            chartSetting.height = Height;
            if (graphicFrame != null)
            {
                graphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = chartSetting.x, Y = chartSetting.y },
                    Extents = new A.Extents { Cx = chartSetting.width, Cy = chartSetting.height }
                };
            }
        }

        #endregion Public Methods

        #region Internal Methods

        internal P.GraphicFrame GetChartGraphicFrame()
        {
            return graphicFrame!;
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
                    numberFormat = Cell?.styleSetting.numberFormat ?? "General",
                    value = Cell?.cellValue,
                    dataType = Cell?.dataType switch
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
                   numberFormat = Cell?.styleSetting.numberFormat ?? "General",
                   value = Cell?.cellValue,
                   dataType = Cell?.dataType switch
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
                    numberFormat = Cell?.styleSetting.numberFormat ?? "General",
                    value = Cell?.cellValue,
                    dataType = Cell?.dataType switch
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
                    numberFormat = Cell?.styleSetting.numberFormat ?? "General",
                    value = Cell?.cellValue,
                    dataType = Cell?.dataType switch
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
                    numberFormat = Cell?.styleSetting.numberFormat ?? "General",
                    value = Cell?.cellValue,
                    dataType = Cell?.dataType switch
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
                    numberFormat = Cell?.styleSetting.numberFormat ?? "General",
                    value = Cell?.cellValue,
                    dataType = Cell?.dataType switch
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
            string? relationshipId = currentSlide.GetSlidePart().GetIdOfPart(GetChartPart());
            P.NonVisualGraphicFrameProperties NonVisualProperties = new()
            {
                NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), Name = "Chart" },
                NonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties(),
                ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties()
            };
            graphicFrame = new()
            {
                NonVisualGraphicFrameProperties = NonVisualProperties,
                Transform = new P.Transform(
                   new A.Offset
                   {
                       X = chartSetting.x,
                       Y = chartSetting.y
                   },
                   new A.Extents
                   {
                       Cx = chartSetting.width,
                       Cy = chartSetting.height
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
            return openXMLChartPart.ChartColorStyleParts.FirstOrDefault()!;
        }

        private ChartPart GetChartPart()
        {
            return openXMLChartPart;
        }

        private ChartStylePart GetChartStylePart()
        {
            return openXMLChartPart.ChartStyleParts.FirstOrDefault()!;
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