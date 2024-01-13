/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    /// <summary>
    /// Represents a slide in a presentation.
    /// </summary>
    public class Slide
    {
        #region Private Fields

        private readonly P.Slide OpenXMLSlide = new();

        #endregion Private Fields

        #region Internal Constructors

        internal Slide(P.Slide? OpenXMLSlide = null)
        {
            if (OpenXMLSlide != null)
            {
                this.OpenXMLSlide = OpenXMLSlide;
            }
            else
            {
                CommonSlideData commonSlideData = new(PresentationConstants.CommonSlideDataType.SLIDE, PresentationConstants.SlideLayoutType.BLANK);
                this.OpenXMLSlide.CommonSlideData = commonSlideData.GetCommonSlideData();
                this.OpenXMLSlide.ColorMapOverride = new P.ColorMapOverride()
                {
                    MasterColorMapping = new A.MasterColorMapping()
                };
                this.OpenXMLSlide.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                this.OpenXMLSlide.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            }
        }

        #endregion Internal Constructors

        #region Public Methods

        /// <summary>
        /// Adds a Area chart to the slide.
        /// </summary>
        /// <param name="DataCells">
        /// </param>
        /// <param name="AreaChartSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public Chart AddChart(Excel.DataCell[][] DataCells, AreaChartSetting AreaChartSetting)
        {
            Chart Chart = new(this, DataCells, AreaChartSetting);
            GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
            return Chart;
        }

        /// <summary>
        /// Adds a Bar chart to the slide.
        /// </summary>
        /// <param name="DataCells">
        /// </param>
        /// <param name="BarChartSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public Chart AddChart(Excel.DataCell[][] DataCells, BarChartSetting BarChartSetting)
        {
            Chart Chart = new(this, DataCells, BarChartSetting);
            GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
            return Chart;
        }

        /// <summary>
        /// Adds a Column chart to the slide.
        /// </summary>
        /// <param name="DataCells">
        /// </param>
        /// <param name="ColumnChartSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public Chart AddChart(Excel.DataCell[][] DataCells, ColumnChartSetting ColumnChartSetting)
        {
            Chart Chart = new(this, DataCells, ColumnChartSetting);
            GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
            return Chart;
        }

        /// <summary>
        /// Adds a Line chart to the slide.
        /// </summary>
        /// <param name="DataCells">
        /// </param>
        /// <param name="LineChartSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public Chart AddChart(Excel.DataCell[][] DataCells, LineChartSetting LineChartSetting)
        {
            Chart Chart = new(this, DataCells, LineChartSetting);
            GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
            return Chart;
        }

        /// <summary>
        /// Adds a Pie chart to the slide.
        /// </summary>
        /// <param name="DataCells">
        /// </param>
        /// <param name="PieChartSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public Chart AddChart(Excel.DataCell[][] DataCells, PieChartSetting PieChartSetting)
        {
            Chart Chart = new(this, DataCells, PieChartSetting);
            GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
            return Chart;
        }

        /// <summary>
        /// Adds a Scatter chart to the slide.
        /// </summary>
        /// <param name="DataCells">
        /// </param>
        /// <param name="ScatterChartSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public Chart AddChart(Excel.DataCell[][] DataCells, ScatterChartSetting ScatterChartSetting)
        {
            Chart Chart = new(this, DataCells, ScatterChartSetting);
            GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
            return Chart;
        }

        /// <summary>
        /// Adds a picture to the slide.
        /// </summary>
        /// <param name="FilePath">
        /// </param>
        /// <param name="PictureSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public Picture AddPicture(string FilePath, PictureSetting PictureSetting)
        {
            Picture Picture = new(FilePath, this, PictureSetting);
            GetSlide().CommonSlideData!.ShapeTree!.Append(Picture.GetPicture());
            return Picture;
        }

        /// <summary>
        /// Adds a picture to the slide.
        /// </summary>
        /// <param name="Stream">
        /// </param>
        /// <param name="PictureSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public Picture AddPicture(Stream Stream, PictureSetting PictureSetting)
        {
            Picture Picture = new(Stream, this, PictureSetting);
            GetSlide().CommonSlideData!.ShapeTree!.Append(Picture.GetPicture());
            return Picture;
        }

        /// <summary>
        /// Adds a table to the slide.
        /// </summary>
        /// <param name="DataCells">
        /// </param>
        /// <param name="TableSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public Table AddTable(TableRow[] DataCells, TableSetting TableSetting)
        {
            Table Table = new(DataCells, TableSetting);
            P.GraphicFrame GraphicFrame = Table.GetTableGraphicFrame();
            GetSlide().CommonSlideData!.ShapeTree!.Append(GraphicFrame);
            return Table;
        }

        /// <summary>
        /// Adds a text box to the slide.
        /// </summary>
        /// <param name="TextBoxSetting">
        /// </param>
        /// <returns>
        /// </returns>
        public TextBox AddTextBox(TextBoxSetting TextBoxSetting)
        {
            TextBox TextBox = new(TextBoxSetting);
            P.Shape Shape = TextBox.GetTextBoxShape();
            GetSlide().CommonSlideData!.ShapeTree!.Append(Shape);
            return TextBox;
        }

        /// <summary>
        /// Finds a shape by its text.
        /// </summary>
        /// <param name="searchText">
        /// </param>
        /// <returns>
        /// </returns>
        public IEnumerable<Shape> FindShapeByText(string searchText)
        {
            IEnumerable<P.Shape> searchResults = GetCommonSlideData().ShapeTree!.Elements<P.Shape>().Where(shape =>
            {
                return shape.InnerText == searchText;
            });
            return searchResults.Select(shape =>
            {
                return new Shape(shape);
            });
        }

        #endregion Public Methods

        #region Internal Methods

        internal string GetNextSlideRelationId()
        {
            return string.Format("rId{0}", GetSlidePart().Parts.Count() + 1);
        }

        internal P.Slide GetSlide()
        {
            return OpenXMLSlide;
        }

        internal SlidePart GetSlidePart()
        {
            return OpenXMLSlide.SlidePart!;
        }

        #endregion Internal Methods

        #region Private Methods

        private P.CommonSlideData GetCommonSlideData()
        {
            return OpenXMLSlide.CommonSlideData!;
        }

        #endregion Private Methods
    }
}