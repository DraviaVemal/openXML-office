// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

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
		private readonly P.Slide openXMLSlide = new();


		internal Slide(P.Slide? OpenXMLSlide = null)
		{
			if (OpenXMLSlide != null)
			{
				openXMLSlide = OpenXMLSlide;
			}
			else
			{
				CommonSlideData commonSlideData = new(PresentationConstants.CommonSlideDataType.SLIDE, PresentationConstants.SlideLayoutType.BLANK);
				openXMLSlide.CommonSlideData = commonSlideData.GetCommonSlideData();
				openXMLSlide.ColorMapOverride = new P.ColorMapOverride()
				{
					MasterColorMapping = new A.MasterColorMapping()
				};
				openXMLSlide.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
				openXMLSlide.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			}
		}





		/// <summary>
		/// Adds a Area chart to the slide.
		/// </summary>
		public Chart AddChart(Excel.DataCell[][] DataCells, AreaChartSetting AreaChartSetting)
		{
			Chart Chart = new(this, DataCells, AreaChartSetting);
			GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}

		/// <summary>
		/// Adds a Bar chart to the slide.
		/// </summary>
		public Chart AddChart(Excel.DataCell[][] DataCells, BarChartSetting BarChartSetting)
		{
			Chart Chart = new(this, DataCells, BarChartSetting);
			GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}

		/// <summary>
		/// Adds a Column chart to the slide.
		/// </summary>
		public Chart AddChart(Excel.DataCell[][] DataCells, ColumnChartSetting ColumnChartSetting)
		{
			Chart Chart = new(this, DataCells, ColumnChartSetting);
			GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}

		/// <summary>
		/// Adds a Line chart to the slide.
		/// </summary>
		public Chart AddChart(Excel.DataCell[][] DataCells, LineChartSetting LineChartSetting)
		{
			Chart Chart = new(this, DataCells, LineChartSetting);
			GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}

		/// <summary>
		/// Adds a Pie chart to the slide.
		/// </summary>
		public Chart AddChart(Excel.DataCell[][] DataCells, PieChartSetting PieChartSetting)
		{
			Chart Chart = new(this, DataCells, PieChartSetting);
			GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}

		/// <summary>
		/// Adds a Scatter chart to the slide.
		/// </summary>
		public Chart AddChart(Excel.DataCell[][] DataCells, ScatterChartSetting ScatterChartSetting)
		{
			Chart Chart = new(this, DataCells, ScatterChartSetting);
			GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}

		/// <summary>
		/// Adds a Combo chart to the slide.
		/// </summary>
		public Chart AddChart(Excel.DataCell[][] DataCells, ComboChartSetting comboChartSetting)
		{
			Chart Chart = new(this, DataCells, comboChartSetting);
			GetSlide().CommonSlideData!.ShapeTree!.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}

		/// <summary>
		/// Adds a picture to the slide.
		/// </summary>
		public Picture AddPicture(string FilePath, PictureSetting PictureSetting)
		{
			Picture Picture = new(FilePath, this, PictureSetting);
			GetSlide().CommonSlideData!.ShapeTree!.Append(Picture.GetPicture());
			return Picture;
		}

		/// <summary>
		/// Adds a picture to the slide.
		/// </summary>
		public Picture AddPicture(Stream Stream, PictureSetting PictureSetting)
		{
			Picture Picture = new(Stream, this, PictureSetting);
			GetSlide().CommonSlideData!.ShapeTree!.Append(Picture.GetPicture());
			return Picture;
		}

		/// <summary>
		/// Adds a table to the slide.
		/// </summary>
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





		internal string GetNextSlideRelationId()
		{
			return string.Format("rId{0}", GetSlidePart().Parts.Count() + 1);
		}

		internal P.Slide GetSlide()
		{
			return openXMLSlide;
		}

		internal SlidePart GetSlidePart()
		{
			return openXMLSlide.SlidePart!;
		}





		private P.CommonSlideData GetCommonSlideData()
		{
			return openXMLSlide.CommonSlideData!;
		}


	}
}
