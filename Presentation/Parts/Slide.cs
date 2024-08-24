// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Spreadsheet_2007;
using OpenXMLOffice.Global_2007;
using OpenXMLOffice.Global_2016;
using P16 = OpenXMLOffice.Presentation_2016;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using System.Linq;
using System.Collections.Generic;
using System.IO;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Represents a slide in a presentation.
	/// </summary>
	public class Slide
	{
		private readonly P.Slide documentSlide = new P.Slide();
		/// <summary>
		///
		/// </summary>
		public bool ShowHideSlide
		{
			get
			{
				return documentSlide.Show;
			}
			set
			{
				documentSlide.Show = value;
			}
		}
		internal Slide(P.Slide OpenXMLSlide = null, SlideModel slideModel = null)
		{
			if (slideModel == null)
			{
				slideModel = new SlideModel();
			}
			if (OpenXMLSlide != null)
			{
				documentSlide = OpenXMLSlide;
			}
			else
			{
				CommonSlideData commonSlideData = new CommonSlideData(PresentationConstants.CommonSlideDataType.SLIDE, PresentationConstants.SlideLayoutType.BLANK);
				documentSlide.CommonSlideData = commonSlideData.GetCommonSlideData();
				documentSlide.ColorMapOverride = new P.ColorMapOverride()
				{
					MasterColorMapping = new A.MasterColorMapping()
				};
				documentSlide.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
				documentSlide.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			}
			documentSlide.Show = !slideModel.hideSlide;
		}

		/// <summary>
		/// Adds a Area chart to the slide.
		/// </summary>
		public Chart<CategoryAxis, ValueAxis, ValueAxis> AddChart(ColumnCell[][] DataCells, AreaChartSetting<PresentationSetting> AreaChartSetting)
		{
			Chart<CategoryAxis, ValueAxis, ValueAxis> Chart = new Chart<CategoryAxis, ValueAxis, ValueAxis>(this, DataCells, AreaChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Bar chart to the slide.
		/// </summary>
		public Chart<ValueAxis, CategoryAxis, ValueAxis> AddChart(ColumnCell[][] DataCells, BarChartSetting<PresentationSetting> BarChartSetting)
		{
			Chart<ValueAxis, CategoryAxis, ValueAxis> Chart = new Chart<ValueAxis, CategoryAxis, ValueAxis>(this, DataCells, BarChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Column chart to the slide.
		/// </summary>
		public Chart<CategoryAxis, ValueAxis, ValueAxis> AddChart(ColumnCell[][] DataCells, ColumnChartSetting<PresentationSetting> ColumnChartSetting)
		{
			Chart<CategoryAxis, ValueAxis, ValueAxis> Chart = new Chart<CategoryAxis, ValueAxis, ValueAxis>(this, DataCells, ColumnChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Line chart to the slide.
		/// </summary>
		public Chart<CategoryAxis, ValueAxis, ValueAxis> AddChart(ColumnCell[][] DataCells, LineChartSetting<PresentationSetting> LineChartSetting)
		{
			Chart<CategoryAxis, ValueAxis, ValueAxis> Chart = new Chart<CategoryAxis, ValueAxis, ValueAxis>(this, DataCells, LineChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Pie chart to the slide.
		/// </summary>
		public Chart<ValueAxis, ValueAxis, ValueAxis> AddChart(ColumnCell[][] DataCells, PieChartSetting<PresentationSetting> PieChartSetting)
		{
			Chart<ValueAxis, ValueAxis, ValueAxis> Chart = new Chart<ValueAxis, ValueAxis, ValueAxis>(this, DataCells, PieChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Scatter chart to the slide.
		/// </summary>
		public Chart<ValueAxis, ValueAxis, ValueAxis> AddChart(ColumnCell[][] DataCells, ScatterChartSetting<PresentationSetting> ScatterChartSetting)
		{
			Chart<ValueAxis, ValueAxis, ValueAxis> Chart = new Chart<ValueAxis, ValueAxis, ValueAxis>(this, DataCells, ScatterChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Combo chart to the slide.
		/// </summary>
		public Chart<XAxisType, YAxisType, ZAxisType> AddChart<XAxisType, YAxisType, ZAxisType>(ColumnCell[][] DataCells, ComboChartSetting<PresentationSetting, XAxisType, YAxisType, ZAxisType> comboChartSetting)
			where XAxisType : class, IAxisTypeOptions, new()
			where YAxisType : class, IAxisTypeOptions, new()
			where ZAxisType : class, IAxisTypeOptions, new()
		{
			Chart<XAxisType, YAxisType, ZAxisType> Chart = new Chart<XAxisType, YAxisType, ZAxisType>(this, DataCells, comboChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Combo chart to the slide.
		/// </summary>
		public P16.Chart AddChart(ColumnCell[][] DataCells, WaterfallChartSetting<PresentationSetting> waterfallChartSetting)
		{
			P16.Chart Chart = new P16.Chart(this, DataCells, waterfallChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetAlternateContent());
			return Chart;
		}
		/// <summary>
		/// Adds a picture to the slide.
		/// </summary>
		public Picture AddPicture(string FilePath, PictureSetting PictureSetting)
		{
			Picture Picture = new Picture(FilePath, this, PictureSetting);
			return Picture;
		}
		/// <summary>
		/// Adds a picture to the slide.
		/// </summary>
		public Picture AddPicture(Stream Stream, PictureSetting PictureSetting)
		{
			Picture Picture = new Picture(Stream, this, PictureSetting);
			return Picture;
		}
		/// <summary>
		/// Adds a table to the slide.
		/// </summary>
		public Table AddTable(TableRow[] DataCells, TableSetting TableSetting)
		{
			Table Table = new Table(DataCells, TableSetting);
			P.GraphicFrame GraphicFrame = Table.GetTableGraphicFrame();
			GetSlide().CommonSlideData.ShapeTree.Append(GraphicFrame);
			return Table;
		}
		/// <summary>
		/// Adds a text box to the slide.
		/// </summary>
		public TextBox AddTextBox(TextBoxSetting TextBoxSetting)
		{
			TextBox TextBox = new TextBox(this, TextBoxSetting);
			return TextBox;
		}
		/// <summary>
		/// Finds a shape by its text.
		/// </summary>
		public IEnumerable<Shape> FindShapeByText(string searchText)
		{
			IEnumerable<P.Shape> searchResults = GetCommonSlideData().ShapeTree.Elements<P.Shape>().Where(shape =>
			{
				return shape.InnerText == searchText;
			});
			return searchResults.Select(shape =>
			{
				return new Shape(shape);
			});
		}
		/// <summary>
		/// Insert Shape into slide
		/// </summary>
		public Shape AddShape<LineColorOption>(LineShapeModel<PresentationSetting, LineColorOption> lineModel)
			where LineColorOption : class, IColorOptions, new()
		{
			P.Shape openXmlShape = new P.Shape();
			GetSlide().CommonSlideData.ShapeTree.Append(openXmlShape);
			Shape shape = new Shape(openXmlShape);
			shape.MakeLine(lineModel);
			return shape;
		}
		/// <summary>
		/// Insert Shape into slide
		/// </summary>
		public Shape AddShape<LineColorOption, FillColorOption>(RectangleShapeModel<PresentationSetting, LineColorOption, FillColorOption> rectangleModel)
			where LineColorOption : class, IColorOptions, new()
			where FillColorOption : class, IColorOptions, new()
		{
			Shape shape = new Shape();
			shape.MakeRectangle(rectangleModel);
			GetSlide().CommonSlideData.ShapeTree.Append(shape.GetDocumentShape());
			return shape;
		}
		/// <summary>
		/// Insert Shape into slide
		/// </summary>
		public Shape AddShape<LineColorOption, FillColorOption>(ArrowShapeModel<PresentationSetting, LineColorOption, FillColorOption> arrowModel)
			where LineColorOption : class, IColorOptions, new()
			where FillColorOption : class, IColorOptions, new()
		{
			P.Shape openXmlShape = new P.Shape();
			GetSlide().CommonSlideData.ShapeTree.Append(openXmlShape);
			Shape shape = new Shape(openXmlShape);
			shape.MakeArrow(arrowModel);
			return shape;
		}
		/// <summary>
		/// 
		/// </summary>
		public bool ExtractSlideIntoExcel()
		{
			return false;
		}
		internal string GetNextSlideRelationId()
		{
			int nextId = GetSlidePart().Parts.Count() + GetSlidePart().ExternalRelationships.Count() + GetSlidePart().HyperlinkRelationships.Count() + GetSlidePart().DataPartReferenceRelationships.Count();
			do
			{
				++nextId;
			} while (GetSlidePart().Parts.Any(item => item.RelationshipId == string.Format("rId{0}", nextId)) ||
			GetSlidePart().ExternalRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)) ||
			GetSlidePart().HyperlinkRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)) ||
			GetSlidePart().DataPartReferenceRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)));
			return string.Format("rId{0}", nextId);
		}
		internal P.Slide GetSlide()
		{
			return documentSlide;
		}
		internal SlidePart GetSlidePart()
		{
			return documentSlide.SlidePart;
		}
		private P.CommonSlideData GetCommonSlideData()
		{
			return documentSlide.CommonSlideData;
		}
	}
}
