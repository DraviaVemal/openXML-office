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
		private readonly P.Slide openXMLSlide = new P.Slide();
		/// <summary>
		///
		/// </summary>
		public bool ShowHideSlide
		{
			get
			{
				return openXMLSlide.Show;
			}
			set
			{
				openXMLSlide.Show = value;
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
				openXMLSlide = OpenXMLSlide;
			}
			else
			{
				CommonSlideData commonSlideData = new CommonSlideData(PresentationConstants.CommonSlideDataType.SLIDE, PresentationConstants.SlideLayoutType.BLANK);
				openXMLSlide.CommonSlideData = commonSlideData.GetCommonSlideData();
				openXMLSlide.ColorMapOverride = new P.ColorMapOverride()
				{
					MasterColorMapping = new A.MasterColorMapping()
				};
				openXMLSlide.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
				openXMLSlide.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			}
			openXMLSlide.Show = !slideModel.hideSlide;
		}

		/// <summary>
		/// Adds a Area chart to the slide.
		/// </summary>
		public Chart<ApplicationSpecificSetting, CategoryAxis, ValueAxis, ValueAxis> AddChart<ApplicationSpecificSetting>(DataCell[][] DataCells, AreaChartSetting<ApplicationSpecificSetting> AreaChartSetting)
			where ApplicationSpecificSetting : PresentationSetting, new()
		{
			Chart<ApplicationSpecificSetting, CategoryAxis, ValueAxis, ValueAxis> Chart = new Chart<ApplicationSpecificSetting, CategoryAxis, ValueAxis, ValueAxis>(this, DataCells, AreaChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Bar chart to the slide.
		/// </summary>
		public Chart<ApplicationSpecificSetting, ValueAxis, CategoryAxis, ValueAxis> AddChart<ApplicationSpecificSetting>(DataCell[][] DataCells, BarChartSetting<ApplicationSpecificSetting> BarChartSetting)
			where ApplicationSpecificSetting : PresentationSetting, new()
		{
			Chart<ApplicationSpecificSetting, ValueAxis, CategoryAxis, ValueAxis> Chart = new Chart<ApplicationSpecificSetting, ValueAxis, CategoryAxis, ValueAxis>(this, DataCells, BarChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Column chart to the slide.
		/// </summary>
		public Chart<ApplicationSpecificSetting, CategoryAxis, ValueAxis, ValueAxis> AddChart<ApplicationSpecificSetting>(DataCell[][] DataCells, ColumnChartSetting<ApplicationSpecificSetting> ColumnChartSetting)
			where ApplicationSpecificSetting : PresentationSetting, new()
		{
			Chart<ApplicationSpecificSetting, CategoryAxis, ValueAxis, ValueAxis> Chart = new Chart<ApplicationSpecificSetting, CategoryAxis, ValueAxis, ValueAxis>(this, DataCells, ColumnChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Line chart to the slide.
		/// </summary>
		public Chart<ApplicationSpecificSetting, CategoryAxis, ValueAxis, ValueAxis> AddChart<ApplicationSpecificSetting>(DataCell[][] DataCells, LineChartSetting<ApplicationSpecificSetting> LineChartSetting)
			where ApplicationSpecificSetting : PresentationSetting, new()
		{
			Chart<ApplicationSpecificSetting, CategoryAxis, ValueAxis, ValueAxis> Chart = new Chart<ApplicationSpecificSetting, CategoryAxis, ValueAxis, ValueAxis>(this, DataCells, LineChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Pie chart to the slide.
		/// </summary>
		public Chart<ApplicationSpecificSetting, ValueAxis, ValueAxis, ValueAxis> AddChart<ApplicationSpecificSetting>(DataCell[][] DataCells, PieChartSetting<ApplicationSpecificSetting> PieChartSetting)
			where ApplicationSpecificSetting : PresentationSetting, new()
		{
			Chart<ApplicationSpecificSetting, ValueAxis, ValueAxis, ValueAxis> Chart = new Chart<ApplicationSpecificSetting, ValueAxis, ValueAxis, ValueAxis>(this, DataCells, PieChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Scatter chart to the slide.
		/// </summary>
		public Chart<ApplicationSpecificSetting, ValueAxis, ValueAxis, ValueAxis> AddChart<ApplicationSpecificSetting>(DataCell[][] DataCells, ScatterChartSetting<ApplicationSpecificSetting> ScatterChartSetting)
			where ApplicationSpecificSetting : PresentationSetting, new()
		{
			Chart<ApplicationSpecificSetting, ValueAxis, ValueAxis, ValueAxis> Chart = new Chart<ApplicationSpecificSetting, ValueAxis, ValueAxis, ValueAxis>(this, DataCells, ScatterChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Combo chart to the slide.
		/// </summary>
		public Chart<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> AddChart<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType>(DataCell[][] DataCells, ComboChartSetting<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> comboChartSetting)
			where ApplicationSpecificSetting : PresentationSetting, new()
			where XAxisType : class, IAxisTypeOptions, new()
			where YAxisType : class, IAxisTypeOptions, new()
			where ZAxisType : class, IAxisTypeOptions, new()
		{
			Chart<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> Chart = new Chart<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType>(this, DataCells, comboChartSetting);
			GetSlide().CommonSlideData.ShapeTree.Append(Chart.GetChartGraphicFrame());
			return Chart;
		}
		/// <summary>
		/// Adds a Combo chart to the slide.
		/// </summary>
		public P16.Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataCell[][] DataCells, WaterfallChartSetting<ApplicationSpecificSetting> waterfallChartSetting)
			where ApplicationSpecificSetting : PresentationSetting, new()
		{
			P16.Chart<ApplicationSpecificSetting> Chart = new P16.Chart<ApplicationSpecificSetting>(this, DataCells, waterfallChartSetting);
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
			return openXMLSlide;
		}
		internal SlidePart GetSlidePart()
		{
			return openXMLSlide.SlidePart;
		}
		private P.CommonSlideData GetCommonSlideData()
		{
			return openXMLSlide.CommonSlideData;
		}
	}
}
