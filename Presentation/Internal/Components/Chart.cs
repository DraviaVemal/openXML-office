// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Excel_2013;
using OpenXMLOffice.Global_2013;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation_2013
{
	/// <summary>
	/// Chart Class Exported out of PPT importing from Global
	/// </summary>
	public class Chart
	{
		private readonly ChartSetting chartSetting;
		private readonly Slide currentSlide;
		private readonly ChartPart openXMLChartPart;
		private P.GraphicFrame? graphicFrame;
		/// <summary>
		/// Create Area Chart with provided settings
		/// </summary>
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
		public Chart(Slide Slide, DataCell[][] DataRows, ScatterChartSetting ScatterChartSetting)
		{
			chartSetting = ScatterChartSetting;
			openXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
			currentSlide = Slide;
			InitialiseChartParts();
			CreateChart(DataRows, ScatterChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart(Slide Slide, DataCell[][] DataRows, ComboChartSetting comboChartSetting)
		{
			chartSetting = comboChartSetting;
			openXMLChartPart = Slide.GetSlidePart().AddNewPart<ChartPart>(Slide.GetNextSlideRelationId());
			currentSlide = Slide;
			InitialiseChartParts();
			CreateChart(DataRows, comboChartSetting);
		}


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





		internal P.GraphicFrame GetChartGraphicFrame()
		{
			return graphicFrame!;
		}

		internal string GetNextChartRelationId()
		{
			return string.Format("rId{0}", GetChartPart().Parts.Count() + 1);
		}





		private void CreateChart(DataCell[][] dataRows, AreaChartSetting areaChartSetting)
		{
			LoadDataToExcel(dataRows);
			// Prepare Excel Data for PPT Cache
			ChartData[][] ChartData = CommonTools.TransposeArray(dataRows).Select(col =>
				col.Select(Cell => new ChartData
				{
					numberFormat = Cell?.styleSetting?.numberFormat ?? "General",
					value = Cell?.cellValue,
					dataType = Cell?.dataType switch
					{
						CellDataType.NUMBER => DataType.NUMBER,
						CellDataType.DATE => DataType.DATE,
						_ => DataType.STRING
					}
				}).ToArray()).ToArray();
			AreaChart areaChart = new(areaChartSetting, ChartData);
			GetChartPart().ChartSpace = areaChart.GetChartSpace();
			GetChartStylePart().ChartStyle = AreaChart.GetChartStyle();
			GetChartColorStylePart().ColorStyle = AreaChart.GetColorStyle();
			CreateChartGraphicFrame();
		}

		private void CreateChart(DataCell[][] dataRows, BarChartSetting barChartSetting)
		{
			LoadDataToExcel(dataRows);
			// Prepare Excel Data for PPT Cache
			ChartData[][] ChartData = CommonTools.TransposeArray(dataRows).Select(col =>
			   col.Select(Cell => new ChartData
			   {
				   numberFormat = Cell?.styleSetting?.numberFormat ?? "General",
				   value = Cell?.cellValue,
				   dataType = Cell?.dataType switch
				   {
					   CellDataType.NUMBER => DataType.NUMBER,
					   CellDataType.DATE => DataType.DATE,
					   _ => DataType.STRING
				   }
			   }).ToArray()).ToArray();
			BarChart barChart = new(barChartSetting, ChartData);
			GetChartPart().ChartSpace = barChart.GetChartSpace();
			GetChartStylePart().ChartStyle = BarChart.GetChartStyle();
			GetChartColorStylePart().ColorStyle = BarChart.GetColorStyle();
			CreateChartGraphicFrame();
		}

		private void CreateChart(DataCell[][] dataRows, ColumnChartSetting columnChartSetting)
		{
			LoadDataToExcel(dataRows);
			// Prepare Excel Data for PPT Cache
			ChartData[][] ChartData = CommonTools.TransposeArray(dataRows).Select(col =>
				col.Select(Cell => new ChartData
				{
					numberFormat = Cell?.styleSetting?.numberFormat ?? "General",
					value = Cell?.cellValue,
					dataType = Cell?.dataType switch
					{
						CellDataType.NUMBER => DataType.NUMBER,
						CellDataType.DATE => DataType.DATE,
						_ => DataType.STRING
					}
				}).ToArray()).ToArray();
			ColumnChart columnChart = new(columnChartSetting, ChartData);
			GetChartPart().ChartSpace = columnChart.GetChartSpace();
			GetChartStylePart().ChartStyle = ColumnChart.GetChartStyle();
			GetChartColorStylePart().ColorStyle = ColumnChart.GetColorStyle();
			CreateChartGraphicFrame();
		}

		private void CreateChart(DataCell[][] dataRows, LineChartSetting lineChartSetting)
		{
			LoadDataToExcel(dataRows);
			// Prepare Excel Data for PPT Cache
			ChartData[][] ChartData = CommonTools.TransposeArray(dataRows).Select(col =>
				col.Select(Cell => new ChartData
				{
					numberFormat = Cell?.styleSetting?.numberFormat ?? "General",
					value = Cell?.cellValue,
					dataType = Cell?.dataType switch
					{
						CellDataType.NUMBER => DataType.NUMBER,
						CellDataType.DATE => DataType.DATE,
						_ => DataType.STRING
					}
				}).ToArray()).ToArray();
			LineChart lineChart = new(lineChartSetting, ChartData);
			GetChartPart().ChartSpace = lineChart.GetChartSpace();
			GetChartStylePart().ChartStyle = LineChart.GetChartStyle();
			GetChartColorStylePart().ColorStyle = LineChart.GetColorStyle();
			CreateChartGraphicFrame();
		}

		private void CreateChart(DataCell[][] dataRows, PieChartSetting pieChartSetting)
		{
			LoadDataToExcel(dataRows);
			// Prepare Excel Data for PPT Cache
			ChartData[][] ChartData = CommonTools.TransposeArray(dataRows).Select(col =>
				col.Select(Cell => new ChartData
				{
					numberFormat = Cell?.styleSetting?.numberFormat ?? "General",
					value = Cell?.cellValue,
					dataType = Cell?.dataType switch
					{
						CellDataType.NUMBER => DataType.NUMBER,
						CellDataType.DATE => DataType.DATE,
						_ => DataType.STRING
					}
				}).ToArray()).ToArray();
			PieChart pieChart = new(pieChartSetting, ChartData);
			GetChartPart().ChartSpace = pieChart.GetChartSpace();
			GetChartStylePart().ChartStyle = PieChart.GetChartStyle();
			GetChartColorStylePart().ColorStyle = PieChart.GetColorStyle();
			CreateChartGraphicFrame();
		}

		private void CreateChart(DataCell[][] dataRows, ScatterChartSetting scatterChartSetting)
		{
			LoadDataToExcel(dataRows);
			// Prepare Excel Data for PPT Cache
			ChartData[][] ChartData = CommonTools.TransposeArray(dataRows).Select(col =>
				col.Select(Cell => new ChartData
				{
					numberFormat = Cell?.styleSetting?.numberFormat ?? "General",
					value = Cell?.cellValue,
					dataType = Cell?.dataType switch
					{
						CellDataType.NUMBER => DataType.NUMBER,
						CellDataType.DATE => DataType.DATE,
						_ => DataType.STRING
					}
				}).ToArray()).ToArray();
			ScatterChart scatterChart = new(scatterChartSetting, ChartData);
			GetChartPart().ChartSpace = scatterChart.GetChartSpace();
			GetChartStylePart().ChartStyle = ScatterChart.GetChartStyle();
			GetChartColorStylePart().ColorStyle = ScatterChart.GetColorStyle();
			CreateChartGraphicFrame();
		}

		private void CreateChart(DataCell[][] dataRows, ComboChartSetting comboChartSetting)
		{
			LoadDataToExcel(dataRows);
			// Prepare Excel Data for PPT Cache
			ChartData[][] ChartData = CommonTools.TransposeArray(dataRows).Select(col =>
				col.Select(Cell => new ChartData
				{
					numberFormat = Cell?.styleSetting?.numberFormat ?? "General",
					value = Cell?.cellValue,
					dataType = Cell?.dataType switch
					{
						CellDataType.NUMBER => DataType.NUMBER,
						CellDataType.DATE => DataType.DATE,
						_ => DataType.STRING
					}
				}).ToArray()).ToArray();
			ComboChart comboChart = new(comboChartSetting, ChartData);
			GetChartPart().ChartSpace = comboChart.GetChartSpace();
			GetChartStylePart().ChartStyle = ComboChart.GetChartStyle();
			GetChartColorStylePart().ColorStyle = ComboChart.GetColorStyle();
			CreateChartGraphicFrame();
		}

		private void CreateChartGraphicFrame()
		{
			// Load Chart Part To Graphics Frame For Export
			string? relationshipId = currentSlide.GetSlidePart().GetIdOfPart(GetChartPart());
			P.NonVisualGraphicFrameProperties nonVisualProperties = new()
			{
				NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), Name = "Chart" },
				NonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties(),
				ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties()
			};
			graphicFrame = new()
			{
				NonVisualGraphicFrameProperties = nonVisualProperties,
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

		private void LoadDataToExcel(DataCell[][] dataRows)
		{
			// Load Data To Embeded Sheet
			Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
			Spreadsheet spreadsheet = new(stream);
			Worksheet worksheet = spreadsheet.AddSheet();
			int rowIndex = 1;
			foreach (DataCell[] dataCells in dataRows)
			{
				worksheet.SetRow(rowIndex, 1, dataCells, new RowProperties());
				++rowIndex;
			}
			spreadsheet.Save();
		}


	}
}
