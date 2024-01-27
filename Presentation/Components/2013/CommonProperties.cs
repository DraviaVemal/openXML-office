// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Excel_2013;
using OpenXMLOffice.Global_2013;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation_2013
{
	/// <summary>
	///
	/// </summary>
	public class ChartProperties
	{
		/// <summary>
		///
		/// </summary>
		protected readonly ChartSetting chartSetting;
		/// <summary>
		///
		/// </summary>
		protected readonly Slide currentSlide;
		/// <summary>
		///
		/// </summary>
		protected P.GraphicFrame? graphicFrame;
		/// <summary>
		///
		/// </summary>
		public ChartProperties(Slide slide, ChartSetting chartSetting)
		{
			this.chartSetting = chartSetting;
			currentSlide = slide;
		}

		/// <summary>
		///
		/// </summary>
		protected void LoadDataToExcel(DataCell[][] dataRows, Stream stream)
		{
			// Load Data To Embeded Sheet
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
		protected void Save()
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

		/// <summary>
		///
		/// </summary>
		/// <param name="dataRows"></param>
		/// <returns></returns>
		protected static ChartData[][] ExcelToPPTdata(DataCell[][] dataRows)
		{
			return CommonTools.TransposeArray(dataRows).Select(col =>
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
		}

		/// <summary>
		///
		/// </summary>
		protected void CreateChartGraphicFrame(string relationshipId, uint id)
		{
			// Load Chart Part To Graphics Frame For Export
			P.NonVisualGraphicFrameProperties nonVisualProperties = new()
			{
				NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = id, Name = "Chart" },
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
		}

		internal P.GraphicFrame GetChartGraphicFrame()
		{
			return graphicFrame!;
		}

	}

}
