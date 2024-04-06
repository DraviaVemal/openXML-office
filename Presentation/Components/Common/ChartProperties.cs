// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Spreadsheet_2007;
using OpenXMLOffice.Global_2007;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	///
	/// </summary>
	public class ChartProperties<ApplicationSpecificSetting> where ApplicationSpecificSetting : PresentationSetting
	{
		/// <summary>
		///
		/// </summary>
		internal readonly ChartSetting<ApplicationSpecificSetting> chartSetting;
		/// <summary>
		///
		/// </summary>
		internal readonly Slide currentSlide;
		/// <summary>
		///
		/// </summary>
		internal P.GraphicFrame? graphicFrame;
		/// <summary>
		///
		/// </summary>
		internal ChartProperties(Slide slide, ChartSetting<ApplicationSpecificSetting> chartSetting)
		{
			this.chartSetting = chartSetting;
			currentSlide = slide;
		}

		/// <summary>
		///
		/// </summary>
		internal void WriteDataToExcel(DataCell[][] dataRows, Stream stream)
		{
			// Load Data To Embeded Sheet
			Excel excel = new(stream, null);
			Worksheet worksheet = excel.AddSheet();
			int rowIndex = 1;
			foreach (DataCell[] dataCells in dataRows)
			{
				worksheet.SetRow(rowIndex, 1, dataCells, new RowProperties());
				++rowIndex;
			}
			excel.Save();
		}

		/// <summary>
		/// </summary>
		/// <returns>
		/// X,Y
		/// </returns>
		internal (uint, uint) GetPosition()
		{
			return (chartSetting.applicationSpecificSetting.x, chartSetting.applicationSpecificSetting.y);
		}

		/// <summary>
		/// </summary>
		/// <returns>
		/// Width,Height
		/// </returns>
		internal (uint, uint) GetSize()
		{
			return (chartSetting.applicationSpecificSetting.width, chartSetting.applicationSpecificSetting.height);
		}

		/// <summary>
		/// Save Chart Part
		/// </summary>
		internal void Save()
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
		public virtual void UpdatePosition(uint X, uint Y)
		{
			chartSetting.applicationSpecificSetting.x = X;
			chartSetting.applicationSpecificSetting.y = Y;
			if (graphicFrame != null)
			{
				graphicFrame.Transform = new P.Transform
				{
					Offset = new A.Offset { X = chartSetting.applicationSpecificSetting.x, Y = chartSetting.applicationSpecificSetting.y },
					Extents = new A.Extents { Cx = chartSetting.applicationSpecificSetting.width, Cy = chartSetting.applicationSpecificSetting.height }
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
		public virtual void UpdateSize(uint Width, uint Height)
		{
			chartSetting.applicationSpecificSetting.width = Width;
			chartSetting.applicationSpecificSetting.height = Height;
			if (graphicFrame != null)
			{
				graphicFrame.Transform = new P.Transform
				{
					Offset = new A.Offset { X = chartSetting.applicationSpecificSetting.x, Y = chartSetting.applicationSpecificSetting.y },
					Extents = new A.Extents { Cx = chartSetting.applicationSpecificSetting.width, Cy = chartSetting.applicationSpecificSetting.height }
				};
			}
		}

		/// <summary>
		///
		/// </summary>
		/// <param name="dataRows"></param>
		/// <returns></returns>
		internal static ChartData[][] ExcelToPPTdata(DataCell[][] dataRows)
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
		internal void CreateChartGraphicFrame(string relationshipId, uint id)
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
					   X = chartSetting.applicationSpecificSetting.x,
					   Y = chartSetting.applicationSpecificSetting.y
				   },
				   new A.Extents
				   {
					   Cx = chartSetting.applicationSpecificSetting.width,
					   Cy = chartSetting.applicationSpecificSetting.height
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
