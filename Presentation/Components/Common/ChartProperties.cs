// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Spreadsheet_2007;
using OpenXMLOffice.Global_2007;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using System.IO;
using System;
using System.Linq;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	///
	/// </summary>
	public class ChartProperties<ApplicationSpecificSetting> : CommonProperties where ApplicationSpecificSetting : PresentationSetting, new()
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
		internal P.GraphicFrame graphicFrame;
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
			Excel excel = new Excel();
			Worksheet worksheet = excel.AddSheet();
			int rowIndex = 1;
			foreach (DataCell[] dataCells in dataRows)
			{
				worksheet.SetRow(rowIndex, 1, dataCells, new RowProperties());
				++rowIndex;
			}
			excel.SaveAs(stream);
		}
		/// <summary>
		/// </summary>
		/// <returns>
		/// X,Y
		/// </returns>
		internal Tuple<uint, uint> GetPosition()
		{
			return Tuple.Create(chartSetting.applicationSpecificSetting.x, chartSetting.applicationSpecificSetting.y);
		}
		/// <summary>
		/// </summary>
		/// <returns>
		/// Width,Height
		/// </returns>
		internal Tuple<uint, uint> GetSize()
		{
			return Tuple.Create(chartSetting.applicationSpecificSetting.width, chartSetting.applicationSpecificSetting.height);
		}
		/// <summary>
		/// Update Chart Position
		/// </summary>
		/// <param name="X"> in pix
		/// </param>
		/// <param name="Y"> in pix
		/// </param>
		public virtual void UpdatePosition(uint X, uint Y)
		{
			chartSetting.applicationSpecificSetting.x = (uint)ConverterUtils.PixelsToEmu((int)X);
			chartSetting.applicationSpecificSetting.y = (uint)ConverterUtils.PixelsToEmu((int)Y);
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
			chartSetting.applicationSpecificSetting.width = (uint)ConverterUtils.PixelsToEmu((int)Width);
			chartSetting.applicationSpecificSetting.height = (uint)ConverterUtils.PixelsToEmu((int)Height);
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
			return CommonTools.TransposeArray(dataRows).Select(col => col.Select(cell =>
			{
				string numberFormat = cell != null && cell.styleSetting != null && cell.styleSetting.numberFormat != null ? cell.styleSetting.numberFormat : "General";
				DataType dataType = DataType.STRING;
				if (cell != null && cell.dataType == CellDataType.NUMBER)
				{
					dataType = DataType.NUMBER;
				}
				else if (cell != null && cell.dataType == CellDataType.DATE)
				{
					dataType = DataType.DATE;
				}
				return new ChartData
				{
					numberFormat = numberFormat,
					value = cell != null ? cell.cellValue : null,
					dataType = dataType
				};
			}).ToArray()).ToArray();
		}
		/// <summary>
		///
		/// </summary>
		internal void CreateChartGraphicFrame(string relationshipId, uint id, HyperlinkProperties hyperlinkProperties)
		{
			// Load Chart Part To Graphics Frame For Export
			P.NonVisualGraphicFrameProperties nonVisualProperties = new P.NonVisualGraphicFrameProperties()
			{
				NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = id, Name = string.Format("Chart {0}", id) },
				NonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties(),
				ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties()
			};
			if (hyperlinkProperties != null)
			{
				nonVisualProperties.NonVisualDrawingProperties.InsertAt(CreateHyperLink(hyperlinkProperties), 0);
			}
			graphicFrame = new P.GraphicFrame()
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
			return graphicFrame;
		}
	}
}
