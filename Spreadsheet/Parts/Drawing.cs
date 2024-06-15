// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLOffice.Global_2007;

namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	///
	/// </summary>
	public class Drawing : SpreadSheetCommonProperties
	{
		/// <summary>
		///
		/// </summary>
		protected DrawingsPart GetDrawingsPart(Worksheet worksheet)
		{
			if (worksheet.GetWorksheetPart().DrawingsPart == null)
			{
				worksheet.GetWorksheetPart().AddNewPart<DrawingsPart>(worksheet.GetNextSheetPartRelationId());
				worksheet.GetWorksheet().Append(new X.Drawing() { Id = worksheet.GetWorksheetPart().GetIdOfPart(worksheet.GetDrawingsPart()) });
				worksheet.GetWorksheetPart().Worksheet.Save();
				if (worksheet.GetWorksheetPart().DrawingsPart.WorksheetDrawing == null)
				{
					worksheet.GetWorksheetPart().DrawingsPart.WorksheetDrawing = new XDR.WorksheetDrawing();
				}
			}
			return worksheet.GetWorksheetPart().DrawingsPart;
		}
		/// <summary>
		///
		/// </summary>
		protected XDR.WorksheetDrawing GetDrawing(Worksheet worksheet)
		{
			return GetDrawingsPart(worksheet).WorksheetDrawing;
		}
		/// <summary>
		///
		/// </summary>
		internal XDR.TwoCellAnchor CreateTwoCellAnchor<TextColorOption>(TwoCellAnchorModel<TextColorOption, NoShape> twoCellAnchorModel)
		where TextColorOption : class, IColorOptions, new()
		{
			XDR.TwoCellAnchor twoCellAnchor = new XDR.TwoCellAnchor(new XDR.ClientData())
			{
				FromMarker = new XDR.FromMarker()
				{
					ColumnId = new XDR.ColumnId(twoCellAnchorModel.from.column.ToString()),
					ColumnOffset = new XDR.ColumnOffset(twoCellAnchorModel.from.columnOffset.ToString()),
					RowId = new XDR.RowId(twoCellAnchorModel.from.row.ToString()),
					RowOffset = new XDR.RowOffset(twoCellAnchorModel.from.rowOffset.ToString())
				},
				ToMarker = new XDR.ToMarker()
				{
					ColumnId = new XDR.ColumnId(twoCellAnchorModel.to.column.ToString()),
					ColumnOffset = new XDR.ColumnOffset(twoCellAnchorModel.to.columnOffset.ToString()),
					RowId = new XDR.RowId(twoCellAnchorModel.to.row.ToString()),
					RowOffset = new XDR.RowOffset(twoCellAnchorModel.to.rowOffset.ToString())
				},
			};
			if (twoCellAnchorModel.anchorEditType != AnchorEditType.NONE)
			{
				switch (twoCellAnchorModel.anchorEditType)
				{
					case AnchorEditType.TWO_CELL:
						twoCellAnchor.EditAs = XDR.EditAsValues.TwoCell;
						break;
					case AnchorEditType.ABSOLUTE:
						twoCellAnchor.EditAs = XDR.EditAsValues.Absolute;
						break;
					default:
						twoCellAnchor.EditAs = XDR.EditAsValues.OneCell;
						break;
				}
			}
			if (twoCellAnchorModel.drawingGraphicFrame != null)
			{
				twoCellAnchor.AddChild(CreateGraphicFrame(twoCellAnchorModel.drawingGraphicFrame));
			}
			if (twoCellAnchorModel.drawingPictureModel != null)
			{
				twoCellAnchor.AddChild(CreatePicture(twoCellAnchorModel.drawingPictureModel));
			}
			if (twoCellAnchorModel.shapeModel != null)
			{
				twoCellAnchor.AddChild(CreateShape(twoCellAnchorModel.shapeModel));
			}
			return twoCellAnchor;
		}

		private static XDR.GraphicFrame CreateGraphicFrame(DrawingGraphicFrame drawingGraphicFrame)
		{
			return new XDR.GraphicFrame()
			{
				Macro = "",
				NonVisualGraphicFrameProperties = new XDR.NonVisualGraphicFrameProperties()
				{
					NonVisualDrawingProperties = new XDR.NonVisualDrawingProperties()
					{
						Id = drawingGraphicFrame.id,
						Name = drawingGraphicFrame.name
					},
					NonVisualGraphicFrameDrawingProperties = new XDR.NonVisualGraphicFrameDrawingProperties()
				},
				Transform = new XDR.Transform()
				{
					Offset = new A.Offset()
					{
						X = 0,
						Y = 0
					},
					Extents = new A.Extents()
					{
						Cx = 0,
						Cy = 0
					}
				},
				Graphic = new A.Graphic()
				{
					GraphicData = new A.GraphicData(new C.ChartReference()
					{
						Id = drawingGraphicFrame.chartId
					})
					{
						Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart",
					}
				},
			};
		}
		private static XDR.Picture CreatePicture(DrawingPictureModel drawingPictureModel)
		{
			XDR.Picture picture = new XDR.Picture()
			{
				NonVisualPictureProperties = new XDR.NonVisualPictureProperties()
				{
					NonVisualDrawingProperties = new XDR.NonVisualDrawingProperties()
					{
						Id = drawingPictureModel.id,
						Name = drawingPictureModel.name
					},
					NonVisualPictureDrawingProperties = new XDR.NonVisualPictureDrawingProperties()
					{
						PictureLocks = new A.PictureLocks() { NoChangeAspect = true }
					}
				},
				BlipFill = new XDR.BlipFill(new A.Stretch(new A.FillRectangle()))
				{
					Blip = new A.Blip()
					{
						Embed = drawingPictureModel.blipEmbed
					}
				},
				ShapeProperties = new XDR.ShapeProperties(new A.PresetGeometry(new A.AdjustValueList())
				{
					Preset = A.ShapeTypeValues.Rectangle
				})
			};
			if (drawingPictureModel.hyperlinkProperties != null)
			{
				picture.NonVisualPictureProperties.NonVisualDrawingProperties.InsertAt(CreateHyperLink(drawingPictureModel.hyperlinkProperties), 0);
			}
			return picture;
		}
	}
}
