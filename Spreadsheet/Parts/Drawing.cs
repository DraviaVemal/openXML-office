// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLOffice.Spreadsheet_2013
{
    /// <summary>
    /// 
    /// </summary>
    public class Drawing
    {
        /// <summary>
        /// 
        /// </summary>
        protected static DrawingsPart GetDrawingsPart(Worksheet worksheet)
        {
            if (worksheet.GetWorksheetPart().DrawingsPart == null)
            {
                worksheet.GetWorksheetPart().AddNewPart<DrawingsPart>(worksheet.GetNextSheetPartRelationId());
                worksheet.GetWorksheet().Append(new X.Drawing() { Id = worksheet.GetWorksheetPart().GetIdOfPart(worksheet.GetDrawingsPart()) });
                worksheet.GetWorksheetPart().Worksheet.Save();
                worksheet.GetWorksheetPart().DrawingsPart!.WorksheetDrawing ??= new();
            }
            return worksheet.GetWorksheetPart().DrawingsPart!;
        }

        /// <summary>
        /// 
        /// </summary>
        protected static XDR.WorksheetDrawing GetDrawing(Worksheet worksheet)
        {
            return GetDrawingsPart(worksheet).WorksheetDrawing;
        }

        /// <summary>
        /// 
        /// </summary>
        internal XDR.TwoCellAnchor CreateTwoCellAnchor(TwoCellAnchorModel twoCellAnchorModel)
        {
            XDR.TwoCellAnchor twoCellAnchor = new(new XDR.ClientData())
            {
                EditAs = XDR.EditAsValues.OneCell,
                FromMarker = new()
                {
                    ColumnId = new XDR.ColumnId(twoCellAnchorModel.from.column.ToString()),
                    ColumnOffset = new XDR.ColumnOffset(twoCellAnchorModel.from.columnOffset.ToString()),
                    RowId = new XDR.RowId(twoCellAnchorModel.from.row.ToString()),
                    RowOffset = new XDR.RowOffset(twoCellAnchorModel.from.rowOffset.ToString())
                },
                ToMarker = new()
                {
                    ColumnId = new XDR.ColumnId(twoCellAnchorModel.to.column.ToString()),
                    ColumnOffset = new XDR.ColumnOffset(twoCellAnchorModel.to.columnOffset.ToString()),
                    RowId = new XDR.RowId(twoCellAnchorModel.to.row.ToString()),
                    RowOffset = new XDR.RowOffset(twoCellAnchorModel.to.rowOffset.ToString())
                },
            };
            if (twoCellAnchorModel.anchorEditType != AnchorEditType.NONE)
            {
                twoCellAnchor.EditAs = twoCellAnchorModel.anchorEditType switch
                {
                    AnchorEditType.TWO_CELL => XDR.EditAsValues.TwoCell,
                    AnchorEditType.ABSOLUTE => XDR.EditAsValues.Absolute,
                    _ => XDR.EditAsValues.OneCell
                };
            }
            if (twoCellAnchorModel.drawingGraphicFrame != null)
            {
                twoCellAnchor.AddChild(CreateGraphicFrame(twoCellAnchorModel.drawingGraphicFrame));
            }
            if (twoCellAnchorModel.drawingPictureModel != null)
            {
                twoCellAnchor.AddChild(CreatePicture(twoCellAnchorModel.drawingPictureModel));
            }
            return twoCellAnchor;
        }

        private static XDR.GraphicFrame CreateGraphicFrame(DrawingGraphicFrame drawingGraphicFrame)
        {
            return new()
            {
                NonVisualGraphicFrameProperties = new()
                {
                    NonVisualDrawingProperties = new()
                    {
                        Id = drawingGraphicFrame.id,
                        Name = drawingGraphicFrame.name
                    },
                    NonVisualGraphicFrameDrawingProperties = new()
                },
                Macro = "",
                Graphic = new()
                {
                    GraphicData = new(new C.ChartReference()
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
            return new()
            {
                NonVisualPictureProperties = new()
                {
                    NonVisualDrawingProperties = new()
                    {
                        Id = drawingPictureModel.id,
                        Name = drawingPictureModel.name
                    },
                    NonVisualPictureDrawingProperties = new()
                    {
                        PictureLocks = new() { NoChangeAspect = true }
                    }
                },
                BlipFill = new(new A.Stretch(new A.FillRectangle()))
                {
                    Blip = new()
                    {
                        Embed = drawingPictureModel.blipEmbed
                    }
                },
                ShapeProperties = new(new A.PresetGeometry(new A.AdjustValueList())
                {
                    Preset = A.ShapeTypeValues.Rectangle
                })
            };
        }
    }

}