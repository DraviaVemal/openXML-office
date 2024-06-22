// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OpenXMLOffice.Spreadsheet_2007
{
    /// <summary>
    /// Shape Class For Presentation shape manipulation
    /// </summary>
    public class Shape : SpreadSheetCommonProperties
    {
        private XDR.TwoCellAnchor twoCellAnchor = new XDR.TwoCellAnchor();
        private readonly Worksheet worksheet;
        internal Shape(Worksheet _worksheet, XDR.TwoCellAnchor _twoCellAnchor = null)
        {
            worksheet = _worksheet;
            if (twoCellAnchor != null)
            {
                twoCellAnchor = _twoCellAnchor;
            }
        }

        /// <summary>
        /// Remove Found Shape
        /// </summary>
        public void RemoveShape()
        {
            twoCellAnchor.Remove();
        }

        internal void MakeLine<LineColorOption>(LineShapeModel<ExcelSetting, LineColorOption> lineModel)
        where LineColorOption : class, IColorOptions, new()
        {
            twoCellAnchor = worksheet.CreateTwoCellAnchor(new TwoCellAnchorModel<NoFillOptions, LineShapeModel<PresentationSetting, LineColorOption>>()
            {
                from = new AnchorPosition()
                {
                    row = lineModel.applicationSpecificSetting.from.row,
                    rowOffset = lineModel.applicationSpecificSetting.from.rowOffset,
                    column = lineModel.applicationSpecificSetting.from.column,
                    columnOffset = lineModel.applicationSpecificSetting.from.columnOffset,
                },
                to = new AnchorPosition()
                {
                    row = lineModel.applicationSpecificSetting.to.row,
                    rowOffset = lineModel.applicationSpecificSetting.to.rowOffset,
                    column = lineModel.applicationSpecificSetting.to.column,
                    columnOffset = lineModel.applicationSpecificSetting.to.columnOffset,
                },
                shapeModel = new ShapeModel<NoFillOptions, LineShapeModel<PresentationSetting, LineColorOption>>()
                {

                }
            });
            if (twoCellAnchor.Parent == null)
            {
                worksheet.GetDrawing().AppendChild(twoCellAnchor);
            }
        }

        internal void MakeRectangle<LineColorOption, FillColorOption, TextColorOption>(RectangleShapeModel<ExcelSetting, LineColorOption, FillColorOption> rectangleModel)
        where LineColorOption : class, IColorOptions, new()
        where FillColorOption : class, IColorOptions, new()
        where TextColorOption : class, IColorOptions, new()
        {
            XDR.TwoCellAnchor twoCellAnchor = worksheet.CreateTwoCellAnchor(new TwoCellAnchorModel<TextColorOption, RectangleShapeModel<PresentationSetting, LineColorOption, FillColorOption>>()
            {
                from = new AnchorPosition()
                {
                    row = rectangleModel.applicationSpecificSetting.from.row,
                    rowOffset = rectangleModel.applicationSpecificSetting.from.rowOffset,
                    column = rectangleModel.applicationSpecificSetting.from.column,
                    columnOffset = rectangleModel.applicationSpecificSetting.from.columnOffset,
                },
                to = new AnchorPosition()
                {
                    row = rectangleModel.applicationSpecificSetting.to.row,
                    rowOffset = rectangleModel.applicationSpecificSetting.to.rowOffset,
                    column = rectangleModel.applicationSpecificSetting.to.column,
                    columnOffset = rectangleModel.applicationSpecificSetting.to.columnOffset,
                },
                shapeModel = new ShapeModel<TextColorOption, RectangleShapeModel<PresentationSetting, LineColorOption, FillColorOption>>()
                {

                }
            });
            if (twoCellAnchor.Parent == null)
            {
                worksheet.GetDrawing().AppendChild(twoCellAnchor);
            }
        }

        internal void MakeArrow<LineColorOption, FillColorOption, TextColorOption>(ArrowShapeModel<ExcelSetting, LineColorOption, FillColorOption> arrowModel)
        where LineColorOption : class, IColorOptions, new()
        where FillColorOption : class, IColorOptions, new()
        where TextColorOption : class, IColorOptions, new()
        {
            XDR.TwoCellAnchor twoCellAnchor = worksheet.CreateTwoCellAnchor(new TwoCellAnchorModel<TextColorOption, ArrowShapeModel<PresentationSetting, LineColorOption, FillColorOption>>()
            {
                from = new AnchorPosition()
                {
                    row = arrowModel.applicationSpecificSetting.from.row,
                    rowOffset = arrowModel.applicationSpecificSetting.from.rowOffset,
                    column = arrowModel.applicationSpecificSetting.from.column,
                    columnOffset = arrowModel.applicationSpecificSetting.from.columnOffset,
                },
                to = new AnchorPosition()
                {
                    row = arrowModel.applicationSpecificSetting.to.row,
                    rowOffset = arrowModel.applicationSpecificSetting.to.rowOffset,
                    column = arrowModel.applicationSpecificSetting.to.column,
                    columnOffset = arrowModel.applicationSpecificSetting.to.columnOffset,
                },
                shapeModel = new ShapeModel<TextColorOption, ArrowShapeModel<PresentationSetting, LineColorOption, FillColorOption>>()
                {

                }
            });
            if (twoCellAnchor.Parent == null)
            {
                worksheet.GetDrawing().AppendChild(twoCellAnchor);
            }
        }
    }
}