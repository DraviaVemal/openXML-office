// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OpenXMLOffice.Spreadsheet_2007
{
    /// <summary>
    /// Shape Class For Presentation shape manipulation
    /// </summary>
    public class Shape : CommonProperties
    {
        private readonly XDR.Shape openXMLShape = new XDR.Shape();
        private readonly Worksheet worksheet;
        internal Shape(Worksheet _worksheet)
        {
            worksheet = _worksheet;
        }

        /// <summary>
        /// Remove Found Shape
        /// </summary>
        public void RemoveShape()
        {
            openXMLShape.Remove();
        }

        internal Shape MakeLine<ApplicationSpecificSetting, LineColorOption>(ShapeLineModel<ApplicationSpecificSetting, LineColorOption> lineModel)
        where ApplicationSpecificSetting : ExcelSetting, new()
        where LineColorOption : class, IColorOptions, new()
        {
            XDR.TwoCellAnchor twoCellAnchor = worksheet.CreateTwoCellAnchor(new TwoCellAnchorModel()
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
            });
            worksheet.GetDrawing().AppendChild(twoCellAnchor);
            return this;
        }

        internal Shape MakeRectangle<ApplicationSpecificSetting, LineColorOption, FillColorOption>(ShapeRectangleModel<ApplicationSpecificSetting, LineColorOption, FillColorOption> rectangleModel)
        where ApplicationSpecificSetting : class, ISizeAndPosition, new()
        where LineColorOption : class, IColorOptions, new()
        where FillColorOption : class, IColorOptions, new()
        {
            XDR.TwoCellAnchor twoCellAnchor = worksheet.CreateTwoCellAnchor(new TwoCellAnchorModel()
            {

            });
            worksheet.GetDrawing().AppendChild(twoCellAnchor);
            return this;
        }

        internal Shape MakeArrow<ApplicationSpecificSetting, LineColorOption, FillColorOption>(ShapeArrowModel<ApplicationSpecificSetting, LineColorOption, FillColorOption> arrowModel)
        where ApplicationSpecificSetting : class, ISizeAndPosition, new()
        where LineColorOption : class, IColorOptions, new()
        where FillColorOption : class, IColorOptions, new()
        {
            XDR.TwoCellAnchor twoCellAnchor = worksheet.CreateTwoCellAnchor(new TwoCellAnchorModel()
            {

            });
            worksheet.GetDrawing().AppendChild(twoCellAnchor);
            return this;
        }
    }
}