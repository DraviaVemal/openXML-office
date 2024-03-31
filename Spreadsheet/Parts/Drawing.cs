// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using DocumentFormat.OpenXml.Packaging;

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
                worksheet.GetWorksheetPart().Worksheet.Save();
                worksheet.GetWorksheetPart().DrawingsPart!.WorksheetDrawing ??= new();
            }
            return worksheet.GetWorksheetPart().DrawingsPart!;
        }
    }

}