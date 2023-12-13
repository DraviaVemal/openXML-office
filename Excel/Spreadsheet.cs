using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLOffice.Excel
{
    /// <summary>
    /// This class serves as a versatile tool for working with Excel spreadsheets, built upon the foundation of the OpenXML SDK. 
    /// This class offers a wide range of functionalities for handling Excel-related objects and operation 
    /// It is designed to simplify tasks related to Excel file manipulation, including the creation of new Excel files, reading and updating existing files, and processing Excel data from stream
    /// </summary>
    public class Spreadsheet
    {
        /// <summary>
        /// Maintain the master OpenXML Spreadsheet document
        /// </summary>
        private readonly SpreadsheetDocument spreadsheetDocument;
        /// <summary>
        /// 
        /// </summary>
        private WorkbookPart? workbookPart;

        private Sheets? sheets;

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class, allowing you to work with Excel spreadsheet 
        /// It accepts a Existing excel file path and a SpreadsheetDocumentType enumeration value as parameters and creates a corresponding SpreadsheetDocument.
        /// This is also used to update as template.
        /// </summary>
        /// <param name="filePath">Excel File path location</param>
        /// <param name="spreadsheetDocumentType">Excel File Type</param>
        /// <param name="autoSave">Defaults to true. The source document gets updated automatically</param>
        public Spreadsheet(string filePath, SpreadsheetDocumentType spreadsheetDocumentType, bool autoSave = true)
        {
            spreadsheetDocument = SpreadsheetDocument.Create(filePath, spreadsheetDocumentType, autoSave);
            PrepareSpreadsheet();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        public Spreadsheet(string filePath)
        {
            spreadsheetDocument = SpreadsheetDocument.CreateFromTemplate(filePath);
            PrepareSpreadsheet();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="isEditable"></param>
        /// <param name="autoSave"></param>
        public Spreadsheet(string filePath, bool isEditable = true, bool autoSave = true)
        {
            spreadsheetDocument = SpreadsheetDocument.Open(filePath, isEditable, new OpenSettings()
            {
                AutoSave = autoSave
            });
            PrepareSpreadsheet();
        }

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class, allowing you to work with Excel spreadsheet 
        /// It accepts a Stream object and a SpreadsheetDocumentType enumeration value as parameters and creates a corresponding SpreadsheetDocument.
        /// </summary>
        /// <param name="stream">Memory stream to use</param>
        /// <param name="spreadsheetDocumentType">Excel File Type</param>
        /// <param name="autoSave">Defaults to true. The source document gets updated automatically</param>
        public Spreadsheet(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType = SpreadsheetDocumentType.Workbook, bool autoSave = true)
        {
            spreadsheetDocument = SpreadsheetDocument.Create(stream, spreadsheetDocumentType, autoSave);
            PrepareSpreadsheet();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="stream">Memory stream to use</param>
        /// <param name="isEditable"></param>
        /// <param name="autoSave"></param>
        public Spreadsheet(Stream stream, bool isEditable = true, bool autoSave = true)
        {
            spreadsheetDocument = SpreadsheetDocument.Open(stream, isEditable, new OpenSettings()
            {
                AutoSave = autoSave
            });
        }
        /// <summary>
        /// Common Spreadsheet perparation process used by all constructor
        /// </summary>
        private void PrepareSpreadsheet()
        {
            workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook ??= new Workbook();
            sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? new Sheets();
            workbookPart.Workbook.AppendChild(sheets);
            workbookPart.Workbook.Save();
        }
        /// <summary>
        /// Return the current max ID from available sheets
        /// </summary>
        /// <returns></returns>
        private UInt32Value GetMaxSheetId()
        {
            return sheets!.Max(sheet => (sheet as Sheet)?.SheetId) ?? 0;
        }
        /// <summary>
        /// Check if sheet name exist in the sheets list
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private bool CheckIfSheetNameExist(string sheetName)
        {
            Sheet? sheet = sheets!.FirstOrDefault(sheet => (sheet as Sheet)?.Name == sheetName) as Sheet;
            return sheet != null;
        }
        /// <summary>
        /// Returns the Sheet ID for the give Sheet Name
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public int? GetSheetId(string sheetName)
        {
            Sheet? sheet = sheets!.FirstOrDefault(sheet => (sheet as Sheet)?.Name == sheetName) as Sheet;
            if (sheet != null)
            {
                return int.Parse(sheet.Id!.Value!);
            }
            return null;
        }
        /// <summary>
        /// Return the Sheet Name for the given Sheet ID
        /// </summary>
        /// <param name="sheetId"></param>
        /// <returns></returns>
        public string? GetSheetName(string sheetId)
        {
            Sheet? sheet = sheets!.FirstOrDefault(sheet => (sheet as Sheet)?.Id?.Value == sheetId) as Sheet;
            if (sheet != null)
            {
                return sheet.Name;
            }
            return null;
        }
        /// <summary>
        /// Creates a new sheet with the specified name and adds its relevant components to the workbook.
        /// Throws an exception if the sheet name is already in use.
        /// </summary>
        /// <param name="sheetName">The name of the new sheet to be created.</param>
        /// <returns>The newly created sheet.</returns>
        /// <exception cref="ArgumentException">Thrown when the sheet name is already in use within the workbook.</exception>

        public Worksheet AddSheet(string? sheetName = null)
        {
            if (!string.IsNullOrEmpty(sheetName) && CheckIfSheetNameExist(sheetName))
            {
                throw new ArgumentException("Sheet with name already exist.");
            }
            // Check If Sheet Already exist
            WorksheetPart worksheetPart = workbookPart!.AddNewPart<WorksheetPart>();
            Sheet sheet = new()
            {
                Id = spreadsheetDocument.WorkbookPart!.GetIdOfPart(worksheetPart),
                SheetId = GetMaxSheetId() + 1,
                Name = string.IsNullOrEmpty(sheetName) ? string.Format("Sheet{0}", GetMaxSheetId() + 1) : sheetName
            };
            sheets!.Append(sheet);
            worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new SheetData());
            return new Worksheet(worksheetPart.Worksheet, sheet);
        }
        /// <summary>
        /// Retrieves a Worksheet object from an OpenXMLOffice, allowing manipulation of the specified target sheet.
        /// </summary>
        /// <param name="sheetName">The name of the target sheet to be retrieved.</param>
        /// <returns>The Worksheet object representing the target sheet for manipulation.</returns>

        public Worksheet? GetWorksheet(string sheetName)
        {
            Sheet? sheet = sheets!.FirstOrDefault(sheet => (sheet as Sheet)?.Name == sheetName) as Sheet;
            if (sheet == null)
                return null;
            if (workbookPart!.GetPartById(sheet.Id!) is not WorksheetPart worksheetPart)
                return null;
            return new Worksheet(worksheetPart.Worksheet, sheet);
        }
        /// <summary>
        /// Renames an existing sheet in the OpenXMLOffice.
        /// </summary>
        /// <param name="oldSheetName">The current name of the sheet to be renamed.</param>
        /// <param name="newSheetName">The new name to assign to the sheet.</param>
        /// <returns>True if the renaming action is successful; otherwise, false.</returns>
        public bool RenameSheet(string oldSheetName, string newSheetName)
        {
            if (CheckIfSheetNameExist(newSheetName))
            {
                throw new ArgumentException("New Sheet with name already exist.");
            }
            Sheet? sheet = sheets!.FirstOrDefault(sheet => (sheet as Sheet)?.Name == oldSheetName) as Sheet;
            if (sheet == null)
                return false;
            sheet.Name = newSheetName;
            return true;
        }

        /// <summary>
        /// Renames an existing sheet in the OpenXMLOffice.
        /// </summary>
        /// <param name="sheetId">The current name of the sheet to be renamed.</param>
        /// <param name="newSheetName">The new name to assign to the sheet.</param>
        /// <returns>True if the renaming action is successful; otherwise, false.</returns>
        public bool RenameSheet(int sheetId, string newSheetName)
        {
            if (CheckIfSheetNameExist(newSheetName))
            {
                throw new ArgumentException("New Sheet with name already exist.");
            }
            Sheet? sheet = sheets!.FirstOrDefault(sheet => (sheet as Sheet)?.Id?.Value == sheetId.ToString()) as Sheet;
            if (sheet == null)
                return false;
            sheet.Name = newSheetName;
            return true;
        }
        /// <summary>
        /// Removes a sheet with the specified name from the OpenXMLOffice
        /// </summary>
        /// <param name="sheetName">The name of the sheet to be removed.</param>
        /// <returns>True if the sheet is successfully removed; otherwise, false.</returns>
        public bool RemoveSheet(string sheetName)
        {
            Sheet? sheet = sheets!.FirstOrDefault(sheet => (sheet as Sheet)?.Name == sheetName) as Sheet;
            if (sheet != null)
            {
                if (workbookPart!.GetPartById(sheet.Id!) is WorksheetPart worksheetPart)
                {
                    workbookPart.DeletePart(worksheetPart);
                }
                sheet.Remove();
                return true;
            }
            return false;
        }
        /// <summary>
        /// Removes a sheet with the specified ID from the OpenXMLOffice
        /// </summary>
        /// <param name="sheetId">The ID of the sheet to be removed.</param>
        /// <returns>True if the sheet with the given ID is successfully removed; otherwise, false.</returns>
        public bool RemoveSheet(int sheetId)
        {
            Sheet? sheet = sheets!.FirstOrDefault(sheet => (sheet as Sheet)?.Id?.Value == sheetId.ToString()) as Sheet;
            if (sheet != null)
            {
                if (workbookPart!.GetPartById(sheet.Id!) is WorksheetPart worksheetPart)
                {
                    workbookPart.DeletePart(worksheetPart);
                }
                sheet.Remove();
                return true;
            }
            return false;
        }

        /// <summary>
        /// Save the active file with all new updates
        /// </summary>
        public void Save()
        {
            spreadsheetDocument.Save();
            spreadsheetDocument.Dispose();
        }
        /// <summary>
        /// Save Copy of the content that updated to the source file
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveAs(string filePath)
        {
            throw new NotImplementedException();
        }

    }
}
