// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLOffice.Excel
{
    /// <summary>
    /// This class serves as a versatile tool for working with Excel spreadsheets, built upon the
    /// foundation of the OpenXML SDK. This class offers a wide range of functionalities for
    /// handling Excel-related objects and operation It is designed to simplify tasks related to
    /// Excel file manipulation, including the creation of new Excel files, reading and updating
    /// existing files, and processing Excel data from stream
    /// </summary>
    public class Spreadsheet : SpreadsheetCore
    {
        #region Public Constructors

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class,
        /// allowing you to work with Excel spreadsheet It accepts a Existing excel file path and a
        /// SpreadsheetDocumentType enumeration value as parameters and creates a corresponding
        /// SpreadsheetDocument. This is also used to update as template.
        /// </summary>
        public Spreadsheet(string filePath) : base(filePath) { }

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class,
        /// </summary>
        public Spreadsheet(string filePath, bool isEditable) : base(filePath, isEditable) { }

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class,
        /// allowing you to work with Excel spreadsheet It accepts a Stream object and a
        /// SpreadsheetDocumentType enumeration value as parameters and creates a corresponding SpreadsheetDocument.
        /// </summary>
        public Spreadsheet(Stream stream) : base(stream) { }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Adds a new sheet to the OpenXMLOffice with the Default Sheet Name Pattern
        /// </summary>
        public Worksheet AddSheet()
        {
            return AddSheet(string.Format("Sheet{0}", GetMaxSheetId() + 1));
        }

        /// <summary>
        /// Adds a new sheet to the OpenXMLOffice with the specified name. Throws an exception if
        /// SheetName already exist.
        /// </summary>
        public Worksheet AddSheet(string sheetName)
        {
            if (CheckIfSheetNameExist(sheetName))
            {
                throw new ArgumentException("Sheet with name already exist.");
            }
            // Check If Sheet Already exist
            WorksheetPart worksheetPart = GetWorkbookPart().AddNewPart<WorksheetPart>();
            Sheet sheet = new()
            {
                Id = GetWorkbookPart().GetIdOfPart(worksheetPart),
                SheetId = GetMaxSheetId() + 1,
                Name = sheetName
            };
            GetSheets().Append(sheet);
            worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new SheetData());
            return new Worksheet(worksheetPart.Worksheet, sheet);
        }

        /// <summary>
        /// Returns the Sheet ID for the give Sheet Name
        /// </summary>
        /// <param name="sheetName">
        /// </param>
        /// <returns>
        /// </returns>
        public int? GetSheetId(string sheetName)
        {
            Sheet? sheet = GetSheets().FirstOrDefault(sheet => (sheet as Sheet)?.Name == sheetName) as Sheet;
            if (sheet != null)
            {
                return int.Parse(sheet.Id!.Value!);
            }
            return null;
        }
        /// <summary>
        /// Use this method to create a new style and get the style id
        /// Use of Style Id instead of Style Setting directly in Worksheet Cell is highly recommended for performance
        /// </summary>
        public static uint GetStyleId(CellStyleSetting CellStyleSetting)
        {
            return Styles.Instance.GetCellStyleId(CellStyleSetting);
        }

        /// <summary>
        /// Return the Sheet Name for the given Sheet ID
        /// </summary>
        /// <param name="sheetId">
        /// </param>
        /// <returns>
        /// </returns>
        public string? GetSheetName(string sheetId)
        {
            Sheet? sheet = GetSheets().FirstOrDefault(sheet => (sheet as Sheet)?.Id?.Value == sheetId) as Sheet;
            if (sheet != null)
            {
                return sheet.Name;
            }
            return null;
        }

        /// <summary>
        /// Retrieves a Worksheet object from an OpenXMLOffice, allowing manipulation of the worksheet.
        /// </summary>
        /// <param name="sheetName">
        /// </param>
        /// <returns>
        /// </returns>
        public Worksheet? GetWorksheet(string sheetName)
        {
            Sheet? sheet = GetSheets().FirstOrDefault(sheet => (sheet as Sheet)?.Name == sheetName) as Sheet;
            if (sheet == null) { return null; }
            if (GetWorkbookPart().GetPartById(sheet.Id!) is not WorksheetPart worksheetPart) { return null; }
            return new Worksheet(worksheetPart.Worksheet, sheet);
        }

        /// <summary>
        /// Removes a sheet with the specified name from the OpenXMLOffice
        /// </summary>
        /// <param name="sheetName">
        /// The name of the sheet to be removed.
        /// </param>
        /// <returns>
        /// True if the sheet is successfully removed; otherwise, false.
        /// </returns>
        public bool RemoveSheet(string sheetName)
        {
            Sheet? sheet = GetSheets().FirstOrDefault(sheet => (sheet as Sheet)?.Name == sheetName) as Sheet;
            if (sheet != null)
            {
                if (GetWorkbookPart().GetPartById(sheet.Id!) is WorksheetPart worksheetPart)
                {
                    GetWorkbookPart().DeletePart(worksheetPart);
                }
                sheet.Remove();
                return true;
            }
            return false;
        }

        /// <summary>
        /// Removes a sheet with the specified ID from the OpenXMLOffice
        /// </summary>
        /// <param name="sheetId">
        /// The ID of the sheet to be removed.
        /// </param>
        /// <returns>
        /// True if the sheet with the given ID is successfully removed; otherwise, false.
        /// </returns>
        public bool RemoveSheet(int sheetId)
        {
            Sheet? sheet = GetSheets().FirstOrDefault(sheet => (sheet as Sheet)?.Id?.Value == sheetId.ToString()) as Sheet;
            if (sheet != null)
            {
                if (GetWorkbookPart().GetPartById(sheet.Id!) is WorksheetPart worksheetPart)
                {
                    GetWorkbookPart().DeletePart(worksheetPart);
                }
                sheet.Remove();
                return true;
            }
            return false;
        }

        /// <summary>
        /// Creates a new sheet with the specified name and adds its relevant components to the
        /// workbook. Throws an exception if the sheet name is already in use.
        /// </summary>
        public bool RenameSheet(string oldSheetName, string newSheetName)
        {
            if (CheckIfSheetNameExist(newSheetName))
            {
                throw new ArgumentException("New Sheet with name already exist.");
            }
            Sheet? sheet = GetSheets().FirstOrDefault(sheet => (sheet as Sheet)?.Name == oldSheetName) as Sheet;
            if (sheet == null)
            {
                return false;
            }
            sheet.Name = newSheetName;
            return true;
        }

        /// <summary>
        /// Renames an existing sheet in the OpenXMLOffice.
        /// </summary>
        public bool RenameSheet(int sheetId, string newSheetName)
        {
            if (CheckIfSheetNameExist(newSheetName))
            {
                throw new ArgumentException("New Sheet with name already exist.");
            }
            Sheet? sheet = GetSheets().FirstOrDefault(sheet => (sheet as Sheet)?.Id?.Value == sheetId.ToString()) as Sheet;
            if (sheet == null)
            {
                return false;
            }
            sheet.Name = newSheetName;
            return true;
        }

        /// <summary>
        /// Save the active file with all new updates
        /// </summary>
        public void Save()
        {
            UpdateStyle();
            UpdateSharedString();
            spreadsheetDocument.Save();
            spreadsheetDocument.Dispose();
        }

        /// <summary>
        /// Save Copy of the content that updated to the source file
        /// </summary>
        public void SaveAs(string filePath)
        {
            throw new NotImplementedException();
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Check if sheet name exist in the sheets list
        /// </summary>
        private bool CheckIfSheetNameExist(string sheetName)
        {
            Sheet? sheet = GetSheets().FirstOrDefault(sheet => (sheet as Sheet)?.Name == sheetName) as Sheet;
            return sheet != null;
        }

        /// <summary>
        /// Return the current max ID from available sheets
        /// </summary>
        /// <returns>
        /// </returns>
        private UInt32Value GetMaxSheetId()
        {
            return GetSheets().Max(sheet => (sheet as Sheet)?.SheetId) ?? 0;
        }

        #endregion Private Methods
    }
}