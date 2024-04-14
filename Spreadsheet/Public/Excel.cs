// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// This class serves as a versatile tool for working with Excel spreadsheets, built upon the
	/// foundation of the OpenXML SDK. This class offers a wide range of functionalities for
	/// handling Excel-related objects and operation It is designed to simplify tasks related to
	/// Excel file manipulation, including the creation of new Excel files, reading and updating
	/// existing files, and processing Excel data from stream
	/// </summary>
	public class Excel
	{
		private readonly Spreadsheet spreadsheet;
		/// <summary>
		/// Create New file in the system
		/// </summary>
		public Excel(SpreadsheetProperties? spreadsheetProperties = null)
		{
			spreadsheet = new(this, spreadsheetProperties);
		}
		/// <summary>
		/// Open and work with existing file
		/// </summary>
		public Excel(string filePath, bool isEditable, SpreadsheetProperties? spreadsheetProperties = null)
		{
			spreadsheet = new(this, filePath, isEditable, spreadsheetProperties);
		}
		/// <summary>
		/// Works with in memory object can be saved to file at later point
		/// </summary>
		public Excel(Stream Stream, bool IsEditable, SpreadsheetProperties? spreadsheetProperties = null)
		{
			spreadsheet = new(this, Stream, IsEditable, spreadsheetProperties);
		}
		/// <summary>
		/// Adds a new sheet to the OpenXMLOffice with the specified name. Throws an exception if
		/// SheetName already exist.
		/// </summary>
		public Worksheet AddSheet(string? sheetName = null)
		{
			return spreadsheet.AddSheet(sheetName);
		}
		/// <summary>
		/// Returns the Sheet ID for the give Sheet Name
		/// </summary>
		public int? GetSheetId(string sheetName)
		{
			return spreadsheet.GetSheetId(sheetName);
		}
		/// <summary>
		/// Use this method to create a new style and get the style id
		/// Use of Style Id instead of Style Setting directly in Worksheet Cell is highly recommended for performance
		/// </summary>
		public uint GetStyleId(CellStyleSetting CellStyleSetting)
		{
			return spreadsheet.GetStyleId(CellStyleSetting);
		}
		internal ShareStringService GetShareStringService()
		{
			return spreadsheet.GetShareStringService();
		}
		internal StylesService GetStyleService()
		{
			return spreadsheet.GetStyleService();
		}
		/// <summary>
		/// Return the Sheet Name for the given Sheet ID
		/// </summary>
		public string? GetSheetName(string sheetId)
		{
			return spreadsheet.GetSheetName(sheetId);
		}
		/// <summary>
		/// Retrieves a Worksheet object from an OpenXMLOffice, allowing manipulation of the worksheet.
		/// </summary>
		public Worksheet? GetWorksheet(string sheetName)
		{
			return spreadsheet.GetWorksheet(sheetName);
		}
		/// <summary>
		/// Removes a sheet with the specified name from the OpenXMLOffice
		/// </summary>
		public bool RemoveSheet(string sheetName)
		{
			return spreadsheet.RemoveSheet(sheetName);
		}
		/// <summary>
		/// Removes a sheet with the specified ID from the OpenXMLOffice
		/// </summary>
		public bool RemoveSheet(int sheetId)
		{
			return spreadsheet.RemoveSheet(sheetId);
		}
		/// <summary>
		/// Creates a new sheet with the specified name and adds its relevant components to the
		/// workbook. Throws an exception if the sheet name is already in use.
		/// </summary>
		public bool RenameSheet(string oldSheetName, string newSheetName)
		{
			return spreadsheet.RenameSheet(oldSheetName, newSheetName);
		}
		/// <summary>
		/// Renames an existing sheet in the OpenXMLOffice.
		/// </summary>
		public bool RenameSheet(int sheetId, string newSheetName)
		{
			return spreadsheet.RenameSheet(sheetId, newSheetName);
		}
		/// <summary>
		/// Save Copy of the content that updated to the source file
		/// </summary>
		public void SaveAs(string filePath)
		{
			spreadsheet.SaveAs(filePath);
		}
		/// <summary>
		/// Save Copy of the content that updated to the source file
		/// </summary>
		public void SaveAs(Stream stream)
		{
			spreadsheet.SaveAs(stream);
		}
	}
}
