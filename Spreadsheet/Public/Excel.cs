// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Reflection;
using System.IO;
using OpenXMLOffice.Global_2007;
using X = DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// This class serves as a versatile tool for working with Excel spreadsheets, built upon the
	/// foundation of the OpenXML SDK. This class offers a wide range of functionalities for
	/// handling Excel-related objects and operation It is designed to simplify tasks related to
	/// Excel file manipulation, including the creation of new Excel files, reading and updating
	/// existing files, and processing Excel data from stream
	/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
	/// </summary>
	public class Excel : PrivacyProperties
	{
		private readonly Spreadsheet spreadsheet;
		/// <summary>
		/// Create New file in the system
		/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
		/// </summary>
		public Excel(ExcelProperties spreadsheetProperties = null)
		{
			spreadsheet = new Spreadsheet(this, spreadsheetProperties);
		}
		/// <summary>
		/// Works with in memory object can be saved to file at later point.
		/// Source file will be cloned and released. hence can be replace by saveAs method if you want to update the same file.
		/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
		/// </summary>
		public Excel(string filePath, bool isEditable, ExcelProperties spreadsheetProperties = null, PrivacyProperties privacyProperties = null)
		{
			isFileEdited = true;
			spreadsheet = new Spreadsheet(this, filePath, isEditable, spreadsheetProperties);
		}
		/// <summary>
		/// Works with in memory object can be saved to file at later point.
		/// Source stream is copied and closed.
		/// Note : Make Clone in your source application if you want to retain the stream handle
		/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
		/// </summary>
		public Excel(Stream Stream, bool IsEditable, ExcelProperties spreadsheetProperties = null, PrivacyProperties privacyProperties = null)
		{
			isFileEdited = true;
			spreadsheet = new Spreadsheet(this, Stream, IsEditable, spreadsheetProperties);
		}
		/// <summary>
		/// Adds a new sheet to the OpenXMLOffice with the specified name. Throws an exception if
		/// SheetName already exist.
		/// </summary>
		public Worksheet AddSheet(string sheetName = null)
		{
			return spreadsheet.AddSheet(sheetName);
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="sheetName">Name of the sheet that needs to be activated</param>
		public void SetActiveSheet(string sheetName)
		{
			uint sheetIndex = 0;
			foreach (X.Sheet workSheet in spreadsheet.GetSheets().Elements<X.Sheet>())
			{
				Worksheet sheet = GetWorksheet(workSheet.Name);
				if (workSheet.Name == sheetName)
				{
					sheet.SetActiveSheet(true);
					X.WorkbookView workBookView = spreadsheet.GetBookViews().Elements<X.WorkbookView>().FirstOrDefault();
					if (workBookView == null)
					{
						workBookView = new X.WorkbookView();
						spreadsheet.GetBookViews().Append(workBookView);
					}
					workBookView.ActiveTab = sheetIndex;
				}
				else
				{
					sheet.SetActiveSheet(false);
				}
				++sheetIndex;
			}
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
		internal CalculationChainService GetCalculationChainService()
		{
			return spreadsheet.GetCalculationChainService();
		}
		internal StylesService GetStyleService()
		{
			return spreadsheet.GetStyleService();
		}
		/// <summary>
		/// Return the Sheet Name for the given Sheet ID
		/// </summary>
		public string GetSheetName(string sheetId)
		{
			return spreadsheet.GetSheetName(sheetId);
		}
		/// <summary>
		/// Retrieves a Worksheet object from an OpenXMLOffice, allowing manipulation of the worksheet.
		/// </summary>
		public Worksheet GetWorksheet(string sheetName)
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
		public bool RenameSheetById(string sheetId, string newSheetName)
		{
			return spreadsheet.RenameSheetById(sheetId, newSheetName);
		}
		/// <summary>
		/// Even on edit file OpenXML-Office Will clone the source and work on top of it to protect the integrity of source file.
		/// You can save the document at the end of lifecycle targeting the edit file to update or new file.
		/// This is supported for both file path and data stream
		/// </summary>
		public void SaveAs(string filePath)
		{
			SendAnonymousSaveStates(Assembly.GetExecutingAssembly().GetName());
			spreadsheet.SaveAs(filePath);
		}
		/// <summary>
		/// Even on edit file OpenXML-Office Will clone the source and work on top of it to protect the integrity of source file.
		/// You can save the document at the end of lifecycle targeting the edit file to update or new file.
		/// This is supported for both file path and data stream
		/// </summary>
		public void SaveAs(Stream stream)
		{
			SendAnonymousSaveStates(Assembly.GetExecutingAssembly().GetName());
			spreadsheet.SaveAs(stream);
		}
	}
}
