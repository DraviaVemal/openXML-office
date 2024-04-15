// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace OpenXMLOffice.Spreadsheet_2007
{
	internal class Spreadsheet : SpreadsheetCore
	{
		internal Spreadsheet(Excel excel, SpreadsheetProperties spreadsheetProperties) : base(excel, spreadsheetProperties) { }
		internal Spreadsheet(Excel excel, string filePath, bool isEditable, SpreadsheetProperties spreadsheetProperties) : base(excel, filePath, isEditable, spreadsheetProperties) { }
		internal Spreadsheet(Excel excel, Stream stream, bool isEditable, SpreadsheetProperties spreadsheetProperties) : base(excel, stream, isEditable, spreadsheetProperties) { }
		internal Worksheet AddSheet(string sheetName)
		{
			sheetName = sheetName ?? string.Format("Sheet{0}", GetMaxSheetId() + 1);
			if (CheckIfSheetNameExist(sheetName))
			{
				throw new ArgumentException("Sheet with name already exist.");
			}
			// Check If Sheet Already exist
			WorksheetPart worksheetPart = GetWorkbookPart().AddNewPart<WorksheetPart>();
			Sheet sheet = new Sheet()
			{
				Id = GetWorkbookPart().GetIdOfPart(worksheetPart),
				SheetId = GetMaxSheetId() + 1,
				Name = sheetName
			};
			GetSheets().Append(sheet);
			worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new SheetData());
			return new Worksheet(excel, worksheetPart.Worksheet, sheet);
		}
		internal int? GetSheetId(string sheetName)
		{
			Sheet sheet = GetSheets().FirstOrDefault(tempSheet => (tempSheet as Sheet).Name == sheetName) as Sheet;
			if (sheet != null)
			{
				return int.Parse(sheet.Id.Value);
			}
			return null;
		}
		internal uint GetStyleId(CellStyleSetting CellStyleSetting)
		{
			return GetStyleService().GetCellStyleId(CellStyleSetting);
		}
		internal string GetSheetName(string sheetId)
		{
			Sheet sheet = GetSheets().FirstOrDefault(tempSheet => (tempSheet as Sheet).Id.Value == sheetId) as Sheet;
			if (sheet != null)
			{
				return sheet.Name;
			}
			return null;
		}
		internal Worksheet GetWorksheet(string sheetName)
		{
			Sheets sheets = GetSheets();
			Sheet sheet = sheets.FirstOrDefault(tempSheet => (tempSheet as Sheet).Name == sheetName) as Sheet;
			if (sheet == null) { return null; }
			WorkbookPart workbookPart = GetWorkbookPart();
			WorksheetPart sheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
			if (sheetPart == null) { return null; }
			return new Worksheet(excel, sheetPart.Worksheet, sheet);
		}
		internal bool RemoveSheet(string sheetName)
		{
			Sheet sheet = GetSheets().FirstOrDefault(tempSheet => (tempSheet as Sheet).Name == sheetName) as Sheet;
			if (sheet != null)
			{
				WorksheetPart sheetPart = GetWorkbookPart().GetPartById(sheet.Id) as WorksheetPart;
				if (sheetPart != null)
				{
					GetWorkbookPart().DeletePart(sheetPart);
				}
				sheet.Remove();
				return true;
			}
			return false;
		}
		internal bool RemoveSheet(int sheetId)
		{
			Sheet sheet = GetSheets().FirstOrDefault(tempSheet => (tempSheet as Sheet).Id.Value == sheetId.ToString()) as Sheet;
			if (sheet != null)
			{
				OpenXmlPart sheetPart = GetWorkbookPart().GetPartById(sheet.Id);
				if (sheetPart != null && sheetPart is WorksheetPart)
				{
					GetWorkbookPart().DeletePart(sheetPart);
				}
				sheet.Remove();
				return true;
			}
			return false;
		}
		internal bool RenameSheet(string oldSheetName, string newSheetName)
		{
			if (CheckIfSheetNameExist(newSheetName))
			{
				throw new ArgumentException("New Sheet with name already exist.");
			}
			Sheet sheet = GetSheets().FirstOrDefault(tempSheet => (tempSheet as Sheet).Name == oldSheetName) as Sheet;
			if (sheet == null)
			{
				return false;
			}
			sheet.Name = newSheetName;
			return true;
		}
		internal bool RenameSheet(int sheetId, string newSheetName)
		{
			if (CheckIfSheetNameExist(newSheetName))
			{
				throw new ArgumentException("New Sheet with name already exist.");
			}
			Sheet sheet = GetSheets().FirstOrDefault(tempSheet => (tempSheet as Sheet).Id.Value == sheetId.ToString()) as Sheet;
			if (sheet == null)
			{
				return false;
			}
			sheet.Name = newSheetName;
			return true;
		}
		internal void SaveAs(string filePath)
		{
			UpdateStyle();
			WriteSharedStringToFile();
			spreadsheetDocument.Save();
			spreadsheetDocument.Clone(filePath).Dispose();
		}
		internal void SaveAs(Stream stream)
		{
			UpdateStyle();
			WriteSharedStringToFile();
			spreadsheetDocument.Save();
			spreadsheetDocument.Clone(stream).Dispose();
		}
		private bool CheckIfSheetNameExist(string sheetName)
		{
			Sheet sheet = GetSheets().FirstOrDefault(tempSheet => (tempSheet as Sheet).Name == sheetName) as Sheet;
			return sheet != null;
		}
		private UInt32Value GetMaxSheetId()
		{
			return GetSheets().Max(sheet => (sheet as Sheet).SheetId) ?? 0;
		}
	}
}
