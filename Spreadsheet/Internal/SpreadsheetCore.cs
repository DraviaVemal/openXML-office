// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLOffice.Global_2013;

namespace OpenXMLOffice.Spreadsheet_2013
{
	/// <summary>
	/// Spreadsheet Core class for initializing the Spreadsheet
	/// </summary>
	internal class SpreadsheetCore
	{

		internal readonly Excel excel;

		internal readonly SpreadsheetDocument spreadsheetDocument;

		internal readonly SpreadsheetInfo spreadsheetInfo = new();

		internal readonly SpreadsheetProperties spreadsheetProperties;

		internal SpreadsheetCore(Excel excel, string filePath, SpreadsheetProperties? spreadsheetProperties)
		{
			this.excel = excel;
			spreadsheetInfo.filePath = filePath;
			this.spreadsheetProperties = spreadsheetProperties ?? new();
			MemoryStream memoryStream = new();
			spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook, true);
			PrepareSpreadsheet(this.spreadsheetProperties);
		}

		internal SpreadsheetCore(Excel excel, string filePath, bool isEditable, SpreadsheetProperties? spreadsheetProperties = null)
		{
			this.excel = excel;
			spreadsheetInfo.filePath = filePath;
			this.spreadsheetProperties = spreadsheetProperties ?? new();
			FileStream reader = new(filePath, FileMode.Open);
			MemoryStream memoryStream = new();
			reader.CopyTo(memoryStream);
			reader.Close();
			spreadsheetDocument = SpreadsheetDocument.Open(memoryStream, isEditable, new OpenSettings()
			{
				AutoSave = true
			});
			if (isEditable)
			{
				spreadsheetInfo.isExistingFile = true;
				PrepareSpreadsheet(this.spreadsheetProperties);
			}
			else
			{
				spreadsheetInfo.isEditable = false;
			}
			ReadDataFromFile();
		}

		internal SpreadsheetCore(Excel excel, Stream stream, SpreadsheetProperties? spreadsheetProperties = null)
		{
			this.excel = excel;
			this.spreadsheetProperties = spreadsheetProperties ?? new();
			spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true);
			PrepareSpreadsheet(this.spreadsheetProperties);
		}

		internal SpreadsheetCore(Excel excel, Stream stream, bool isEditable, SpreadsheetProperties? spreadsheetProperties = null)
		{
			this.excel = excel;
			this.spreadsheetProperties = spreadsheetProperties ?? new();
			spreadsheetDocument = SpreadsheetDocument.Open(stream, isEditable, new OpenSettings()
			{
				AutoSave = true
			});
			if (isEditable)
			{
				spreadsheetInfo.isExistingFile = true;
				PrepareSpreadsheet(this.spreadsheetProperties);
			}
			else
			{
				spreadsheetInfo.isEditable = false;
			}
			ReadDataFromFile();
		}

		/// <summary>
		/// Read Data from exiting file
		/// </summary>
		internal void ReadDataFromFile()
		{
			LoadShareStringFromFileToCache();
			LoadStyleFromFileToCache();
		}

		/// <summary>
		/// Return the next relation id for the Spreadsheet
		/// </summary>
		internal string GetNextSpreadSheetRelationId()
		{
			return string.Format("rId{0}", GetWorkbookPart().Parts.Count() + 1);
		}

		/// <summary>
		/// Return the Shared String Table for the Spreadsheet
		/// </summary>
		internal SharedStringTable GetShareString()
		{
			SharedStringTablePart? sharedStringPart = GetWorkbookPart().GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
			if (sharedStringPart == null)
			{
				sharedStringPart = GetWorkbookPart().AddNewPart<SharedStringTablePart>();
				sharedStringPart.SharedStringTable = new SharedStringTable();
			}
			return sharedStringPart.SharedStringTable;
		}

		/// <summary>
		/// Return the Sheets for the Spreadsheet
		/// </summary>
		internal Sheets GetSheets()
		{
			Sheets? Sheets = GetWorkbookPart().Workbook.GetFirstChild<Sheets>();
			if (Sheets == null)
			{
				Sheets = new Sheets();
				GetWorkbookPart().Workbook.AppendChild(new Sheets());
			}
			return Sheets;
		}

		/// <summary>
		/// Return Woorkbook Part for the Spreadsheet
		/// </summary>
		internal WorkbookPart GetWorkbookPart()
		{
			if (spreadsheetDocument.WorkbookPart == null)
			{
				return spreadsheetDocument.AddWorkbookPart();
			}
			return spreadsheetDocument.WorkbookPart;
		}

		/// <summary>
		/// Load the Shared String to the Cache (aka in memeory database lightdb)
		/// </summary>
		internal void LoadShareStringFromFileToCache()
		{
			List<string> Records = new();
			GetShareString().Elements<SharedStringItem>().ToList().ForEach(rec =>
			{
				Text? text = rec.GetFirstChild<Text>();
				if (text != null)
				{
					Records.Add(text.Text);
				}
			});
			ShareString.Instance.InsertBulk(Records);
		}

		/// <summary>
		/// Load Exisiting Style from the Sheet
		/// </summary>
		internal void LoadStyleFromFileToCache()
		{
			Styles.Instance.LoadStyleFromSheet(GetWorkbookPart().WorkbookStylesPart!.Stylesheet);
		}

		/// <summary>
		/// Update the cache data into spreadsheet
		/// </summary>
		internal void WriteSharedStringToFile()
		{
			GetShareString().RemoveAllChildren<SharedStringItem>();
			ShareString.Instance.GetRecords().ForEach(Value =>
			{
				GetShareString().Append(new SharedStringItem(new Text(Value)));
			});
			GetShareString().Count = (uint)GetShareString().ChildElements.Count;
			GetShareString().UniqueCount = (uint)GetShareString().ChildElements.Count;
		}

		/// <summary>
		/// Load The DB Style Cache to Style Sheet
		/// </summary>
		internal void UpdateStyle()
		{
			Styles.Instance.SaveStyleProps(GetWorkbookPart().WorkbookStylesPart!.Stylesheet);
		}

		private void InitialiseStyle()
		{
			if (GetWorkbookPart().WorkbookStylesPart == null)
			{
				GetWorkbookPart().AddNewPart<WorkbookStylesPart>();
				GetWorkbookPart().WorkbookStylesPart!.Stylesheet = new();
			}
			else
			{
				GetWorkbookPart().WorkbookStylesPart!.Stylesheet ??= new();
			}
		}

		/// <summary>
		/// Common Spreadsheet perparation process used by all constructor
		/// </summary>
		private void PrepareSpreadsheet(SpreadsheetProperties SpreadsheetProperties)
		{
			GetWorkbookPart().Workbook ??= new Workbook();
			Sheets? sheets = GetWorkbookPart().Workbook.GetFirstChild<Sheets>();
			if (sheets == null)
			{
				sheets = new Sheets();
				GetWorkbookPart().Workbook.AppendChild(sheets);
			}
			if (GetWorkbookPart().ThemePart == null)
			{
				GetWorkbookPart().AddNewPart<ThemePart>(GetNextSpreadSheetRelationId());
			}
			Theme theme = new(SpreadsheetProperties?.theme);
			GetWorkbookPart().ThemePart!.Theme = theme.GetTheme();
			InitialiseStyle();
			GetWorkbookPart().Workbook.Save();
		}


	}
}
