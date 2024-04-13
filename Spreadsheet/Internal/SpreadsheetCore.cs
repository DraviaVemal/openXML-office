// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLOffice.Global_2007;

namespace OpenXMLOffice.Spreadsheet_2007
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

		private readonly StylesService stylesService = new();

		private readonly ShareStringService shareStringService = new();

		internal SpreadsheetCore(Excel excel, SpreadsheetProperties? spreadsheetProperties = null)
		{
			this.excel = excel;
			this.spreadsheetProperties = spreadsheetProperties ?? new();
			MemoryStream memoryStream = new();
			spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook, true);
			InitialiseSpreadsheet(this.spreadsheetProperties);
		}

		internal SpreadsheetCore(Excel excel, string filePath, bool isEditable, SpreadsheetProperties? spreadsheetProperties = null)
		{
			this.excel = excel;
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
				InitialiseSpreadsheet(this.spreadsheetProperties);
			}
			else
			{
				spreadsheetInfo.isEditable = false;
			}
			ReadDataFromFile();
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
				InitialiseSpreadsheet(this.spreadsheetProperties);
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
		internal SharedStringTable GetExcelShareString()
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
			GetExcelShareString().Elements<SharedStringItem>().ToList().ForEach(rec =>
			{
				Text? text = rec.GetFirstChild<Text>();
				if (text != null)
				{
					Records.Add(text.Text);
				}
			});
			GetShareStringService().InsertBulk(Records);
		}

		/// <summary>
		/// Load Exisiting Style from the Sheet
		/// </summary>
		internal void LoadStyleFromFileToCache()
		{
			GetStyleService().LoadStyleFromSheet(GetWorkbookPart().WorkbookStylesPart!.Stylesheet);
		}

		/// <summary>
		/// Update the cache data into spreadsheet
		/// </summary>
		internal void WriteSharedStringToFile()
		{
			GetExcelShareString().RemoveAllChildren<SharedStringItem>();
			GetShareStringService().GetRecords().ForEach(Value =>
			{
				GetExcelShareString().Append(new SharedStringItem(new Text(Value)));
			});
			GetExcelShareString().Count = (uint)GetExcelShareString().ChildElements.Count;
			GetExcelShareString().UniqueCount = (uint)GetExcelShareString().ChildElements.Count;
		}

		/// <summary>
		/// Load The DB Style Cache to Style Sheet
		/// </summary>
		internal void UpdateStyle()
		{
			GetStyleService().SaveStyleProps(GetWorkbookPart().WorkbookStylesPart!.Stylesheet);
		}

		internal StylesService GetStyleService()
		{
			return stylesService;
		}

		internal ShareStringService GetShareStringService()
		{
			return shareStringService;
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
		private void InitialiseSpreadsheet(SpreadsheetProperties SpreadsheetProperties)
		{
			if (spreadsheetDocument.CoreFilePropertiesPart == null)
			{
				spreadsheetDocument.AddCoreFilePropertiesPart();
				XMLHelper.AddOrUpdateCoreProperties(spreadsheetDocument.CoreFilePropertiesPart!.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite));
			}
			if (spreadsheetDocument.CustomFilePropertiesPart == null)
			{
				spreadsheetDocument.AddCustomFilePropertiesPart();
				XMLHelper.AddOrUpdateOpenXMLProperties(spreadsheetDocument.CustomFilePropertiesPart!.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite));
			}
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
