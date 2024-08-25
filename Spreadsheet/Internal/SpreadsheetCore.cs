// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using G = OpenXMLOffice.Global_2007;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// Spreadsheet Core class for initializing the Spreadsheet
	/// </summary>
	internal class SpreadsheetCore
	{
		internal readonly Excel excel;
		internal readonly SpreadsheetDocument spreadsheetDocument;
		internal readonly ExcelInfo spreadsheetInfo = new ExcelInfo();
		internal readonly ExcelProperties spreadsheetProperties;
		private readonly StylesService stylesService = new StylesService();
		private readonly ShareStringService shareStringService = new ShareStringService();
		private readonly CalculationChainService calculationChainService = new CalculationChainService();
		internal SpreadsheetCore(Excel excel, ExcelProperties spreadsheetProperties = null)
		{
			this.excel = excel;
			this.spreadsheetProperties = spreadsheetProperties ?? new ExcelProperties();
			MemoryStream memoryStream = new MemoryStream();
			spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook, true);
			InitializeSpreadsheet(this.spreadsheetProperties);
			ReadDataFromFile();
		}
		internal SpreadsheetCore(Excel excel, string filePath, bool isEditable, ExcelProperties spreadsheetProperties = null)
		{
			this.excel = excel;
			this.spreadsheetProperties = spreadsheetProperties ?? new ExcelProperties();
			FileStream reader = new FileStream(filePath, FileMode.Open);
			MemoryStream memoryStream = new MemoryStream();
			reader.CopyTo(memoryStream);
			reader.Close();
			spreadsheetDocument = SpreadsheetDocument.Open(memoryStream, isEditable, new OpenSettings()
			{
				AutoSave = true
			});
			if (isEditable)
			{
				spreadsheetInfo.isExistingFile = true;
				InitializeSpreadsheet(this.spreadsheetProperties);
			}
			else
			{
				spreadsheetInfo.isEditable = false;
			}
			ReadDataFromFile();
		}
		internal SpreadsheetCore(Excel excel, Stream stream, bool isEditable, ExcelProperties spreadsheetProperties = null)
		{
			this.excel = excel;
			this.spreadsheetProperties = spreadsheetProperties ?? new ExcelProperties();
			MemoryStream memoryStream = new MemoryStream();
			stream.CopyTo(memoryStream);
			stream.Dispose();
			spreadsheetDocument = SpreadsheetDocument.Open(memoryStream, isEditable, new OpenSettings()
			{
				AutoSave = true
			});
			if (isEditable)
			{
				spreadsheetInfo.isExistingFile = true;
				InitializeSpreadsheet(this.spreadsheetProperties);
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
			LoadCalculationChainFromFileToCache();
			LoadShareStringFromFileToCache();
			LoadStyleFromFileToCache();
		}
		/// <summary>
		/// Return the next relation id for the Spreadsheet
		/// </summary>
		internal string GetNextSpreadSheetRelationId()
		{
			int nextId = GetWorkbookPart().Parts.Count() + GetWorkbookPart().ExternalRelationships.Count() + GetWorkbookPart().HyperlinkRelationships.Count() + GetWorkbookPart().DataPartReferenceRelationships.Count();
			do
			{
				++nextId;
			} while (GetWorkbookPart().Parts.Any(item => item.RelationshipId == string.Format("rId{0}", nextId)) ||
			GetWorkbookPart().ExternalRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)) ||
			GetWorkbookPart().HyperlinkRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)) ||
			GetWorkbookPart().DataPartReferenceRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)));
			return string.Format("rId{0}", nextId);
		}
		/// <summary>
		/// Return the Shared String Table for the Spreadsheet
		/// </summary>
		internal SharedStringTablePart GetExcelShareStringPart()
		{
			SharedStringTablePart sharedStringPart = GetWorkbookPart().GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
			if (sharedStringPart == null)
			{
				sharedStringPart = GetWorkbookPart().AddNewPart<SharedStringTablePart>(GetNextSpreadSheetRelationId());
				sharedStringPart.SharedStringTable = new SharedStringTable();
			}
			return sharedStringPart;
		}
		/// <summary>
		/// Return the Calculation Chain for the Spreadsheet
		/// </summary>
		internal CalculationChainPart GetExcelCalculationChainPart()
		{
			CalculationChainPart calculationChainPart = GetWorkbookPart().GetPartsOfType<CalculationChainPart>().FirstOrDefault();
			if (calculationChainPart == null)
			{
				calculationChainPart = GetWorkbookPart().AddNewPart<CalculationChainPart>(GetNextSpreadSheetRelationId());
				calculationChainPart.CalculationChain = new CalculationChain();
			}
			return calculationChainPart;
		}
		/// <summary>
		/// Return the Shared String Table for the Spreadsheet
		/// </summary>
		internal SharedStringTable GetExcelShareString()
		{
			return GetExcelShareStringPart().SharedStringTable;
		}
		/// <summary>
		/// Return the Calculation Chain for the Spreadsheet
		/// </summary>
		internal CalculationChain GetExcelCalculationChain()
		{
			return GetExcelCalculationChainPart().CalculationChain;
		}

		/// <summary>
		/// Remove Share string part if not required
		/// </summary>
		internal void DeleteShareStringPart()
		{
			GetWorkbookPart().DeletePart(GetExcelShareStringPart());
		}

		/// <summary>
		/// Remove calculation part if not required
		/// </summary>
		internal void DeleteCalculationPart()
		{
			GetWorkbookPart().DeletePart(GetExcelCalculationChainPart());
		}
		/// <summary>
		/// Return the Sheets for the Spreadsheet
		/// </summary>
		internal Sheets GetSheets()
		{
			Sheets Sheets = GetWorkbookPart().Workbook.GetFirstChild<Sheets>();
			if (Sheets == null)
			{
				Sheets = new Sheets();
				GetWorkbookPart().Workbook.AppendChild(new Sheets());
			}
			return Sheets;
		}
		/// <summary>
		/// Return Workbook Part for the Spreadsheet
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
		/// Load the Shared String to the Cache (aka in memory database lightdb)
		/// </summary>
		internal void LoadCalculationChainFromFileToCache()
		{
			List<CalculationRecord> Records = new List<CalculationRecord>();
			GetExcelCalculationChain().Elements<CalculationCell>().ToList().ForEach(rec =>
			{
				Records.Add(new CalculationRecord(rec.CellReference, rec.SheetId));
			});
			GetCalculationChainService().InsertBulk(Records);
		}
		/// <summary>
		/// Load the Shared String to the Cache (aka in memory database lightdb)
		/// </summary>
		internal void LoadShareStringFromFileToCache()
		{
			List<string> Records = new List<string>();
			GetExcelShareString().Elements<SharedStringItem>().ToList().ForEach(rec =>
			{
				Text text = rec.GetFirstChild<Text>();
				if (text != null)
				{
					Records.Add(text.Text);
				}
			});
			GetShareStringService().InsertBulk(Records);
		}
		/// <summary>
		/// Load Existing Style from the Sheet
		/// </summary>
		internal void LoadStyleFromFileToCache()
		{
			GetStyleService().LoadStyleFromSheet(GetWorkbookPart().WorkbookStylesPart.Stylesheet);
		}
		/// <summary>
		/// Update the cache data into spreadsheet
		/// </summary>
		internal void WriteSharedStringToFile()
		{
			GetExcelShareString().RemoveAllChildren<SharedStringItem>();
			if (GetShareStringService().GetRecords().Count > 0)
			{
				GetShareStringService().GetRecords().ForEach(Value =>
				{
					GetExcelShareString().Append(new SharedStringItem(new Text(Value)));
				});
				GetExcelShareString().Count = (uint)GetExcelShareString().ChildElements.Count;
				GetExcelShareString().UniqueCount = (uint)GetExcelShareString().ChildElements.Count;
			}
			else
			{
				DeleteShareStringPart();
			}
		}
		/// <summary>
		/// Update the calculation chain into spreadsheet
		/// </summary>
		internal void WriteCalculationChainToFile()
		{
			GetExcelCalculationChain().RemoveAllChildren<CalculationCell>();
			if (GetCalculationChainService().GetRecords().Count > 0)
			{
				GetCalculationChainService().GetRecords().ForEach(Value =>
				{
					GetExcelCalculationChain().Append(new CalculationCell()
					{
						CellReference = Value.CellId,
						SheetId = Value.SheetIndex
					});
				});
			}
			else
			{
				DeleteCalculationPart();
			}
		}
		/// <summary>
		/// Load The DB Style Cache to Style Sheet
		/// </summary>
		internal void UpdateStyle()
		{
			GetStyleService().SaveStyleProps(GetWorkbookPart().WorkbookStylesPart.Stylesheet);
		}
		internal StylesService GetStyleService()
		{
			return stylesService;
		}
		internal ShareStringService GetShareStringService()
		{
			return shareStringService;
		}
		internal CalculationChainService GetCalculationChainService()
		{
			return calculationChainService;
		}
		private void InitializeStyle()
		{
			if (GetWorkbookPart().WorkbookStylesPart == null)
			{
				GetWorkbookPart().AddNewPart<WorkbookStylesPart>(GetNextSpreadSheetRelationId());
				GetWorkbookPart().WorkbookStylesPart.Stylesheet = new Stylesheet();
			}
			else
			{
				if (GetWorkbookPart().WorkbookStylesPart.Stylesheet == null)
				{
					GetWorkbookPart().WorkbookStylesPart.Stylesheet = new Stylesheet();
				}
			}
		}
		/// <summary>
		/// Common Spreadsheet preparation process used by all constructor
		/// </summary>
		private void InitializeSpreadsheet(ExcelProperties excelProperties)
		{
			if (spreadsheetDocument.CoreFilePropertiesPart == null)
			{
				spreadsheetDocument.AddCoreFilePropertiesPart();
				using (Stream stream = spreadsheetDocument.CoreFilePropertiesPart.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite))
				{
					G.CoreProperties.AddCoreProperties(stream, excelProperties.coreProperties);
				}
			}
			else
			{
				using (Stream stream = spreadsheetDocument.CoreFilePropertiesPart.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite))
				{
					G.CoreProperties.UpdateModifiedDetails(stream, excelProperties.coreProperties);
				}
			}
			if (GetWorkbookPart().Workbook == null)
			{
				GetWorkbookPart().Workbook = new Workbook();
			}
			Sheets sheets = GetWorkbookPart().Workbook.GetFirstChild<Sheets>();
			if (sheets == null)
			{
				sheets = new Sheets();
				GetWorkbookPart().Workbook.AppendChild(sheets);
			}
			if (GetWorkbookPart().ThemePart == null)
			{
				GetWorkbookPart().AddNewPart<ThemePart>(GetNextSpreadSheetRelationId());
			}
			if (GetWorkbookPart().ThemePart.Theme == null)
			{
				G.Theme theme = new G.Theme(excelProperties.theme);
				GetWorkbookPart().ThemePart.Theme = theme.GetTheme();
			}
			InitializeStyle();
			GetWorkbookPart().Workbook.Save();
		}
	}
}
