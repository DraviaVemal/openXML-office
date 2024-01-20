// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLOffice.Global;

namespace OpenXMLOffice.Excel
{
    /// <summary>
    /// Spreadsheet Core class for initializing the Spreadsheet
    /// </summary>
    public class SpreadsheetCore
    {
        #region Protected Fields

        /// <summary>
        /// Maintain the master OpenXML Spreadsheet document
        /// </summary>
        protected readonly SpreadsheetDocument spreadsheetDocument;

        /// <summary>
        /// Maintain the Spreadsheet Properties
        /// </summary>
        protected readonly SpreadsheetProperties spreadsheetProperties;

        #endregion Protected Fields

        #region Protected Constructors

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class
        /// </summary>
        protected SpreadsheetCore(string filePath)
        {
            spreadsheetProperties = new();
            spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook, true);
            PrepareSpreadsheet(spreadsheetProperties);
            InitialiseStyle();
            LoadStyle();
        }

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class,
        /// allowing you to work with Excel spreadsheet It accepts a Existing excel file path and a
        /// SpreadsheetDocumentType enumeration value as parameters and creates a corresponding
        /// SpreadsheetDocument. This is also used to update as template.
        /// </summary>
        protected SpreadsheetCore(string filePath, SpreadsheetProperties spreadsheetProperties)
        {
            this.spreadsheetProperties = spreadsheetProperties;
            spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook, true);
            PrepareSpreadsheet(this.spreadsheetProperties);
            InitialiseStyle();
            LoadStyle();
        }

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class
        /// </summary>
        protected SpreadsheetCore(string filePath, bool isEditable)
        {
            spreadsheetProperties = new();
            spreadsheetDocument = SpreadsheetDocument.Open(filePath, isEditable, new OpenSettings
            {
                AutoSave = true
            });
            PrepareSpreadsheet(spreadsheetProperties);
            InitialiseStyle();
            LoadStyle();
        }

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class.
        /// </summary>
        protected SpreadsheetCore(string filePath, bool isEditable, SpreadsheetProperties spreadsheetProperties)
        {
            this.spreadsheetProperties = spreadsheetProperties;
            spreadsheetDocument = SpreadsheetDocument.Open(filePath, isEditable, new OpenSettings
            {
                AutoSave = true
            });
            PrepareSpreadsheet(this.spreadsheetProperties);
            InitialiseStyle();
            LoadStyle();
        }

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class,
        /// allowing you to work with Excel spreadsheet It accepts a Stream object and a
        /// SpreadsheetDocumentType enumeration value as parameters and creates a corresponding SpreadsheetDocument.
        /// </summary>
        protected SpreadsheetCore(Stream stream)
        {
            spreadsheetProperties = new();
            spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true);
            PrepareSpreadsheet(spreadsheetProperties);
            InitialiseStyle();
            LoadStyle();
        }

        /// <summary>
        /// This public constructor method initializes a new instance of the Spreadsheet class,
        /// allowing you to work with Excel spreadsheet It accepts a Stream object and a
        /// SpreadsheetDocumentType enumeration value as parameters and creates a corresponding SpreadsheetDocument.
        /// </summary>
        protected SpreadsheetCore(Stream stream, SpreadsheetProperties spreadsheetProperties)
        {
            this.spreadsheetProperties = spreadsheetProperties;
            spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true);
            PrepareSpreadsheet(this.spreadsheetProperties);
            InitialiseStyle();
            LoadStyle();
        }

        #endregion Protected Constructors

        #region Protected Methods

        /// <summary>
        /// Return the next relation id for the Spreadsheet
        /// </summary>
        /// <returns>
        /// </returns>
        protected string GetNextSpreadSheetRelationId()
        {
            return string.Format("rId{0}", GetWorkbookPart().Parts.Count() + 1);
        }

        /// <summary>
        /// Return the Shared String Table for the Spreadsheet
        /// </summary>
        protected SharedStringTable GetShareString()
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
        protected Sheets GetSheets()
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
        protected WorkbookPart GetWorkbookPart()
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
        protected void LoadShareStringToCache()
        {
            List<string> Records = new();
            GetShareString().ChildElements.ToList().ForEach(rec =>
            {
                // TODO : File Open Implementation
                //Records.Add("");
            });
            ShareString.Instance.InsertBulk(Records);
        }

        /// <summary>
        /// Load Exisiting Style from the Sheet
        /// </summary>
        protected void LoadStyle()
        {
            Styles.Instance.LoadStyleFromSheet(GetWorkbookPart().WorkbookStylesPart!.Stylesheet);
        }

        /// <summary>
        /// Update the cache data into spreadsheet
        /// </summary>
        protected void UpdateSharedString()
        {
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
        protected void UpdateStyle()
        {
            Styles.Instance.SaveStyleProps(GetWorkbookPart().WorkbookStylesPart!.Stylesheet);
        }

        #endregion Protected Methods

        #region Private Methods

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
            Sheets sheets = GetWorkbookPart().Workbook.GetFirstChild<Sheets>() ?? new Sheets();
            GetWorkbookPart().Workbook.AppendChild(sheets);
            if (GetWorkbookPart().ThemePart == null)
            {
                GetWorkbookPart().AddNewPart<ThemePart>(GetNextSpreadSheetRelationId());
            }
            Theme theme = new(SpreadsheetProperties?.theme);
            GetWorkbookPart().ThemePart!.Theme = theme.GetTheme();
            LoadShareStringToCache();
            GetWorkbookPart().Workbook.Save();
        }

        #endregion Private Methods
    }
}