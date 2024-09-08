// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using G = OpenXMLOffice.Global_2007;

namespace OpenXMLOffice.Document_2007
{
    /// <summary>
    /// Spreadsheet Core class for initializing the Spreadsheet
    /// </summary>
    internal class DocumentCore
    {
        internal readonly Word word;
        internal readonly WordprocessingDocument wordDocument;
        internal readonly WordInfo wordInfo = new WordInfo();
        internal readonly WordProperties wordProperties;
        internal MemoryStream documentStream = new MemoryStream();
        internal DocumentCore(Word word, WordProperties wordProperties = null)
        {
            this.word = word;
            this.wordProperties = wordProperties ?? new WordProperties();
            wordDocument = WordprocessingDocument.Create(documentStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true);
            InitializeDocument(this.wordProperties);
        }
        internal DocumentCore(Word word, string filePath, bool isEditable, WordProperties wordProperties = null)
        {
            this.word = word;
            this.wordProperties = wordProperties ?? new WordProperties();
            FileStream reader = new FileStream(filePath, FileMode.Open);
            reader.CopyTo(documentStream);
            reader.Close();
            wordDocument = WordprocessingDocument.Open(documentStream, isEditable, new OpenSettings()
            {
                AutoSave = true
            });
            if (isEditable)
            {
                wordInfo.isExistingFile = true;
                InitializeDocument(this.wordProperties);
            }
            else
            {
                wordInfo.isEditable = false;
            }
        }
        internal DocumentCore(Word word, Stream stream, bool isEditable, WordProperties wordProperties = null)
        {
            this.word = word;
            this.wordProperties = wordProperties ?? new WordProperties();
            stream.CopyTo(documentStream);
            stream.Dispose();
            wordDocument = WordprocessingDocument.Open(documentStream, isEditable, new OpenSettings()
            {
                AutoSave = true
            });
            if (isEditable)
            {
                wordInfo.isExistingFile = true;
                InitializeDocument(this.wordProperties);
            }
            else
            {
                wordInfo.isEditable = false;
            }
        }

        internal FontTablePart GetFontTablePart()
        {
            if (wordDocument.MainDocumentPart.FontTablePart == null)
            {
                wordDocument.MainDocumentPart.AddNewPart<FontTablePart>(GetNextSpreadSheetRelationId());
                wordDocument.MainDocumentPart.FontTablePart.Fonts = new W.Fonts();
            }
            return wordDocument.MainDocumentPart.FontTablePart;
        }

        internal DocumentSettingsPart GetDocumentSettingPart()
        {
            if (wordDocument.MainDocumentPart.DocumentSettingsPart == null)
            {
                wordDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>(GetNextSpreadSheetRelationId());
                wordDocument.MainDocumentPart.DocumentSettingsPart.Settings = new W.Settings();
            }
            return wordDocument.MainDocumentPart.DocumentSettingsPart;
        }

        internal StyleDefinitionsPart GetStylesPart()
        {
            if (wordDocument.MainDocumentPart.StyleDefinitionsPart == null)
            {
                wordDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>(GetNextSpreadSheetRelationId());
                wordDocument.MainDocumentPart.StyleDefinitionsPart.Styles = new W.Styles();
            }
            return wordDocument.MainDocumentPart.StyleDefinitionsPart;
        }

        internal WebSettingsPart GetWebSettingsPart()
        {
            if (wordDocument.MainDocumentPart.WebSettingsPart == null)
            {
                wordDocument.MainDocumentPart.AddNewPart<WebSettingsPart>(GetNextSpreadSheetRelationId());
                wordDocument.MainDocumentPart.WebSettingsPart.WebSettings = new W.WebSettings();
            }
            return wordDocument.MainDocumentPart.WebSettingsPart;
        }

        /// <summary>
        /// Read file setting data into local file DB 
        /// </summary>
        internal void ReadDataFromFile()
        {
            LoadFontTable();
            LoadSettings();
            LoadStyles();
            LoadWebSetting();
        }
        /// <summary>
        /// 
        /// </summary>
        internal void LoadFontTable()
        {
            GetFontTablePart();
        }
        /// <summary>
        /// 
        /// </summary>
        internal void LoadSettings()
        {
            GetDocumentSettingPart();
        }
        /// <summary>
        /// 
        /// </summary>
        internal void LoadStyles()
        {
            GetStylesPart();
        }
        /// <summary>
        /// 
        /// </summary>
        internal void LoadWebSetting()
        {
            GetWebSettingsPart();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        protected W.Document GetDocument()
        {
            return GetMainDocumentPart().Document;
        }

        internal MainDocumentPart GetMainDocumentPart()
        {
            return wordDocument.MainDocumentPart;
        }
        /// <summary>
        /// Return the next relation id for the Spreadsheet
        /// </summary>
        internal string GetNextSpreadSheetRelationId()
        {
            int nextId = GetMainDocumentPart().Parts.Count() + GetMainDocumentPart().ExternalRelationships.Count() + GetMainDocumentPart().HyperlinkRelationships.Count() + GetMainDocumentPart().DataPartReferenceRelationships.Count();
            do
            {
                ++nextId;
            } while (GetMainDocumentPart().Parts.Any(item => item.RelationshipId == string.Format("rId{0}", nextId)) ||
            GetMainDocumentPart().ExternalRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)) ||
            GetMainDocumentPart().HyperlinkRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)) ||
            GetMainDocumentPart().DataPartReferenceRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)));
            return string.Format("rId{0}", nextId);
        }
        private void InitializeDocument(WordProperties wordProperties)
        {
            if (wordDocument.CoreFilePropertiesPart == null)
            {
                wordDocument.AddCoreFilePropertiesPart();
                using (Stream stream = wordDocument.CoreFilePropertiesPart.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    G.CoreProperties.AddCoreProperties(stream, wordProperties.coreProperties);
                }
            }
            else
            {
                using (Stream stream = wordDocument.CoreFilePropertiesPart.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    G.CoreProperties.UpdateModifiedDetails(stream, wordProperties.coreProperties);
                }
            }
            if (GetMainDocumentPart() == null)
            {
                wordDocument.AddMainDocumentPart();
            }
            if (GetMainDocumentPart().Document == null)
            {
                GetMainDocumentPart().Document = new W.Document();
                GetMainDocumentPart().Document.Save();
            }
            if (GetMainDocumentPart().ThemePart == null)
            {
                GetMainDocumentPart().AddNewPart<ThemePart>(GetNextSpreadSheetRelationId());
            }
            if (GetMainDocumentPart().ThemePart.Theme == null)
            {
                G.Theme theme = new G.Theme(wordProperties.theme);
                GetMainDocumentPart().ThemePart.Theme = theme.GetTheme();
            }
            ReadDataFromFile();
        }
    }
}