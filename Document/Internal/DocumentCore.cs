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
        internal readonly WordprocessingDocument wordprocessingDocument;
        internal readonly WordInfo wordInfo = new WordInfo();
        internal readonly WordProperties wordProperties;
        internal MemoryStream documentStream = new MemoryStream();
        internal DocumentCore(Word word, WordProperties wordProperties = null)
        {
            this.word = word;
            this.wordProperties = wordProperties ?? new WordProperties();
            wordprocessingDocument = WordprocessingDocument.Create(documentStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true);
            InitializeDocument(this.wordProperties);
        }
        internal DocumentCore(Word word, string filePath, bool isEditable, WordProperties wordProperties = null)
        {
            this.word = word;
            this.wordProperties = wordProperties ?? new WordProperties();
            FileStream reader = new FileStream(filePath, FileMode.Open);
            reader.CopyTo(documentStream);
            reader.Close();
            wordprocessingDocument = WordprocessingDocument.Open(documentStream, isEditable, new OpenSettings()
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
            wordprocessingDocument = WordprocessingDocument.Open(documentStream, isEditable, new OpenSettings()
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
        protected W.Document GetDocument()
        {
            return GetMainDocumentPart().Document;
        }
        internal MainDocumentPart GetMainDocumentPart()
        {
            return wordprocessingDocument.MainDocumentPart;
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
            if (wordprocessingDocument.CoreFilePropertiesPart == null)
            {
                wordprocessingDocument.AddCoreFilePropertiesPart();
                using (Stream stream = wordprocessingDocument.CoreFilePropertiesPart.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    G.CoreProperties.AddCoreProperties(stream, wordProperties.coreProperties);
                }
            }
            else
            {
                using (Stream stream = wordprocessingDocument.CoreFilePropertiesPart.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    G.CoreProperties.UpdateModifiedDetails(stream, wordProperties.coreProperties);
                }
            }
            if (GetMainDocumentPart() == null)
            {
                wordprocessingDocument.AddMainDocumentPart();
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
        }
    }
}