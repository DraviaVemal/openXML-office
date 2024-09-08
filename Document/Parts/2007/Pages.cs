// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLOffice.Document_2007
{
    /// <summary>
    /// This class keeps the page content organized for insert, edit and delete
    /// </summary>
    public class Pages
    {
        private readonly MainDocumentPart mainDocumentPart;

        internal Pages(MainDocumentPart mainDocumentPart)
        {
            this.mainDocumentPart = mainDocumentPart;
        }
        /// <summary>
        /// Append New Paragraph to the end of the document
        /// </summary>
        public Paragraph AppendParagraph(ParagraphSetting paragraphSetting)
        {
            return new Paragraph(paragraphSetting);
        }
    }
}