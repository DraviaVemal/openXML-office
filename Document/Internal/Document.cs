// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.IO;

namespace OpenXMLOffice.Document_2007
{
    internal class Document : DocumentCore
    {
        internal Document(Word word, WordProperties wordProperties) : base(word, wordProperties) { }
        internal Document(Word word, string filePath, bool isEditable, WordProperties wordProperties) : base(word, filePath, isEditable, wordProperties) { }
        internal Document(Word word, Stream stream, bool isEditable, WordProperties wordProperties) : base(word, stream, isEditable, wordProperties) { }
        private void SaveAs()
        {
            wordDocument.Save();
        }
        internal void SaveAs(string filePath)
        {
            SaveAs();
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                documentStream.WriteTo(fileStream);
            }
        }
        internal void SaveAs(Stream stream)
        {
            SaveAs();
            documentStream.CopyTo(stream);
        }
    }
}