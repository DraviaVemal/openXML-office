// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.IO;
using System.Reflection;
using OpenXMLOffice.Global_2007;

namespace OpenXMLOffice.Document_2007
{
	/// <summary>
	/// Word Class
	/// </summary>
	public class Word : PrivacyProperties
	{
		private readonly Document document;
		/// <summary>
		/// Create New file in the system
		/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
		/// </summary>
		public Word(WordProperties wordProperties = null)
		{
			document = new Document(this, wordProperties);
		}

		/// <summary>
		/// Works with in memory object can be saved to file at later point.
		/// Source file will be cloned and released. hence can be replace by saveAs method if you want to update the same file.
		/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
		/// </summary>
		public Word(string filePath, bool isEditable, WordProperties wordProperties = null, PrivacyProperties privacyProperties = null)
		{
			isFileEdited = true;
			document = new Document(this, filePath, isEditable, wordProperties);
		}

		/// <summary>
		/// Works with in memory object can be saved to file at later point.
		/// Source stream is copied and closed.
		/// Note : Make Clone in your source application if you want to retain the stream handle
		/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
		/// </summary>
		public Word(Stream Stream, bool IsEditable, WordProperties wordProperties = null, PrivacyProperties privacyProperties = null)
		{
			isFileEdited = true;
			document = new Document(this, Stream, IsEditable, wordProperties);
		}

		/// <summary>
		/// Even on edit file OpenXML-Office Will clone the source and work on top of it to protect the integrity of source file.
		/// You can save the document at the end of lifecycle targeting the edit file to update or new file.
		/// This is supported for both file path and data stream
		/// </summary>
		public void SaveAs(string filePath)
		{
			SendAnonymousSaveStates(Assembly.GetExecutingAssembly().GetName());
			document.SaveAs(filePath);
		}

		/// <summary>
		/// Even on edit file OpenXML-Office Will clone the source and work on top of it to protect the integrity of source file.
		/// You can save the document at the end of lifecycle targeting the edit file to update or new file.
		/// This is supported for both file path and data stream
		/// </summary>
		public void SaveAs(Stream stream)
		{
			SendAnonymousSaveStates(Assembly.GetExecutingAssembly().GetName());
			document.SaveAs(stream);
		}
	}
}
