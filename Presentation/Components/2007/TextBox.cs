// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using OpenXMLOffice.Global_2007;
using P = DocumentFormat.OpenXml.Presentation;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Textbox Class
	/// </summary>
	public class TextBox : TextBoxBase
	{
		/// <summary>
		/// Create Textbox with provided settings
		/// </summary>
		/// <param name="TextBoxSetting">
		/// </param>
		public TextBox(TextBoxSetting TextBoxSetting) : base(TextBoxSetting) { }
		/// <summary>
		/// Return OpenXML Shape
		/// </summary>
		/// <returns>
		/// </returns>
		internal P.Shape GetTextBoxShape()
		{
			return GetTextBoxBaseShape();
		}
	}
}
