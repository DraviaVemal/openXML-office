// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
namespace OpenXMLOffice.Spreadsheet_2007
{

	/// <summary>
	/// Represents Textbox class to build on
	/// </summary>
	public class TextBox : CommonProperties
	{
		private readonly TextBoxSetting textBoxSetting;
		private XDR.Shape openXMLShape;
		private readonly Worksheet worksheet;
		/// <summary>
		/// Create Textbox with provided settings
		/// </summary>
		internal TextBox(TextBoxSetting TextBoxSetting)
		{
			textBoxSetting = TextBoxSetting;
		}
		/// <summary>
		/// Create Textbox with provided settings
		/// </summary>
		public TextBox(Worksheet Worksheet, TextBoxSetting TextBoxSetting)
		{
			worksheet = Worksheet;
			textBoxSetting = TextBoxSetting;
		}
		/// <summary>
		/// Get Textbox Shape
		/// </summary>
		internal XDR.Shape GetTextBoxShape()
		{
			return openXMLShape;
		}
	}
}
