// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using OpenXMLOffice.Global_2007;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	///
	/// </summary>
	public class ShapeTextModel
	{
		/// <summary>
		///
		/// </summary>
		public string text;
		/// <summary>
		///
		/// </summary>
		public string fontColor;
		/// <summary>
		///
		/// </summary>
		public string fontFamily = "(Calibri (Body))";
		/// <summary>
		///
		/// </summary>
		public int fontSize = 8;
		/// <summary>
		///
		/// </summary>
		public bool? isBold;
		/// <summary>
		///
		/// </summary>
		public bool? isItalic;
		/// <summary>
		///
		/// </summary>
		public UnderLineValues? underline = null;
	}
}
