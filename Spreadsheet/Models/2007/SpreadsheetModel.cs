// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using OpenXMLOffice.Global_2007;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// Represents the properties of a column in a worksheet.
	/// </summary>
	public class ExcelProperties
	{
		/// <summary>
		/// Spreadsheet settings
		/// </summary>
		public ExcelSettings settings = new ExcelSettings();
		/// <summary>
		/// Spreadsheet theme settings
		/// </summary>
		public ThemePallet theme = new ThemePallet();
		/// <summary>
		/// Add Meta Data Details to File
		/// </summary>
		public CorePropertiesModel coreProperties = new CorePropertiesModel();
	}
	/// <summary>
	/// Represents the settings of a spreadsheet.
	/// </summary>
	public class ExcelSettings
	{
	}
	internal class ExcelInfo
	{
		public bool isEditable = true;
		public bool isExistingFile = false;
	}
}
