// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;
namespace OpenXMLOffice.Document_2007
{
	/// <summary>
	/// 
	/// </summary>
	public class WordProperties
	{
		/// <summary>
		/// Spreadsheet settings
		/// </summary>
		public WordSettings settings = new WordSettings();
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
	/// 
	/// </summary>
	public class WordSettings
	{
	}
	internal class WordInfo
	{
		public bool isEditable = true;
		public bool isExistingFile;
	}
}
