// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	///
	/// </summary>
	public class ExcelPictureSetting
	{
		/// <summary>
		///
		/// </summary>
		public HyperlinkProperties hyperlinkProperties = null;
		/// <summary>
		/// The type of image.
		/// </summary>
		public ImageType imageType = ImageType.JPEG;
		/// <summary>
		///
		/// </summary>
		public AnchorEditType anchorEditType = AnchorEditType.NONE;
		/// <summary>
		///
		/// </summary>
		public AnchorPosition from = new AnchorPosition();
		/// <summary>
		///
		/// </summary>
		public AnchorPosition to = new AnchorPosition();
	}
}
