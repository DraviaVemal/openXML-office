// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.IO;
using OpenXMLOffice.Global_2007;
using DocumentFormat.OpenXml.Packaging;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// Excel Picture
	/// </summary>
	public class Picture : CommonProperties
	{
		private readonly ExcelPictureSetting excelPictureSetting;
		private readonly Worksheet currentWorksheet;
		/// <summary>
		/// Initializes a new instance of the <see cref="Picture"/> class.
		/// </summary>
		internal Picture(Worksheet worksheet, Stream stream, ExcelPictureSetting excelPictureSetting)
		{
			this.excelPictureSetting = excelPictureSetting;
			currentWorksheet = worksheet;
			AddImageToDrawing(stream);
		}
		private void AddImageToDrawing(Stream stream)
		{
			string embedId = currentWorksheet.GetNextSheetPartRelationId();
			ImagePart imagePart;
			switch (excelPictureSetting.imageType)
			{
				case ImageType.PNG:
					imagePart = currentWorksheet.GetDrawingsPart().AddNewPart<ImagePart>("image/png", embedId);
					break;
				case ImageType.GIF:
					imagePart = currentWorksheet.GetDrawingsPart().AddNewPart<ImagePart>("image/gif", embedId);
					break;
				case ImageType.TIFF:
					imagePart = currentWorksheet.GetDrawingsPart().AddNewPart<ImagePart>("image/tiff", embedId);
					break;
				default:
					imagePart = currentWorksheet.GetDrawingsPart().AddNewPart<ImagePart>("image/jpeg", embedId);
					break;
			}
			imagePart.FeedData(stream);
			if (excelPictureSetting.hyperlinkProperties != null)
			{
				string relationId = currentWorksheet.GetNextDrawingPartRelationId();
				switch (excelPictureSetting.hyperlinkProperties.hyperlinkPropertyType)
				{
					case HyperlinkPropertyTypeValues.EXISTING_FILE:
						excelPictureSetting.hyperlinkProperties.relationId = relationId;
						excelPictureSetting.hyperlinkProperties.action = "ppaction://hlinkfile";
						currentWorksheet.GetDrawingsPart().AddHyperlinkRelationship(new Uri(excelPictureSetting.hyperlinkProperties.value), true, relationId);
						break;
					case HyperlinkPropertyTypeValues.TARGET_SHEET:
						excelPictureSetting.hyperlinkProperties.relationId = relationId;
						excelPictureSetting.hyperlinkProperties.action = "ppaction://hlinksldjump";
						//TODO: Update Target Slide Prop
						currentWorksheet.GetDrawingsPart().AddHyperlinkRelationship(new Uri(excelPictureSetting.hyperlinkProperties.value), true, relationId);
						break;
					case HyperlinkPropertyTypeValues.TARGET_SLIDE:
					case HyperlinkPropertyTypeValues.FIRST_SLIDE:
					case HyperlinkPropertyTypeValues.LAST_SLIDE:
					case HyperlinkPropertyTypeValues.NEXT_SLIDE:
					case HyperlinkPropertyTypeValues.PREVIOUS_SLIDE:
						throw new ArgumentException("This Option is valid only for Powerpoint Files");
					default:// Web URL
						excelPictureSetting.hyperlinkProperties.relationId = relationId;
						currentWorksheet.GetDrawingsPart().AddHyperlinkRelationship(new Uri(excelPictureSetting.hyperlinkProperties.value), true, relationId);
						break;
				}
			}
			XDR.TwoCellAnchor twoCellAnchor = currentWorksheet.CreateTwoCellAnchor(new TwoCellAnchorModel()
			{
				anchorEditType = AnchorEditType.ONE_CELL,
				from = excelPictureSetting.from,
				to = excelPictureSetting.to,
				drawingPictureModel = new DrawingPictureModel()
				{
					id = 2U,
					name = "Picture 1",
					noChangeAspectRatio = true,
					blipEmbed = embedId,
					hyperlinkProperties = excelPictureSetting.hyperlinkProperties
				}
			});
			currentWorksheet.GetDrawing().AppendChild(twoCellAnchor);
		}
	}
}
