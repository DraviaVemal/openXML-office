// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using OpenXMLOffice.Global_2007;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// Excel Picture
	/// </summary>
	public class Picture
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
			currentWorksheet.CreateTwoCellAnchor(new TwoCellAnchorModel()
			{
				anchorEditType = AnchorEditType.ONE_CELL,
				from = new AnchorPosition()
				{
					column = excelPictureSetting.fromCol,
					columnOffset = excelPictureSetting.fromColOff,
					row = excelPictureSetting.fromRow,
					rowOffset = excelPictureSetting.fromRowOff,
				},
				to = new AnchorPosition()
				{
					column = excelPictureSetting.toCol,
					columnOffset = excelPictureSetting.toColOff,
					row = excelPictureSetting.toRow,
					rowOffset = excelPictureSetting.toRowOff
				},
				drawingPictureModel = new DrawingPictureModel()
				{
					id = 2U,
					name = "Picture 1",
					noChangeAspectRatio = true,
					blipEmbed = embedId
				}
			});
		}
	}
}
