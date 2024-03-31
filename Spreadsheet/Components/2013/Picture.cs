// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLOffice.Spreadsheet_2013
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
			ImagePart imagePart = currentWorksheet.GetDrawingsPart().AddNewPart<ImagePart>(excelPictureSetting.imageType switch
			{
				ImageType.PNG => "image/png",
				ImageType.GIF => "image/gif",
				ImageType.TIFF => "image/tiff",
				_ => "image/jpeg"
			}, embedId);
			imagePart.FeedData(stream);
			currentWorksheet.CreateTwoCellAnchor(new()
			{
				from = new()
				{
					column = excelPictureSetting.fromCol,
					columnOffset = excelPictureSetting.fromColOff,
					row = excelPictureSetting.fromRow,
					rowOffset = excelPictureSetting.fromRowOff,
				},
				to = new()
				{
					column = excelPictureSetting.toCol,
					columnOffset = excelPictureSetting.toColOff,
					row = excelPictureSetting.toRow,
					rowOffset = excelPictureSetting.toRowOff
				},
				drawingPictureModel = new()
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
