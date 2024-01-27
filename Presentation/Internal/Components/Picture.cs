// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global_2013;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation_2013
{
	/// <summary>
	/// Picture Import Class
	/// </summary>
	public class Picture : CommonProperties
	{
		private readonly Slide currentSlide;

		private readonly P.Picture openXMLPicture;

		private readonly PictureSetting pictureSetting;

		/// <summary>
		/// Create Picture Object with provided settings
		/// </summary>
		public Picture(Stream Stream, Slide Slide, PictureSetting PictureSetting)
		{
			currentSlide = Slide;
			string EmbedId = currentSlide.GetNextSlideRelationId();
			pictureSetting = PictureSetting;
			openXMLPicture = new();
			CreatePicture(EmbedId);
			ImagePart ImagePart = currentSlide.GetSlide().SlidePart!.AddNewPart<ImagePart>(PictureSetting.imageType switch
			{
				ImageType.PNG => "image/png",
				ImageType.GIF => "image/gif",
				ImageType.TIFF => "image/tiff",
				_ => "image/jpeg"
			}, EmbedId);
			ImagePart.FeedData(Stream);
		}

		/// <summary>
		/// Create Picture Object with provided settings
		/// </summary>
		public Picture(string FilePath, Slide Slide, PictureSetting PictureSetting)
		{
			currentSlide = Slide;
			string EmbedId = currentSlide.GetNextSlideRelationId();
			pictureSetting = PictureSetting;
			openXMLPicture = new();
			CreatePicture(EmbedId);
			ImagePart ImagePart = currentSlide.GetSlide().SlidePart!.AddNewPart<ImagePart>(PictureSetting.imageType switch
			{
				ImageType.PNG => "image/png",
				ImageType.GIF => "image/gif",
				ImageType.TIFF => "image/tiff",
				_ => "image/jpeg"
			}, EmbedId);
			ImagePart.FeedData(new FileStream(FilePath, FileMode.Open, FileAccess.Read));
		}


		/// <summary>
		/// </summary>
		/// <returns>
		/// X,Y
		/// </returns>
		public (uint, uint) GetPosition()
		{
			return (pictureSetting.x, pictureSetting.y);
		}

		/// <summary>
		/// </summary>
		/// <returns>
		/// Width,Height
		/// </returns>
		public (uint, uint) GetSize()
		{
			return (pictureSetting.width, pictureSetting.height);
		}

		/// <summary>
		/// Update Picture Position
		/// </summary>
		/// <param name="X">
		/// </param>
		/// <param name="Y">
		/// </param>
		public void UpdatePosition(uint X, uint Y)
		{
			pictureSetting.x = X;
			pictureSetting.y = Y;
			if (openXMLPicture != null)
			{
				openXMLPicture.ShapeProperties!.Transform2D = new A.Transform2D
				{
					Offset = new A.Offset { X = pictureSetting.x, Y = pictureSetting.y },
					Extents = new A.Extents { Cx = pictureSetting.width, Cy = pictureSetting.height }
				};
			}
		}

		/// <summary>
		/// Update Picture Size
		/// </summary>
		/// <param name="Width">
		/// </param>
		/// <param name="Height">
		/// </param>
		public void UpdateSize(uint Width, uint Height)
		{
			pictureSetting.width = Width;
			pictureSetting.height = Height;
			if (openXMLPicture != null)
			{
				openXMLPicture.ShapeProperties!.Transform2D = new A.Transform2D
				{
					Offset = new A.Offset { X = pictureSetting.x, Y = pictureSetting.y },
					Extents = new A.Extents { Cx = pictureSetting.width, Cy = pictureSetting.height }
				};
			}
		}

		internal P.Picture GetPicture()
		{
			return openXMLPicture;
		}

		private void CreatePicture(string EmbedId)
		{
			GetPicture().NonVisualPictureProperties = new P.NonVisualPictureProperties(
				new P.NonVisualDrawingProperties()
				{
					Id = 1,
					Name = "Picture"
				},
				new P.NonVisualPictureDrawingProperties(
					new A.PictureLocks()
					{
						NoChangeAspect = true
					}
				),
				new P.ApplicationNonVisualDrawingProperties()
			);
			GetPicture().ShapeProperties = new P.ShapeProperties(
				new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
			)
			{
				Transform2D = new A.Transform2D()
				{
					Offset = new A.Offset() { X = pictureSetting.x, Y = pictureSetting.y },
					Extents = new A.Extents() { Cx = pictureSetting.width, Cy = pictureSetting.height }
				}
			};
			GetPicture().BlipFill = new()
			{
				Blip = new A.Blip() { Embed = EmbedId }
			};
			GetPicture().BlipFill!.Append(new A.Stretch(new A.FillRectangle()));
		}
	}
}
