// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global_2007;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
namespace OpenXMLOffice.Presentation_2007
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
		public Picture(Stream stream, Slide slide, PictureSetting pictureSetting)
		{
			currentSlide = slide;
			string EmbedId = currentSlide.GetNextSlideRelationId();
			this.pictureSetting = pictureSetting;
			openXMLPicture = new P.Picture();
			ImagePart ImagePart;
			if (pictureSetting.imageType == ImageType.PNG)
			{
				ImagePart = currentSlide.GetSlidePart().AddNewPart<ImagePart>("image/png", EmbedId);
			}
			else if (pictureSetting.imageType == ImageType.GIF)
			{
				ImagePart = currentSlide.GetSlidePart().AddNewPart<ImagePart>("image/gif", EmbedId);
			}
			else if (pictureSetting.imageType == ImageType.TIFF)
			{
				ImagePart = currentSlide.GetSlidePart().AddNewPart<ImagePart>("image/tiff", EmbedId);
			}
			else
			{
				ImagePart = currentSlide.GetSlidePart().AddNewPart<ImagePart>("image/jpeg", EmbedId);
			}
			// Add Hyperlink Relationships to slide
			if (pictureSetting.hyperlinkProperties != null)
			{
				string relationId = slide.GetNextSlideRelationId();
				switch (pictureSetting.hyperlinkProperties.hyperlinkPropertyType)
				{
					case HyperlinkPropertyType.EXISTING_FILE:
						pictureSetting.hyperlinkProperties.relationId = relationId;
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinkfile";
						slide.GetSlidePart().AddHyperlinkRelationship(new Uri(pictureSetting.hyperlinkProperties.value), true, relationId);
						break;
					case HyperlinkPropertyType.TARGET_SLIDE:
						pictureSetting.hyperlinkProperties.relationId = relationId;
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinksldjump";
						//TODO: Update Target Slide Prop
						slide.GetSlidePart().AddHyperlinkRelationship(new Uri(pictureSetting.hyperlinkProperties.value), true, relationId);
						break;
					case HyperlinkPropertyType.FIRST_SLIDE:
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=firstslide";
						break;
					case HyperlinkPropertyType.LAST_SLIDE:
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=lastslide";
						break;
					case HyperlinkPropertyType.NEXT_SLIDE:
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=nextslide";
						break;
					case HyperlinkPropertyType.PREVIOUS_SLIDE:
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=previousslide";
						break;
					default:// Web URL
						pictureSetting.hyperlinkProperties.relationId = relationId;
						slide.GetSlidePart().AddHyperlinkRelationship(new Uri(pictureSetting.hyperlinkProperties.value), true, relationId);
						break;
				}
			}
			CreatePicture(EmbedId, pictureSetting.hyperlinkProperties);
			slide.GetSlide().CommonSlideData.ShapeTree.Append(GetPicture());
			ImagePart.FeedData(stream);
		}
		/// <summary>
		/// Create Picture Object with provided settings
		/// </summary>
		public Picture(string filePath, Slide slide, PictureSetting pictureSetting)
		{
			using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
			{
				new Picture(fileStream, slide, pictureSetting);
			}
		}
		/// <summary>
		/// X,Y
		/// </summary>
		public Tuple<uint, uint> GetPosition()
		{
			return Tuple.Create(pictureSetting.x, pictureSetting.y);
		}
		/// <summary>
		/// Width,Height
		/// </summary>
		public Tuple<uint, uint> GetSize()
		{
			return Tuple.Create(pictureSetting.width, pictureSetting.height);
		}
		/// <summary>
		/// Update Picture Position
		/// </summary>
		public void UpdatePosition(uint X, uint Y)
		{
			pictureSetting.x = X;
			pictureSetting.y = Y;
			if (openXMLPicture != null)
			{
				openXMLPicture.ShapeProperties.Transform2D = new A.Transform2D
				{
					Offset = new A.Offset { X = pictureSetting.x, Y = pictureSetting.y },
					Extents = new A.Extents { Cx = pictureSetting.width, Cy = pictureSetting.height }
				};
			}
		}
		/// <summary>
		/// Update Picture Size
		/// </summary>
		public void UpdateSize(uint Width, uint Height)
		{
			pictureSetting.width = Width;
			pictureSetting.height = Height;
			if (openXMLPicture != null)
			{
				openXMLPicture.ShapeProperties.Transform2D = new A.Transform2D
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
		private void CreatePicture(string EmbedId, HyperlinkProperties hyperlinkProperties)
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
			if (hyperlinkProperties != null)
			{
				GetPicture().NonVisualPictureProperties.NonVisualDrawingProperties.InsertAt(CreateHyperLink(hyperlinkProperties), 0);
			}
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
			GetPicture().BlipFill = new P.BlipFill()
			{
				Blip = new A.Blip() { Embed = EmbedId }
			};
			GetPicture().BlipFill.Append(new A.Stretch(new A.FillRectangle()));
		}
	}
}
