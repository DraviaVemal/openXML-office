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
	public class Picture : PresentationCommonProperties
	{
		private readonly Slide currentSlide;
		private readonly P.Picture openXMLPicture;
		private readonly PictureSetting pictureSetting;

		/// <summary>
		/// Create Picture Object with provided settings
		/// </summary>
		public Picture(string filePath, Slide slide, PictureSetting pictureSetting)
		{
			currentSlide = slide;
			this.pictureSetting = pictureSetting;
			openXMLPicture = new P.Picture();
			using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
			{
				Initialize(fileStream, slide, pictureSetting);
			}
		}

		/// <summary>
		/// Create Picture Object with provided settings
		/// </summary>
		public Picture(Stream stream, Slide slide, PictureSetting pictureSetting)
		{
			currentSlide = slide;
			this.pictureSetting = pictureSetting;
			openXMLPicture = new P.Picture();
			Initialize(stream, slide, pictureSetting);
		}

		private void Initialize(Stream stream, Slide slide, PictureSetting pictureSetting)
		{
			string EmbedId = currentSlide.GetNextSlideRelationId();
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
					case HyperlinkPropertyTypeValues.EXISTING_FILE:
						pictureSetting.hyperlinkProperties.relationId = relationId;
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinkfile";
						var unused2 = slide.GetSlidePart().AddHyperlinkRelationship(new Uri(pictureSetting.hyperlinkProperties.value), true, relationId);
						break;
					case HyperlinkPropertyTypeValues.TARGET_SLIDE:
						pictureSetting.hyperlinkProperties.relationId = relationId;
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinksldjump";
						//TODO: Update Target Slide Prop
						var unused1 = slide.GetSlidePart().AddHyperlinkRelationship(new Uri(pictureSetting.hyperlinkProperties.value), true, relationId);
						break;
					case HyperlinkPropertyTypeValues.TARGET_SHEET:
						throw new ArgumentException("This Option is valid only for Excel Files");
					case HyperlinkPropertyTypeValues.FIRST_SLIDE:
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=firstslide";
						break;
					case HyperlinkPropertyTypeValues.LAST_SLIDE:
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=lastslide";
						break;
					case HyperlinkPropertyTypeValues.NEXT_SLIDE:
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=nextslide";
						break;
					case HyperlinkPropertyTypeValues.PREVIOUS_SLIDE:
						pictureSetting.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=previousslide";
						break;
					default:// Web URL
						pictureSetting.hyperlinkProperties.relationId = relationId;
						var unused = slide.GetSlidePart().AddHyperlinkRelationship(new Uri(pictureSetting.hyperlinkProperties.value), true, relationId);
						break;
				}
			}
			CreatePicture(EmbedId, pictureSetting.hyperlinkProperties);
			slide.GetSlide().CommonSlideData.ShapeTree.Append(GetPicture());
			ImagePart.FeedData(stream);
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
			pictureSetting.x = (uint)ConverterUtils.PixelsToEmu((int)X);
			pictureSetting.y = (uint)ConverterUtils.PixelsToEmu((int)Y);
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
			pictureSetting.width = (uint)ConverterUtils.PixelsToEmu((int)Width);
			pictureSetting.height = (uint)ConverterUtils.PixelsToEmu((int)Height);
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
				var unused = GetPicture().NonVisualPictureProperties.NonVisualDrawingProperties.InsertAt(CreateHyperLink(hyperlinkProperties), 0);
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
