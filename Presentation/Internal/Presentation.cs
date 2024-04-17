// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
namespace OpenXMLOffice.Presentation_2007
{
	internal class Presentation : PresentationCore
	{
		internal Presentation(PowerPointProperties presentationProperties = null)
		: base(presentationProperties) { }
		internal Presentation(string filePath, bool isEditable, PowerPointProperties presentationProperties = null, bool autosave = true)
		: base(filePath, isEditable, presentationProperties) { }
		internal Presentation(Stream stream, bool isEditable, PowerPointProperties presentationProperties = null)
		: base(stream, isEditable, presentationProperties) { }
		internal Slide AddSlide(PresentationConstants.SlideLayoutType slideLayoutType)
		{
			SlidePart slidePart = GetPresentationPart().AddNewPart<SlidePart>(GetNextPresentationRelationId());
			Slide slide = new Slide();
			slidePart.Slide = slide.GetSlide();
			slidePart.AddPart(GetSlideLayoutPart(slideLayoutType));
			P.SlideIdList slideIdList = GetSlideIdList();
			P.SlideId slideId = new P.SlideId() { Id = GetNextSlideId(), RelationshipId = GetPresentationPart().GetIdOfPart(slidePart) };
			slideIdList.Append(slideId);
			return slide;
		}
		internal Slide GetSlideByIndex(int SlideIndex)
		{
			if (SlideIndex >= 0 && GetSlideIdList().Count() > SlideIndex)
			{
				P.SlideId SlideId = (P.SlideId)GetSlideIdList().ElementAt(SlideIndex);
				SlidePart SlidePart = (SlidePart)GetPresentationPart().GetPartById(SlideId.RelationshipId.Value);
				return new Slide(SlidePart.Slide);
			}
			else
			{
				throw new IndexOutOfRangeException("The specified slide index is out of range.");
			}
		}
		internal int GetSlideCount()
		{
			return GetSlideIdList().Count();
		}
		internal void MoveSlideByIndex(int SourceIndex, int TargetIndex)
		{
			if (SourceIndex >= 0 && GetSlideIdList().Count() > SourceIndex &&
			TargetIndex >= 0 && GetSlideIdList().Count() > TargetIndex)
			{
				P.SlideId SourceSlideId = (P.SlideId)GetSlideIdList().ElementAt(SourceIndex);
				P.SlideId TargetSlideId = (P.SlideId)GetSlideIdList().ElementAt(TargetIndex);
				GetSlideIdList().RemoveChild(SourceSlideId);
				GetSlideIdList().InsertBefore(SourceSlideId, TargetSlideId);
				presentationDocument.Save();
			}
			else
			{
				throw new IndexOutOfRangeException("The specified slide index is out of range.");
			}
		}
		internal void RemoveSlideByIndex(int SlideIndex)
		{
			if (SlideIndex >= 0 && GetSlideIdList().Count() > SlideIndex)
			{
				P.SlideId SlideId = (P.SlideId)GetSlideIdList().ElementAt(SlideIndex);
				SlidePart SlidePart = (SlidePart)GetPresentationPart().GetPartById(SlideId.RelationshipId.Value);
				GetSlideIdList().RemoveChild(SlideId);
				try
				{
					GetPresentationPart().DeleteReferenceRelationship(SlideId.RelationshipId.Value);
				}
				catch (ArgumentNullException)
				{
					//ignore
				}
				catch (Exception ex)
				{
					throw ex;
				}
				GetPresentationPart().DeletePart(SlidePart);
			}
			else
			{
				throw new IndexOutOfRangeException("The specified slide index is out of range.");
			}
		}
		internal void SaveAs(string filePath)
		{
			presentationDocument.Clone(filePath).Dispose();
			presentationDocument.Dispose();
		}
	}
}
