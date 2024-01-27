// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global_2013;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation_2013
{
	internal class PresentationCore
	{


		protected readonly PresentationDocument presentationDocument;

		protected readonly PresentationInfo presentationInfo = new();

		protected readonly PresentationProperties presentationProperties;

		protected ExtendedFilePropertiesPart? extendedFilePropertiesPart;


		//#### Presentation Constants ####//
		private readonly uint slideIdStart = 255;

		private readonly uint slideMasterIdStart = 2147483647; public PresentationCore(string filePath, PresentationProperties? presentationProperties = null)
		{
			presentationInfo.filePath = filePath;
			this.presentationProperties = presentationProperties ?? new();
			MemoryStream memoryStream = new();
			presentationDocument = PresentationDocument.Create(memoryStream, PresentationDocumentType.Presentation, true);
			InitialisePresentation(this.presentationProperties);
		}

		public PresentationCore(string filePath, bool isEditable, PresentationProperties? presentationProperties = null)
		{
			presentationInfo.filePath = filePath;
			this.presentationProperties = presentationProperties ?? new();
			FileStream reader = new(filePath, FileMode.Open);
			MemoryStream memoryStream = new();
			reader.CopyTo(memoryStream);
			reader.Close();
			presentationDocument = PresentationDocument.Open(memoryStream, isEditable, new OpenSettings()
			{
				AutoSave = true
			});
			if (isEditable)
			{
				presentationInfo.isExistingFile = true;
			}
			else
			{
				presentationInfo.isEditable = false;
			}
		}

		public PresentationCore(Stream stream, PresentationProperties? presentationProperties = null)
		{
			this.presentationProperties = presentationProperties ?? new();
			presentationDocument = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, true);
			InitialisePresentation(this.presentationProperties);
		}

		public PresentationCore(Stream stream, bool isEditable, PresentationProperties? presentationProperties = null)
		{
			this.presentationProperties = presentationProperties ?? new();
			presentationDocument = PresentationDocument.Open(stream, isEditable, new OpenSettings()
			{
				AutoSave = true
			});
			InitialisePresentation(this.presentationProperties);
		}


		protected string GetNextPresentationRelationId()
		{
			return string.Format("rId{0}", GetPresentationPart().Parts.Count() + 1);
		}

		protected uint GetNextSlideId()
		{
			return (uint)(slideIdStart + GetSlideIdList().Count() + 1);
		}

		protected uint GetNextSlideMasterId()
		{
			return (uint)(slideMasterIdStart + GetSlideMasterIdList().Count() + 1);
		}

		protected PresentationPart GetPresentationPart()
		{
			return presentationDocument.PresentationPart!;
		}

		protected P.SlideIdList GetSlideIdList()
		{
			return GetPresentationPart().Presentation.SlideIdList!;
		}

		protected SlideLayoutPart GetSlideLayoutPart(PresentationConstants.SlideLayoutType slideLayoutType)
		{
			// TODO: Multi Slide Master Use
			SlideMasterPart slideMasterPart = GetPresentationPart().SlideMasterParts.FirstOrDefault()!;
			return slideMasterPart.SlideLayoutParts
				   .FirstOrDefault(sl => sl.SlideLayout.CommonSlideData!.Name == PresentationConstants.GetSlideLayoutType(slideLayoutType))!;
		}

		protected P.SlideMasterIdList GetSlideMasterIdList()
		{
			return GetPresentationPart().Presentation.SlideMasterIdList!;
		}





		private static P.DefaultTextStyle CreateDefaultTextStyle()
		{
			P.DefaultTextStyle defaultTextStyle = new();
			A.DefaultParagraphProperties defaultParagraphProperties = new();
			A.DefaultRunProperties defaultRunProperties = new() { Language = "en-IN" };
			defaultParagraphProperties.Append(defaultRunProperties);
			defaultTextStyle.Append(defaultParagraphProperties);
			A.Level1ParagraphProperties levelParagraphProperties = new()
			{
				Alignment = A.TextAlignmentTypeValues.Left,
				DefaultTabSize = 914400,
				EastAsianLineBreak = true,
				LatinLineBreak = false,
				LeftMargin = 457200,
				RightToLeft = false
			};
			A.DefaultRunProperties levelRunProperties = new()
			{
				Kerning = 1200,
				FontSize = 1800
			};
			A.SolidFill solidFill = new();
			A.SchemeColor schemeColor = new() { Val = A.SchemeColorValues.Text1 };
			solidFill.Append(schemeColor);
			levelRunProperties.Append(solidFill);
			A.LatinFont latinTypeface = new() { Typeface = "+mn-lt" };
			A.EastAsianFont eastAsianTypeface = new() { Typeface = "+mn-ea" };
			A.ComplexScriptFont complexScriptTypeface = new() { Typeface = "+mn-cs" };
			levelRunProperties.Append(latinTypeface, eastAsianTypeface, complexScriptTypeface);
			levelParagraphProperties.Append(levelRunProperties);
			defaultTextStyle.Append(levelParagraphProperties);
			return defaultTextStyle;
		}

		private void InitialisePresentation(PresentationProperties? powerPointProperties)
		{
			SlideMaster slideMaster = new();
			SlideLayout slideLayout = new();
			PresentationPart presentationPart = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();
			presentationPart.Presentation ??= new P.Presentation();
			presentationPart.Presentation.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			presentationPart.Presentation.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			presentationPart.Presentation.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
			if (presentationPart.Presentation.GetFirstChild<P.SlideMasterIdList>() == null)
			{
				presentationPart.Presentation.AppendChild(new P.SlideMasterIdList());
			}
			if (presentationPart.Presentation.SlideIdList == null)
			{
				presentationPart.Presentation.AppendChild(new P.SlideIdList());
			}
			if (presentationPart.Presentation.GetFirstChild<P.SlideSize>() == null)
			{
				presentationPart.Presentation.AppendChild(new P.SlideSize { Cx = 12192000, Cy = 6858000 });
			}
			if (presentationPart.Presentation.GetFirstChild<P.NotesSize>() == null)
			{
				presentationPart.Presentation.AppendChild(new P.NotesSize { Cx = 6858000, Cy = 6858000 });
			}
			if (presentationPart.Presentation.GetFirstChild<P.DefaultTextStyle>() == null)
			{
				presentationPart.Presentation.AppendChild(CreateDefaultTextStyle());
			}
			if (presentationPart.ViewPropertiesPart == null)
			{
				ViewProperties viewProperties = new();
				ViewPropertiesPart viewPropertiesPart = presentationPart.AddNewPart<ViewPropertiesPart>(GetNextPresentationRelationId());
				viewPropertiesPart.ViewProperties = viewProperties.GetViewProperties();
				viewPropertiesPart.ViewProperties.Save();
			}
			if (presentationPart.PresentationPropertiesPart == null)
			{
				PresentationPropertiesPart presentationPropertiesPart = presentationPart.AddNewPart<PresentationPropertiesPart>(GetNextPresentationRelationId());
				presentationPropertiesPart.PresentationProperties ??= new P.PresentationProperties();
				presentationPropertiesPart.PresentationProperties.Save();
			}
			if (!presentationPart.SlideMasterParts.Any())
			{
				SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>(GetNextPresentationRelationId());
				slideMasterPart.SlideMaster = slideMaster.GetSlideMaster();
				P.SlideMasterId slideMasterId = new() { Id = GetNextSlideMasterId(), RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) };
				GetSlideMasterIdList().Append(slideMasterId);
				SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>(GetNextPresentationRelationId());
				slideMaster.AddSlideLayoutIdToList(slideMasterPart.GetIdOfPart(slideLayoutPart));
				slideLayoutPart.SlideLayout = slideLayout.GetSlideLayout();
				slideLayout.UpdateRelationship(slideMasterPart, presentationPart.GetIdOfPart(slideMasterPart));
				slideLayoutPart.SlideLayout.Save();
				slideMasterPart.SlideMaster.Save();
			}
			if (presentationDocument.ExtendedFilePropertiesPart == null)
			{
				extendedFilePropertiesPart = presentationDocument.AddExtendedFilePropertiesPart();
				extendedFilePropertiesPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
				extendedFilePropertiesPart.Properties.Save();
			}
			if (presentationPart.TableStylesPart == null)
			{
				TableStylesPart tableStylesPart = presentationPart.AddNewPart<TableStylesPart>(GetNextPresentationRelationId());
				tableStylesPart.TableStyleList ??= new A.TableStyleList()
				{
					Default = GeneratorUtils.GenerateNewGUID()
				};
				tableStylesPart.TableStyleList.Save();
			}
			if (presentationPart.ThemePart == null)
			{
				presentationPart.AddNewPart<ThemePart>(GetNextPresentationRelationId());
			}
			Theme theme = new(powerPointProperties?.theme);
			presentationPart.ThemePart!.Theme = theme.GetTheme();
			slideMaster.UpdateRelationship(presentationPart.ThemePart, presentationPart.GetIdOfPart(presentationPart.ThemePart));
			presentationPart.Presentation.Save();
		}


	}
}
