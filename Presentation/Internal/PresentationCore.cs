// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global_2007;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
namespace OpenXMLOffice.Presentation_2007
{
	internal class PresentationCore
	{
		internal readonly PresentationDocument presentationDocument;
		internal readonly PowerPointInfo presentationInfo = new PowerPointInfo();
		internal readonly PowerPointProperties presentationProperties;
		internal ExtendedFilePropertiesPart extendedFilePropertiesPart;
		//#### Presentation Constants ####//
		private readonly uint slideIdStart = 255;
		private readonly uint slideMasterIdStart = 2147483647;
		public PresentationCore(PowerPointProperties presentationProperties = null)
		{
			this.presentationProperties = presentationProperties ?? new PowerPointProperties();
			using (MemoryStream memoryStream = new MemoryStream())
			{
				presentationDocument = PresentationDocument.Create(memoryStream, PresentationDocumentType.Presentation, true);
				InitialisePresentation(this.presentationProperties);
			}
		}
		public PresentationCore(string filePath, bool isEditable = true, PowerPointProperties presentationProperties = null)
		{
			presentationInfo.isEditable = isEditable;
			this.presentationProperties = presentationProperties ?? new PowerPointProperties();
			FileStream reader = new FileStream(filePath, FileMode.Open);
			using (MemoryStream memoryStream = new MemoryStream())
			{
				reader.CopyTo(memoryStream);
				reader.Close();
				presentationDocument = PresentationDocument.Open(memoryStream, isEditable, new OpenSettings()
				{
					AutoSave = true
				});
				if (presentationInfo.isEditable)
				{
					InitialisePresentation(this.presentationProperties);
					presentationInfo.isExistingFile = true;
				}
			}
		}
		internal PresentationCore(Stream stream, bool isEditable = true, PowerPointProperties presentationProperties = null)
		{
			presentationInfo.isEditable = isEditable;
			this.presentationProperties = presentationProperties ?? new PowerPointProperties();
			using (MemoryStream memoryStream = new MemoryStream())
			{
				stream.CopyTo(memoryStream);
				stream.Dispose();
				presentationDocument = PresentationDocument.Open(memoryStream, isEditable, new OpenSettings()
				{
					AutoSave = true
				});
				if (presentationInfo.isEditable)
				{
					InitialisePresentation(this.presentationProperties);
				}
			}
		}
		internal string GetNextPresentationRelationId()
		{
			return string.Format("rId{0}", GetPresentationPart().Parts.Count() + 1);
		}
		internal uint GetNextSlideId()
		{
			return (uint)(slideIdStart + GetSlideIdList().Count() + 1);
		}
		internal uint GetNextSlideMasterId()
		{
			return (uint)(slideMasterIdStart + GetSlideMasterIdList().Count() + 1);
		}
		internal PresentationPart GetPresentationPart()
		{
			return presentationDocument.PresentationPart;
		}
		internal P.SlideIdList GetSlideIdList()
		{
			return GetPresentationPart().Presentation.SlideIdList;
		}
		internal SlideLayoutPart GetSlideLayoutPart(PresentationConstants.SlideLayoutType slideLayoutType)
		{
			// TODO: Multi Slide Master Use
			SlideMasterPart slideMasterPart = GetPresentationPart().SlideMasterParts.FirstOrDefault();
			return slideMasterPart.SlideLayoutParts
				   .FirstOrDefault(sl => sl.SlideLayout.CommonSlideData.Name == PresentationConstants.GetSlideLayoutType(slideLayoutType));
		}
		internal P.SlideMasterIdList GetSlideMasterIdList()
		{
			return GetPresentationPart().Presentation.SlideMasterIdList;
		}
		private static P.DefaultTextStyle CreateDefaultTextStyle()
		{
			P.DefaultTextStyle defaultTextStyle = new P.DefaultTextStyle();
			A.DefaultParagraphProperties defaultParagraphProperties = new A.DefaultParagraphProperties();
			A.DefaultRunProperties defaultRunProperties = new A.DefaultRunProperties() { Language = "en-IN" };
			defaultParagraphProperties.Append(defaultRunProperties);
			defaultTextStyle.Append(defaultParagraphProperties);
			A.Level1ParagraphProperties levelParagraphProperties = new A.Level1ParagraphProperties()
			{
				Alignment = A.TextAlignmentTypeValues.Left,
				DefaultTabSize = 914400,
				EastAsianLineBreak = true,
				LatinLineBreak = false,
				LeftMargin = 457200,
				RightToLeft = false
			};
			A.DefaultRunProperties levelRunProperties = new A.DefaultRunProperties()
			{
				Kerning = 1200,
				FontSize = 1800
			};
			A.SolidFill solidFill = new A.SolidFill();
			A.SchemeColor schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			solidFill.Append(schemeColor);
			levelRunProperties.Append(solidFill);
			A.LatinFont latinTypeface = new A.LatinFont() { Typeface = "+mn-lt" };
			A.EastAsianFont eastAsianTypeface = new A.EastAsianFont() { Typeface = "+mn-ea" };
			A.ComplexScriptFont complexScriptTypeface = new A.ComplexScriptFont() { Typeface = "+mn-cs" };
			levelRunProperties.Append(latinTypeface, eastAsianTypeface, complexScriptTypeface);
			levelParagraphProperties.Append(levelRunProperties);
			defaultTextStyle.Append(levelParagraphProperties);
			return defaultTextStyle;
		}
		private void InitialisePresentation(PowerPointProperties powerPointProperties)
		{
			SlideMaster slideMaster = new SlideMaster();
			SlideLayout slideLayout = new SlideLayout();
			if (presentationDocument.CoreFilePropertiesPart == null)
			{
				presentationDocument.AddCoreFilePropertiesPart();
				using (Stream stream = presentationDocument.CoreFilePropertiesPart.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite))
				{
					CoreProperties.AddCoreProperties(stream, powerPointProperties.coreProperties);
				}
			}
			else
			{
				using (Stream stream = presentationDocument.CoreFilePropertiesPart.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite))
				{
					CoreProperties.UpdateModifiedDetails(stream, powerPointProperties.coreProperties);
				}
			}
			PresentationPart presentationPart = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();
			if (presentationPart.Presentation == null)
			{
				presentationPart.Presentation = new P.Presentation();
				presentationPart.Presentation.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
				presentationPart.Presentation.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
				presentationPart.Presentation.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
			}
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
				ViewProperties viewProperties = new ViewProperties();
				ViewPropertiesPart viewPropertiesPart = presentationPart.AddNewPart<ViewPropertiesPart>(GetNextPresentationRelationId());
				viewPropertiesPart.ViewProperties = viewProperties.GetViewProperties();
				viewPropertiesPart.ViewProperties.Save();
			}
			if (presentationPart.PresentationPropertiesPart == null)
			{
				PresentationPropertiesPart presentationPropertiesPart = presentationPart.AddNewPart<PresentationPropertiesPart>(GetNextPresentationRelationId());
				if (presentationPropertiesPart.PresentationProperties == null)
				{
					presentationPropertiesPart.PresentationProperties = new P.PresentationProperties();
				}
				presentationPropertiesPart.PresentationProperties.Save();
			}
			if (!presentationPart.SlideMasterParts.Any())
			{
				SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>(GetNextPresentationRelationId());
				slideMasterPart.SlideMaster = slideMaster.GetSlideMaster();
				P.SlideMasterId slideMasterId = new P.SlideMasterId() { Id = GetNextSlideMasterId(), RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) };
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
				if (extendedFilePropertiesPart.Properties == null)
				{
					extendedFilePropertiesPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
				}
				extendedFilePropertiesPart.Properties.Save();
			}
			if (presentationPart.TableStylesPart == null)
			{
				TableStylesPart tableStylesPart = presentationPart.AddNewPart<TableStylesPart>(GetNextPresentationRelationId());
				if (tableStylesPart.TableStyleList == null)
				{
					tableStylesPart.TableStyleList = new A.TableStyleList()
					{
						Default = GeneratorUtils.GenerateNewGUID()
					};
				}
				tableStylesPart.TableStyleList.Save();
			}
			if (presentationPart.ThemePart == null)
			{
				presentationPart.AddNewPart<ThemePart>(GetNextPresentationRelationId());
			}
			Theme theme = new Theme(powerPointProperties.theme);
			presentationPart.ThemePart.Theme = theme.GetTheme();
			slideMaster.UpdateRelationship(presentationPart.ThemePart, presentationPart.GetIdOfPart(presentationPart.ThemePart));
			presentationPart.Presentation.Save();
		}
	}
}
