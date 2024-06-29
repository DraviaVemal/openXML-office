// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
namespace OpenXMLOffice.Presentation_2007
{
	internal class SlideMaster
	{
		internal P.SlideMaster openXMLSlideMaster = new P.SlideMaster();
		internal P.SlideLayoutIdList slideLayoutIdList;
		private readonly CommonSlideData commonSlideData = new CommonSlideData(PresentationConstants.CommonSlideDataType.SLIDE_MASTER, PresentationConstants.SlideLayoutType.BLANK); public SlideMaster()
		{
			CreateSlideMaster();
		}
		public void AddSlideLayoutIdToList(string relationshipId)
		{
			slideLayoutIdList.AppendChild(new P.SlideLayoutId()
			{
				Id = (uint)(2147483649 + slideLayoutIdList.Count() + 1),
				RelationshipId = relationshipId
			});
		}
		public P.SlideMaster GetSlideMaster()
		{
			return openXMLSlideMaster;
		}
		public string UpdateRelationship(OpenXmlPart openXmlPart, string RelationshipId)
		{
			if (openXMLSlideMaster.SlideMasterPart != null)
			{
				return openXMLSlideMaster.SlideMasterPart.CreateRelationshipToPart(openXmlPart, RelationshipId);
			}
			return null;
		}
		private void CreateSlideMaster()
		{
			slideLayoutIdList = new P.SlideLayoutIdList();
			openXMLSlideMaster = new P.SlideMaster(commonSlideData.GetCommonSlideData());
			openXMLSlideMaster.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			openXMLSlideMaster.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			openXMLSlideMaster.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
			openXMLSlideMaster.AppendChild(new P.ColorMap()
			{
				Accent1 = A.ColorSchemeIndexValues.Accent1,
				Accent2 = A.ColorSchemeIndexValues.Accent2,
				Accent3 = A.ColorSchemeIndexValues.Accent3,
				Accent4 = A.ColorSchemeIndexValues.Accent4,
				Accent5 = A.ColorSchemeIndexValues.Accent5,
				Accent6 = A.ColorSchemeIndexValues.Accent6,
				Background1 = A.ColorSchemeIndexValues.Light1,
				Text1 = A.ColorSchemeIndexValues.Dark1,
				Background2 = A.ColorSchemeIndexValues.Light2,
				Text2 = A.ColorSchemeIndexValues.Dark2,
				Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
				FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
			});
			openXMLSlideMaster.AppendChild(slideLayoutIdList);
			openXMLSlideMaster.AppendChild(CreateTextStyles());
		}
		private static P.TextStyles CreateTextStyles()
		{
			P.TextStyles textStyles = new P.TextStyles();
			P.TitleStyle titleStyle = new P.TitleStyle();
			A.Level1ParagraphProperties titleLevel1ParagraphProperties = new A.Level1ParagraphProperties()
			{
				Alignment = A.TextAlignmentTypeValues.Left,
				DefaultTabSize = 914400,
				EastAsianLineBreak = true,
				LatinLineBreak = false,
				RightToLeft = false
			};
			A.LineSpacing lineSpacing = new A.LineSpacing()
			{
				SpacingPercent = new A.SpacingPercent()
				{
					Val = 90000
				}
			};
			A.SpaceBefore spaceBefore = new A.SpaceBefore()
			{
				SpacingPoints = new A.SpacingPoints()
				{
					Val = 0
				}
			};
			A.NoBullet bulletNone = new A.NoBullet();
			A.DefaultRunProperties titleRunProperties = new A.DefaultRunProperties()
			{
				Kerning = 1200,
				FontSize = 4400
			};
			A.SolidFill solidFill = new A.SolidFill();
			A.SchemeColor schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			solidFill.Append(schemeColor);
			titleRunProperties.Append(solidFill);
			A.LatinFont latinTypeface = new A.LatinFont() { Typeface = "+mj-lt" };
			A.EastAsianFont eastAsianTypeface = new A.EastAsianFont() { Typeface = "+mj-ea" };
			A.ComplexScriptFont complexScriptTypeface = new A.ComplexScriptFont() { Typeface = "+mj-cs" };
			titleRunProperties.Append(latinTypeface, eastAsianTypeface, complexScriptTypeface);
			titleLevel1ParagraphProperties.Append(lineSpacing, spaceBefore, bulletNone, titleRunProperties);
			titleStyle.Append(titleLevel1ParagraphProperties);
			P.BodyStyle bodyStyle = new P.BodyStyle();
			A.Level1ParagraphProperties bodyLevelParagraphProperties = new A.Level1ParagraphProperties()
			{
				Alignment = A.TextAlignmentTypeValues.Left,
				DefaultTabSize = 914400,
				EastAsianLineBreak = true,
				Indent = -228600,
				LatinLineBreak = false,
				RightToLeft = false
			};
			lineSpacing = new A.LineSpacing()
			{
				SpacingPercent = new A.SpacingPercent()
				{
					Val = 90000
				}
			};
			spaceBefore = new A.SpaceBefore()
			{
				SpacingPoints = new A.SpacingPoints()
				{
					Val = 0
				}
			};
			A.BulletFont bulletFont = new A.BulletFont() { CharacterSet = 0, Panose = "020B0604020202020204", PitchFamily = 34, Typeface = "Arial" };
			A.CharacterBullet bulletChar = new A.CharacterBullet() { Char = "•" };
			A.DefaultRunProperties bodyRunProperties = new A.DefaultRunProperties()
			{
				Kerning = 1200,
				FontSize = 2800
			};
			solidFill = new A.SolidFill();
			schemeColor = new A.SchemeColor { Val = A.SchemeColorValues.Text1 };
			solidFill.Append(schemeColor);
			bodyRunProperties.Append(solidFill);
			latinTypeface = new A.LatinFont { Typeface = "+mn-lt" };
			eastAsianTypeface = new A.EastAsianFont { Typeface = "+mn-ea" };
			complexScriptTypeface = new A.ComplexScriptFont { Typeface = "+mn-cs" };
			bodyRunProperties.Append(latinTypeface, eastAsianTypeface, complexScriptTypeface);
			bodyLevelParagraphProperties.Append(lineSpacing, spaceBefore, bulletFont, bulletChar, bodyRunProperties);
			bodyStyle.Append(bodyLevelParagraphProperties);
			P.OtherStyle otherStyle = new P.OtherStyle();
			A.DefaultParagraphProperties otherDefaultParagraphProperties = new A.DefaultParagraphProperties();
			A.DefaultRunProperties otherDefaultRunProperties = new A.DefaultRunProperties() { Language = "en-IN" };
			solidFill = new A.SolidFill();
			schemeColor = new A.SchemeColor { Val = A.SchemeColorValues.Text1 };
			solidFill.Append(schemeColor);
			otherDefaultRunProperties.Append(solidFill);
			latinTypeface = new A.LatinFont { Typeface = "+mn-lt" };
			eastAsianTypeface = new A.EastAsianFont { Typeface = "+mn-ea" };
			complexScriptTypeface = new A.ComplexScriptFont { Typeface = "+mn-cs" };
			otherDefaultRunProperties.Append(latinTypeface, eastAsianTypeface, complexScriptTypeface);
			otherDefaultParagraphProperties.Append(otherDefaultRunProperties);
			otherStyle.Append(otherDefaultParagraphProperties);
			textStyles.Append(titleStyle, bodyStyle, otherStyle);
			return textStyles;
		}
	}
}
