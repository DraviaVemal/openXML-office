/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    internal class SlideMaster
    {
        #region Protected Fields

        protected P.SlideMaster OpenXMLSlideMaster = new();
        protected P.SlideLayoutIdList? slideLayoutIdList;

        #endregion Protected Fields

        #region Private Fields

        private readonly CommonSlideData commonSlideData = new(PresentationConstants.CommonSlideDataType.SLIDE_MASTER, PresentationConstants.SlideLayoutType.BLANK);

        #endregion Private Fields

        #region Public Constructors

        public SlideMaster()
        {
            CreateSlideMaster();
        }

        #endregion Public Constructors

        #region Public Methods

        public void AddSlideLayoutIdToList(string relationshipId)
        {
            slideLayoutIdList!.AppendChild(new P.SlideLayoutId()
            {
                Id = (uint)(2147483649 + slideLayoutIdList.Count() + 1),
                RelationshipId = relationshipId
            });
        }

        public P.SlideMaster GetSlideMaster()
        {
            return OpenXMLSlideMaster;
        }

        public string UpdateRelationship(OpenXmlPart openXmlPart, string RelationshipId)
        {
            return OpenXMLSlideMaster.SlideMasterPart!.CreateRelationshipToPart(openXmlPart, RelationshipId);
        }

        #endregion Public Methods

        #region Private Methods

        private void CreateSlideMaster()
        {
            slideLayoutIdList = new();
            OpenXMLSlideMaster = new(commonSlideData.GetCommonSlideData());
            OpenXMLSlideMaster.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            OpenXMLSlideMaster.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            OpenXMLSlideMaster.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            OpenXMLSlideMaster.AppendChild(new P.ColorMap()
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
            OpenXMLSlideMaster.AppendChild(slideLayoutIdList);
            OpenXMLSlideMaster.AppendChild(CreateTextStyles());
        }

        private P.TextStyles CreateTextStyles()
        {
            P.TextStyles textStyles = new();

            P.TitleStyle titleStyle = new();
            A.Level1ParagraphProperties titleLevel1ParagraphProperties = new()
            {
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                RightToLeft = false
            };
            A.LineSpacing lineSpacing = new()
            {
                SpacingPercent = new A.SpacingPercent()
                {
                    Val = 90000
                }
            };
            A.SpaceBefore spaceBefore = new()
            {
                SpacingPoints = new A.SpacingPoints()
                {
                    Val = 0
                }
            };
            A.NoBullet bulletNone = new();
            A.DefaultRunProperties titleRunProperties = new()
            {
                Kerning = 1200,
                FontSize = 4400
            };
            A.SolidFill solidFill = new();
            A.SchemeColor schemeColor = new() { Val = A.SchemeColorValues.Text1 };
            solidFill.Append(schemeColor);
            titleRunProperties.Append(solidFill);
            A.LatinFont latinTypeface = new() { Typeface = "+mj-lt" };
            A.EastAsianFont eastAsianTypeface = new() { Typeface = "+mj-ea" };
            A.ComplexScriptFont complexScriptTypeface = new() { Typeface = "+mj-cs" };
            titleRunProperties.Append(latinTypeface, eastAsianTypeface, complexScriptTypeface);
            titleLevel1ParagraphProperties.Append(lineSpacing, spaceBefore, bulletNone, titleRunProperties);
            titleStyle.Append(titleLevel1ParagraphProperties);
            P.BodyStyle bodyStyle = new();
            A.Level1ParagraphProperties bodyLevelParagraphProperties = new()
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
            A.BulletFont bulletFont = new() { CharacterSet = 0, Panose = "020B0604020202020204", PitchFamily = 34, Typeface = "Arial" };
            A.CharacterBullet bulletChar = new() { Char = "•" };
            A.DefaultRunProperties bodyRunProperties = new()
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
            P.OtherStyle otherStyle = new();
            A.DefaultParagraphProperties otherDefaultParagraphProperties = new();
            A.DefaultRunProperties otherDefaultRunProperties = new() { Language = "en-US" };
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

        #endregion Private Methods
    }
}