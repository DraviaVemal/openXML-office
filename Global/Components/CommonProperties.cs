// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Common Properties organised in one place to get inherited by child classes
    /// </summary>
    public class CommonProperties
    {
        #region Protected Constructors

        /// <summary>
        /// Class is only for inheritance purposes.
        /// </summary>
        protected CommonProperties() { }

        #endregion Protected Constructors

        #region Protected Methods

        /// <summary>
        /// Create Soild Fill XML Property
        /// </summary>
        protected static A.SolidFill CreateSolidFill(SolidFillModel solidFillModel)
        {
            if (solidFillModel.hexColor == null && solidFillModel.schemeColorModel == null)
            {
                throw new ArgumentException("Solid Fill Color Error");
            }
            if (solidFillModel.hexColor != null)
            {
                return new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = solidFillModel.hexColor } };
            }
            else
            {
                A.SchemeColor schemeColor = new()
                { Val = new A.SchemeColorValues(SolidFillModel.GetSchemeColorValuesText(solidFillModel.schemeColorModel!.themeColorValues)) };
                if (solidFillModel.schemeColorModel.tint != null)
                {
                    schemeColor.Append(new A.Tint() { Val = solidFillModel.schemeColorModel.tint });
                }
                if (solidFillModel.schemeColorModel.shade != null)
                {
                    schemeColor.Append(new A.Shade() { Val = solidFillModel.schemeColorModel.shade });
                }
                if (solidFillModel.schemeColorModel.saturationModulation != null)
                {
                    schemeColor.Append(new A.SaturationModulation() { Val = solidFillModel.schemeColorModel.saturationModulation });
                }
                if (solidFillModel.schemeColorModel.saturationOffset != null)
                {
                    schemeColor.Append(new A.SaturationOffset() { Val = solidFillModel.schemeColorModel.saturationOffset });
                }
                if (solidFillModel.schemeColorModel.luminanceModulation != null)
                {
                    schemeColor.Append(new A.LuminanceModulation() { Val = solidFillModel.schemeColorModel.luminanceModulation });
                }
                if (solidFillModel.schemeColorModel.luminanceOffset != null)
                {
                    schemeColor.Append(new A.LuminanceOffset() { Val = solidFillModel.schemeColorModel.luminanceOffset });
                }
                return new A.SolidFill(schemeColor);
            }
        }

        /// <summary>
        /// Create Shape Properties With Default Settings
        /// </summary>
        /// <returns></returns>
        protected C.ShapeProperties CreateChartShapeProperties()
        {
            return CreateChartShapeProperties(new());
        }

        /// <summary>
        /// Create Shape Properties
        /// </summary>
        /// <returns></returns>
        protected C.ShapeProperties CreateChartShapeProperties(ShapePropertiesModel shapePropertiesModel)
        {
            C.ShapeProperties shapeProperties = new();
            if (shapePropertiesModel.solidFill != null)
            {
                shapeProperties.Append(CreateSolidFill(shapePropertiesModel.solidFill));
            }
            else
            {
                shapeProperties.Append(new A.NoFill());
            }
            shapeProperties.Append(CreateOutline(shapePropertiesModel.outline));
            if (shapePropertiesModel.effectList != null)
            {
                shapeProperties.Append(CreateEffectList(shapePropertiesModel.effectList));
            }
            else
            {
                shapeProperties.Append(new A.EffectList());
            }
            return shapeProperties;
        }

        /// <summary>
        /// Create Effect List
        /// </summary>
        protected static A.EffectList CreateEffectList(EffectListModel effectListModel)
        {
            return new();
        }

        /// <summary>
        /// Create Outline
        /// </summary>
        protected A.Outline CreateOutline(OutlineModel outlineModel)
        {
            A.Outline outline = new();
            if (outlineModel.solidFill != null)
            {
                outline.Append(CreateSolidFill(outlineModel.solidFill));
            }
            else
            {
                outline.Append(new A.NoFill());
            }
            if (outlineModel.width != null)
            {
                outline.Width = outlineModel.width;
            }
            if (outlineModel.outlineCapTypeValues != null)
            {
                outline.CapType = OutlineModel.GetLineCapValues((OutlineCapTypeValues)outlineModel.outlineCapTypeValues);
            }
            if (outlineModel.outlineLineTypeValues != null)
            {
                outline.CompoundLineType = OutlineModel.GetLineTypeValues((OutlineLineTypeValues)outlineModel.outlineLineTypeValues);
            }
            if (outlineModel.outlineAlignmentValues != null)
            {
                outline.Alignment = OutlineModel.GetLineAlignmentValues((OutlineAlignmentValues)outlineModel.outlineAlignmentValues);
            }
            return outline;
        }
        /// <summary>
        /// Create Default Run Properties
        /// </summary>
        protected A.DefaultRunProperties CreateDefaultRunProperties()
        {
            return CreateDefaultRunProperties(new());
        }
        /// <summary>
        ///     Create Default Run Properties
        /// </summary>
        protected A.DefaultRunProperties CreateDefaultRunProperties(DefaultRunPropertiesModel defaultRunPropertiesModel)
        {
            A.DefaultRunProperties DefaultRunProperties = new();
            if (defaultRunPropertiesModel.solidFill != null)
            {
                DefaultRunProperties.Append(CreateSolidFill(defaultRunPropertiesModel.solidFill));
            }
            if (defaultRunPropertiesModel.latinFont != null)
            {
                DefaultRunProperties.Append(new A.LatinFont { Typeface = defaultRunPropertiesModel.latinFont });
            }
            if (defaultRunPropertiesModel.eastAsianFont != null)
            {
                DefaultRunProperties.Append(new A.EastAsianFont { Typeface = defaultRunPropertiesModel.eastAsianFont });
            }
            if (defaultRunPropertiesModel.complexScriptFont != null)
            {
                DefaultRunProperties.Append(new A.ComplexScriptFont { Typeface = defaultRunPropertiesModel.complexScriptFont });
            }
            if (defaultRunPropertiesModel.fontSize != null)
            {
                DefaultRunProperties.FontSize = defaultRunPropertiesModel.fontSize;
            }
            if (defaultRunPropertiesModel.bold != null)
            {
                DefaultRunProperties.Bold = defaultRunPropertiesModel.bold;
            }
            if (defaultRunPropertiesModel.italic != null)
            {
                DefaultRunProperties.Italic = defaultRunPropertiesModel.italic;
            }
            if (defaultRunPropertiesModel.underline != null)
            {
                DefaultRunProperties.Underline = DefaultRunPropertiesModel.GetTextUnderlineValues((UnderLineValues)defaultRunPropertiesModel.underline);
            }
            if (defaultRunPropertiesModel.strike != null)
            {
                DefaultRunProperties.Strike = DefaultRunPropertiesModel.GetTextStrikeValues((StrikeValues)defaultRunPropertiesModel.strike);
            }
            if (defaultRunPropertiesModel.kerning != null)
            {
                DefaultRunProperties.Kerning = defaultRunPropertiesModel.kerning;
            }
            if (defaultRunPropertiesModel.baseline != null)
            {
                DefaultRunProperties.Baseline = defaultRunPropertiesModel.baseline;
            }
            return DefaultRunProperties;
        }

        /// <summary>
        /// 
        /// </summary>
        protected A.Paragraph CreateDrawingParagraph(DrawingParagraphModel drawingParagraphModel)
        {
            A.Paragraph paragraph = new();
            if (drawingParagraphModel.paragraphPropertiesModel != null)
            {
                paragraph.Append(
                    CreateDrawingParagraphProperties(drawingParagraphModel.paragraphPropertiesModel),
                    new A.EndParagraphRunProperties() { Language = "en-US" });
            }
            return paragraph;
        }

        /// <summary>
        /// 
        /// </summary>
        private A.ParagraphProperties CreateDrawingParagraphProperties(ParagraphPropertiesModel paragraphPropertiesModel)
        {
            A.ParagraphProperties paragraphProperties = new();
            if (paragraphPropertiesModel.defaultRunProperties != null)
            {
                paragraphProperties.Append(CreateDefaultRunProperties(paragraphPropertiesModel.defaultRunProperties));
            }
            return paragraphProperties;
        }

        /// <summary>
        /// 
        /// </summary>
        protected static A.ListStyle CreateDrawingListStyle()
        {
            return new();
        }
        /// <summary>
        ///     Create Chart Text Properties
        /// </summary>
        protected C.TextProperties CreateChartTextProperties(ChartTextPropertiesModel chartTextPropertiesModel)
        {
            C.TextProperties textProperties = new();
            if (chartTextPropertiesModel.bodyProperties != null)
            {
                textProperties.Append(CreateDrawingBodyProperties(chartTextPropertiesModel.bodyProperties));
            }
            return textProperties;
        }
        /// <summary>
        ///    Create Drawing Body Properties
        /// </summary>
        /// <param name="drawingBodyPropertiesModel"></param>
        /// <returns></returns>
        private static A.BodyProperties CreateDrawingBodyProperties(DrawingBodyPropertiesModel drawingBodyPropertiesModel)
        {
            A.BodyProperties bodyProperties = new(new A.ShapeAutoFit())
            {
                Rotation = drawingBodyPropertiesModel.rotation
            };
            if (drawingBodyPropertiesModel.leftInset != null)
            {
                bodyProperties.LeftInset = drawingBodyPropertiesModel.leftInset;
            }
            if (drawingBodyPropertiesModel.topInset != null)
            {
                bodyProperties.TopInset = drawingBodyPropertiesModel.topInset;
            }
            if (drawingBodyPropertiesModel.rightInset != null)
            {
                bodyProperties.RightInset = drawingBodyPropertiesModel.rightInset;
            }
            if (drawingBodyPropertiesModel.bottomInset != null)
            {
                bodyProperties.BottomInset = drawingBodyPropertiesModel.bottomInset;
            }
            if (drawingBodyPropertiesModel.useParagraphSpacing != null)
            {
                bodyProperties.UseParagraphSpacing = drawingBodyPropertiesModel.useParagraphSpacing;
            }
            if (drawingBodyPropertiesModel.verticalOverflow != null)
            {
                bodyProperties.VerticalOverflow = DrawingBodyPropertiesModel.GetTextVerticalOverflowValues((TextVerticalOverflowValues)drawingBodyPropertiesModel.verticalOverflow);
            }
            if (drawingBodyPropertiesModel.vertical != null)
            {
                bodyProperties.Vertical = DrawingBodyPropertiesModel.GetTextVerticalAlignmentValues((TextVerticalAlignmentValues)drawingBodyPropertiesModel.vertical);
            }
            if (drawingBodyPropertiesModel.wrap != null)
            {
                bodyProperties.Wrap = DrawingBodyPropertiesModel.GetWrapingValues((TextWrappingValues)drawingBodyPropertiesModel.wrap);
            }
            if (drawingBodyPropertiesModel.anchor != null)
            {
                bodyProperties.Anchor = DrawingBodyPropertiesModel.GetAnchorValues((TextAnchoringValues)drawingBodyPropertiesModel.anchor);
            }
            return bodyProperties;
        }

        #endregion Protected Methods
    }
}