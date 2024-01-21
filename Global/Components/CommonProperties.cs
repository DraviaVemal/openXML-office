// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

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
        protected A.SolidFill CreateSolidFill(SolidFillModel solidFillModel)
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
                { Val = new A.SchemeColorValues(solidFillModel.GetSchemeColorValuesText(solidFillModel.schemeColorModel!.themeColorValues)) };
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
        /// Create Paragraph
        /// </summary>
        protected A.Paragraph CreateParagraph()
        {
            return new();
        }
        /// <summary>
        /// Create Effect List
        /// </summary>
        protected A.EffectList CreateEffectList(EffectListModel effectListModel)
        {
            return new();
        }
        /// <summary>
        /// Create Text Properties
        /// </summary>
        protected C.TextProperties CreateTextProperties()
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
                outline.CapType = outlineModel.GetLineCapValues((OutlineCapTypeValues)outlineModel.outlineCapTypeValues);
            }
            if (outlineModel.outlineLineTypeValues != null)
            {
                outline.CompoundLineType = outlineModel.GetLineTypeValues((OutlineLineTypeValues)outlineModel.outlineLineTypeValues);
            }
            if (outlineModel.outlineAlignmentValues != null)
            {
                outline.Alignment = outlineModel.GetLineAlignmentValues((OutlineAlignmentValues)outlineModel.outlineAlignmentValues);
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
                DefaultRunProperties.Underline = defaultRunPropertiesModel.GetTextUnderlineValues((UnderLineValues)defaultRunPropertiesModel.underline);
            }
            if (defaultRunPropertiesModel.strike != null)
            {
                DefaultRunProperties.Strike = defaultRunPropertiesModel.GetTextStrikeValues((StrikeValues)defaultRunPropertiesModel.strike);
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

        #endregion Protected Methods
    }
}