// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OpenXMLOffice.Global_2007;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Represents TextBox class to build on
	/// </summary>
	public class TextBox : PresentationCommonProperties
	{
		private readonly TextBoxSetting textBoxSetting;
		private P.Shape documentShape;
		private readonly Slide slide;
		/// <summary>
		/// Create TextBox with provided settings
		/// </summary>
		internal TextBox(TextBoxSetting TextBoxSetting)
		{
			textBoxSetting = TextBoxSetting;
			CreateTextBox();
		}
		/// <summary>
		/// Create TextBox with provided settings
		/// </summary>
		public TextBox(Slide Slide, TextBoxSetting TextBoxSetting)
		{
			slide = Slide;
			textBoxSetting = TextBoxSetting;
			CreateTextBox();
			slide.GetSlide().CommonSlideData.ShapeTree.Append(GetTextBoxShape());
		}
		/// <summary>
		/// Get TextBox Shape
		/// </summary>
		internal P.Shape GetTextBoxShape()
		{
			return documentShape;
		}
		/// <summary>
		/// Update TextBox Position
		/// </summary>
		public void UpdatePosition(uint X, uint Y)
		{
			textBoxSetting.x = (int)ConverterUtils.PixelsToEmu((int)X);
			textBoxSetting.y = (int)ConverterUtils.PixelsToEmu((int)Y);
			if (documentShape != null)
			{
				documentShape.ShapeProperties.Transform2D = new A.Transform2D
				{
					Offset = new A.Offset { X = textBoxSetting.x, Y = textBoxSetting.y },
					Extents = new A.Extents { Cx = textBoxSetting.width, Cy = textBoxSetting.height }
				};
			}
		}
		/// <summary>
		/// Update TextBox Size
		/// </summary>
		public void UpdateSize(uint Width, uint Height)
		{
			textBoxSetting.width = (int)ConverterUtils.PixelsToEmu((int)Width);
			textBoxSetting.height = (int)ConverterUtils.PixelsToEmu((int)Height);
			if (documentShape != null)
			{
				documentShape.ShapeProperties.Transform2D = new A.Transform2D
				{
					Offset = new A.Offset { X = textBoxSetting.x, Y = textBoxSetting.y },
					Extents = new A.Extents { Cx = textBoxSetting.width, Cy = textBoxSetting.height }
				};
			}
		}
		/// <summary>
		///
		/// </summary>
		public void UpdateShapeStyle(P.ShapeStyle shapeStyle)
		{
			GetTextBoxShape().ShapeStyle = shapeStyle;
		}
		private P.Shape CreateTextBox()
		{
			P.ShapeProperties ShapeProperties = new P.ShapeProperties(
							new A.Transform2D(
								new A.Offset { X = textBoxSetting.x, Y = textBoxSetting.y },
								new A.Extents { Cx = textBoxSetting.width, Cy = textBoxSetting.height }),
							new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

			if (textBoxSetting.shapeBackground != null)
			{
				ShapeProperties.Append(CreateColorComponent(new ColorOptionModel<SolidOptions>()
				{
					colorOption = new SolidOptions()
					{
						hexColor = textBoxSetting.shapeBackground
					}
				}));
			}
			else
			{
				ShapeProperties.Append(CreateColorComponent(new ColorOptionModel<NoFillOptions>()));
			}
			List<DrawingRunModel<SolidOptions>> drawingRunModels = new List<DrawingRunModel<SolidOptions>>();
			List<DrawingParagraphModel<SolidOptions>> paragraphsModels = new List<DrawingParagraphModel<SolidOptions>>();
			DrawingParagraphModel<SolidOptions> drawingParagraphModel = new DrawingParagraphModel<SolidOptions>();
			foreach (TextBlock textBlock in textBoxSetting.textBlocks)
			{
				// Add Hyperlink Relationships to slide
				if (textBlock.hyperlinkProperties != null)
				{
					string relationId = slide.GetNextSlideRelationId();
					switch (textBlock.hyperlinkProperties.hyperlinkPropertyType)
					{
						case HyperlinkPropertyTypeValues.EXISTING_FILE:
							textBlock.hyperlinkProperties.relationId = relationId;
							textBlock.hyperlinkProperties.action = "ppaction://hlinkfile";
							slide.GetSlidePart().AddHyperlinkRelationship(new Uri(textBlock.hyperlinkProperties.value), true, relationId);
							break;
						case HyperlinkPropertyTypeValues.TARGET_SLIDE:
							textBlock.hyperlinkProperties.relationId = relationId;
							textBlock.hyperlinkProperties.action = "ppaction://hlinksldjump";
							//TODO: Update Target Slide Prop
							slide.GetSlidePart().AddHyperlinkRelationship(new Uri(textBlock.hyperlinkProperties.value), true, relationId);
							break;
						case HyperlinkPropertyTypeValues.TARGET_SHEET:
							throw new ArgumentException("This Option is valid only for Excel Files");
						case HyperlinkPropertyTypeValues.FIRST_SLIDE:
							textBlock.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=firstslide";
							break;
						case HyperlinkPropertyTypeValues.LAST_SLIDE:
							textBlock.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=lastslide";
							break;
						case HyperlinkPropertyTypeValues.NEXT_SLIDE:
							textBlock.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=nextslide";
							break;
						case HyperlinkPropertyTypeValues.PREVIOUS_SLIDE:
							textBlock.hyperlinkProperties.action = "ppaction://hlinkshowjump?jump=previousslide";
							break;
						default:// Web URL
							textBlock.hyperlinkProperties.relationId = relationId;
							slide.GetSlidePart().AddHyperlinkRelationship(new Uri(textBlock.hyperlinkProperties.value), true, relationId);
							break;
					}
				}
				ColorOptionModel<SolidOptions> textColorOption = new ColorOptionModel<SolidOptions>()
				{
					colorOption = new SolidOptions()
					{
						schemeColorModel = new SchemeColorModel()
						{
							themeColorValues = ThemeColorValues.TEXT_1
						}
					}
				};
				if (textBlock.textColor != null)
				{
					textColorOption.colorOption.hexColor = textBlock.textColor;
					textColorOption.colorOption.schemeColorModel = null;
				}
				DrawingRunModel<SolidOptions> drawingRunModel = new DrawingRunModel<SolidOptions>()
				{
					text = textBlock.textValue,
					textHighlight = textBlock.textBackground,
					drawingRunProperties = new DrawingRunPropertiesModel<SolidOptions>()
					{
						textColorOption = textColorOption,
						fontFamily = textBlock.fontFamily,
						fontSize = textBlock.fontSize,
						isBold = textBlock.isBold,
						isItalic = textBlock.isItalic,
						underLineValues = textBlock.isUnderline ? UnderLineValues.SINGLE : UnderLineValues.NONE,
						hyperlinkProperties = textBlock.hyperlinkProperties
					}
				};
				if (textBlock.bulletsAndNumbering != null && textBlock.bulletsAndNumbering != BulletsAndNumberingValues.NONE)
				{
					drawingParagraphModel.drawingRuns = drawingRunModels.ToArray();
					paragraphsModels.Add(drawingParagraphModel);
					drawingParagraphModel = new DrawingParagraphModel<SolidOptions>()
					{
						paragraphPropertiesModel = new ParagraphPropertiesModel<SolidOptions>()
						{
							horizontalAlignment = textBoxSetting.horizontalAlignment,
							bulletsAndNumbering = textBlock.bulletsAndNumbering
						}
					};
					drawingRunModels = new List<DrawingRunModel<SolidOptions>>()
					{
						drawingRunModel
					};
				}
				else if (textBlock.isEndParagraph)
				{
					drawingParagraphModel.drawingRuns = drawingRunModels.ToArray();
					paragraphsModels.Add(drawingParagraphModel);
					drawingRunModels = new List<DrawingRunModel<SolidOptions>>();
					drawingParagraphModel = new DrawingParagraphModel<SolidOptions>()
					{
						paragraphPropertiesModel = new ParagraphPropertiesModel<SolidOptions>()
						{
							horizontalAlignment = textBoxSetting.horizontalAlignment,
							bulletsAndNumbering = textBlock.bulletsAndNumbering
						}
					};
				}
				else
				{
					drawingRunModels.Add(drawingRunModel);
				}
			}
			if (drawingRunModels.Count > 0)
			{
				drawingParagraphModel.drawingRuns = drawingRunModels.ToArray();
				paragraphsModels.Add(drawingParagraphModel);
				drawingParagraphModel = new DrawingParagraphModel<SolidOptions>()
				{
					paragraphPropertiesModel = new ParagraphPropertiesModel<SolidOptions>()
					{
						horizontalAlignment = textBoxSetting.horizontalAlignment
					}
				};
			}
			IEnumerable<OpenXmlElement> drawingParagraphs = Enumerable
				.Repeat((OpenXmlElement)new A.BodyProperties(), 1)
				.Concat(Enumerable.Repeat((OpenXmlElement)new A.ListStyle(), 1))
				.Concat(paragraphsModels.Select(item =>
					{
						return CreateDrawingParagraph(item);
					}));
			documentShape = new P.Shape()
			{
				NonVisualShapeProperties = new P.NonVisualShapeProperties(
				new P.NonVisualDrawingProperties()
				{
					Id = 10,
					Name = "Text Box"
				},
				new P.NonVisualShapeDrawingProperties(),
				new P.ApplicationNonVisualDrawingProperties()),
				ShapeProperties = ShapeProperties,
				TextBody = new P.TextBody(drawingParagraphs),
			};
			return documentShape;
		}
	}
}
