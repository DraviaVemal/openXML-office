// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.Collections.Generic;
using OpenXMLOffice.Global_2007;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Represents Textbox class to build on
	/// </summary>
	public class TextBox : CommonProperties
	{
		private readonly TextBoxSetting textBoxSetting;
		private P.Shape openXMLShape;
		private readonly Slide slide;
		/// <summary>
		/// Create Textbox with provided settings
		/// </summary>
		internal TextBox(TextBoxSetting TextBoxSetting)
		{
			textBoxSetting = TextBoxSetting;
			CreateTextBox();
		}
		/// <summary>
		/// Create Textbox with provided settings
		/// </summary>
		public TextBox(Slide Slide, TextBoxSetting TextBoxSetting)
		{
			slide = Slide;
			textBoxSetting = TextBoxSetting;
			CreateTextBox();
			slide.GetSlide().CommonSlideData.ShapeTree.Append(GetTextBoxShape());
		}
		/// <summary>
		/// Get Textbox Shape
		/// </summary>
		internal P.Shape GetTextBoxShape()
		{
			return openXMLShape;
		}
		/// <summary>
		/// Update Textbox Position
		/// </summary>
		public void UpdatePosition(uint X, uint Y)
		{
			textBoxSetting.x = (uint)ConverterUtils.PixelsToEmu((int)X);
			textBoxSetting.y = (uint)ConverterUtils.PixelsToEmu((int)Y);
			if (openXMLShape != null)
			{
				openXMLShape.ShapeProperties.Transform2D = new A.Transform2D
				{
					Offset = new A.Offset { X = textBoxSetting.x, Y = textBoxSetting.y },
					Extents = new A.Extents { Cx = textBoxSetting.width, Cy = textBoxSetting.height }
				};
			}
		}
		/// <summary>
		/// Update Textbox Size
		/// </summary>
		public void UpdateSize(uint Width, uint Height)
		{
			textBoxSetting.width = (uint)ConverterUtils.PixelsToEmu((int)Width);
			textBoxSetting.height = (uint)ConverterUtils.PixelsToEmu((int)Height);
			if (openXMLShape != null)
			{
				openXMLShape.ShapeProperties.Transform2D = new A.Transform2D
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
				ShapeProperties.Append(CreateSolidFill(new SolidFillModel() { hexColor = textBoxSetting.shapeBackground }));
			}
			else
			{
				ShapeProperties.Append(new A.NoFill());
			}
			List<DrawingRunModel> drawingRunModels = new List<DrawingRunModel>();
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
							slide.GetSlidePart().AddHyperlinkRelationship(new System.Uri(textBlock.hyperlinkProperties.value), true, relationId);
							break;
						case HyperlinkPropertyTypeValues.TARGET_SLIDE:
							textBlock.hyperlinkProperties.relationId = relationId;
							textBlock.hyperlinkProperties.action = "ppaction://hlinksldjump";
							//TODO: Update Target Slide Prop
							slide.GetSlidePart().AddHyperlinkRelationship(new System.Uri(textBlock.hyperlinkProperties.value), true, relationId);
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
							slide.GetSlidePart().AddHyperlinkRelationship(new System.Uri(textBlock.hyperlinkProperties.value), true, relationId);
							break;
					}
				}
				SolidFillModel solidFillModel = new SolidFillModel()
				{
					schemeColorModel = new SchemeColorModel()
					{
						themeColorValues = ThemeColorValues.TEXT_1
					}
				};
				if (textBlock.textColor != null)
				{
					solidFillModel.hexColor = textBlock.textColor;
					solidFillModel.schemeColorModel = null;
				}
				DrawingRunModel drawingRunModel = new DrawingRunModel()
				{
					text = textBlock.text,
					textHightlight = textBlock.textBackground,
					drawingRunProperties = new DrawingRunPropertiesModel()
					{
						solidFill = solidFillModel,
						fontFamily = textBlock.fontFamily,
						fontSize = textBlock.fontSize,
						isBold = textBlock.isBold,
						isItalic = textBlock.isItalic,
						underline = textBlock.isUnderline ? UnderLineValues.SINGLE : UnderLineValues.NONE,
						hyperlinkProperties = textBlock.hyperlinkProperties
					}
				};
				drawingRunModels.Add(drawingRunModel);
			}
			openXMLShape = new P.Shape()
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
				TextBody = new P.TextBody(
						new A.BodyProperties(),
						new A.ListStyle(),
						CreateDrawingParagraph(new DrawingParagraphModel()
						{
							paragraphPropertiesModel = new ParagraphPropertiesModel()
							{
								horizontalAlignment = textBoxSetting.horizontalAlignment
							},
							drawingRuns = drawingRunModels.ToArray()
						})),
			};
			return openXMLShape;
		}
	}
}
