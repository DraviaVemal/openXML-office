// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents Textbox base class to build on
	/// </summary>
	public class TextBoxBase : CommonProperties
	{
		private readonly TextBoxSetting textBoxSetting;
		private P.Shape? openXMLShape;

		/// <summary>
		/// Create Textbox with provided settings
		/// </summary>
		public TextBoxBase(TextBoxSetting TextBoxSetting)
		{
			textBoxSetting = TextBoxSetting;
			CreateTextBox();
		}

		/// <summary>
		/// Get Textbox Shape
		/// </summary>
		public P.Shape GetTextBoxBaseShape()
		{
			return openXMLShape!;
		}

		/// <summary>
		/// Update Textbox Position
		/// </summary>
		public void UpdatePosition(uint X, uint Y)
		{
			textBoxSetting.x = X;
			textBoxSetting.y = Y;
			if (openXMLShape != null)
			{
				openXMLShape.ShapeProperties!.Transform2D = new A.Transform2D
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
			textBoxSetting.width = Width;
			textBoxSetting.height = Height;
			if (openXMLShape != null)
			{
				openXMLShape.ShapeProperties!.Transform2D = new A.Transform2D
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
			GetTextBoxBaseShape().ShapeStyle = shapeStyle;
		}

		private P.Shape CreateTextBox()
		{
			SolidFillModel solidFillModel = new()
			{
				schemeColorModel = new()
				{
					themeColorValues = ThemeColorValues.TEXT_1
				}
			};
			if (textBoxSetting.textColor != null)
			{
				solidFillModel.hexColor = textBoxSetting.textColor;
				solidFillModel.schemeColorModel = null;
			}
			openXMLShape = new()
			{
				NonVisualShapeProperties = new P.NonVisualShapeProperties(
				new P.NonVisualDrawingProperties()
				{
					Id = 10,
					Name = "Text Box"
				},
				new P.NonVisualShapeDrawingProperties(),
				new P.ApplicationNonVisualDrawingProperties()),
				ShapeProperties = new P.ShapeProperties(
				new A.Transform2D(
					new A.Offset { X = textBoxSetting.x, Y = textBoxSetting.y },
					new A.Extents { Cx = textBoxSetting.width, Cy = textBoxSetting.height }),
				new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
				textBoxSetting.shapeBackground != null ? CreateSolidFill(new() { hexColor = textBoxSetting.shapeBackground }) : new A.NoFill()),
				TextBody = new P.TextBody(
						new A.BodyProperties(),
						new A.ListStyle(),
						CreateDrawingParagraph(new()
						{
							paragraphPropertiesModel = new()
							{
								horizontalAlignment = textBoxSetting.horizontalAlignment
							},
							drawingRun = new()
							{
								text = textBoxSetting.text,
								textHightlight = textBoxSetting.textBackground,
								drawingRunProperties = new()
								{
									solidFill = solidFillModel,
									fontFamily = textBoxSetting.fontFamily,
									fontSize = textBoxSetting.fontSize,
									isBold = textBoxSetting.isBold,
									isItalic = textBoxSetting.isItalic,
									underline = textBoxSetting.isUnderline ? UnderLineValues.SINGLE : UnderLineValues.NONE,
								}
							}
						})),
			};
			return openXMLShape;
		}

	}
}
