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
		/// Get Textbox Run
		/// </summary>
		public A.Run GetTextBoxBaseRun()
		{
			return CreateTextRun();
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
				textBoxSetting.shapeBackground != null ? CommonProperties.CreateSolidFill(new() { hexColor = textBoxSetting.shapeBackground }) : new A.NoFill()),
				TextBody = new P.TextBody(
						new A.BodyProperties(),
						new A.ListStyle(),
						new A.Paragraph(CreateTextRun())),
			};
			return openXMLShape;
		}

		private A.Run CreateTextRun()
		{
			A.Run Run = new(new A.RunProperties(CreateSolidFill(new() { hexColor = textBoxSetting.textColor }),
						new A.LatinFont { Typeface = textBoxSetting.fontFamily },
						new A.EastAsianFont { Typeface = textBoxSetting.fontFamily },
						new A.ComplexScriptFont { Typeface = textBoxSetting.fontFamily })
			{
				FontSize = ConverterUtils.FontSizeToFontSize(textBoxSetting.fontSize),
				Bold = textBoxSetting.isBold,
				Italic = textBoxSetting.isItalic,
				Underline = textBoxSetting.isUnderline ? A.TextUnderlineValues.Single : A.TextUnderlineValues.None,
				Dirty = false
			}, new A.Text(textBoxSetting.text));
			if (textBoxSetting.textBackground != null)
			{
				Run.Append(new A.Highlight(new A.RgbColorModelHex { Val = textBoxSetting.textBackground }));
			}
			return Run;
		}


	}
}
