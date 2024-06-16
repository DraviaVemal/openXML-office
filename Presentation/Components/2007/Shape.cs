// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.Collections.Generic;
using OpenXMLOffice.Global_2007;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P16 = OpenXMLOffice.Presentation_2016;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Shape Class For Presentation shape manipulation
	/// </summary>
	public class Shape : PresentationCommonProperties
	{
		private readonly P.Shape documentShape = new P.Shape();
		internal Shape(P.Shape shape = null)
		{
			if (shape != null)
			{
				documentShape = shape;
			}
		}

		/// <summary>
		/// Remove Found Shape
		/// </summary>
		public void RemoveShape()
		{
			documentShape.Remove();
		}

		internal P.Shape GetDocumentShape()
		{
			return documentShape;
		}

		internal Shape AddLine<LineColorOption>(LineShapeModel<PresentationSetting, LineColorOption> lineModel)
			where LineColorOption : class, IColorOptions, new()
		{
			return this;
		}

		internal Shape AddRectangle<LineColorOption, FillColorOption>(RectangleShapeModel<PresentationSetting, LineColorOption, FillColorOption> rectangleModel)
			where LineColorOption : class, IColorOptions, new()
			where FillColorOption : class, IColorOptions, new()
		{
			return this;
		}

		internal Shape AddArrow<LineColorOption, FillColorOption>(ArrowShapeModel<PresentationSetting, LineColorOption, FillColorOption> arrowModel)
			where LineColorOption : class, IColorOptions, new()
			where FillColorOption : class, IColorOptions, new()
		{
			return this;
		}

		/// <summary>
		/// Replace Chart for the source Shape
		/// </summary>
		public Chart<XAxisType, YAxisType, ZAxisType> ReplaceChart<XAxisType, YAxisType, ZAxisType>(Chart<XAxisType, YAxisType, ZAxisType> chart)
			where XAxisType : class, IAxisTypeOptions, new()
			where YAxisType : class, IAxisTypeOptions, new()
			where ZAxisType : class, IAxisTypeOptions, new()
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = documentShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (documentShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = documentShape.ShapeProperties.Transform2D;
				chart.UpdateSize((uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cx), (uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cy));
				chart.UpdatePosition((uint)ConverterUtils.EmuToPixels(oldTransform.Offset.X), (uint)ConverterUtils.EmuToPixels(oldTransform.Offset.Y));
			}
			if (chart.GetChartGraphicFrame().Parent == null)
			{
				parent.InsertBefore(chart.GetChartGraphicFrame(), documentShape);
			}
			documentShape.Remove();
			return chart;
		}

		/// <summary>
		/// Replace 2016 Support Chart for the source Shape
		/// </summary>
		public P16.Chart ReplaceChart(P16.Chart chart)
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = documentShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (documentShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = documentShape.ShapeProperties.Transform2D;
				chart.UpdateSize((uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cx), (uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cy));
				chart.UpdatePosition((uint)ConverterUtils.EmuToPixels(oldTransform.Offset.X), (uint)ConverterUtils.EmuToPixels(oldTransform.Offset.Y));
			}
			if (chart.GetAlternateContent().Parent == null)
			{
				parent.InsertBefore(chart.GetAlternateContent(), documentShape);
			}
			documentShape.Remove();
			return chart;
		}

		/// <summary>
		/// Replace Picture for the source Shape
		/// </summary>
		public Picture ReplacePicture(Picture picture)
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = documentShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (documentShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = documentShape.ShapeProperties.Transform2D;
				picture.UpdateSize((uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cx), (uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cy));
				picture.UpdatePosition((uint)ConverterUtils.EmuToPixels(oldTransform.Offset.X), (uint)ConverterUtils.EmuToPixels(oldTransform.Offset.Y));
			}
			if (picture.GetPicture().Parent == null)
			{
				parent.InsertBefore(picture.GetPicture(), documentShape);
			}
			documentShape.Remove();
			return picture;
		}

		/// <summary>
		/// Replace Table for the source Shape
		/// </summary>
		public Table ReplaceTable(Table table)
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = documentShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (documentShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = documentShape.ShapeProperties.Transform2D;
				table.UpdateSize((uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cx), (uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cy));
				table.UpdatePosition((uint)ConverterUtils.EmuToPixels(oldTransform.Offset.X), (uint)ConverterUtils.EmuToPixels(oldTransform.Offset.Y));
			}
			if (table.GetTableGraphicFrame().Parent == null)
			{
				parent.InsertBefore(table.GetTableGraphicFrame(), documentShape);
			}
			documentShape.Remove();
			return table;
		}

		/// <summary>
		/// Replace Text box for the source Shape
		/// </summary>
		public TextBox ReplaceTextBox(TextBox textBox)
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = documentShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (documentShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = documentShape.ShapeProperties.Transform2D;
				textBox.UpdateSize((uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cx), (uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cy));
				textBox.UpdatePosition((uint)ConverterUtils.EmuToPixels(oldTransform.Offset.X), (uint)ConverterUtils.EmuToPixels(oldTransform.Offset.Y));
				if (documentShape.ShapeStyle != null)
				{
					P.ShapeStyle ShapeStyle = (P.ShapeStyle)documentShape.ShapeStyle.Clone();
					textBox.UpdateShapeStyle(ShapeStyle);
				}
			}
			if (textBox.GetTextBoxShape().Parent == null)
			{
				parent.InsertBefore(textBox.GetTextBoxShape(), documentShape);
			}
			documentShape.Remove();
			return textBox;
		}

		/// <summary>
		/// Update Shape Text without changing any other properties
		/// </summary>
		public void UpdateShape(ShapeTextModel shapeTextModel)
		{
			if (documentShape.TextBody != null)
			{
				A.Paragraph paragraph = documentShape.TextBody.GetFirstChild<A.Paragraph>();
				if (paragraph != null)
				{
					paragraph.RemoveAllChildren<A.Run>();
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
					if (shapeTextModel.fontColor != null)
					{
						textColorOption.colorOption.hexColor = shapeTextModel.fontColor;
						textColorOption.colorOption.schemeColorModel = null;
					}
					paragraph.Append(CreateDrawingRun(new List<DrawingRunModel<SolidOptions>>()
					{
						new DrawingRunModel<SolidOptions>(){
							text = shapeTextModel.textValue,
						drawingRunProperties = new DrawingRunPropertiesModel<SolidOptions>()
						{
							textColorOption = textColorOption,
							fontFamily = shapeTextModel.fontFamily,
							fontSize = shapeTextModel.fontSize,
							isBold = shapeTextModel.isBold,
							isItalic = shapeTextModel.isItalic,
							underLineValues = shapeTextModel.underLineValues
						}
						}
					}.ToArray()));
				}
			}
		}
	}
}
