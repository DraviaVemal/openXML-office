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
		private readonly P.Shape openXMLShape = new P.Shape();
		internal Shape(P.Shape shape = null)
		{
			if (shape != null)
			{
				openXMLShape = shape;
			}
		}

		/// <summary>
		/// Remove Found Shape
		/// </summary>
		public void RemoveShape()
		{
			openXMLShape.Remove();
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
			DocumentFormat.OpenXml.OpenXmlElement parent = openXMLShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (openXMLShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				chart.UpdateSize((uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cx), (uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cy));
				chart.UpdatePosition((uint)ConverterUtils.EmuToPixels(oldTransform.Offset.X), (uint)ConverterUtils.EmuToPixels(oldTransform.Offset.Y));
			}
			if (chart.GetChartGraphicFrame().Parent == null)
			{
				var unused = parent.InsertBefore(chart.GetChartGraphicFrame(), openXMLShape);
			}
			openXMLShape.Remove();
			return chart;
		}

		/// <summary>
		/// Replace 2016 Support Chart for the source Shape
		/// </summary>
		public P16.Chart ReplaceChart(P16.Chart chart)
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = openXMLShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (openXMLShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				chart.UpdateSize((uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cx), (uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cy));
				chart.UpdatePosition((uint)ConverterUtils.EmuToPixels(oldTransform.Offset.X), (uint)ConverterUtils.EmuToPixels(oldTransform.Offset.Y));
			}
			if (chart.GetAlternateContent().Parent == null)
			{
				var unused = parent.InsertBefore(chart.GetAlternateContent(), openXMLShape);
			}
			openXMLShape.Remove();
			return chart;
		}

		/// <summary>
		/// Replace Picture for the source Shape
		/// </summary>
		public Picture ReplacePicture(Picture picture)
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = openXMLShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (openXMLShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				picture.UpdateSize((uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cx), (uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cy));
				picture.UpdatePosition((uint)ConverterUtils.EmuToPixels(oldTransform.Offset.X), (uint)ConverterUtils.EmuToPixels(oldTransform.Offset.Y));
			}
			if (picture.GetPicture().Parent == null)
			{
				var unused = parent.InsertBefore(picture.GetPicture(), openXMLShape);
			}
			openXMLShape.Remove();
			return picture;
		}

		/// <summary>
		/// Replace Table for the source Shape
		/// </summary>
		public Table ReplaceTable(Table table)
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = openXMLShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (openXMLShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				table.UpdateSize((uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cx), (uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cy));
				table.UpdatePosition((uint)ConverterUtils.EmuToPixels(oldTransform.Offset.X), (uint)ConverterUtils.EmuToPixels(oldTransform.Offset.Y));
			}
			if (table.GetTableGraphicFrame().Parent == null)
			{
				var unused = parent.InsertBefore(table.GetTableGraphicFrame(), openXMLShape);
			}
			openXMLShape.Remove();
			return table;
		}

		/// <summary>
		/// Replace Text box for the source Shape
		/// </summary>
		public TextBox ReplaceTextBox(TextBox textBox)
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = openXMLShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (openXMLShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				textBox.UpdateSize((uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cx), (uint)ConverterUtils.EmuToPixels(oldTransform.Extents.Cy));
				textBox.UpdatePosition((uint)ConverterUtils.EmuToPixels(oldTransform.Offset.X), (uint)ConverterUtils.EmuToPixels(oldTransform.Offset.Y));
				if (openXMLShape.ShapeStyle != null)
				{
					P.ShapeStyle ShapeStyle = (P.ShapeStyle)openXMLShape.ShapeStyle.Clone();
					textBox.UpdateShapeStyle(ShapeStyle);
				}
			}
			if (textBox.GetTextBoxShape().Parent == null)
			{
				var unused = parent.InsertBefore(textBox.GetTextBoxShape(), openXMLShape);
			}
			openXMLShape.Remove();
			return textBox;
		}

		/// <summary>
		/// Update Shape Text without changing any other properties
		/// </summary>
		public void UpdateShape(ShapeTextModel shapeTextModel)
		{
			if (openXMLShape.TextBody != null)
			{
				A.Paragraph paragraph = openXMLShape.TextBody.GetFirstChild<A.Paragraph>();
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
