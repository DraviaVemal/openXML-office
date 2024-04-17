// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System;
using OpenXMLOffice.Global_2007;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P16 = OpenXMLOffice.Presentation_2016;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Shape Class For Presentation shape manipulation
	/// </summary>
	public class Shape : CommonProperties
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
		/// <summary>
		/// Replace Chart for the source Shape
		/// </summary>
		public Chart<ApplicationSpecificSetting> ReplaceChart<ApplicationSpecificSetting>(Chart<ApplicationSpecificSetting> chart) where ApplicationSpecificSetting : PresentationSetting, new()
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = openXMLShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (openXMLShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				chart.UpdateSize((uint)oldTransform.Extents.Cx, (uint)oldTransform.Extents.Cy);
				chart.UpdatePosition((uint)oldTransform.Offset.X, (uint)oldTransform.Offset.Y);
			}
			if (chart.GetChartGraphicFrame().Parent == null)
			{
				parent.InsertBefore(chart.GetChartGraphicFrame(), openXMLShape);
			}
			openXMLShape.Remove();
			return chart;
		}
		/// <summary>
		/// Replace 2016 Support Chart for the source Shape
		/// </summary>
		public P16.Chart<ApplicationSpecificSetting> ReplaceChart<ApplicationSpecificSetting>(P16.Chart<ApplicationSpecificSetting> chart) where ApplicationSpecificSetting : PresentationSetting, new()
		{
			DocumentFormat.OpenXml.OpenXmlElement parent = openXMLShape.Parent;
			if (parent == null)
			{
				throw new InvalidOperationException("Old shape must have a parent.");
			}
			if (openXMLShape.ShapeProperties.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				chart.UpdateSize((uint)oldTransform.Extents.Cx, (uint)oldTransform.Extents.Cy);
				chart.UpdatePosition((uint)oldTransform.Offset.X, (uint)oldTransform.Offset.Y);
			}
			if (chart.GetAlternateContent().Parent == null)
			{
				parent.InsertBefore(chart.GetAlternateContent(), openXMLShape);
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
				picture.UpdateSize((uint)oldTransform.Extents.Cx, (uint)oldTransform.Extents.Cy);
				picture.UpdatePosition((uint)oldTransform.Offset.X, (uint)oldTransform.Offset.Y);
			}
			if (picture.GetPicture().Parent == null)
			{
				parent.InsertBefore(picture.GetPicture(), openXMLShape);
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
				table.UpdateSize((uint)oldTransform.Extents.Cx, (uint)oldTransform.Extents.Cy);
				table.UpdatePosition((uint)oldTransform.Offset.X, (uint)oldTransform.Offset.Y);
			}
			if (table.GetTableGraphicFrame().Parent == null)
			{
				parent.InsertBefore(table.GetTableGraphicFrame(), openXMLShape);
			}
			openXMLShape.Remove();
			return table;
		}
		/// <summary>
		/// Replace Textbox for the source Shape
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
				textBox.UpdateSize((uint)oldTransform.Extents.Cx, (uint)oldTransform.Extents.Cy);
				textBox.UpdatePosition((uint)oldTransform.Offset.X, (uint)oldTransform.Offset.Y);
				if (openXMLShape.ShapeStyle != null)
				{
					P.ShapeStyle ShapeStyle = (P.ShapeStyle)openXMLShape.ShapeStyle.Clone();
					textBox.UpdateShapeStyle(ShapeStyle);
				}
			}
			if (textBox.GetTextBoxShape().Parent == null)
			{
				parent.InsertBefore(textBox.GetTextBoxShape(), openXMLShape);
			}
			openXMLShape.Remove();
			return textBox;
		}
		internal P.Shape GetShape()
		{
			return openXMLShape;
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
					SolidFillModel solidFillModel = new SolidFillModel()
					{
						schemeColorModel = new SchemeColorModel()
						{
							themeColorValues = ThemeColorValues.TEXT_1
						}
					};
					if (shapeTextModel.fontColor != null)
					{
						solidFillModel.hexColor = shapeTextModel.fontColor;
						solidFillModel.schemeColorModel = null;
					}
					paragraph.Append(CreateDrawingRun(new DrawingRunModel()
					{
						text = shapeTextModel.text,
						drawingRunProperties = new DrawingRunPropertiesModel()
						{
							solidFill = solidFillModel,
							fontFamily = shapeTextModel.fontFamily,
							fontSize = shapeTextModel.fontSize,
							isBold = shapeTextModel.isBold,
							isItalic = shapeTextModel.isItalic,
							underline = shapeTextModel.underline
						}
					}));
				}
			}
		}
	}
}
