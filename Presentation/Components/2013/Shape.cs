// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation_2013
{
	/// <summary>
	/// Shape Class For Presentation shape manipulation
	/// </summary>
	public class Shape
	{
		private readonly P.Shape openXMLShape = new();

		internal Shape(P.Shape? shape = null)
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
		public Chart ReplaceChart(Chart chart)
		{
			DocumentFormat.OpenXml.OpenXmlElement? parent = openXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
			if (openXMLShape.ShapeProperties?.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				chart.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
				chart.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
			}
			if (chart.GetChartGraphicFrame().Parent == null)
			{
				parent.InsertBefore(chart.GetChartGraphicFrame(), openXMLShape);
			}
			openXMLShape.Remove();
			return chart;
		}

		/// <summary>
		/// Replace Picture for the source Shape
		/// </summary>
		public Picture ReplacePicture(Picture picture)
		{
			DocumentFormat.OpenXml.OpenXmlElement? parent = openXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
			if (openXMLShape.ShapeProperties?.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				picture.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
				picture.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
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
			DocumentFormat.OpenXml.OpenXmlElement? parent = openXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
			if (openXMLShape.ShapeProperties?.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				table.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
				table.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
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
			DocumentFormat.OpenXml.OpenXmlElement? parent = openXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
			if (openXMLShape.ShapeProperties?.Transform2D != null)
			{
				A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
				textBox.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
				textBox.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
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


	}
}
