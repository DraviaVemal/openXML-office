// Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License. See License in
// the project root for license information.
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    public class Shape
    {
        #region Private Fields

        private readonly P.Shape OpenXMLShape = new();

        #endregion Private Fields

        #region Internal Constructors

        internal Shape(P.Shape? shape = null)
        {
            if (shape != null)
            {
                OpenXMLShape = shape;
            }
        }

        #endregion Internal Constructors

        #region Public Methods

        public void RemoveShape()
        {
            OpenXMLShape.Remove();
        }

        public Chart ReplaceChart(Chart Chart)
        {
            DocumentFormat.OpenXml.OpenXmlElement? parent = OpenXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
            if (OpenXMLShape.ShapeProperties?.Transform2D != null)
            {
                A.Transform2D oldTransform = OpenXMLShape.ShapeProperties.Transform2D;
                Chart.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
                Chart.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
            }
            parent.InsertBefore(Chart.GetChartGraphicFrame(), OpenXMLShape);
            OpenXMLShape.Remove();
            return Chart;
        }

        public Picture ReplacePicture(Picture Picture)
        {
            DocumentFormat.OpenXml.OpenXmlElement? parent = OpenXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
            if (OpenXMLShape.ShapeProperties?.Transform2D != null)
            {
                A.Transform2D oldTransform = OpenXMLShape.ShapeProperties.Transform2D;
                Picture.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
                Picture.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
            }
            parent.InsertBefore(Picture.GetPicture(), OpenXMLShape);
            OpenXMLShape.Remove();
            return Picture;
        }

        public Table ReplaceTable(Table Table)
        {
            DocumentFormat.OpenXml.OpenXmlElement? parent = OpenXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
            if (OpenXMLShape.ShapeProperties?.Transform2D != null)
            {
                A.Transform2D oldTransform = OpenXMLShape.ShapeProperties.Transform2D;
                Table.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
                Table.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
            }
            parent.InsertBefore(Table.GetTableGraphicFrame(), OpenXMLShape);
            OpenXMLShape.Remove();
            return Table;
        }

        public TextBox ReplaceTextBox(TextBox TextBox)
        {
            DocumentFormat.OpenXml.OpenXmlElement? parent = OpenXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
            if (OpenXMLShape.ShapeProperties?.Transform2D != null)
            {
                A.Transform2D oldTransform = OpenXMLShape.ShapeProperties.Transform2D;
                TextBox.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
                TextBox.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
            }
            parent.InsertBefore(TextBox.GetTextBoxShape(), OpenXMLShape);
            OpenXMLShape.Remove();
            return TextBox;
        }

        #endregion Public Methods

        #region Internal Methods

        internal P.Shape GetShape()
        {
            return OpenXMLShape;
        }

        #endregion Internal Methods
    }
}