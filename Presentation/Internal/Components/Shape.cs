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

        public void ReplaceShape(P.Shape Shape)
        {
            DocumentFormat.OpenXml.OpenXmlElement? parent = OpenXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
            if (OpenXMLShape.ShapeProperties?.Transform2D != null)
            {
                A.Transform2D oldTransform = OpenXMLShape.ShapeProperties.Transform2D;
                Shape.ShapeProperties ??= new P.ShapeProperties();
                Shape.ShapeProperties.Transform2D = new A.Transform2D
                {
                    Offset = new A.Offset { X = oldTransform.Offset!.X, Y = oldTransform.Offset.Y },
                    Extents = new A.Extents { Cx = oldTransform.Extents!.Cx, Cy = oldTransform.Extents.Cy }
                };
            }
            parent.InsertBefore(Shape, OpenXMLShape);
            OpenXMLShape.Remove();
        }

        public void ReplaceShape(P.GraphicFrame GraphicFrame)
        {
            DocumentFormat.OpenXml.OpenXmlElement? parent = OpenXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
            if (OpenXMLShape.ShapeProperties?.Transform2D != null)
            {
                A.Transform2D oldTransform = OpenXMLShape.ShapeProperties.Transform2D;
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = oldTransform.Offset!.X, Y = oldTransform.Offset.Y },
                    Extents = new A.Extents { Cx = oldTransform.Extents!.Cx, Cy = oldTransform.Extents.Cy }
                };
            }
            parent.InsertBefore(GraphicFrame, OpenXMLShape);
            OpenXMLShape.Remove();
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