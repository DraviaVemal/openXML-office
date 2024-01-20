// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    /// <summary>
    /// Shape Class For Presentation shape manipulation
    /// </summary>
    public class Shape
    {
        #region Private Fields

        private readonly P.Shape openXMLShape = new();

        #endregion Private Fields

        #region Internal Constructors

        internal Shape(P.Shape? shape = null)
        {
            if (shape != null)
            {
                openXMLShape = shape;
            }
        }

        #endregion Internal Constructors

        #region Public Methods

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
        /// <param name="Chart">
        /// </param>
        /// <returns>
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// </exception>
        public Chart ReplaceChart(Chart Chart)
        {
            DocumentFormat.OpenXml.OpenXmlElement? parent = openXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
            if (openXMLShape.ShapeProperties?.Transform2D != null)
            {
                A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
                Chart.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
                Chart.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
            }
            if (Chart.GetChartGraphicFrame().Parent == null)
            {
                parent.InsertBefore(Chart.GetChartGraphicFrame(), openXMLShape);
            }
            openXMLShape.Remove();
            return Chart;
        }

        /// <summary>
        /// Replace Picture for the source Shape
        /// </summary>
        /// <param name="Picture">
        /// </param>
        /// <returns>
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// </exception>
        public Picture ReplacePicture(Picture Picture)
        {
            DocumentFormat.OpenXml.OpenXmlElement? parent = openXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
            if (openXMLShape.ShapeProperties?.Transform2D != null)
            {
                A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
                Picture.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
                Picture.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
            }
            if (Picture.GetPicture().Parent == null)
            {
                parent.InsertBefore(Picture.GetPicture(), openXMLShape);
            }
            openXMLShape.Remove();
            return Picture;
        }

        /// <summary>
        /// Replace Table for the source Shape
        /// </summary>
        /// <param name="Table">
        /// </param>
        /// <returns>
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// </exception>
        public Table ReplaceTable(Table Table)
        {
            DocumentFormat.OpenXml.OpenXmlElement? parent = openXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
            if (openXMLShape.ShapeProperties?.Transform2D != null)
            {
                A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
                Table.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
                Table.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
            }
            if (Table.GetTableGraphicFrame().Parent == null)
            {
                parent.InsertBefore(Table.GetTableGraphicFrame(), openXMLShape);
            }
            openXMLShape.Remove();
            return Table;
        }

        /// <summary>
        /// Replace Textbox for the source Shape
        /// </summary>
        /// <param name="TextBox">
        /// </param>
        /// <returns>
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// </exception>
        public TextBox ReplaceTextBox(TextBox TextBox)
        {
            DocumentFormat.OpenXml.OpenXmlElement? parent = openXMLShape.Parent ?? throw new InvalidOperationException("Old shape must have a parent.");
            if (openXMLShape.ShapeProperties?.Transform2D != null)
            {
                A.Transform2D oldTransform = openXMLShape.ShapeProperties.Transform2D;
                TextBox.UpdateSize((uint)oldTransform.Extents!.Cx!, (uint)oldTransform.Extents!.Cy!);
                TextBox.UpdatePosition((uint)oldTransform.Offset!.X!, (uint)oldTransform.Offset!.Y!);
            }
            if (TextBox.GetTextBoxShape().Parent == null)
            {
                parent.InsertBefore(TextBox.GetTextBoxShape(), openXMLShape);
            }
            openXMLShape.Remove();
            return TextBox;
        }

        #endregion Public Methods

        #region Internal Methods

        internal P.Shape GetShape()
        {
            return openXMLShape;
        }

        #endregion Internal Methods
    }
}