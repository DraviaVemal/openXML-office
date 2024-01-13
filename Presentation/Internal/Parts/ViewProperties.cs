/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    /// <summary>
    /// PPT View Properties Part Handling
    /// </summary>
    public class ViewProperties
    {
        #region Private Fields

        private readonly P.ViewProperties OpenXMLViewProperties = new();

        #endregion Private Fields

        #region Public Constructors
        /// <summary>
        /// Create New View Properties
        /// </summary>
        public ViewProperties()
        {
            OpenXMLViewProperties.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            OpenXMLViewProperties.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            OpenXMLViewProperties.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            OpenXMLViewProperties.AppendChild(CreateNormalViewProperties());
            OpenXMLViewProperties.AppendChild(CreateSlideViewProperties());
            OpenXMLViewProperties.AppendChild(CreateNotesTextViewProperties());
            OpenXMLViewProperties.AppendChild(new P.GridSpacing()
            {
                Cx = 72008,
                Cy = 72008
            });
        }

        #endregion Public Constructors

        #region Public Methods
        /// <summary>
        /// Return OpenXML View Properties
        /// </summary>
        /// <returns></returns>
        public P.ViewProperties GetViewProperties()
        {
            return OpenXMLViewProperties;
        }

        #endregion Public Methods

        #region Private Methods

        private static P.NotesTextViewProperties CreateNotesTextViewProperties()
        {
            P.NotesTextViewProperties notesTextViewProperties = new(
                new P.CommonViewProperties
                {
                    ScaleFactor = new P.ScaleFactor(new A.ScaleX()
                    {
                        Denominator = 1,
                        Numerator = 1
                    }, new A.ScaleY()
                    {
                        Denominator = 1,
                        Numerator = 1
                    }),
                    Origin = new P.Origin()
                    {
                        X = 0,
                        Y = 0,
                    },
                }
            );
            return notesTextViewProperties;
        }

        private P.NormalViewProperties CreateNormalViewProperties()
        {
            P.NormalViewProperties normalViewProperties = new()
            {
                HorizontalBarState = P.SplitterBarStateValues.Maximized,
                RestoredLeft = new P.RestoredLeft { AutoAdjust = false, Size = 16014 },
                RestoredTop = new P.RestoredTop { Size = 94660 }
            };
            return normalViewProperties;
        }

        private P.SlideViewProperties CreateSlideViewProperties()
        {
            var slideViewProperties = new P.SlideViewProperties(
                new P.CommonSlideViewProperties(
                    new P.CommonViewProperties
                    {
                        VariableScale = true,
                        ScaleFactor = new P.ScaleFactor(new A.ScaleX()
                        {
                            Denominator = 100,
                            Numerator = 159
                        }, new A.ScaleY()
                        {
                            Denominator = 100,
                            Numerator = 159
                        }),
                        Origin = new P.Origin()
                        {
                            X = 306,
                            Y = 138,
                        },
                    },
                    new P.GuideList()
                )
                {
                    SnapToGrid = false
                }
            );
            return slideViewProperties;
        }

        #endregion Private Methods
    }
}