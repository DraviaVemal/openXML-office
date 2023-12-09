
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenXMLOffice.Presentation;
public class Global
{
    protected CommonSlideData CreateCommonSlideData(bool isAddBackground = false)
    {
        Background background = new()
        {
            BackgroundStyleReference = new BackgroundStyleReference(new A.SchemeColor()
            {
                Val = A.SchemeColorValues.PhColor
            })
            {
                Index = 1001
            }
        };
        ShapeTree shapeTree = new()
        {
            GroupShapeProperties = new GroupShapeProperties()
            {
                TransformGroup = new A.TransformGroup()
                {
                    Offset = new A.Offset()
                    {
                        X = 0,
                        Y = 0
                    },
                    Extents = new A.Extents()
                    {
                        Cx = 0,
                        Cy = 0
                    },
                    ChildOffset = new A.ChildOffset()
                    {
                        X = 0,
                        Y = 0
                    },
                    ChildExtents = new A.ChildExtents()
                    {
                        Cx = 0,
                        Cy = 0
                    }
                }
            },
            NonVisualGroupShapeProperties = new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties() { Id = 1, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        )
        };
        if (isAddBackground)
        {
            return new CommonSlideData(background, shapeTree);
        }
        else
        {
            return new CommonSlideData(shapeTree);
        }
    }
}
