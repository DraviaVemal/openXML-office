
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenXMLOffice.Presentation;
public class Global
{
    protected CommonSlideData CreateCommonSlideData()
    {
        return new CommonSlideData(new ShapeTree()
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
        });
    }
}
