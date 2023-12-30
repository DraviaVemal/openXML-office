using A = DocumentFormat.OpenXml.Drawing;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global;

internal class ChartColor
{
    #region Public Methods

    public CS.ColorStyle CreateColorStyles()
    {
        CS.ColorStyle colorStyle = new { Method = "cycle", Id = 10 };
        colorStyle.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent1
        });
        colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent2
        }); colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent3
        }); colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent4
        }); colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent5
        }); colorStyle.Append(new A.SchemeColor()
        {
            Val = A.SchemeColorValues.Accent6
        });
        colorStyle.Append(new CS.ColorStyleVariation());
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 60000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 80000
        }, new A.LuminanceOffset()
        {
            Val = 20000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 80000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 60000
        }, new A.LuminanceOffset()
        {
            Val = 40000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 50000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 70000
        }, new A.LuminanceOffset()
        {
            Val = 30000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 70000
        }));
        colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
        {
            Val = 50000
        }, new A.LuminanceOffset()
        {
            Val = 50000
        }));
        return colorStyle;
    }

    #endregion Public Methods
}