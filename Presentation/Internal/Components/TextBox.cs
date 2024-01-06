using OpenXMLOffice.Global;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    public class TextBox : TextBoxBase
    {
        public TextBox(TextBoxSetting TextBoxSetting) : base(TextBoxSetting)
        {
        }

        internal A.Run GetTextBoxRun()
        {
            return base.GetTextBoxRun();
        }

        internal P.Shape GetTextBoxShape()
        {
            return base.GetTextBoxShape();
        }
    }
}