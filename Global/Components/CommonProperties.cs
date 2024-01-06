using A = DocumentFormat.OpenXml.Drawing;

namespace OpenXMLOffice.Global
{
    public class CommonProperties
    {
        #region Protected Methods

        protected A.SolidFill CreateSolidFill(List<string> FillColors, int index)
        {
            if (FillColors.Count > 0)
            {
                return new A.SolidFill(new A.RgbColorModelHex() { Val = FillColors[index % FillColors.Count] });
            }
            return new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(index % 6) + 1}") });
        }

        #endregion Protected Methods
    }
}