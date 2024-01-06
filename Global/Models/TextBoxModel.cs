namespace OpenXMLOffice.Global
{
    public class TextBoxSetting
    {
        #region Public Fields
        public uint X = 0;
        public uint Y = 0;
        public uint Height = 100;
        public uint Width = 100;
        public string FontFamily = "Calibri (Body)";
        public int FontSize = 18;
        public bool IsBold = false;
        public bool IsItalic = false;
        public bool IsUnderline = false;
        public string ShapeBackground = "FFFFFF";
        public string Text = "Text Box";
        public string? TextBackground;
        public string TextColor = "000000";

        #endregion Public Fields
    }
}