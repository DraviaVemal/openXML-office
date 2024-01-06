namespace OpenXMLOffice.Global
{
    public enum DataType
    {
        DATE,
        NUMBER,
        STRING
    }

    public class ChartData
    {
        #region Public Fields

        public DataType DataType = DataType.STRING;
        public string? Value;

        #endregion Public Fields
    }
}