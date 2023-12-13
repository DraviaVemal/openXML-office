namespace OpenXMLOffice.Global
{
    public static class GeneratorUtils
    {
        #region Public Methods

        public static string GenerateNewGUID()
        {
            return string.Format("{{{0}}}", Guid.NewGuid().ToString("D").ToUpper());
        }

        #endregion Public Methods
    }
}