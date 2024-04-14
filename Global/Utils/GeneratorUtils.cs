// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
namespace OpenXMLOffice.Global_2007
{
    /// <summary>
    /// Generator Utils
    /// </summary>
    public static class GeneratorUtils
    {
        /// <summary>
        /// Generate new GUID
        /// </summary>
        /// <returns>
        /// </returns>
        public static string GenerateNewGUID()
        {
            return string.Format("{{{0}}}", Guid.NewGuid().ToString("D").ToUpper());
        }
    }
}
