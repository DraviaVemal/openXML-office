/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Generator Utils
    /// </summary>
    public static class GeneratorUtils
    {
        #region Public Methods
        /// <summary>
        /// Generate new GUID
        /// </summary>
        /// <returns></returns>
        public static string GenerateNewGUID()
        {
            return string.Format("{{{0}}}", Guid.NewGuid().ToString("D").ToUpper());
        }

        #endregion Public Methods
    }
}