// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
namespace OpenXMLOffice.Global_2007
{
    /// <summary>
    /// Generator Utils
    /// </summary>
    public static class Validation
    {
        /// <summary>
        /// Validate given coordinate are within range
        /// </summary>
        public static bool IsWithinRange(int x, int y, int topLeftX, int topLeftY, int bottomRightX, int bottomRightY)
        {
            return x >= topLeftX && x <= bottomRightX && y >= topLeftY && y <= bottomRightY;
        }
    }
}
