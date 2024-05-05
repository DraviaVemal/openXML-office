// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
namespace OpenXMLOffice.Global_2007
{
    /// <summary>
    /// Generator Utils
    /// </summary>
    public static class LogUtils
    {
        /// <summary>
        /// Print Warning Message
        /// </summary>
        /// <param name="message">Warning Message</param>
        public static void ShowWarning(string message)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("[WARNING] " + message);
            Console.ResetColor();
        }

        /// <summary>
        /// Print Error Message
        /// </summary>
        /// <param name="message">Error Message</param>

        public static void ShowError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("[ERROR] " + message);
            Console.ResetColor();
        }
    }
}
