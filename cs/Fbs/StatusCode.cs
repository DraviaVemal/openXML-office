// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;

namespace OpenXMLOffice.Global_2007
{
    public enum StatusCodeValues : sbyte
    {
        UnknownError = -1,
        Success = 0,
        InvalidArgument = 1,
        FlatBufferError = 2,
    }
    public static class StatusCode
    {
        public static void ProcessStatusCode(sbyte statusCode)
        {
            StatusCodeValues status = (StatusCodeValues)statusCode;
            switch (status)
            {
                case StatusCodeValues.UnknownError:
                    throw new Exception("Unknow Error");
                case StatusCodeValues.InvalidArgument:
                    throw new ArgumentException("Method argument failed");
                case StatusCodeValues.FlatBufferError:
                    throw new Exception("Flatbuffer parsing error");
                default:
                    return;
            }
        }
    }
}
