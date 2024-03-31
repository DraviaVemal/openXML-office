// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;

namespace OpenXMLOffice.Spreadsheet_2013
{

    /// <summary>
    /// 
    /// </summary>
    public class DrawingPictureModel
    {
        /// <summary>
        /// 
        /// </summary>
        public required uint id;

        /// <summary>
        /// 
        /// </summary>
        public required string name;

        /// <summary>
        /// 
        /// </summary>
        public bool noChangeAspectRatio = true;

        /// <summary>
        /// 
        /// </summary>
        public required string blipEmbed;
    }

    /// <summary>
    /// 
    /// </summary>
    public class DrawingGraphicFrame
    {
        /// <summary>
        /// 
        /// </summary>
        public required uint id;

        /// <summary>
        /// 
        /// </summary>
        public required string name;

        /// <summary>
        /// 
        /// </summary>
        public required string chartId;
    }

    /// <summary>
    /// 
    /// </summary>
    public class TwoCellAnchorModel
    {

        /// <summary>
        /// 
        /// </summary>
        public AnchorPosition from = new();

        /// <summary>
        /// 
        /// </summary>
        public AnchorPosition to = new();

        /// <summary>
        /// 
        /// </summary>
        public DrawingPictureModel? drawingPictureModel;

        /// <summary>
        /// 
        /// </summary>
        public DrawingGraphicFrame? drawingGraphicFrame;

        /// <summary>
        /// 
        /// </summary>
        public uint x = 0;
        /// <summary>
        /// 
        /// </summary>
        public uint y = 0;
        /// <summary>
        /// 
        /// </summary>
        public uint height = 0;
        /// <summary>
        /// 
        /// </summary>
        public uint weight = 0;
    }

}