// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2007
{
    /// <summary>
    ///
    /// </summary>
    public enum ShapeRectangleTypes
    {
        /// <summary>
        ///
        /// </summary>
        RECTANGLE,
        /// <summary>
        ///
        /// </summary>
        OVAL,
        /// <summary>
        ///
        /// </summary>
        TRIANGLE,
    }
    /// <summary>
    ///
    /// </summary>
    public enum ShapeLineTypes
    {
        /// <summary>
        ///
        /// </summary>
        LINE,
        /// <summary>
        ///
        /// </summary>
        LINE_ARROW,
        /// <summary>
        ///
        /// </summary>
        LINE_DOUBLE_ARROW,
        /// <summary>
        ///
        /// </summary>
        CONNECTOR_STRAIGHT,
        /// <summary>
        ///
        /// </summary>
        CONNECTOR_ELBOW,
        /// <summary>
        ///
        /// </summary>
        CONNECTOR_CURVED,
    }

    /// <summary>
    ///
    /// </summary>
    public enum ShapeArrowTypes
    {

        /// <summary>
        ///
        /// </summary>
        LEFT,

        /// <summary>
        ///
        /// </summary>
        RIGHT,

        /// <summary>
        ///
        /// </summary>
        UP,

        /// <summary>
        ///
        /// </summary>
        DOWN,

        /// <summary>
        ///
        /// </summary>
        LEFT_RIGHT,

        /// <summary>
        ///
        /// </summary>
        UP_DOWN,

        /// <summary>
        ///
        /// </summary>
        QUAD,

        /// <summary>
        ///
        /// </summary>
        LEFT_RIGHT_UP,

        /// <summary>
        ///
        /// </summary>
        CURVED_LEFT,

        /// <summary>
        ///
        /// </summary>
        CURVED_RIGHT
    }
    /// <summary>
    ///
    /// </summary>
    public class ShapeLineModel<ApplicationSpecificSetting>
        where ApplicationSpecificSetting : class, ISizeAndPosition, new()
    {
        /// <summary>
        ///
        /// </summary>
        public ShapeLineTypes lineTypes = ShapeLineTypes.LINE;
        /// <summary>
        ///
        /// </summary>
        public ApplicationSpecificSetting applicationSpecificSetting = new ApplicationSpecificSetting();
    }
    /// <summary>
    ///
    /// </summary>
    public class ShapeRectangleModel<ApplicationSpecificSetting>
        where ApplicationSpecificSetting : class, ISizeAndPosition, new()
    {
        /// <summary>
        ///
        /// </summary>
        public ShapeRectangleTypes rectangleType = ShapeRectangleTypes.RECTANGLE;
        /// <summary>
        ///
        /// </summary>
        public TextOptions text = null;
        /// <summary>
        ///
        /// </summary>
        public ApplicationSpecificSetting applicationSpecificSetting = new ApplicationSpecificSetting();
    }
    /// <summary>
    ///
    /// </summary>
    public class ShapeArrowModel<ApplicationSpecificSetting>
        where ApplicationSpecificSetting : class, ISizeAndPosition, new()
    {
        /// <summary>
        ///
        /// </summary>
        public ShapeArrowTypes rectangleType = ShapeArrowTypes.LEFT;
        /// <summary>
        ///
        /// </summary>
        public TextOptions text = null;
        /// <summary>
        ///
        /// </summary>
        public ApplicationSpecificSetting applicationSpecificSetting = new ApplicationSpecificSetting();
    }
}