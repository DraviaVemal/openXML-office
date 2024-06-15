// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;

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
    public class ShapeLineModel<ApplicationSpecificSetting, LineColorOption>
        where ApplicationSpecificSetting : class, ISizeAndPosition, new()
        where LineColorOption : class, IColorOptions, new()
    {
        /// <summary>
        ///
        /// </summary>
        public ShapeLineTypes lineTypes = ShapeLineTypes.LINE;
        /// <summary>
        ///
        /// </summary>
        public ApplicationSpecificSetting applicationSpecificSetting = new ApplicationSpecificSetting();
        /// <summary>
        ///
        /// </summary>
        public LineColorOption lineColorOption = new LineColorOption();
    }
    /// <summary>
    ///
    /// </summary>
    public class ShapeRectangleModel<ApplicationSpecificSetting, LineColorOption, FillColorOption>
        where ApplicationSpecificSetting : class, ISizeAndPosition, new()
        where LineColorOption : class, IColorOptions, new()
        where FillColorOption : class, IColorOptions, new()
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
        /// <summary>
        ///
        /// </summary>
        public LineColorOption lineColorOption = new LineColorOption();
        /// <summary>
        ///
        /// </summary>
        public FillColorOption fillColorOption = new FillColorOption();
    }
    /// <summary>
    ///
    /// </summary>
    public class ShapeArrowModel<ApplicationSpecificSetting, LineColorOption, FillColorOption>
        where ApplicationSpecificSetting : class, ISizeAndPosition, new()
        where LineColorOption : class, IColorOptions, new()
        where FillColorOption : class, IColorOptions, new()
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
        /// <summary>
        ///
        /// </summary>
        public LineColorOption lineColorOption = new LineColorOption();
        /// <summary>
        ///
        /// </summary>
        public FillColorOption fillColorOption = new FillColorOption();
    }
    /// <summary>
    ///
    /// </summary>
    public class ShapePropertiesModel
    {
        /// <summary>
        ///
        /// </summary>
        public int X = 1562100;
        /// <summary>
        ///
        /// </summary>
        public int Y = 1524000;
        /// <summary>
        ///
        /// </summary>
        public int Cx = 4743450;
        /// <summary>
        ///
        /// </summary>
        public int Cy = 1419225;
    }
    /// <summary>
    ///
    /// </summary>
    public class ShapeModel<TextColorOption>
    where TextColorOption : class, IColorOptions, new()
    {
        /// <summary>
        ///
        /// </summary>
        public string Name = "";
        /// <summary>
        ///
        /// </summary>
        public ShapePropertiesModel shapePropertiesModel = new ShapePropertiesModel();
        /// <summary>
        ///
        /// </summary>
        public DrawingParagraphModel<TextColorOption> drawingParagraph;
    }
}