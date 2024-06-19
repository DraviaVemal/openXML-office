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
    public interface IShapeTypeDetailsModel { }
    /// <summary>
    ///
    /// </summary>
    public class NoShape<ApplicationSpecificSetting> : IShapeTypeDetailsModel
        where ApplicationSpecificSetting : class, ISizeAndPosition, new()
    {
        /// <summary>
        ///
        /// </summary>
        public ApplicationSpecificSetting applicationSpecificSetting = new ApplicationSpecificSetting();
    }
    /// <summary>
    ///
    /// </summary>
    public class LineShapeModel<ApplicationSpecificSetting, LineColorOption> : IShapeTypeDetailsModel
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
    public class RectangleShapeModel<ApplicationSpecificSetting, LineColorOption, FillColorOption> : IShapeTypeDetailsModel
        where ApplicationSpecificSetting : class, ISizeAndPosition, new()
        where LineColorOption : class, IColorOptions, new()
        where FillColorOption : class, IColorOptions, new()
    {
        /// <summary>
        ///
        /// </summary>
        public TextOptions text = null;
        /// <summary>
        ///
        /// </summary>
        public ShapeRectangleTypes rectangleType = ShapeRectangleTypes.RECTANGLE;
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
    public class ArrowShapeModel<ApplicationSpecificSetting, LineColorOption, FillColorOption> : IShapeTypeDetailsModel
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
        public double x = 1562100L;
        /// <summary>
        ///
        /// </summary>
        public double y = 1524000L;
        /// <summary>
        ///
        /// </summary>
        public double cx = 4743450L;
        /// <summary>
        ///
        /// </summary>
        public double cy = 1419225L;
    }
    /// <summary>
    ///
    /// </summary>
    public class ShapeModel<TextColorOption, ShapeTypeOptions>
    where TextColorOption : class, IColorOptions, new()
    where ShapeTypeOptions : class, IShapeTypeDetailsModel, new()
    {
        /// <summary>
        ///
        /// </summary>
        public uint id = 1;
        /// <summary>
        ///
        /// </summary>
        public string name = "";
        /// <summary>
        ///
        /// </summary>
        public ShapeTypeOptions shapeTypeOptions = new ShapeTypeOptions();
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