namespace OpenXMLOffice.Presentation;
public static class PresentationConstants
{
    public enum SlideLayoutType
    {
        BLANK
    }

    public enum CommonSlideDataType
    {
        SLIDE_MASTER,
        SLIDE_LAYOUT,
        SLIDE
    }

    public static string GetSlideLayoutType(SlideLayoutType value)
    {
        return value switch
        {
            _ => "Blank",
        };
    }
}
