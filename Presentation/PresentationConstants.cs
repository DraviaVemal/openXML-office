namespace OpenXMLOffice.Presentation;
public static class PresentationConstants
{
    public enum SlideLayoutType
    {
        BLANK
    }

    public enum CommonSlideDataType
    {
        SLIDE_LAYOUT,
        SLIDE_MASTER,
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
