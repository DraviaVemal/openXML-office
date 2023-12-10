namespace OpenXMLOffice.Presentation;
public class PresentationProperties
{
    public PresentationTheme Theme = new();
    public PresentationSettings Settings = new();
    /// <summary>
    /// TODO : Multi Theme Slide Master Support
    /// </summary>
    public Dictionary<string, PresentationSlideMaster>? SlideMasters;
}
/// <summary>
/// TODO : Multi Theme Slide Master Support
/// </summary>
public class PresentationSlideMaster
{
    public PresentationTheme Theme = new();
}

public class PresentationTheme
{
    public string Dark1 = "000000";
    public string Light1 = "FFFFFF";
    public string Dark2 = "44546A";
    public string Light2 = "E7E6E6";
    public string Accent1 = "4472C4";
    public string Accent2 = "ED7D31";
    public string Accent3 = "A5A5A5";
    public string Accent4 = "FFC000";
    public string Accent5 = "5B9BD5";
    public string Accent6 = "70AD47";
    public string Hyperlink = "0563C1";
    public string FollowedHyperlink = "954F72";
}

public class PresentationSettings
{
    public bool IsMultiSlideMasterPartPresentation = false;
    public bool IsMultiThemePresentation = false;
}

internal class PresentationInfo
{
    public string? FilePath;
    public bool IsExistingFile = false;
    public bool IsEditable = true;
}