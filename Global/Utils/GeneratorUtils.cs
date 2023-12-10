namespace OpenXMLOffice.Global;

public static class GeneratorUtils
{

    public static string GenerateNewGUID()
    {
        return string.Format("{{{0}}}", Guid.NewGuid().ToString("D").ToUpper());
    }

}