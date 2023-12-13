namespace OpenXMLOffice.Global;

public class ChartData
{
    public string? Value;
}

public class ChartSeriesSetting
{
    public string? NumberFormat;
}

public class ChartSetting
{
    public List<ChartSeriesSetting>? SeriesSettings;
}