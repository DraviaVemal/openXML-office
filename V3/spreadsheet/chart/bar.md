---
layout:
  title:
    visible: true
  description:
    visible: false
  tableOfContents:
    visible: true
  outline:
    visible: true
  pagination:
    visible: true
---

# Bar

Add chart method present in worksheet component. By default the anchor is at 1,1 aka A1 cell.

<figure><img src="../../.gitbook/assets/Screenshot 2024-04-04 102754.png" alt=""><figcaption></figcaption></figure>

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
Worksheet worksheet = excel1.AddSheet("Bar Chart");
	worksheet.AddChart(new()
	{
		cellIdStart = "A1",
		cellIdEnd = "D4"
	}, new BarChartSetting<ExcelSetting>()
	{
		applicationSpecificSetting = new()
		{
			from = new()
			{
				row = 5,
				column = 5
			},
			to = new()
			{
				row = 20,
				column = 20
			}
		}
	});
```
{% endtab %}
{% endtabs %}

### `BarChartSetting` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-less-than-excelsetting-greater-than-options) that are specific to bar chart.

<table><thead><tr><th width="238">Property</th><th width="262">Type</th><th>Details</th></tr></thead><tbody><tr><td>barChartDataLabel</td><td><a href="bar.md#barchartdatalabel-options">BarChartDataLabel</a></td><td>General Data label option applied for all series</td></tr><tr><td>barChartSeriesSettings</td><td>List&#x3C;<a href="bar.md#barchartseriessetting-options">BarChartSeriesSetting</a>?></td><td>Data Series specific options are used from the list. The position on the list is matched with the data series position. you can use null to skip a series</td></tr><tr><td>barChartTypes</td><td>BarChartTypes</td><td>Type of chart</td></tr><tr><td>chartAxesOptions</td><td><a href="../../presentation/chart/#chartaxesoptions-options">ChartAxesOptions</a></td><td>Chart axes options</td></tr><tr><td>barGraphicsSetting</td><td><a href="bar.md#bargraphicssetting-options">BarGraphicsSetting</a></td><td>Set properties related to bar placement</td></tr></tbody></table>

### `BarChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to bar chart.

<table><thead><tr><th width="188">Property</th><th width="231">Type</th><th>Details</th></tr></thead><tbody><tr><td>dataLabelPosition</td><td>DataLabelPositionValues</td><td>Data Label placement options.</td></tr></tbody></table>

### `barChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to bar chart.

<table><thead><tr><th width="255"></th><th width="285"></th><th></th></tr></thead><tbody><tr><td>barChartDataLabel</td><td><a href="bar.md#barchartdatalabel-options">BarChartDataLabel</a></td><td>Data Label Option specific to one series</td></tr><tr><td>fillColor</td><td>string?</td><td>Fill color specific to one series</td></tr><tr><td>barChartDataPointSettings</td><td>List&#x3C;<a href="bar.md#barchartdatapointsetting-options">BarChartDataPointSetting</a>?></td><td>Data point specific options are used from the list. The position on the list is matched with the data point position. you can use null to skip a data point.</td></tr></tbody></table>

### `BarGraphicsSetting` Options only applied in cluster type

<table><thead><tr><th width="150">Property</th><th width="83">Type</th><th>Details</th></tr></thead><tbody><tr><td>categoryGap</td><td>int</td><td>Gap between Category. Default : 219</td></tr><tr><td>seriesGap</td><td>int</td><td>Gap between Series. Default : -27</td></tr></tbody></table>

### `BarChartDataPointSetting` Options

Contains options details extended from [`ChartDataPointSetting`](../../presentation/chart/#chartdatapointsettings-options) that are specific to bar chart.
