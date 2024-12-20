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

# Line

Add chart method present in worksheet component. By default the anchor is at 1,1 aka A1 cell.

<figure><img src="../../.gitbook/assets/Screenshot 2024-04-04 102903.png" alt=""><figcaption></figcaption></figure>

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
Worksheet worksheet = excel1.AddSheet("Line Chart");
	worksheet.AddChart(new()
	{
		cellIdStart = "A1",
		cellIdEnd = "D4"
	}, new LineChartSetting<ExcelSetting>()
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

### `LineChartSetting` Options

Contains options details extended from [`ChartSetting`](../../presentation/chart/#chartsetting-options) that are specific to line chart.

<table><thead><tr><th width="231">Property</th><th width="262">Type</th><th>Details</th></tr></thead><tbody><tr><td>lineChartDataLabel</td><td><a href="line.md#linechartdatalabel-options">LineChartDataLabel</a></td><td>General Data label option applied for all series</td></tr><tr><td>lineChartSeriesSettings</td><td>List&#x3C;<a href="line.md#linechartseriessetting-options">LineChartSeriesSetting</a>?></td><td>Data Series specific options are used from the list. The position on the list is matched with the data series position. you can use null to skip a series</td></tr><tr><td>lineChartTypes</td><td>LineChartTypes</td><td>Type of chart</td></tr><tr><td>chartAxesOptions</td><td><a href="../../presentation/chart/#chartaxesoptions-options">ChartAxesOptions</a></td><td>Chart axes options</td></tr></tbody></table>

### `LineChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](../../presentation/chart/#chartdatalabel-options) that are specific to line chart.

<table><thead><tr><th width="191">Property</th><th width="222">Type</th><th>Details</th></tr></thead><tbody><tr><td>dataLabelPosition</td><td>DataLabelPositionValues</td><td>Data Label placement options.</td></tr></tbody></table>

### `LineChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](../../presentation/chart/#chartseriessetting-options) that are specific to column chart.

<table><thead><tr><th width="258">Property</th><th width="292">Type</th><th>Details</th></tr></thead><tbody><tr><td>lineChartDataLabel</td><td><a href="line.md#linechartdatalabel-options">LineChartDataLabel</a></td><td>Data Label Option specific to one series</td></tr><tr><td>lineChartLineFormat</td><td><a href="line.md#linechartlineformat-options">LineChartLineFormat</a></td><td></td></tr><tr><td>lineChartDataPointSettings</td><td>List&#x3C;LineChartDataPointSetting?></td><td>TODO</td></tr></tbody></table>

### `LineChartLineFormat` Options

<table><thead><tr><th width="213">Property</th><th width="270">Type</th><th>Details</th></tr></thead><tbody><tr><td>transparency</td><td>int?</td><td></td></tr><tr><td>width</td><td>int?</td><td></td></tr><tr><td>outlineCapTypeValues</td><td>OutlineCapTypeValues?</td><td></td></tr><tr><td>outlineLineTypeValues</td><td>OutlineLineTypeValues?</td><td></td></tr><tr><td>beginArrowValues</td><td>DrawingBeginArrowValues?</td><td></td></tr><tr><td>endArrowValues</td><td>DrawingEndArrowValues?</td><td></td></tr><tr><td>dashType</td><td>DrawingPresetLineDashValues?</td><td></td></tr><tr><td>lineStartWidth</td><td>LineWidthValues?</td><td></td></tr><tr><td>lineEndWidth</td><td>LineWidthValues?</td><td></td></tr></tbody></table>
