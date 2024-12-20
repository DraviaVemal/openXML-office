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

# Pie

Add chart method present in worksheet component. By default the anchor is at 1,1 aka A1 cell.

<figure><img src="../../.gitbook/assets/Screenshot 2024-04-04 103025.png" alt=""><figcaption></figcaption></figure>

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
Worksheet worksheet = excel1.AddSheet("Line Chart");
	worksheet.AddChart(new()
	{
		cellIdStart = "A1",
		cellIdEnd = "D4"
	}, new PieChartSetting<ExcelSetting>()
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

### `PieChartSetting` Options

Contains options details extended from [`ChartSetting`](../../presentation/chart/#chartsetting-options) that are specific to pie chart.

<table><thead><tr><th width="238">Property</th><th width="262">Type</th><th>Details</th></tr></thead><tbody><tr><td>pieChartDataLabel</td><td><a href="pie.md#piechartdatalabel-options">PieChartDataLabel</a></td><td>General Data label option applied for all series</td></tr><tr><td>pieChartSeriesSettings</td><td>List&#x3C;<a href="pie.md#piechartseriessetting-options">PieChartSeriesSetting</a>?></td><td>Data Series specific options are used from the list. The position on the list is matched with the data series position. you can use null to skip a series</td></tr><tr><td>pieChartTypes</td><td>PieChartTypes</td><td>Type of chart</td></tr><tr><td>doughnutHoleSize</td><td>uint</td><td></td></tr><tr><td>angleOfFirstSlice</td><td>uint</td><td></td></tr><tr><td>pointExplosion</td><td>uint</td><td></td></tr></tbody></table>

### `PieChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](../../presentation/chart/#chartdatalabel-options) that are specific to pie chart.

<table><thead><tr><th width="194">Property</th><th width="220">Type</th><th>Details</th></tr></thead><tbody><tr><td>dataLabelPosition</td><td>DataLabelPositionValues</td><td>Data Label placement options.</td></tr></tbody></table>

### `PieChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](../../presentation/chart/#chartseriessetting-options) that are specific to pie chart.

<table><thead><tr><th width="206">Property</th><th width="188">Type</th><th>Details</th></tr></thead><tbody><tr><td>pieChartDataLabel</td><td><a href="pie.md#piechartdatalabel-options">PieChartDataLabel</a></td><td>Data Label Option specific to one series</td></tr><tr><td>fillColor</td><td>string?</td><td>Fill color specific to one series</td></tr><tr><td>borderColor</td><td>string?</td><td>Border color specific to one series</td></tr><tr><td>pieChartDataPointSettings</td><td>List&#x3C;PieChartDataPointSetting?></td><td>TODO</td></tr></tbody></table>
