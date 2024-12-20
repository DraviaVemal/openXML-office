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

Add chart method present in slide component or you can replace the chart using shape componenet.

<figure><img src="../../.gitbook/assets/Screenshot 2024-04-04 104842.png" alt=""><figcaption></figcaption></figure>

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
// Bare minimum
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.PieChartSetting<G.PresentationSetting>());
// Some additional samples
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.PieChartSetting<G.PresentationSetting>()
	{
		pieChartTypes = G.PieChartTypes.DOUGHNUT,
		pieChartDataLabel = new()
		{
			dataLabelPosition = G.PieChartDataLabel.DataLabelPositionValues.SHOW,
			showCategoryName = true,
			showValue = true,
			separator = ". "
		}
	});
```
{% endtab %}
{% endtabs %}

### `PieChartSetting<G.PresentationSetting>` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-options) that are specific to pie chart.

<table><thead><tr><th width="238">Property</th><th width="262">Type</th><th>Details</th></tr></thead><tbody><tr><td>pieChartDataLabel</td><td><a href="pie.md#piechartdatalabel-options">PieChartDataLabel</a></td><td>General Data label option applied for all series</td></tr><tr><td>pieChartSeriesSettings</td><td>List&#x3C;<a href="pie.md#piechartseriessetting-options">PieChartSeriesSetting</a>?></td><td>Data Series specific options are used from the list. The position on the list is matched with the data series position. you can use null to skip a series</td></tr><tr><td>pieChartTypes</td><td>PieChartTypes</td><td>Type of chart</td></tr><tr><td>doughnutHoleSize</td><td>uint</td><td></td></tr><tr><td>angleOfFirstSlice</td><td>uint</td><td></td></tr><tr><td>pointExplosion</td><td>uint</td><td></td></tr></tbody></table>

### `PieChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to pie chart.

<table><thead><tr><th width="194">Property</th><th width="220">Type</th><th>Details</th></tr></thead><tbody><tr><td>dataLabelPosition</td><td>DataLabelPositionValues</td><td>Data Label placement options.</td></tr></tbody></table>

### `PieChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to pie chart.

<table><thead><tr><th width="206">Property</th><th width="188">Type</th><th>Details</th></tr></thead><tbody><tr><td>pieChartDataLabel</td><td><a href="pie.md#piechartdatalabel-options">PieChartDataLabel</a></td><td>Data Label Option specific to one series</td></tr><tr><td>fillColor</td><td>string?</td><td>Fill color specific to one series</td></tr><tr><td>borderColor</td><td>string?</td><td>Border color specific to one series</td></tr><tr><td>pieChartDataPointSettings</td><td>List&#x3C;PieChartDataPointSetting?></td><td>TODO</td></tr></tbody></table>
