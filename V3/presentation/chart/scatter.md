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

# Scatter

Add chart method present in slide component or you can replace the chart using shape componenet.

<figure><img src="../../.gitbook/assets/Screenshot 2024-04-04 104930.png" alt=""><figcaption></figcaption></figure>

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
// Bare minimum
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.ScatterChartSetting<G.PresentationSetting>());
// Some additional samples
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.ScatterChartSetting<G.PresentationSetting>()
	{
		scatterChartTypes = G.ScatterChartTypes.BUBBLE
	});
```
{% endtab %}
{% endtabs %}

### `ScatterChartSetting<G.PresentationSetting>` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-options) that are specific to scatter chart.

<table><thead><tr><th width="251">Property</th><th width="287">Type</th><th>Details</th></tr></thead><tbody><tr><td>scatterChartDataLabel</td><td><a href="scatter.md#scatterchartdatalabel-options">ScatterChartDataLabel</a></td><td>General Data label option applied for all series</td></tr><tr><td>scatterChartSeriesSettings</td><td>List&#x3C;<a href="scatter.md#scatterchartseriessetting-options">ScatterChartSeriesSetting</a>?></td><td>Data Series specific options are used from the list. The position on the list is matched with the data series position. you can use null to skip a series</td></tr><tr><td>scatterChartTypes</td><td>ScatterChartTypes</td><td>Type of chart</td></tr><tr><td>chartAxesOptions</td><td><a href="./#chartaxesoptions-options">ChartAxesOptions</a></td><td>Chart axes options</td></tr></tbody></table>

### `ScatterChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to scatter chart.

<table><thead><tr><th width="194">Property</th><th width="220">Type</th><th>Details</th></tr></thead><tbody><tr><td>dataLabelPosition</td><td>DataLabelPositionValues</td><td>Data Label placement options.</td></tr></tbody></table>

### `ScatterChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to scatter chart.

<table><thead><tr><th width="226">Property</th><th width="209">Type</th><th>Details</th></tr></thead><tbody><tr><td>ScatterChartDataLabel</td><td><a href="scatter.md#scatterchartdatalabel-options">ScatterChartDataLabel</a></td><td>Data Label Option specific to one series</td></tr></tbody></table>
