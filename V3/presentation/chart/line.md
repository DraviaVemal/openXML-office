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

Add chart method present in slide component or you can replace the chart using shape componenet.

<figure><img src="../../.gitbook/assets/Screenshot 2024-04-04 104353.png" alt=""><figcaption></figcaption></figure>

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
<pre class="language-csharp"><code class="lang-csharp"><strong>// Bare minimum
</strong>owerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.LineChartSetting&#x3C;G.PresentationSetting>());
// Some additional samples
<strong>powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
</strong>	.AddChart(CreateDataCellPayload(), new G.LineChartSetting&#x3C;G.PresentationSetting>()
	{
		lineChartSeriesSettings = new(){
			new(){
				lineChartLineFormat = new(){
					dashType = G.DrawingPresetLineDashValues.DASH_DOT,
					lineColor = "FF0000",
					beginArrowValues= G.DrawingBeginArrowValues.ARROW,
					endArrowValues= G.DrawingEndArrowValues.TRIANGLE,
					lineStartWidth = G.LineWidthValues.MEDIUM,
					lineEndWidth = G.LineWidthValues.LARGE,
					outlineCapTypeValues = G.OutlineCapTypeValues.ROUND,
					outlineLineTypeValues = G.OutlineLineTypeValues.DOUBLE,
					width = 5
				}
			}
		}
	});
</code></pre>
{% endtab %}
{% endtabs %}

### `LineChartSetting<G.PresentationSetting>` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-options) that are specific to line chart.

<table><thead><tr><th width="231">Property</th><th width="262">Type</th><th>Details</th></tr></thead><tbody><tr><td>lineChartDataLabel</td><td><a href="line.md#linechartdatalabel-options">LineChartDataLabel</a></td><td>General Data label option applied for all series</td></tr><tr><td>lineChartSeriesSettings</td><td>List&#x3C;<a href="line.md#linechartseriessetting-options">LineChartSeriesSetting</a>?></td><td>Data Series specific options are used from the list. The position on the list is matched with the data series position. you can use null to skip a series</td></tr><tr><td>lineChartTypes</td><td>LineChartTypes</td><td>Type of chart</td></tr><tr><td>chartAxesOptions</td><td><a href="./#chartaxesoptions-options">ChartAxesOptions</a></td><td>Chart axes options</td></tr></tbody></table>

### `LineChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to line chart.

<table><thead><tr><th width="191">Property</th><th width="222">Type</th><th>Details</th></tr></thead><tbody><tr><td>dataLabelPosition</td><td>DataLabelPositionValues</td><td>Data Label placement options.</td></tr></tbody></table>

### `LineChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to column chart.

<table><thead><tr><th width="258">Property</th><th width="292">Type</th><th>Details</th></tr></thead><tbody><tr><td>lineChartDataLabel</td><td><a href="line.md#linechartdatalabel-options">LineChartDataLabel</a></td><td>Data Label Option specific to one series</td></tr><tr><td>lineChartLineFormat</td><td><a href="line.md#linechartlineformat-options">LineChartLineFormat</a></td><td></td></tr><tr><td>lineChartDataPointSettings</td><td>List&#x3C;LineChartDataPointSetting?></td><td>TODO</td></tr></tbody></table>

### `LineChartLineFormat` Options

<table><thead><tr><th width="213">Property</th><th width="270">Type</th><th>Details</th></tr></thead><tbody><tr><td>transparency</td><td>int?</td><td></td></tr><tr><td>width</td><td>int?</td><td></td></tr><tr><td>outlineCapTypeValues</td><td>OutlineCapTypeValues?</td><td></td></tr><tr><td>outlineLineTypeValues</td><td>OutlineLineTypeValues?</td><td></td></tr><tr><td>beginArrowValues</td><td>DrawingBeginArrowValues?</td><td></td></tr><tr><td>endArrowValues</td><td>DrawingEndArrowValues?</td><td></td></tr><tr><td>dashType</td><td>DrawingPresetLineDashValues?</td><td></td></tr><tr><td>lineStartWidth</td><td>LineWidthValues?</td><td></td></tr><tr><td>lineEndWidth</td><td>LineWidthValues?</td><td></td></tr></tbody></table>
