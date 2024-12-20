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

# Column

Add chart method present in slide component or you can replace the chart using shape componenet.

<figure><img src="../../.gitbook/assets/Screenshot 2024-04-04 104300.png" alt=""><figcaption></figcaption></figure>

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
// Bare minimum
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.ColumnChartSetting<G.PresentationSetting>());
// Some additional samples
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.ColumnChartSetting<G.PresentationSetting>()
	{
		titleOptions = new()
		{
			title = "Column Chart"
		},
		chartLegendOptions = new G.ChartLegendOptions()
		{
			legendPosition = G.ChartLegendOptions.LegendPositionValues.TOP,
			fontSize = 5
		},
		columnChartSeriesSettings = new(){
			null,
			new(){
				columnChartDataPointSettings = new(){
				null,
				new(){
					fillColor = "FF0000"
				},
				new(){
					fillColor = "00FF00"
				},
			},
				fillColor= "AABBCC"
			},
			new(){
				fillColor= "CCBBAA"
			}
		}
	});
```
{% endtab %}
{% endtabs %}

### `ColumnChartSetting<G.PresentationSetting>` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-options) that are specific to column chart.

<table><thead><tr><th width="251">Property</th><th width="287">Type</th><th>Details</th></tr></thead><tbody><tr><td>columnChartDataLabel</td><td><a href="column.md#columnchartdatalabel-options">ColumnChartDataLabel</a></td><td>General Data label option applied for all series</td></tr><tr><td>columnChartSeriesSettings</td><td>List&#x3C;<a href="column.md#columnchartsetting-options">ColumnChartSeriesSetting</a>?></td><td>Data Series specific options are used from the list. The position on the list is matched with the data series position. you can use null to skip a series</td></tr><tr><td>columnChartTypes</td><td>ColumnChartTypes</td><td>Type of chart</td></tr><tr><td>chartAxesOptions</td><td><a href="./#chartaxesoptions-options">ChartAxesOptions</a></td><td>Chart axes options</td></tr><tr><td>columnGraphicsSetting</td><td><a href="column.md#columngraphicssetting-options">ColumnGraphicsSetting</a></td><td>Set properties related to bar placement</td></tr></tbody></table>

### `ColumnChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to column chart.

<table><thead><tr><th width="194">Property</th><th width="223">Type</th><th>Details</th></tr></thead><tbody><tr><td>dataLabelPosition</td><td>DataLabelPositionValues</td><td>Data Label placement options.</td></tr></tbody></table>

### `ColumnChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to column chart.

<table><thead><tr><th width="281"></th><th width="311"></th><th></th></tr></thead><tbody><tr><td>columnChartDataLabel</td><td><a href="column.md#columnchartdatalabel-options">ColumnChartDataLabel</a></td><td>Data Label Option specific to one series</td></tr><tr><td>fillColor</td><td>string?</td><td>Fill color specific to one series</td></tr><tr><td>columnChartDataPointSettings</td><td>List&#x3C;<a href="column.md#columnchartdatapointsetting-options">ColumnChartDataPointSetting</a>?></td><td>Data point specific options are used from the list. The position on the list is matched with the data point position. you can use null to skip a data point.</td></tr></tbody></table>

### `ColumnGraphicsSetting` Options only applied in cluster type

<table><thead><tr><th width="165">Property</th><th width="82">Type</th><th>Details</th></tr></thead><tbody><tr><td>categoryGap</td><td>int</td><td>Gap between Category. Default : 219</td></tr><tr><td>seriesGap</td><td>int</td><td>Gap between Series. Default : -27</td></tr></tbody></table>

### `ColumnChartDataPointSetting` Options

Contains options details extended from [`ChartDataPointSetting`](./#chartdatapointsettings-options) that are specific to column chart.
