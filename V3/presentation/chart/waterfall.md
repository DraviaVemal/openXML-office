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

# Waterfall

Add chart method present in slide component or you can replace the chart using shape componenet.\
Base supported version for this type of chart is office 2016&#x20;

<figure><img src="../../.gitbook/assets/Screenshot 2024-04-04 105039.png" alt=""><figcaption></figcaption></figure>

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
    .AddChart(data, new WaterfallChartSetting<G.PresentationSetting>());
```
{% endtab %}
{% endtabs %}

### `WaterfallChartSetting<G.PresentationSetting>` Options

At this moment waterfall supports base setting from[`ChartSetting`](./#chartsetting-options) future updates will get updated below.
