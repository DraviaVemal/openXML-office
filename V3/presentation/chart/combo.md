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

# Combo

Add chart method present in slide component or you can replace the chart using shape componenet.\
This type is bit different from previous core chart types. Combo setting act as organiser for adding other base type chart to single chart componenet along with additional options

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
G.ComboChartSetting<G.PresentationSetting> comboChartSetting = new()
{
	titleOptions = new()
	{
		title = "Combo Chart"
	},
};
comboChartSetting.AddComboChartsSetting(new G.AreaChartSetting<G.PresentationSetting>());
comboChartSetting.AddComboChartsSetting(new G.BarChartSetting<G.PresentationSetting>());
comboChartSetting.AddComboChartsSetting(new G.ColumnChartSetting<G.PresentationSetting>());
comboChartSetting.AddComboChartsSetting(new G.LineChartSetting<G.PresentationSetting>()
{
	isSecondaryAxis = true
});
comboChartSetting.AddComboChartsSetting(new G.PieChartSetting<G.PresentationSetting>());

powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(10), comboChartSetting);
```

Above code is example to add each data series as different type of chart grapics combined as combo chart
{% endtab %}
{% endtabs %}

### `ComboChartSetting<G.PresentationSetting>` Options

| Property              | Type             | Details                                                                                                                                                                                                          |
| --------------------- | ---------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| secondaryAxisPosition | AxisPosition     |                                                                                                                                                                                                                  |
| chartAxesOptions      | ChartAxesOptions |                                                                                                                                                                                                                  |
| AddComboChartsSetting | function         | Method accepts core chart type's setting as input. This provides fexlibilty to handle all options of each chart style. [Supported List](combo.md#list-of-supported-chart-that-can-be-inserted-into-combo-chart). |

<details>

<summary>List of supported chart that can be inserted into combo chart</summary>

* [Area Chart](area.md)
* [Bar Chart](bar.md)
* [Column Chart](column.md)
* [Line Chart](line.md)
* [Pie Chart](pie.md)
* Scatter Chart (TODO)

</details>
