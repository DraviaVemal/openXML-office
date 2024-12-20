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

Add chart method present in worksheet component.\
This type is bit different from previous core chart types. Combo setting act as organiser for adding other base type chart to single chart componenet along with additional options

<figure><img src="../../.gitbook/assets/Screenshot 2024-04-04 103511.png" alt=""><figcaption></figcaption></figure>

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
Worksheet worksheet = excel.AddSheet("Combo Chart");
CreateDataCellPayload().ToList().ForEach(rowData =>
{
	worksheet.SetRow(ConverterUtils.ConvertToExcelCellReference(++row, 1), rowData, new());
});
ComboChartSetting<ExcelSetting> comboChartSetting = new()
{
	secondaryAxisPosition = AxisPosition.TOP,
	applicationSpecificSetting = new()
	{
		from = new()
		{
			row = 21,
			column = 5
		},
		to = new()
		{
			row = 41,
			column = 20
		}
	}
};
comboChartSetting.AddComboChartsSetting(new LineChartSetting<ExcelSetting>()
{
	applicationSpecificSetting = new()
});
comboChartSetting.AddComboChartsSetting(new BarChartSetting<ExcelSetting>()
{
	isSecondaryAxis = true,
	applicationSpecificSetting = new()
});
comboChartSetting.AddComboChartsSetting(new ColumnChartSetting<ExcelSetting>()
{
	applicationSpecificSetting = new()
});
worksheet.AddChart(new()
{
	cellIdStart = "A1",
	cellIdEnd = "D4"
}, comboChartSetting);
```

Above code is example to add each data series as different type of chart grapics combined as combo chart
{% endtab %}
{% endtabs %}

### `ComboChartSetting` Options

| Property              | Type             | Details                                                                                                                                                                                                          |
| --------------------- | ---------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| secondaryAxisPosition | AxisPosition     |                                                                                                                                                                                                                  |
| chartAxesOptions      | ChartAxesOptions |                                                                                                                                                                                                                  |
| AddComboChartsSetting | function         | Method accepts core chart type's setting as input. This provides fexlibilty to handle all options of each chart style. [Supported List](combo.md#list-of-supported-chart-that-can-be-inserted-into-combo-chart). |

<details>

<summary>List of supported chart that can be inserted into combo chart</summary>

* [Area Chart](../../presentation/chart/area.md)
* [Bar Chart](../../presentation/chart/bar.md)
* [Column Chart](../../presentation/chart/column.md)
* [Line Chart](../../presentation/chart/line.md)
* [Pie Chart](../../presentation/chart/pie.md)
* Scatter Chart (TODO)

</details>
